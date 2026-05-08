require('dotenv').config();

const express = require('express');
const path = require('path');
const fs = require('fs');
const axios = require('axios');
const { HttpsProxyAgent } = require('https-proxy-agent');
const XLSX = require('xlsx');

const PORT = process.env.PORT || 3000;

function loadFileConfig() {
  const p = path.join(__dirname, 'config.json');
  try {
    if (fs.existsSync(p)) {
      return JSON.parse(fs.readFileSync(p, 'utf8'));
    }
  } catch (e) {
    console.warn('[config] 读取 config.json 失败:', e.message);
  }
  return {};
}

const fileConfig = loadFileConfig();

/** 优先级：环境变量 > config.json。YouTube 需在 Google Cloud 启用 YouTube Data API v3 */
const YOUTUBE_API_KEY =
  process.env.YOUTUBE_API_KEY ||
  process.env.GOOGLE_API_KEY ||
  String(fileConfig.youtubeApiKey || fileConfig.googleApiKey || '').trim();

/** Gemini：简介无正则邮箱时辅助提取（可选） */
const GEMINI_API_KEY =
  process.env.GEMINI_API_KEY || String(fileConfig.geminiApiKey || '').trim();

/** 访问 Google API 超时（毫秒）；网络差或走代理时可适当加大 */
const GOOGLE_API_TIMEOUT_MS = Number(process.env.GOOGLE_API_TIMEOUT_MS || 60000);

/**
 * search.list 每次分页约消耗 100 点每日配额；原先默认 40 页极易打满免费额度。
 * 可通过环境变量调大，例如 YOUTUBE_SEARCH_MAX_PAGES=15
 */
const YOUTUBE_SEARCH_MAX_PAGES = (() => {
  const n = Number(process.env.YOUTUBE_SEARCH_MAX_PAGES ?? 6);
  if (!Number.isFinite(n)) return 6;
  return Math.min(50, Math.max(1, Math.floor(n)));
})();

/** 同一 IP 两次搜索完成之间的最短间隔（毫秒），0 表示不限制；可减轻连点把日配额打光 */
const YOUTUBE_SEARCH_MIN_INTERVAL_MS = Math.max(
  0,
  Number(process.env.YOUTUBE_SEARCH_MIN_INTERVAL_MS || 0)
);

/** 每次 search.list 翻页后延迟再请求下一页（毫秒），0 表示不延迟；略拉长总耗时、降低突发 QPS */
const YOUTUBE_SEARCH_PAGE_DELAY_MS = Math.max(
  0,
  Number(process.env.YOUTUBE_SEARCH_PAGE_DELAY_MS || 0)
);

/** YouTube 项目默认每日配额（点），用于估算；以 Google Cloud 控制台为准 */
const YOUTUBE_QUOTA_UNITS_PER_DAY = Number(process.env.YOUTUBE_QUOTA_UNITS_PER_DAY || 10000);

const searchingByIp = new Set();
const lastSearchCompleteByIp = new Map();

function clientIp(req) {
  const xf = req.headers['x-forwarded-for'];
  if (typeof xf === 'string' && xf.trim()) return xf.split(',')[0].trim();
  return req.socket?.remoteAddress || 'unknown';
}

const httpsProxyUrl = String(
  process.env.HTTPS_PROXY ||
    process.env.HTTP_PROXY ||
    fileConfig.httpsProxy ||
    ''
)
  .trim()
  .replace(/\/$/, '');

const httpClient = httpsProxyUrl
  ? axios.create({
      httpsAgent: new HttpsProxyAgent(httpsProxyUrl),
      proxy: false,
      timeout: GOOGLE_API_TIMEOUT_MS,
    })
  : axios.create({ timeout: GOOGLE_API_TIMEOUT_MS });

if (httpsProxyUrl) {
  console.log('[proxy] 已启用 HTTPS 代理（YouTube / Gemini）');
}

const app = express();
app.use(express.json({ limit: '2mb' }));

function stripHtml(s) {
  return String(s || '')
    .replace(/<[^>]*>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function formatYoutubeApiError(err) {
  const status = err.response?.status;
  const data = err.response?.data;
  const gErr = data?.error;
  const first = gErr?.errors?.[0];
  const reason = first?.reason || '';

  if (reason === 'quotaExceeded' || reason === 'dailyLimitExceeded') {
    return (
      'YouTube Data API daily quota exceeded. search.list costs about 100 units per page; ' +
      'wait until quota resets (Pacific midnight) or raise the limit in Google Cloud Console → ' +
      'APIs & Services → YouTube Data API v3 → Quotas. You can also lower usage by setting a smaller ' +
      'YOUTUBE_SEARCH_MAX_PAGES in your server env (default is ' +
      YOUTUBE_SEARCH_MAX_PAGES +
      ').'
    );
  }

  let msg = '';
  if (typeof gErr === 'string') msg = gErr;
  else if (gErr?.message) msg = gErr.message;
  else if (first?.message) msg = `${first.message}${first.reason ? ` (${first.reason})` : ''}`;
  else if (typeof data === 'string' && data.trim()) msg = data.slice(0, 500);
  else if (err.message && err.message !== 'Error') msg = err.message;
  else if (err.code) msg = `Network error: ${err.code}`;
  else msg = 'YouTube request failed';

  msg = stripHtml(msg);
  if (/quota/i.test(msg) && status === 403) {
    return (
      'YouTube Data API quota exceeded (403). ' +
      'Reduce searches or increase quota in Google Cloud Console; see YOUTUBE_SEARCH_MAX_PAGES (default ' +
      YOUTUBE_SEARCH_MAX_PAGES +
      ').'
    );
  }
  return msg;
}

const EMAIL_RE = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;

function extractEmailsFromText(text) {
  if (!text || typeof text !== 'string') return '';
  const matches = text.match(EMAIL_RE);
  if (!matches || !matches.length) return '';
  const filtered = [...new Set(matches.map((e) => e.toLowerCase()))].filter(
    (e) => !e.endsWith('@youtube.com') && !e.endsWith('@google.com')
  );
  return filtered[0] || '';
}

async function geminiExtractEmail(description, channelTitle) {
  if (!GEMINI_API_KEY || !description) return '';
  const prompt = `你是信息提取助手。下面是一段 YouTube 频道简介，请只输出该频道公开的商务/联系用邮箱（单个），若没有明确邮箱则输出单词 NONE。不要解释。\n频道名：${channelTitle}\n简介：\n${description.slice(0, 8000)}`;
  try {
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${encodeURIComponent(GEMINI_API_KEY)}`;
    const { data } = await httpClient.post(url, {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.1, maxOutputTokens: 64 },
    });
    const text =
      data?.candidates?.[0]?.content?.parts?.map((p) => p.text).join('')?.trim() || '';
    if (!text || /^none$/i.test(text)) return '';
    const m = text.match(EMAIL_RE);
    return m ? m[0] : '';
  } catch {
    return '';
  }
}

/**
 * 按关键词搜索「视频」（YouTube 会匹配标题、说明等；非频道名搜索）。
 * 保留分页顺序，每条结果一条视频（去重 videoId）。
 */
async function youtubeSearchVideoList(keyword, apiKey) {
  const list = [];
  const seenVideoIds = new Set();
  let pageToken = '';
  const maxPages = YOUTUBE_SEARCH_MAX_PAGES;
  let searchPageCount = 0;

  for (let page = 0; page < maxPages; page++) {
    const params = {
      part: 'snippet',
      type: 'video',
      q: keyword,
      order: 'relevance',
      maxResults: 50,
      key: apiKey,
    };
    if (pageToken) params.pageToken = pageToken;

    const { data } = await httpClient.get('https://www.googleapis.com/youtube/v3/search', {
      params,
    });
    searchPageCount += 1;

    for (const item of data.items || []) {
      const videoId = item?.id?.videoId;
      const cid = item?.snippet?.channelId;
      if (!videoId || !cid || seenVideoIds.has(videoId)) continue;
      seenVideoIds.add(videoId);
      list.push({
        videoId,
        videoTitle: item?.snippet?.title || '',
        channelId: cid,
        channelTitle: item?.snippet?.channelTitle || '',
      });
    }

    pageToken = data.nextPageToken;
    if (!pageToken) break;
    if (YOUTUBE_SEARCH_PAGE_DELAY_MS > 0 && page + 1 < maxPages) {
      await new Promise((r) => setTimeout(r, YOUTUBE_SEARCH_PAGE_DELAY_MS));
    }
  }

  return { rows: list, searchPageCount };
}

async function youtubeChannelsBatch(ids, apiKey) {
  const results = [];
  const chunkSize = 50;
  for (let i = 0; i < ids.length; i += chunkSize) {
    const chunk = ids.slice(i, i + chunkSize);
    const { data } = await httpClient.get('https://www.googleapis.com/youtube/v3/channels', {
      params: {
        part: 'snippet,statistics',
        id: chunk.join(','),
        key: apiKey,
      },
    });
    for (const ch of data.items || []) {
      results.push(ch);
    }
  }
  return results;
}

function channelPageUrl(ch) {
  const id = ch.id;
  const custom = ch?.snippet?.customUrl;
  if (custom) {
    const c = String(custom).replace(/^@/, '');
    return `https://www.youtube.com/@${c}`;
  }
  return `https://www.youtube.com/channel/${id}`;
}

app.get('/api/search', async (req, res) => {
  const q = String(req.query.q || '').trim();
  if (!q) return res.status(400).json({ error: '请输入关键词' });
  if (!YOUTUBE_API_KEY) {
    return res.status(500).json({
      error:
        '未配置 YouTube API 密钥。请设置环境变量 YOUTUBE_API_KEY（在 Google Cloud 启用 YouTube Data API v3）。',
    });
  }

  const ip = clientIp(req);
  if (searchingByIp.has(ip)) {
    return res.status(429).json({
      error:
        'A search is already running for your connection. Please wait until it finishes before starting another.',
      code: 'search_in_progress',
    });
  }
  if (YOUTUBE_SEARCH_MIN_INTERVAL_MS > 0) {
    const last = lastSearchCompleteByIp.get(ip) || 0;
    const elapsed = Date.now() - last;
    if (last > 0 && elapsed < YOUTUBE_SEARCH_MIN_INTERVAL_MS) {
      const retryAfterSeconds = Math.ceil((YOUTUBE_SEARCH_MIN_INTERVAL_MS - elapsed) / 1000);
      res.setHeader('Retry-After', String(retryAfterSeconds));
      return res.status(429).json({
        error: `Please wait ${retryAfterSeconds}s before searching again (spacing requests saves daily YouTube API quota).`,
        code: 'rate_limited',
        retryAfterSeconds,
      });
    }
  }

  searchingByIp.add(ip);
  try {
    const { rows: videoRows, searchPageCount } = await youtubeSearchVideoList(q, YOUTUBE_API_KEY);
    if (!videoRows.length) {
      return res.json({
        keyword: q,
        videos: [],
        message: '未找到相关视频',
        usage: {
          searchPages: searchPageCount,
          channelCalls: 0,
          estimatedUnits: searchPageCount * 100,
          quotaUnitsPerDay: YOUTUBE_QUOTA_UNITS_PER_DAY,
        },
      });
    }

    const uniqueChannelIds = [...new Set(videoRows.map((v) => v.channelId))];
    const channels = await youtubeChannelsBatch(uniqueChannelIds, YOUTUBE_API_KEY);
    const byChannelId = new Map(channels.map((ch) => [ch.id, ch]));

    const channelMeta = new Map();
    for (const ch of channels) {
      const title = ch?.snippet?.title || '';
      const description = ch?.snippet?.description || '';
      let subs = ch?.statistics?.subscriberCount;
      if (subs === undefined || subs === null) subs = '';
      channelMeta.set(ch.id, {
        channelName: title,
        channelUrl: channelPageUrl(ch),
        email: extractEmailsFromText(description) || '',
        subscribers: subs === '' ? '' : Number(subs),
        _description: description,
        _geminiTitle: title,
      });
    }

    const needGemini = [...channelMeta.entries()]
      .filter(([, m]) => !m.email && m._description)
      .slice(0, 30);
    const gchunk = 5;
    for (let i = 0; i < needGemini.length; i += gchunk) {
      const slice = needGemini.slice(i, i + gchunk);
      await Promise.all(
        slice.map(async ([, m]) => {
          const g = await geminiExtractEmail(m._description, m._geminiTitle);
          if (g) m.email = g;
        })
      );
    }

    for (const m of channelMeta.values()) {
      delete m._description;
      delete m._geminiTitle;
    }

    const videos = [];
    for (const v of videoRows) {
      const ch = byChannelId.get(v.channelId);
      const meta = channelMeta.get(v.channelId);
      if (!ch || !meta) continue;
      videos.push({
        videoId: v.videoId,
        videoTitle: v.videoTitle,
        videoUrl: `https://www.youtube.com/watch?v=${v.videoId}`,
        channelId: v.channelId,
        channelName: v.channelTitle || meta.channelName,
        channelUrl: meta.channelUrl,
        email: meta.email || '',
        subscribers: meta.subscribers,
      });
    }

    const channelCalls = Math.ceil(uniqueChannelIds.length / 50);
    const estimatedUnits = searchPageCount * 100 + channelCalls * 1;
    const approxSearchesPerDay = Math.max(1, Math.floor(YOUTUBE_QUOTA_UNITS_PER_DAY / estimatedUnits));

    res.json({
      keyword: q,
      videos,
      usage: {
        searchPages: searchPageCount,
        channelCalls,
        estimatedUnits,
        quotaUnitsPerDay: YOUTUBE_QUOTA_UNITS_PER_DAY,
        approxFullSearchesPerDay: approxSearchesPerDay,
      },
    });
  } catch (err) {
    const msg = formatYoutubeApiError(err);
    const status = err.response?.status;
    console.error('[api/search]', status || '', msg, err.response?.data || err.message);
    res.status(502).json({
      error: status ? `[HTTP ${status}] ${msg}` : msg,
    });
  } finally {
    searchingByIp.delete(ip);
    lastSearchCompleteByIp.set(ip, Date.now());
  }
});

app.post('/api/export', (req, res) => {
  const { videos, creators } = req.body || {};
  const list = Array.isArray(videos) && videos.length ? videos : creators;
  if (!Array.isArray(list) || !list.length) {
    return res.status(400).json({ error: '没有可导出的数据' });
  }

  const isVideoShape = list[0] && Object.prototype.hasOwnProperty.call(list[0], 'videoUrl');
  const rows = isVideoShape
    ? list.map((r) => ({
        视频标题: r.videoTitle || '',
        视频链接: r.videoUrl || '',
        频道名称: r.channelName || '',
        频道主页: r.channelUrl || '',
        公开邮箱: r.email || '',
        粉丝数: r.subscribers === '' || r.subscribers == null ? '' : r.subscribers,
      }))
    : list.map((c) => ({
        频道名称: c.channelName || '',
        频道主页链接: c.channelUrl || '',
        公开邮箱: c.email || '',
        粉丝数: c.subscribers === '' || c.subscribers == null ? '' : c.subscribers,
      }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = isVideoShape
    ? [{ wch: 36 }, { wch: 44 }, { wch: 22 }, { wch: 40 }, { wch: 26 }, { wch: 12 }]
    : [{ wch: 28 }, { wch: 42 }, { wch: 28 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, ws, isVideoShape ? '视频与博主' : '达人');

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const name = `youtube-influencers-${Date.now()}.xlsx`;
  res.setHeader(
    'Content-Type',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  );
  res.setHeader(
    'Content-Disposition',
    `attachment; filename="influencers.xlsx"; filename*=UTF-8''${encodeURIComponent(name)}`
  );
  res.send(buf);
});

const publicDir = path.join(__dirname, 'public');
const indexHtmlPath = path.resolve(publicDir, 'index.html');

function sendIndexHtml(_req, res, next) {
  res.sendFile(indexHtmlPath, (err) => {
    if (err) next(err);
  });
}

app.get('/', sendIndexHtml);
app.get('/app', sendIndexHtml);
app.get('/login', (_req, res) => res.redirect(301, '/'));

app.use(express.static(publicDir));

app.listen(PORT, () => {
  console.log(`http://localhost:${PORT}`);
  if (!YOUTUBE_API_KEY) {
    console.warn('[warn] 未设置 YOUTUBE_API_KEY，搜索不可用');
  } else {
    const worstUnitsPerSearch = YOUTUBE_SEARCH_MAX_PAGES * 100 + YOUTUBE_SEARCH_MAX_PAGES;
    const approx = Math.max(1, Math.floor(YOUTUBE_QUOTA_UNITS_PER_DAY / worstUnitsPerSearch));
    console.log(
      `[youtube] max pages=${YOUTUBE_SEARCH_MAX_PAGES} (search.list ≈100 units/page); worst-case ≈${worstUnitsPerSearch} units/search; ` +
        `quota≈${YOUTUBE_QUOTA_UNITS_PER_DAY}/day → about ${approx} heavy searches/day. ` +
        `minIntervalMs=${YOUTUBE_SEARCH_MIN_INTERVAL_MS} pageDelayMs=${YOUTUBE_SEARCH_PAGE_DELAY_MS}`
    );
  }
});
