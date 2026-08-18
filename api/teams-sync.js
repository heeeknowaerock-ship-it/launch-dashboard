function stripHtml(html) {
  return (html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .trim();
}

const KNOWN_PLATFORMS = [
  '네이버 시리즈', '카카오페이지', '카카오웹툰', '원스토어', '미스터블루',
  '봄툰', '리디북스', '리디', '조아라', '문피아', '시리즈', '톡소다', '북팔',
];

const KNOWN_LABELS = ['에이블', '원티드', '비올렛', '라노체', '이브'];

function normalizeDate(raw) {
  if (!raw) return null;
  let y, m, d;
  let mo;
  if ((mo = raw.match(/(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/))) {
    y = +mo[1]; m = +mo[2]; d = +mo[3];
  } else if ((mo = raw.match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일/))) {
    m = +mo[1]; d = +mo[2];
  } else if ((mo = raw.match(/(\d{1,2})[.\/](\d{1,2})(?!\d)/))) {
    m = +mo[1]; d = +mo[2];
  } else {
    return null;
  }
  if (!y) {
    const now = new Date();
    y = now.getFullYear();
    const candidate = new Date(y, m - 1, d);
    if (candidate.getTime() < now.getTime() - 30 * 86400000) y += 1;
  }
  if (m < 1 || m > 12 || d < 1 || d > 31) return null;
  const dt = new Date(y, m - 1, d);
  return dt.toISOString().slice(0, 10);
}

const DATE_PATTERN = /(\d{4}[-.\/]\d{1,2}[-.\/]\d{1,2})|(\d{1,2}\s*월\s*\d{1,2}\s*일)|(\d{1,2}[.\/]\d{1,2}(?!\d))/g;

function extractDates(text) {
  let match;
  let dueRaw = null, pubRaw = null;
  while ((match = DATE_PATTERN.exec(text)) !== null) {
    const dateStr = match[0];
    const contextStart = Math.max(0, match.index - 6);
    const context = text.slice(contextStart, match.index);
    if (/마감/.test(context)) {
      if (!dueRaw) dueRaw = dateStr;
    } else if (!pubRaw) {
      pubRaw = dateStr;
    }
  }
  return {
    pubDate: normalizeDate(pubRaw),
    dueDate: normalizeDate(dueRaw),
  };
}

function parseLaunchFields(text) {
  const titleMatch = text.match(/<([^<>]+)>/);
  const title = titleMatch ? titleMatch[1].trim() : '';

  const before = titleMatch ? text.slice(0, titleMatch.index) : '';
  const after = titleMatch ? text.slice(titleMatch.index + titleMatch[0].length) : '';

  const beforeTokens = before.trim().split(/\s+/).filter(Boolean);

  let label = '';
  let authorBefore = '';
  if (beforeTokens.length && KNOWN_LABELS.includes(beforeTokens[0])) {
    label = beforeTokens[0];
    authorBefore = beforeTokens.slice(1).join(' ');
  } else {
    authorBefore = beforeTokens.join(' ');
  }

  let platform = '';
  let platformIndex = -1;
  for (const p of KNOWN_PLATFORMS) {
    const idx = after.indexOf(p);
    if (idx !== -1) { platform = p; platformIndex = idx; break; }
  }
  const authorAfter = platformIndex !== -1 ? after.slice(0, platformIndex).trim() : '';

  const author = authorBefore || authorAfter;
  const dates = extractDates(text);

  return { title, author, label, kind: '', platform, pubDate: dates.pubDate, dueDate: dates.dueDate };
}

const DONE_MARKERS = ['완료', '완입니다', '완료입니다', '완료됨', '됐습니다', '됐어요', '됨', '끝', 'ok', 'OK', '오케이', '컨펌'];
const NEGATION_PATTERN = /아직|안\s*(됐|됨|끝|완료)|못\s*(했|함|끝)|미완료|아니요|아뇨|보류/;

const CHECK_RULES = [
  { field: 'biblio', keywords: ['서지정보', '서지'] },
  { field: 'manuscript', keywords: ['원고', '완고'] },
  { field: 'cover', keywords: ['표지'] },
];

function scanReplyForChecks(text) {
  const result = {};
  if (NEGATION_PATTERN.test(text)) return result;
  for (const rule of CHECK_RULES) {
    const hasKeyword = rule.keywords.some((kw) => text.includes(kw));
    const hasDone = DONE_MARKERS.some((mk) => text.includes(mk));
    if (hasKeyword && hasDone) result[rule.field] = true;
  }
  if (/등록\s*완료/.test(text) || (text.includes('등록') && DONE_MARKERS.some((mk) => text.includes(mk)) && !text.includes('마감'))) {
    result.registered = true;
  }
  if (/승인\s*(대기|검토|확인)/.test(text)) {
    result.approving = true;
  }
  return result;
}

async function mapLimit(items, limit, fn) {
  const results = new Array(items.length);
  let idx = 0;
  async function worker() {
    while (idx < items.length) {
      const current = idx++;
      results[current] = await fn(items[current], current);
    }
  }
  const workers = Array.from({ length: Math.min(limit, items.length) }, worker);
  await Promise.all(workers);
  return results;
}

async function fetchRepliesChecks(teamId, channelId, messageId, accessToken) {
  try {
    const repRes = await fetch(
      `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies?$top=50`,
      { headers: { Authorization: `Bearer ${accessToken}` } }
    );
    const repData = await repRes.json();
    if (!repRes.ok) return { checks: {}, replyCount: 0, error: (repData && repData.error && repData.error.message) || 'unknown' };

    const merged = {};
    const replies = repData.value || [];
    for (const r of replies) {
      const text = stripHtml(r.body && r.body.content);
      if (!text) continue;
      Object.assign(merged, scanReplyForChecks(text));
    }
    return { checks: merged, replyCount: replies.length };
  } catch (err) {
    return { checks: {}, replyCount: 0, error: String(err) };
  }
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'POST만 지원합니다.' });
    return;
  }

  const tenantId = (process.env.TEAMS_TENANT_ID || '').trim();
  const clientId = (process.env.TEAMS_CLIENT_ID || '').trim();
  const clientSecret = (process.env.TEAMS_CLIENT_SECRET || '').trim();

  if (!tenantId || !clientId || !clientSecret) {
    res.status(500).json({ error: '서버에 TEAMS_TENANT_ID / TEAMS_CLIENT_ID / TEAMS_CLIENT_SECRET 환경변수가 설정되어 있지 않습니다.' });
    return;
  }

  if (req.body && req.body.debug === true) {
    res.status(200).json({
      debug: true,
      tenantIdLength: tenantId.length,
      clientIdLength: clientId.length,
      clientSecretLength: clientSecret.length,
      clientSecretFirst4: clientSecret.slice(0, 4),
      clientSecretLast4: clientSecret.slice(-4),
      clientSecretHasWhitespace: /\s/.test(process.env.TEAMS_CLIENT_SECRET || ''),
    });
    return;
  }

  const { teamId: rawTeamId, channelId: rawChannelId, top } = req.body || {};
  if (!rawTeamId || !rawChannelId) {
    res.status(400).json({ error: 'teamId와 channelId를 입력하세요.' });
    return;
  }

  function normalizeId(v) {
    let s = String(v).trim();
    try { s = decodeURIComponent(s); } catch (e) {}
    return encodeURIComponent(s);
  }
  const teamId = normalizeId(rawTeamId);
  const channelId = normalizeId(rawChannelId);

  try {
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          scope: 'https://graph.microsoft.com/.default',
          grant_type: 'client_credentials',
        }),
      }
    );
    const tokenData = await tokenRes.json();
    if (!tokenData.access_token) {
      res.status(502).json({ error: '토큰 발급 실패: ' + (tokenData.error_description || tokenData.error || JSON.stringify(tokenData)) });
      return;
    }

    const msgRes = await fetch(
      `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages?$top=${top || 50}`,
      { headers: { Authorization: `Bearer ${tokenData.access_token}` } }
    );
    const msgData = await msgRes.json();
    if (!msgRes.ok) {
      const msg = (msgData && msgData.error && msgData.error.message) || JSON.stringify(msgData);
      res.status(502).json({ error: 'Teams 메시지 조회 실패: ' + msg });
      return;
    }

    const threads = (msgData.value || []).map((m) => {
      const text = stripHtml(m.body && m.body.content);
      const fields = parseLaunchFields(text);
      const reactions = Array.isArray(m.reactions) ? m.reactions : [];
      const hasThumbsUp = reactions.some((r) => r.reactionType === 'like' || r.reactionType === 'thumbsup' || r.reactionType === 'thumbsUp');
      return {
        messageId: m.id,
        createdDateTime: m.createdDateTime,
        webUrl: m.webUrl,
        author: (m.from && m.from.user && m.from.user.displayName) || fields.author || '',
        rawText: text,
        title: fields.title || '',
        label: fields.label || '',
        kind: fields.kind || '',
        platform: fields.platform || '',
        pubDateRaw: fields.pubDate || '',
        dueDateRaw: fields.dueDate || '',
        hasThumbsUp,
      };
    }).filter((t) => t.title || t.rawText);

    // 답글에서 준비 체크(서지/원고/표지) 및 마감 체크(등록완료/승인대기) 자동 인식 — 실제 작품 타래(title 있는 것)만 조회
    const launchThreads = threads.filter((t) => t.title);
    await mapLimit(launchThreads, 5, async (t) => {
      const { checks, replyCount } = await fetchRepliesChecks(teamId, channelId, t.messageId, tokenData.access_token);
      t.checks = checks;
      t.replyCount = replyCount;
    });

    res.status(200).json({
      threads,
      fetchedAt: new Date().toISOString(),
      totalMessagesFound: (msgData.value || []).length,
      threadsWithRepliesScanned: launchThreads.length,
      sampleRawTexts: (msgData.value || []).slice(0, 3).map((m) => stripHtml(m.body && m.body.content).slice(0, 300)),
    });
  } catch (err) {
    res.status(500).json({ error: '동기화 중 오류가 발생했습니다.', detail: String(err) });
  }
};
