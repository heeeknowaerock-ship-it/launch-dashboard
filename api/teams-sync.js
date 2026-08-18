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

  return { title, author, label, kind: '', platform };
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
        pubDateRaw: fields.pubDateRaw || '',
        hasThumbsUp,
      };
    }).filter((t) => t.title || t.rawText);

    res.status(200).json({
      threads,
      fetchedAt: new Date().toISOString(),
      totalMessagesFound: (msgData.value || []).length,
      sampleRawTexts: (msgData.value || []).slice(0, 3).map((m) => stripHtml(m.body && m.body.content).slice(0, 300)),
    });
  } catch (err) {
    res.status(500).json({ error: '동기화 중 오류가 발생했습니다.', detail: String(err) });
  }
};
