function stripHtml(html) {
  return (html || '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .trim();
}

function extractField(text, labels) {
  for (const label of labels) {
    const re = new RegExp(label + '\\s*[:：]\\s*(.+)', 'i');
    const m = text.match(re);
    if (m) return m[1].trim();
  }
  return '';
}

function parseLaunchFields(text) {
  return {
    title: extractField(text, ['작품명']),
    author: extractField(text, ['작가명']),
    label: extractField(text, ['레이블 태그', '레이블']),
    kind: extractField(text, ['구분\\(연재/단행\\)', '구분']),
    platform: extractField(text, ['출간 플랫폼', '플랫폼']),
    pubDateRaw: extractField(text, ['출간 일정', '출간일']),
  };
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
