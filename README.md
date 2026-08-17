# 론칭 타래 대시보드

Teams 론칭 타래를 상태 기반 일정 관리 표로 바꾸는 대시보드입니다. 빌드 과정 없이 `index.html` 하나로 동작합니다.

## GitHub Pages로 배포하기

1. 이 폴더(`index.html`, `README.md`)를 새 GitHub 저장소에 올립니다.
   ```bash
   git init
   git add .
   git commit -m "launch dashboard"
   git branch -M main
   git remote add origin <레포 주소>
   git push -u origin main
   ```
2. 저장소 **Settings → Pages** 로 이동합니다.
3. **Source**를 `Deploy from a branch`로 설정하고, 브랜치는 `main`, 폴더는 `/ (root)`로 지정합니다.
4. 잠시 뒤 `https://<계정>.github.io/<레포이름>/` 주소로 접속하면 대시보드가 뜹니다.

Netlify나 Vercel에 올리는 경우도 동일합니다 — 빌드 명령어 없이 `index.html`을 그대로 정적 파일로 배포하면 됩니다.

## 지금 상태 (이번 해커톤 기준)

- 작품 목록 표, 상태 자동 계산(서지/원고/표지/등록/승인/완료), 마감일 자동 계산(영업일 7일 전), 마감 임박 강조 → **완료**
- 데이터는 브라우저 `localStorage`에만 저장됩니다. 새로고침해도 유지되지만, **다른 사람 브라우저와는 공유되지 않습니다.**
- Teams 자동 알림 → **미구현** (다음 단계)
- Teams 연동 → **읽기 전용 화면만 구현**, 실제 데이터 연동은 안 됨

## 다음 단계 1: 팀 전체가 같이 보게 하려면

지금은 각자 브라우저에만 데이터가 저장됩니다. 여러 사람이 같은 데이터를 보게 하려면 아래 중 하나가 필요합니다.

- 가장 간단: **Supabase** 같은 무료 백엔드에 표를 하나 만들고, `fetch`로 읽고 쓰기
- 이미 쓰는 도구가 있다면: **Notion API**, **Google Sheets API**로 대체
- 자체 서버가 있다면: 작은 REST API 하나 붙이기

## 다음 단계 2: Teams 연동을 실제로 붙이려면

지금 화면의 "Teams 현황 불러오기" 버튼은 **의도적으로 막혀 있습니다.** 이유는 두 가지입니다.

1. **CORS** — 브라우저는 `graph.microsoft.com`에 직접 요청을 보낼 수 없습니다.
2. **보안** — Client Secret을 브라우저 JS에 넣으면 개발자도구만 열어도 누구나 그 값을 볼 수 있습니다. 절대 배포용 코드에 Secret을 하드코딩하지 마세요.

그래서 실제 연동은 **작은 백엔드(서버리스 함수) 하나**를 거쳐야 합니다. 흐름은 이렇게 됩니다.

```
브라우저 (index.html) → 내 서버리스 함수 → Microsoft Graph API
```

Vercel이나 Netlify를 쓴다면 아래처럼 API 라우트 하나만 추가하면 됩니다 (Vercel 예시, `/api/teams-sync.js`):

```js
export default async function handler(req, res) {
  const { tenantId, clientId, clientSecret, teamId, channelId, top } = req.body;

  const tokenRes = await fetch(
    `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        scope: "https://graph.microsoft.com/.default",
        grant_type: "client_credentials",
      }),
    }
  );
  const { access_token } = await tokenRes.json();

  const msgRes = await fetch(
    `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages?$top=${top || 50}`,
    { headers: { Authorization: `Bearer ${access_token}` } }
  );
  const data = await msgRes.json();
  res.status(200).json(data);
}
```

그리고 `index.html`의 `handleSync()` 함수에서 `throw new Error(...)` 부분을 아래처럼 바꾸면 됩니다.

```js
const resp = await fetch('/api/teams-sync', {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({ tenantId, clientId, clientSecret, teamId, channelId, top: fetchCount }),
});
const data = await resp.json();
// data.value 안의 메시지를 파싱해서 items 배열에 반영
```

Client Secret은 이 서버리스 함수 안에서만 쓰고, 환경변수(예: Vercel의 Environment Variables)에 저장하세요 — 브라우저로는 절대 보내지 않습니다.

### Entra ID 앱 등록에 필요한 권한

- API 권한(애플리케이션 권한, 관리자 동의 필요): `ChannelMessage.Read.All`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`

## 다음 단계 3: 자동 알림

마감 3일 전, 미완료 건에 대해 Teams로 알림을 보내려면, 위에서 만든 서버리스 함수를 하나 더 추가해서 **정기적으로(cron) 실행**하고, Teams의 [Incoming Webhook](https://learn.microsoft.com/microsoftteams/platform/webhooks-and-connectors/how-to/add-incoming-webhook) 또는 Graph의 `chatMessage` 전송 API로 알림을 발송하면 됩니다. Vercel Cron, GitHub Actions 스케줄, 또는 별도 서버의 cron 중 편한 방식을 고르시면 됩니다.
