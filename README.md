론칭 타래 대시보드
Teams 론칭 타래를 상태 기반 일정 관리 표로 바꾸는 대시보드입니다. `index.html` + `api/teams-sync.js` 두 부분으로 구성되어 있고, Vercel에 배포하면 바로 동작합니다.
배포 (Vercel)
이 폴더 전체(`index.html`, `api/teams-sync.js`, `README.md`)를 GitHub 저장소에 올립니다.
vercel.com → Add New → Project → 해당 저장소 Import → Deploy
배포된 프로젝트 → Settings → Environment Variables 에 아래 3개 추가 후 Redeploy:
Name	값
`TEAMS_TENANT_ID`	Microsoft Entra tenant ID
`TEAMS_CLIENT_ID`	앱 등록 client ID
`TEAMS_CLIENT_SECRET`	앱 등록 client secret
Client Secret은 절대 index.html이나 브라우저 입력창에 넣지 않습니다. 서버 환경변수에만 저장되고, `api/teams-sync.js` 안에서만 사용됩니다.
사용 방법
페이지 상단 "Teams 읽기 전용 연동" 카드에 Team ID, Channel ID를 입력
"Teams 현황 불러오기" 클릭
`api/teams-sync.js`가 서버에서 Microsoft Graph를 호출해 채널의 최근 메시지를 가져오고, 각 메시지 본문에서 다음 라벨을 정규식으로 추출합니다:
`작품명`, `작가명`, `레이블 태그`(또는 `레이블`), `구분(연재/단행)`, `출간 플랫폼`, `출간 일정`
추출된 값으로 작품 카드가 자동으로 추가/갱신됩니다. 이미 있는 작품(제목이 같은 항목)은 새로 만들지 않고 갱신만 됩니다.
메시지에 👍(thumbsup/like) 반응이 있으면 완료 처리로 반영됩니다.
Team ID / Channel ID 구하는 법
Teams에서 대상 팀 → "..." → 팀 링크 가져오기 → URL의 `groupId=` 뒤 값이 Team ID
대상 채널 → "..." → 채널 링크 가져오기 → URL 안의 `19%3a`로 시작하는 인코딩 값이 Channel ID
Entra ID 앱에 필요한 권한
애플리케이션 권한(위임된 권한 아님) + 관리자 동의 필요:
`ChannelMessage.Read.All`
`Team.ReadBasic.All`
`Channel.ReadBasic.All`
알아두어야 할 것
본문 파싱 규칙이 우리 팀 타래 양식에 맞게 되어 있는지 먼저 테스트하세요. `api/teams-sync.js`의 `parseLaunchFields` 함수가 찾는 라벨 문구가 실제 타래에서 쓰는 표현과 다르면 못 읽어옵니다. 필요하면 라벨 목록(`extractField`에 넘기는 배열)에 실제 쓰는 표현을 추가하세요.
👍 반응 감지는 Graph API 응답 구조에 따라 조정이 필요할 수 있습니다. `reactions` 필드가 기대한 형태로 오지 않으면, 서버 함수에서 한 번 `console.log(msgData)`로 실제 응답 구조를 확인한 뒤 `hasThumbsUp` 판정 로직을 맞춰야 합니다.
이 페이지는 로그인 보호가 없습니다. URL을 아는 사람은 누구나 동기화 버튼을 누를 수 있어요. 사내 전용으로만 쓰실 거라면 Vercel의 Password Protection 기능(Pro 플랜)이나 간단한 접근 코드 체크를 추가하는 걸 권장합니다.
작품별 체크박스 상태(서지/원고/표지/제작/등록/승인/완료)는 Teams에 없는 정보라 여전히 대시보드에서 직접 체크해야 합니다. Teams 동기화는 작품 기본 정보(제목/작가/레이블/플랫폼)를 채워주는 역할만 합니다.
데이터는 브라우저 `localStorage`에 저장됩니다. 팀 전체가 같은 데이터를 보게 하려면 Supabase 같은 DB를 붙여 `items` 배열을 서버에 저장하도록 바꿔야 합니다 (다음 단계 과제).
다음 단계: 자동 알림
마감 3일 전 미완료 건에 대해 자동 알림을 보내려면, `api/teams-sync.js`처럼 함수를 하나 더 만들어 Vercel Cron으로 주기 실행하고, Teams Incoming Webhook으로 메시지를 보내면 됩니다.
