# GPT 풀이와 음성 글 변환 연결

현재 GitHub Pages는 정적 화면만 제공하므로 아래 Node 서버를 HTTPS 호스팅에 별도로 실행해야 합니다. 서버와 OpenAI 프로젝트 API 키가 연결되기 전에는 앱에 **서버 연결 필요**가 표시됩니다. 이 저장소에는 실제 키와 배포 주소가 없습니다.

## 실행

Node.js 22 이상. 추가 패키지 없음. 호스팅의 비밀 환경변수 설정에 다음을 입력하세요. 키를 index.html, GitHub 파일, 브라우저 저장소 또는 채팅에 붙여 넣지 마세요.

- `OPENAI_API_KEY`: OpenAI 프로젝트의 API 키
- `AI_ACCESS_TOKEN`: 32자 이상의 무작위 서버 접속 코드. API 키와 별개입니다.
- `ALLOWED_ORIGIN`: `https://msjung75.github.io`
- `OPENAI_MODEL`: 기본 `gpt-4.1-mini`
- `OPENAI_TRANSCRIBE_MODEL`: 기본 `gpt-4o-mini-transcribe`
- `AI_HOURLY_LIMIT`: 기본 30회, 서버 프로세스당 풀이와 전사 합계
- `PORT`: 기본 3000 (호스팅 제공 포트 사용 가능)

시작 명령: `node server/ai-server.mjs`

HTTPS 리버스 프록시에서 24MB 이상의 요청과 120초 타임아웃을 지원해야 합니다. 이 파일은 독립 Node 서버입니다. Vercel 서버리스 함수에 그대로 넣는 형태가 아닙니다. 여러 인스턴스 운영 시 공유 저장소 기반 사용량 제한과 사용자별 인증을 추가해야 합니다. 현재는 개인/소규모 상담용 공유 접속 코드 방식입니다. OpenAI 프로젝트 사용 한도도 함께 설정하세요. 이 시간당 제한은 프로세스 재시작 시 초기화되므로 결제 한도를 대신하지 않습니다.

앱 설정 → GPT 서버 주소에 HTTPS 주소 입력 → 별도의 서버 접속 코드 입력 → 연결 확인. 접속 코드는 메모리에만 보관하므로 새로 열면 다시 입력합니다. `/health`는 접속 코드와 서버 환경설정을 확인하며, OpenAI의 결제/모델 접근 권한은 첫 실제 요청에서 확인됩니다.

## 동작과 데이터

- 풀이: 선택 연도의 앱 계산 명식, 생년월일·성별·출생시간만 전송합니다. 이름과 메모는 보내지 않습니다. 실패 시 이전 풀이를 유지합니다. Responses API의 `store:false`를 사용합니다. 이것이 OpenAI의 모든 로그 보관을 비활성화한다는 뜻은 아닙니다.
- 녹음: 기기 IndexedDB에 원본 저장 → 파일별 **글로 변환** → 전송 확인 → 글 검토·수정 → **메모에 추가**. 실시간 Web Speech의 자동 재시작을 제거했습니다. OS 자체의 녹음 알림음은 앱이 끌 수 없습니다.
- 기존 메모·연도별 기록·녹음은 유지됩니다. JSON 백업은 녹음 원본을 포함하지 않습니다. 파일별 **파일 저장**으로 원본을 보관하세요.
- 서버는 명식/음성을 파일이나 로그로 남기지 않습니다. OpenAI 서비스와 호스팅 제공자의 데이터 처리 정책은 별도 적용됩니다.
- API 오류·인증 실패·빈 응답은 저장 결과를 덮어쓰지 않습니다. 24MB 넘는 녹음은 원본을 받아 나눈 뒤 처리해야 합니다.

## 검증

`npm test` 및 `node --test tests/ai-server.test.mjs`. 자동 검증은 모의 OpenAI 응답으로 실행하며 비용이나 개인정보 전송을 발생시키지 않습니다. 실제 마이크·블루투스·통화 중단·백그라운드 녹음은 대상 휴대폰에서 추가 확인해야 합니다.

공식 문서: https://developers.openai.com/api/docs/guides/migrate-to-responses · https://developers.openai.com/api/docs/guides/speech-to-text
