# 지출 알림/스레드 운영 가이드

## 1) 현재 프로젝트 구성 요약

- 실행 환경: Google Apps Script
- 데이터 소스: 스프레드시트 `원장DB`
- 알림 채널: Slack Bot (`chat.postMessage`, `chat.update`)
- 주요 기능:
  - DM/채널 요약 메시지 발송
  - 스레드 카드 생성 및 상태 버튼(`완료`, `취소`, `대기전환`)
  - 버튼 클릭 시 DB 상태 동기화 + 메시지 갱신

## 1-1) 간단 운영 매뉴얼(실무용)

- 매일 오전 9시(스크립트 타임존 기준)에 `dailyDigestJob()`이 자동 실행됩니다.
- 자동 발송 내용은 `/cdv 예정` 성격의 "예정 요약"입니다.
- `DAILY_DM_USER_IDS`가 설정되어 있으면 해당 사용자에게 DM으로 발송됩니다.
- `DAILY_DM_USER_IDS`가 비어 있으면 `SLACK_CHANNEL_ID` 채널로 발송됩니다.
- 큐 상태 리포트(`dailyHealthCheck`)는 현재 비활성화되어 자동 발송되지 않습니다.

자주 쓰는 수동 명령:

- `/cdv 예정` : 예정 요약/스레드 조회
- `/cdv 완료` : 완료 건 조회
- `/cdv 취소` : 취소 건 조회
- `/cdv 예정 전체` : 기간 제한 없이 전체 조회
- `/cdv 완료 7` : 최근 7일 완료 건 조회
- `/cdv init` : 액션 큐 초기화

운영 점검(필요 시 Apps Script에서 수동 실행):

- `inspectSlackActionQueueStatus(false)` : 로그로 큐 상태 확인
- `inspectSlackActionQueueStatus(true)` : 채널로 큐 상태 전송
- `nudgeSlackActionQueue()` : 지연 항목 즉시 처리 대상으로 당김

## 2) 현재 운영 안정화 상태

현재 코드는 ACK-비동기 큐 구조로 전환되어, Slack 인터랙션 타임아웃(3초)을 구조적으로 완화합니다.

- `doPost`는 버튼 이벤트를 즉시 ACK 후 큐 적재
- `processSlackActionQueue`가 비동기로 실제 처리
- 단발 워커 + 1분 주기 워커 하이브리드
- 재시도 지수 백오프(`nextAttemptAt`) + jitter
- 큐 길이 상한(trim) 및 중복 적재 방지

## 3) 운영 가시성(이미 구현됨)

- `inspectSlackActionQueueStatus(notify)`
  - 큐 길이, 대기/지연/재시도 건수, 트리거 상태, 최근 에러 조회
- `nudgeSlackActionQueue()`
  - 지연된 큐 항목을 즉시 처리 대상으로 당김
- `dailyHealthCheck()`
  - 일일 큐 상태 리포트 발송(중복 방지 포함)

## 4) 트리거 운영(권장)

최초 1회 실행:

- `setupOperationsTriggers()`

내부적으로 아래를 설정:

- `setupSlackActionQueueTrigger()` : 1분 주기 큐 워커
- `setupDailyHealthCheckTrigger()` : 매일 오전 9시 상태 리포트
- `setupDailyDigestTrigger()` : 매일 오전 9시 5분 DM 요약 1회 발송

중복 생성 방지 로직이 있으므로 여러 번 실행해도 안전합니다.

## 5) 향후 자동화 로드맵

### 5.1 DM 발송(1회) 자동화

목표:

- 매일 지정 시각에 요약 DM 1회 발송
- 동일 날짜 중복 발송 방지

권장 구현:

- `dailyDigestJob()` 구현 완료
- Properties의 `DAILY_DIGEST_SENT_DATE`로 일일 중복 발송 방지
- 수신자는 `DAILY_DM_USER_IDS`(쉼표 구분 Slack User ID 목록) 사용

확장 옵션:

- 발송 대상 사용자(고정 1명 / 다수) 관리 시트 추가
- 테스트/운영 모드 분리(`DRY_RUN=true`)

### 5.2 Slack 채널 명령으로 최신 정보 DM 발송

목표:

- 채널에서 명령 입력 시 최신 상태를 DM으로 수신

권장 방식(택1):

- Slash Command(`/kfg-status`) + Apps Script Webhook
- 멘션 명령(`@봇 최신요약`) 파싱

요청 파라미터 예시:

- 범위: `전체`, `n(숫자 일수)`, `직접입력`
- 직접입력: `from=2026-04-01 to=2026-04-30`
- 상태필터: `완료만`, `대기만`, `전체`

처리 흐름:

1. 명령 수신/검증
2. 필터 파싱(기간/상태)
3. 원장DB 조회
4. DM 메시지 구성
5. 요청자에게 DM 발송

### 5.3 이미 완료된 건만 DM 발송(기간 필터 포함)

추천 파라미터 설계:

- `scope=all|7d|30d|custom`
- `from`, `to`(custom일 때 필수)

권장 데이터 기준:

- 상태 컬럼(L열) = `완료`
- 완료일 컬럼(M열) 기준으로 기간 필터

메시지 예시:

- 제목: `[완료 건 요약]`
- 본문: 기간, 건수, 합계
- 상세: 상위 N건 + 스레드 링크

## 6) 안정성/보안 체크리스트

- Bot token, channel id는 Script Properties 우선 사용
- Slack API 에러(`ratelimited`, quota 계열) 재시도
- 큐 처리 실패 로그 모니터링(일일 리포트 + 수동 점검)
- 대량 전송 시 배치 크기/재시도 간격 튜닝

## 7) 추천 다음 작업 순서

1. `setupOperationsTriggers()` 1회 실행
2. Script Properties에 `DAILY_DM_USER_IDS` 설정
3. Slack 명령 기반 DM(`전체/7일/30일/직접입력`) 구현
4. 완료건 전용 DM 템플릿 분리
5. 운영 대시보드 시트(최근 실패/큐 길이) 추가

## 7-1) 새 시트 + 새 채널 이관 방법

아래 순서대로 진행하면 안전하게 이관할 수 있습니다.

1. 새 스프레드시트 준비
   - 기존 원장 시트 구조를 동일하게 복사하고 시트 이름을 `원장DB`로 맞춥니다.
   - 헤더/열 순서(C열 결재번호, L열 상태 등)는 기존과 동일해야 합니다.
2. Apps Script 프로젝트 연결
   - 새 시트에서 Apps Script를 열고 현재 `Code.js`를 그대로 반영합니다.
   - 웹 앱 URL을 사용 중이라면 배포 버전을 최신으로 재배포합니다.
3. Slack 앱/권한 점검
   - Slash Command(`/cdv`)의 Request URL을 새 웹 앱 URL로 변경합니다.
   - Bot Token에 `chat:write`, `commands` 등 기존 권한이 유지되는지 확인합니다.
4. Script Properties 설정
   - `SLACK_BOT_TOKEN`: 운영 Bot 토큰
   - `SLACK_CHANNEL_ID`: 새 기본 채널 ID
   - `DAILY_DM_USER_IDS`: DM 수신자(쉼표 구분, 선택)
   - `MAX_THREAD_MESSAGES_PER_RUN`: 초기에는 `10` 권장
5. 트리거 재설정
   - `setupOperationsTriggers()`를 1회 실행해 트리거를 생성/정리합니다.
   - 기대 상태:
     - `processSlackActionQueue`: 1분 주기
     - `dailyDigestJob`: 매일 오전 9시
     - `dailyHealthCheck`: 없음(자동 비활성)
6. 운영 검증(필수)
   - `/cdv 예정` 수동 실행으로 메시지 발송/버튼 동작 확인
   - 테스트 건 1개로 `완료/취소/대기전환` 상태 동기화 확인
   - 다음날 오전 9시 자동 발송 여부 확인

## 8) 필수 설정 값

- `SLACK_BOT_TOKEN` : Slack Bot OAuth Token
- `SLACK_CHANNEL_ID` : 기본 발송 채널 ID
- `DAILY_DM_USER_IDS` : (선택) 일일 DM 수신자 Slack User ID 목록(예: `U0123,U0456`)
  - 미설정 시 `SLACK_CHANNEL_ID` 채널로 자동 발송
- `REQUESTER_SLACK_MAP` : (선택, fallback) 기안자명-슬랙ID 매핑 JSON
  - 예: `{"홍길동":"U0123ABCD","김철수":"U0456EFGH"}`

### 8-1) 기안자 개인 DM 매핑 방식(권장)

상태 변경(완료/취소/대기전환) 시 기안자에게 DM 안내를 보내기 위해 아래 우선순위로 Slack ID를 찾습니다.

1. 원장DB `V열`(슬랙ID) 값 사용 **(권장)**
2. `REQUESTER_SLACK_MAP` Script Properties 값 사용(fallback)

시트 컬럼 운영 권장:

- `U열`: 기안자명
- `V열`: 해당 기안자의 Slack User ID (`U...` 형식)

예시:

- `U열=홍길동`, `V열=U0123ABCD` → 상태 변경 시 `U0123ABCD`로 개인 DM 발송
- `V열`이 비어 있으면 `REQUESTER_SLACK_MAP`에서 `홍길동` 키를 찾아 발송
- 둘 다 없으면 DM은 스킵되고 로그에만 기록

## 9) 현재 안정 설정값(권장)

- `MAX_THREAD_MESSAGES_PER_RUN=10` (초기 안정 운영값, 필요 시 점진 상향)
- `DAILY_DM_USER_IDS` 미설정 시 채널 자동 발송 방식 유지
- 트리거는 `setupOperationsTriggers()`로 일괄 관리
- 버튼 상태변경은 동기 즉시 처리 경로 유지(큐는 보조/운영용)

## 10) 슬래시 명령 사용법

`doPost` 엔드포인트는 Slack Slash Command(`/cdv`)를 지원합니다.

### 10.1 명령 텍스트

- `예정` : 예정 스레드 발송 실행
- `완료` : 완료 목록 스레드 발송 실행
- `취소` : 취소 목록 스레드 발송 실행
- `init` : 액션 큐 비우기(초기화)
- `help` 또는 `도움말` : 사용 가능한 명령 안내

기간 지정(선택):

- `전체`
- `7` (최근 7일)
- `30` (최근 30일)
- `YYYY-MM-DD~YYYY-MM-DD`

예시:

- `예정 전체`
- `완료 7`
- `취소 2026-04-01~2026-04-30`

실행 방식:

- `/cd`는 즉시 접수 응답 후 백그라운드 실행
- 대상건수가 작으면(기본 5건 이하) 즉시 실행
- 완료 후 성공/실패 건수를 에페메랄로 재안내

### 10.2 Slack 설정 방법

1. Slack App 관리 페이지 접속
2. 좌측 `Slash Commands` 메뉴에서 `Create New Command` 선택
3. Command 입력 (예: `/cdv`)
4. Request URL에 Apps Script Web App URL 입력
5. Description/Usage Hint 입력 후 저장
6. Slack 워크스페이스에 앱 재설치(Install/ Reinstall) 수행

### 10.3 Apps Script 배포 체크

- `배포 > 새 배포 > 웹 앱`
- 실행 사용자: `나`
- 액세스 권한: `모든 사용자`(또는 Slack 호출 가능한 범위)
- 새 배포 후 나온 웹앱 URL을 Slash Command의 Request URL에 반영
- 코드 수정 시 재배포 후 URL/버전 갱신 여부 확인
