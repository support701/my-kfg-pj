# 지출 알림/스레드 운영 가이드

## 1) 현재 프로젝트 구성 요약

- 실행 환경: Google Apps Script
- 데이터 소스: 스프레드시트 `원장DB`
- 알림 채널: Slack Bot (`chat.postMessage`, `chat.update`)
- 주요 기능:
  - DM/채널 요약 메시지 발송
  - 스레드 카드 생성 및 상태 버튼(`완료`, `취소`, `대기전환`)
  - 버튼 클릭 시 DB 상태 동기화 + 메시지 갱신

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

중복 생성 방지 로직이 있으므로 여러 번 실행해도 안전합니다.

## 5) 향후 자동화 로드맵

### 5.1 DM 발송(1회) 자동화

목표:

- 매일 지정 시각에 요약 DM 1회 발송
- 동일 날짜 중복 발송 방지

권장 구현:

- `dailyDigestJob()` 함수 추가
- Properties에 `LAST_DM_SENT_DATE_YYYYMMDD` 기록
- 이미 발송된 날짜면 skip, 아니면 발송 후 기록

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

- 범위: `전체`, `최근7일`, `최근한달`, `직접입력`
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
2. `dailyDigestJob()`(1회 자동 DM) 구현
3. Slack 명령 기반 DM(`전체/7일/30일/직접입력`) 구현
4. 완료건 전용 DM 템플릿 분리
5. 운영 대시보드 시트(최근 실패/큐 길이) 추가
