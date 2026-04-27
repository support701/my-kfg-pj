const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const SLACK_BOT_TOKEN = SCRIPT_PROPS.getProperty('SLACK_BOT_TOKEN') || '';
const SLACK_CHANNEL_ID = SCRIPT_PROPS.getProperty('SLACK_CHANNEL_ID') || 'C0AV006BC3U';

/**
 * 1. 지출 알림 전송 (Block Kit 적용)
 */
function sendDailyReminders(mode = 'pending') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('원장DB');
  const lastRow = sheet.getLastRow();
  if (lastRow < 10) return;

  const data = sheet.getRange(1, 1, lastRow, 19).getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const targetDays = [10, 7, 5, 3, 1];
  let groups = {}; 
  let grandTotal = 0;

  for (let i = 9; i < data.length; i++) {
    const row = data[i];
    const status = String(row[11] || "").trim();
    if (mode === 'pending') {
      if (['완료', '보류', '취소'].includes(status)) continue;
    } else {
      if (status !== '완료') continue;
    }

    const dueDateRaw = row[6];
    if (!dueDateRaw) continue;
    const dueDate = new Date(dueDateRaw);
    dueDate.setHours(0, 0, 0, 0);
    const diffDays = Math.round((dueDate - today) / (1000 * 60 * 60 * 24));

    if (mode === 'completed' || targetDays.includes(diffDays)) {
      const dayLabels = ['일', '월', '화', '수', '목', '금', '토'];
      const dateStr = `${Utilities.formatDate(dueDate, "GMT+9", "yyyy-MM-dd")}(${dayLabels[dueDate.getDay()]})`;
      const desc = String(row[5] || "내역 없음").trim();
      if (!groups[dateStr]) groups[dateStr] = { items: {}, isUrgent: (diffDays <= 7) };
      if (!groups[dateStr].items[desc]) groups[dateStr].items[desc] = [];
      groups[dateStr].items[desc].push({ data: row, rowNum: i + 1, diffDays: diffDays });
      grandTotal += Number(row[4] || 0);
    }
  }

  const sortedDates = Object.keys(groups).sort();
  if (sortedDates.length === 0) return;

  const title = mode === 'pending' ? `*[지출 집행 예정 요약]*` : `*[최근 완료 내역]*`;
  const summaryLines = buildParentSummaryLines(groups, mode, grandTotal);
  const parentMsg = `${title}\n\`\`\`\n${summaryLines}\n\`\`\`\n※ 상세내역은 아래 스래드를 확인하세요`;
  
  const parentTs = postToSlack(parentMsg);
  
  if (parentTs) {
    let itemIndex = 1;
    sortedDates.forEach(date => {
      const descGroups = groups[date].items;
      Object.keys(descGroups).forEach(desc => {
        descGroups[desc].forEach(item => {
          const blocks = buildDetailBlocks(item.data, (mode === 'completed'), "", itemIndex++, mode === 'completed' ? '완료' : '예정');
          try {
            const childTs = postToSlackBlocks(blocks, parentTs);
            if (childTs) sheet.getRange(item.rowNum, 17).setValue("'" + childTs);
          } catch (err) {
            Logger.log(`[Slack 전송 실패] row=${item.rowNum}, error=${err}`);
          }
        });
      });
    });
  }
}

/**
 * 2. 상세 메시지 Block Kit 빌더 (버튼 포함)
 */
function buildDetailBlocks(row, isComplete, userName, index, currentStatus) {
  const bank = String(row[8] || '').trim();
  const account = String(row[9] || '').trim();
  const owner = String(row[10] || '').trim();
  const itemTitle = String(row[5] || '내역 없음').trim();
  const target = String(row[7] || '-').trim();
  const approvalNo = String(row[2] || '-').trim();
  const memo = String(row[15] || '-').trim();
  const prefix = index ? `${index}. ` : "";
  const dueInfo = getDueInfo(row[6]);
  const amountRaw = Number(row[4] || 0);
  const amount = `${amountRaw.toLocaleString()}원`;
  const titleLine = `${prefix}[*${itemTitle}*] \`${approvalNo}\``;
  const detailCodeBlock = `\`\`\`\n - 집행예정: ${dueInfo.dateLabel} | ${dueInfo.ddayLabel}\n - 집행대상: ${target}\n - 계좌정보: ${bank} | ${account} | ${owner}\n - 담당메모: ${memo}\n\`\`\``;
  
  let mainText = "";
  if (currentStatus === "완료") {
    const approveDate = row[12] ? Utilities.formatDate(new Date(row[12]), "GMT+9", "yyyy-MM-dd HH:mm") : "-";
    const handlerName = extractDisplayName(userName || row[13] || '시스템');
    const doneTitle = `${index || "-"} ${itemTitle}_${target}_${amount}_${approvalNo}`;
    mainText = `\n*\`${doneTitle}\`*\n> ✅ 완료 | ${handlerName} | ${approveDate}`;
  } else if (currentStatus === "취소") {
    const cancelledAt = row[18] ? Utilities.formatDate(new Date(row[18]), "GMT+9", "yyyy-MM-dd HH:mm") : "-";
    const handlerName = extractDisplayName(userName || row[17] || '시스템');
    const cancelTitle = `${index || "-"} ${itemTitle}_${target}_${amount}_${approvalNo}`;
    mainText = `\n*\`${cancelTitle}\`*\n> ⛔ 취소 | ${handlerName} | ${cancelledAt}`;
  } else {
    const pendingTitleLine = `${prefix}${itemTitle} | \`${approvalNo}\``;
    mainText = `\n${pendingTitleLine}\n${detailCodeBlock}`;
  }

  // 버튼 섹션 추가
  const blocks = [
    { "type": "section", "text": { "type": "mrkdwn", "text": mainText } },
    {
      "type": "actions",
      "elements": [
        { "type": "button", "text": { "type": "plain_text", "text": "완료" }, "style": "primary", "action_id": "status_done", "value": String(index || "") },
        { "type": "button", "text": { "type": "plain_text", "text": "취소" }, "style": "danger", "action_id": "status_cancel", "value": String(index || "") },
        { "type": "button", "text": { "type": "plain_text", "text": "대기전환" }, "action_id": "status_pending", "value": String(index || "") }
      ]
    }
  ];
  return blocks;
}

/**
 * 3. 슬랙 인터랙션 수신 (doPost)
 */
function doPost(e) {
  if (!e || !e.postData) return;
  const payload = e.parameter.payload ? JSON.parse(e.parameter.payload) : JSON.parse(e.postData.contents);
  
  // URL 검증용
  if (payload.type === "url_verification") return ContentService.createTextOutput(payload.challenge);

  // 버튼 클릭(block_actions) 처리
  if (payload.type === "block_actions") {
    const action = payload.actions[0];
    const actionId = action.action_id;
    const ts = payload.container.message_ts;
    const userId = payload.user.id;
    const clickedUserName = payload.user && payload.user.name ? payload.user.name : userId;
    const displayIndex = extractDisplayIndex(action, payload.message && payload.message.blocks);

    let targetStatus = "";
    if (actionId === "status_done") targetStatus = "완료";
    else if (actionId === "status_cancel") targetStatus = "취소";
    else if (actionId === "status_pending") targetStatus = "예정";

    if (targetStatus) enqueueSlackStatusAction(ts, targetStatus, userId, clickedUserName, displayIndex);
    return ContentService.createTextOutput("");
  }
  
  // 기존 리액션 방식도 백업으로 유지
  if (payload.event && payload.event.type === "reaction_added") {
    const event = payload.event;
    if (event.reaction === "white_check_mark") {
      enqueueSlackStatusAction(event.item.ts, "완료", event.user, event.user, null);
    }
  }

  return ContentService.createTextOutput("ok");
}

/**
 * 4. 상태 업데이트 로직
 */
function handleStatusUpdate(ts, targetStatus, slackUserId, clickedUserName, displayIndex, maxApiAttempts, fastMode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('원장DB');
  const data = sheet.getRange(1, 1, sheet.getLastRow(), 19).getValues();
  const searchTs = String(ts).trim();

  for (let i = 9; i < data.length; i++) {
    if (String(data[i][16] || "").trim() === searchTs) {
      const userName = clickedUserName || getSlackUserName(slackUserId);
      const now = new Date();

      if (targetStatus === "완료") {
        sheet.getRange(i + 1, 12).setValue("완료");
        sheet.getRange(i + 1, 13).setValue(now);
        sheet.getRange(i + 1, 14).setValue(userName);
      } else {
        sheet.getRange(i + 1, 12).setValue(targetStatus);
        sheet.getRange(i + 1, 18).setValue(userName); // R열: 변경자
        sheet.getRange(i + 1, 19).setValue(now);      // S열: 변경일시
        if (targetStatus === "예정" && String(data[i][11] || "").trim() === "완료") {
          sheet.getRange(i + 1, 13).clearContent();   // M열: 집행일 초기화
          sheet.getRange(i + 1, 14).clearContent();   // N열: 처리자 초기화
        }
      }

      const updatedRow = sheet.getRange(i + 1, 1, 1, 19).getValues()[0];
      const updatedBlocks = buildDetailBlocks(updatedRow, (targetStatus === "완료"), userName, displayIndex, targetStatus);
      updateSlackMessageBlocks(ts, updatedBlocks, maxApiAttempts || 5, !!fastMode);
      break;
    }
  }
}

function enqueueSlackStatusAction(ts, targetStatus, slackUserId, clickedUserName, displayIndex) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const queue = getSlackActionQueue_();
    const action = {
      id: Utilities.getUuid(),
      ts: String(ts || "").trim(),
      targetStatus: String(targetStatus || "").trim(),
      slackUserId: String(slackUserId || "").trim(),
      clickedUserName: String(clickedUserName || "").trim(),
      displayIndex: displayIndex ? String(displayIndex).trim() : "",
      attempts: 0,
      enqueuedAt: new Date().toISOString(),
      nextAttemptAt: Date.now()
    };

    // 동일 상태의 중복 클릭 적재를 최소화
    const duplicate = queue.some(item =>
      item.ts === action.ts &&
      item.targetStatus === action.targetStatus &&
      item.slackUserId === action.slackUserId &&
      (item.displayIndex || "") === action.displayIndex
    );
    if (!duplicate) queue.push(action);
    const trimmedQueue = trimSlackQueue_(queue, 150);
    saveSlackActionQueue_(trimmedQueue);
  } finally {
    lock.releaseLock();
  }
  scheduleSlackActionQueueWorkerSoon_();
}

/**
 * 슬랙 인터랙션 비동기 처리용 워커
 * - 시간기반 트리거(예: 1분 간격)에서 실행 권장
 */
function processSlackActionQueue() {
  const maxBatch = 10;
  const maxAttemptsPerItem = 5;
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  let batch = [];
  let hasDeferredItems = false;
  try {
    const queue = getSlackActionQueue_();
    if (!queue.length) return;
    const now = Date.now();
    const remaining = [];
    queue.forEach(item => {
      const nextAttemptAt = Number(item.nextAttemptAt || 0);
      const isReady = !nextAttemptAt || nextAttemptAt <= now;
      if (isReady && batch.length < maxBatch) {
        batch.push(item);
      } else {
        remaining.push(item);
        if (!isReady) hasDeferredItems = true;
      }
    });
    saveSlackActionQueue_(remaining);
  } finally {
    lock.releaseLock();
  }

  const retryItems = [];
  batch.forEach(item => {
    try {
      handleStatusUpdate(
        item.ts,
        item.targetStatus,
        item.slackUserId,
        item.clickedUserName,
        item.displayIndex || null,
        3,
        false
      );
    } catch (err) {
      const nextAttempts = Number(item.attempts || 0) + 1;
      if (nextAttempts < maxAttemptsPerItem) {
        item.attempts = nextAttempts;
        item.lastError = String(err);
        item.nextAttemptAt = Date.now() + getQueueRetryDelayMs_(nextAttempts);
        retryItems.push(item);
      } else {
        Logger.log(`[Queue Drop] ts=${item.ts}, status=${item.targetStatus}, error=${err}`);
      }
    }
  });

  if (retryItems.length) {
    const appendLock = LockService.getScriptLock();
    appendLock.waitLock(5000);
    try {
      const queue = getSlackActionQueue_();
      saveSlackActionQueue_(trimSlackQueue_(retryItems.concat(queue), 150));
    } finally {
      appendLock.releaseLock();
    }
  }

  if (retryItems.length || hasDeferredItems) {
    scheduleSlackActionQueueWorkerSoon_();
  }
}

/**
 * 즉시 처리용 단발 워커
 * - 버튼 클릭 후 수초 내 처리를 위해 사용
 */
function processSlackActionQueueOnce() {
  processSlackActionQueue();
  cleanupOneShotQueueTriggers_();
}

function setupSlackActionQueueTrigger() {
  const handler = "processSlackActionQueue";
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(t => t.getHandlerFunction() === handler);
  if (exists) return;

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyMinutes(1)
    .create();
}

/**
 * 5. 슬랙 API 보조 함수 (Block 전송용)
 */
function postToSlackBlocks(blocks, threadTs = null) {
  const payload = { channel: SLACK_CHANNEL_ID, blocks: blocks };
  if (threadTs) payload.thread_ts = threadTs;
  const result = slackApiCall("https://slack.com/api/chat.postMessage", payload);
  if (!result.ok) throw new Error(`Slack post 실패: ${result.error || 'unknown_error'}`);
  return result.ts;
}

function updateSlackMessageBlocks(ts, blocks, maxApiAttempts, fastMode) {
  const result = slackApiCall(
    "https://slack.com/api/chat.update",
    { channel: SLACK_CHANNEL_ID, ts: ts, blocks: blocks },
    maxApiAttempts,
    { fastMode: !!fastMode }
  );
  if (!result.ok) throw new Error(`Slack update 실패: ${result.error || 'unknown_error'}`);
}

// 나머지 기존 보조 함수(getSlackUserName, postToSlack, getSlackMessageText 등)는 그대로 유지
function getSlackUserName(userId) {
  try {
    const response = UrlFetchApp.fetch(`https://slack.com/api/users.info?user=${userId}`, {
      headers: { "Authorization": "Bearer " + SLACK_BOT_TOKEN }
    });
    const resData = JSON.parse(response.getContentText());
    return resData.ok ? (resData.user.real_name || resData.user.name) : userId;
  } catch (e) { return "조회실패"; }
}

function postToSlack(message) {
  const result = slackApiCall("https://slack.com/api/chat.postMessage", { channel: SLACK_CHANNEL_ID, text: message });
  if (!result.ok) throw new Error(`Slack post 실패: ${result.error || 'unknown_error'}`);
  return result.ts;
}

function getSlackMessageText(ts) {
  try {
    const response = UrlFetchApp.fetch(`https://slack.com/api/conversations.replies?channel=${SLACK_CHANNEL_ID}&ts=${ts}&limit=1`, {
      headers: { "Authorization": "Bearer " + SLACK_BOT_TOKEN }
    });
    const resData = JSON.parse(response.getContentText());
    if (resData.ok && resData.messages.length > 0) return resData.messages[0].text || "1.";
  } catch (e) {}
  return "1.";
}

function buildParentSummaryLines(groups, mode, grandTotal) {
  const dates = Object.keys(groups).sort();
  const lines = [];

  dates.forEach(date => {
    const descGroups = groups[date].items;
    Object.keys(descGroups).forEach(desc => {
      const rows = descGroups[desc];
      const total = rows.reduce((acc, item) => acc + Number(item.data[4] || 0), 0);
      const isUrgent = mode === 'pending' && rows.some(item => Number(item.diffDays) <= 7);
      const urgentPrefix = isUrgent ? '[🚨임박] ' : '';
      lines.push(`${urgentPrefix}${date} | ${desc} | ${total.toLocaleString()}원 (총 ${rows.length}건)`);
    });
  });

  if (!lines.length) return "요약 데이터 없음";

  const separator = "--------------------------------";
  const totalLabel = mode === 'pending' ? "[💰 지출 집행 예정 총계]" : "✅ [최근 완료 총계]";
  return `${lines.join('\n')}\n\n${separator}\n${totalLabel} : ${Number(grandTotal || 0).toLocaleString()}원`;
}

function extractDisplayName(rawName) {
  const name = String(rawName || '').trim();
  if (!name) return '-';
  return name.split(/\s|\(/)[0];
}

function getDueInfo(dueDateRaw) {
  if (!dueDateRaw) {
    return { dateLabel: '-', ddayLabel: 'D-?' };
  }

  const dueDate = new Date(dueDateRaw);
  dueDate.setHours(0, 0, 0, 0);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const dayLabels = ['일', '월', '화', '수', '목', '금', '토'];
  const dateLabel = `${Utilities.formatDate(dueDate, "GMT+9", "yyyy-MM-dd")}(${dayLabels[dueDate.getDay()]})`;
  const diffDays = Math.round((dueDate - today) / (1000 * 60 * 60 * 24));
  const ddayLabel = diffDays >= 0 ? `D-${diffDays}` : `D+${Math.abs(diffDays)}`;
  return { dateLabel: dateLabel, ddayLabel: ddayLabel };
}

function slackApiCall(url, payload, maxAttempts, opts) {
  if (!SLACK_BOT_TOKEN) {
    throw new Error("SLACK_BOT_TOKEN이 설정되지 않았습니다. Script Properties에 등록하세요.");
  }

  const attempts = maxAttempts || 5;
  const options = opts || {};
  const fastMode = !!options.fastMode;
  let waitMs = 1200;

  for (let i = 1; i <= attempts; i++) {
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      headers: { "Authorization": "Bearer " + SLACK_BOT_TOKEN },
      contentType: "application/json",
      muteHttpExceptions: true,
      payload: JSON.stringify(payload)
    });

    const statusCode = response.getResponseCode();
    const bodyText = response.getContentText();
    let body = {};
    try {
      body = JSON.parse(bodyText || "{}");
    } catch (e) {
      body = { ok: false, error: `invalid_json:${bodyText}` };
    }

    // Slack rate limit(429) 또는 quota 계열 에러일 때 백오프 후 재시도
    if (statusCode === 429 || isRetryableSlackError(body.error)) {
      if (i === attempts) return body;
      const retryAfterSec = Number(response.getHeaders()["Retry-After"] || 0);
      const sleepMs = retryAfterSec > 0 ? retryAfterSec * 1000 : (fastMode ? 250 : waitMs);
      Utilities.sleep(sleepMs);
      waitMs = Math.min(waitMs * 2, 12000);
      continue;
    }

    // 정상 혹은 비재시도 에러는 즉시 반환
    if (body.ok || statusCode < 500) {
      if (!fastMode) Utilities.sleep(350); // 인터랙션 경로에서는 지연 최소화
      return body;
    }

    if (i === attempts) return body;
    Utilities.sleep(waitMs);
    waitMs = Math.min(waitMs * 2, 12000);
  }

  return { ok: false, error: "unknown_retry_exhausted" };
}

function isRetryableSlackError(errorCode) {
  const code = String(errorCode || "");
  return [
    "ratelimited",
    "rate_limited",
    "bandwidth_quota_exceeded",
    "request_timeout",
    "internal_error"
  ].indexOf(code) >= 0;
}

function extractItemIndexFromBlocks(blocks) {
  try {
    if (!blocks || !blocks.length) return null;
    const first = blocks[0];
    const text = first && first.text ? String(first.text.text || "") : "";
    const match = text.match(/(?:^|\n)\s*(\d+)\.\s*/);
    return match ? match[1] : null;
  } catch (e) {
    return null;
  }
}

function extractDisplayIndex(action, blocks) {
  const fromButton = action && action.value ? String(action.value).trim() : "";
  if (/^\d+$/.test(fromButton)) return fromButton;
  return extractItemIndexFromBlocks(blocks);
}

function getSlackActionQueue_() {
  const raw = SCRIPT_PROPS.getProperty("SLACK_ACTION_QUEUE") || "[]";
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function saveSlackActionQueue_(queue) {
  SCRIPT_PROPS.setProperty("SLACK_ACTION_QUEUE", JSON.stringify(queue || []));
}

function trimSlackQueue_(queue, maxSize) {
  const limit = Number(maxSize || 150);
  if (!Array.isArray(queue)) return [];
  if (queue.length <= limit) return queue;
  const trimmed = queue.slice(queue.length - limit);
  Logger.log(`[Queue Trim] ${queue.length} -> ${trimmed.length}`);
  return trimmed;
}

function getQueueRetryDelayMs_(attempts) {
  const base = 2000;
  const maxDelay = 60000;
  const jitter = Math.floor(Math.random() * 700);
  const delay = Math.min(base * Math.pow(2, Math.max(0, attempts - 1)) + jitter, maxDelay);
  return delay;
}

function scheduleSlackActionQueueWorkerSoon_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const now = Date.now();
    const cooldownMs = 15000; // 단발 트리거 과생성 방지
    const lastScheduled = Number(SCRIPT_PROPS.getProperty("SLACK_QUEUE_ONESHOT_SCHEDULED_AT") || 0);
    if (lastScheduled && now - lastScheduled < cooldownMs) return;

    const exists = ScriptApp.getProjectTriggers()
      .some(t => t.getHandlerFunction() === "processSlackActionQueueOnce");
    if (!exists) {
      ScriptApp.newTrigger("processSlackActionQueueOnce")
        .timeBased()
        .after(5000)
        .create();
    }
    SCRIPT_PROPS.setProperty("SLACK_QUEUE_ONESHOT_SCHEDULED_AT", String(now));
  } finally {
    lock.releaseLock();
  }
}

function cleanupOneShotQueueTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === "processSlackActionQueueOnce") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * 운영 가시성: 큐 상태 요약 조회
 * - Apps Script 실행 로그에서 바로 확인 가능
 * - 필요시 notify=true 로 Slack 채널에 상태 전송
 */
function inspectSlackActionQueueStatus(notify) {
  const now = Date.now();
  const queue = getSlackActionQueue_();

  const readyItems = [];
  const delayedItems = [];
  const retryItems = [];
  let maxAttempts = 0;

  queue.forEach(item => {
    const attempts = Number(item.attempts || 0);
    const nextAttemptAt = Number(item.nextAttemptAt || 0);
    if (!nextAttemptAt || nextAttemptAt <= now) readyItems.push(item);
    else delayedItems.push(item);
    if (attempts > 0) retryItems.push(item);
    if (attempts > maxAttempts) maxAttempts = attempts;
  });

  const oldestEnqueuedAt = queue
    .map(item => Number(new Date(item.enqueuedAt || 0)))
    .filter(v => Number.isFinite(v) && v > 0)
    .sort((a, b) => a - b)[0] || null;

  const nearestReadyAt = delayedItems
    .map(item => Number(item.nextAttemptAt || 0))
    .filter(v => Number.isFinite(v) && v > 0)
    .sort((a, b) => a - b)[0] || null;

  const recentErrors = queue
    .filter(item => item.lastError)
    .slice(-3)
    .map(item => `- ts:${item.ts}, status:${item.targetStatus}, attempts:${item.attempts}, err:${String(item.lastError).slice(0, 80)}`);

  const status = {
    checkedAt: new Date().toISOString(),
    queueSize: queue.length,
    readyCount: readyItems.length,
    delayedCount: delayedItems.length,
    retryCount: retryItems.length,
    maxAttempts: maxAttempts,
    oldestWaitSeconds: oldestEnqueuedAt ? Math.floor((now - oldestEnqueuedAt) / 1000) : 0,
    nextReadyInSeconds: nearestReadyAt ? Math.max(0, Math.floor((nearestReadyAt - now) / 1000)) : 0,
    hasOneShotTrigger: ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "processSlackActionQueueOnce"),
    hasMinuteTrigger: ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "processSlackActionQueue"),
    recentErrors: recentErrors
  };

  const text = buildQueueStatusText_(status);
  Logger.log(text);
  if (notify) {
    try {
      postToSlack(`*[Queue 상태 점검]*\n\`\`\`\n${text}\n\`\`\``);
    } catch (e) {
      Logger.log(`[Queue Status Notify Failed] ${e}`);
    }
  }
  return status;
}

/**
 * 운영 가시성: 실패/지연 건 수동 재가속
 * - nextAttemptAt을 현재 시각으로 당겨 즉시 처리 대상으로 전환
 */
function nudgeSlackActionQueue() {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const now = Date.now();
    const queue = getSlackActionQueue_().map(item => {
      item.nextAttemptAt = now;
      return item;
    });
    saveSlackActionQueue_(queue);
  } finally {
    lock.releaseLock();
  }
  scheduleSlackActionQueueWorkerSoon_();
}

function buildQueueStatusText_(status) {
  const lines = [
    `checkedAt: ${status.checkedAt}`,
    `queueSize: ${status.queueSize}`,
    `readyCount: ${status.readyCount}`,
    `delayedCount: ${status.delayedCount}`,
    `retryCount: ${status.retryCount}`,
    `maxAttempts: ${status.maxAttempts}`,
    `oldestWaitSeconds: ${status.oldestWaitSeconds}`,
    `nextReadyInSeconds: ${status.nextReadyInSeconds}`,
    `hasMinuteTrigger: ${status.hasMinuteTrigger}`,
    `hasOneShotTrigger: ${status.hasOneShotTrigger}`
  ];
  if (status.recentErrors && status.recentErrors.length) {
    lines.push("recentErrors:");
    status.recentErrors.forEach(line => lines.push(line));
  }
  return lines.join('\n');
}

/**
 * 매일 오전 9시(스크립트 타임존 기준) 큐 상태 자동 리포트
 * - 같은 날짜에 중복 실행되더라도 1회만 발송
 */
function dailyHealthCheck() {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const todayKey = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
    const sentKey = SCRIPT_PROPS.getProperty("DAILY_HEALTHCHECK_SENT_DATE");
    if (sentKey === todayKey) return;

    const status = inspectSlackActionQueueStatus(false);
    const text = buildQueueStatusText_(status);
    postToSlack(`*[일일 Queue 상태 리포트]*\n\`\`\`\n${text}\n\`\`\``);
    SCRIPT_PROPS.setProperty("DAILY_HEALTHCHECK_SENT_DATE", todayKey);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 운영 트리거 일괄 설정
 * - 큐 워커(1분)
 * - 일일 상태 리포트(오전 9시)
 */
function setupOperationsTriggers() {
  setupSlackActionQueueTrigger();
  setupDailyHealthCheckTrigger();
}

function setupDailyHealthCheckTrigger() {
  const handler = "dailyHealthCheck";
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === handler);
  if (exists) return;

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .nearMinute(0)
    .create();
}