const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const SLACK_BOT_TOKEN = SCRIPT_PROPS.getProperty('SLACK_BOT_TOKEN') || '';
const SLACK_CHANNEL_ID = SCRIPT_PROPS.getProperty('SLACK_CHANNEL_ID') || 'C0B0MAD3QQZ';

/**
 * 1. 지출 알림 전송 (Block Kit 적용)
 */
function sendDailyReminders(mode = 'pending', opts) {
  const options = opts || {};
  const period = options.period || { type: "default" };
  const stats = {
    mode: mode,
    ok: false,
    skipped: false,
    reason: "",
    parentSent: false,
    parentTs: "",
    candidateCount: 0,
    sentCount: 0,
    failedCount: 0,
    quotaHit: false
  };

  const runLock = LockService.getScriptLock();
  if (!runLock.tryLock(3000)) {
    Logger.log(`[sendDailyReminders] 다른 실행이 진행 중이라 스킵: mode=${mode}`);
    stats.skipped = true;
    stats.reason = "locked";
    return stats;
  }
  try {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('원장DB');
  const lastRow = sheet.getLastRow();
  if (lastRow < 10) {
    stats.ok = true;
    stats.reason = "no_data";
    return stats;
  }

  const data = sheet.getRange(1, 1, lastRow, 22).getValues();
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
    } else if (mode === 'completed') {
      if (status !== '완료') continue;
    } else if (mode === 'cancelled') {
      if (status !== '취소') continue;
    } else {
      continue;
    }

    const dueDateRaw = row[6];
    if (!dueDateRaw) continue;
    const dueDate = new Date(dueDateRaw);
    dueDate.setHours(0, 0, 0, 0);
    const diffDays = Math.round((dueDate - today) / (1000 * 60 * 60 * 24));
    const referenceDate = getReferenceDateForMode_(row, mode, dueDate);
    if (shouldIncludeByPeriod_(mode, diffDays, referenceDate, period, targetDays, today)) {
      const dayLabels = ['일', '월', '화', '수', '목', '금', '토'];
      const dateStr = `${Utilities.formatDate(referenceDate, "GMT+9", "yyyy-MM-dd")}(${dayLabels[referenceDate.getDay()]})`;
      const desc = String(row[5] || "내역 없음").trim();
      if (!groups[dateStr]) groups[dateStr] = { items: {}, isUrgent: (diffDays <= 7) };
      if (!groups[dateStr].items[desc]) groups[dateStr].items[desc] = [];
      groups[dateStr].items[desc].push({ data: row, rowNum: i + 1, diffDays: diffDays });
      grandTotal += Number(row[4] || 0);
    }
  }

  const sortedDates = Object.keys(groups).sort();
  if (sortedDates.length === 0) {
    stats.ok = true;
    stats.reason = "no_target_items";
    return stats;
  }

  sortedDates.forEach(date => {
    const descGroups = groups[date].items;
    Object.keys(descGroups).forEach(desc => {
      stats.candidateCount += descGroups[desc].length;
    });
  });

  const title = mode === 'pending' ? `*[지출 집행 예정 요약]*` : (mode === 'completed' ? `*[최근 완료 내역]*` : `*[최근 취소 내역]*`);
  const summaryLines = buildParentSummaryLines(groups, mode, grandTotal);
  const parentMsg = `${title}\n\`\`\`\n${summaryLines}\n\`\`\`\n※ 상세내역은 아래 스래드를 확인하세요`;
  
  let parentTs = null;
  try {
    parentTs = postToSlack(parentMsg);
    stats.parentSent = true;
    stats.parentTs = String(parentTs || "");
  } catch (e) {
    Logger.log(`[sendDailyReminders] parent message 전송 실패: ${e}`);
    stats.reason = "parent_post_failed";
    return stats;
  }
  
  if (parentTs) {
    let itemIndex = 1;
    const configuredLimit = Number(SCRIPT_PROPS.getProperty("MAX_THREAD_MESSAGES_PER_RUN") || 25);
    const maxThreadMessagesPerRun = (!isNaN(configuredLimit) && configuredLimit > 0) ? Math.floor(configuredLimit) : 25;
    let sentCount = 0;
    let quotaHit = false;
    let failedCount = 0;
    sortedDates.forEach(date => {
      if (quotaHit) return;
      const descGroups = groups[date].items;
      Object.keys(descGroups).forEach(desc => {
        if (quotaHit) return;
        descGroups[desc].forEach(item => {
          if (quotaHit) return;
          if (sentCount >= maxThreadMessagesPerRun) {
            quotaHit = true;
            Logger.log(`[sendDailyReminders] 실행당 상한 도달: ${maxThreadMessagesPerRun}`);
            return;
          }
          const initialStatus = mode === 'completed' ? '완료' : (mode === 'cancelled' ? '취소' : '예정');
          const blocks = buildDetailBlocks(item.data, (mode === 'completed'), "", itemIndex++, initialStatus);
          try {
            const childTs = postToSlackBlocks(blocks, parentTs);
            if (childTs) {
              sheet.getRange(item.rowNum, 17).setValue("'" + childTs);
              registerApprovalThreadTs_(String(item.data[2] || "").trim(), childTs);
              sentCount += 1;
            } else {
              quotaHit = true;
            }
          } catch (err) {
            Logger.log(`[Slack 전송 실패] row=${item.rowNum}, error=${err}`);
            const errText = String(err || "");
            if (errText.indexOf("Bandwidth quota exceeded") >= 0) {
              quotaHit = true;
            } else {
              // 블록 전송 실패 시 텍스트 전송으로 1회 폴백
              try {
                const fallbackText = buildDetailFallbackText_(item.data, itemIndex - 1, mode === 'completed' ? '완료' : '예정');
                const fallbackTs = postToSlack(fallbackText, parentTs);
                if (fallbackTs) {
                  sheet.getRange(item.rowNum, 17).setValue("'" + fallbackTs);
                  registerApprovalThreadTs_(String(item.data[2] || "").trim(), fallbackTs);
                  sentCount += 1;
                } else {
                  failedCount += 1;
                }
              } catch (fallbackErr) {
                failedCount += 1;
                Logger.log(`[Slack 폴백 전송 실패] row=${item.rowNum}, error=${fallbackErr}`);
              }
            }
          }
          Utilities.sleep(120); // 짧은 간격으로 burst 전송 완화
        });
      });
    });

    if (failedCount > 0) {
      try {
        postToSlack(`※ 일부 스레드 전송 실패: ${failedCount}건 (로그 확인 필요)`, parentTs);
      } catch (e) {
        Logger.log(`[sendDailyReminders] 실패 요약 알림 전송 실패: ${e}`);
      }
    }
    stats.sentCount = sentCount;
    stats.failedCount = failedCount;
    stats.quotaHit = quotaHit;
    stats.ok = true;
    stats.reason = quotaHit ? "quota_or_limit_stop" : "done";
  }
  return stats;
  } finally {
    runLock.releaseLock();
  }
}

/**
 * 수동 실행용: 지출 예정 스레드 발송
 */
function runPendingRemindersNow() {
  return sendDailyReminders("pending");
}

/**
 * 수동 실행용: 완료 목록 스레드 발송
 */
function runCompletedRemindersNow() {
  return sendDailyReminders("completed");
}

function runCancelledRemindersNow() {
  return sendDailyReminders("cancelled");
}

/**
 * 수동 실행용: 결재번호 기준 상태 변경
 * - approvalNo: 원장DB C열 결재번호
 * - targetStatus: "완료" | "취소" | "예정"
 * - actorName: 처리자명(선택)
 */
function changeStatusByApprovalNo(approvalNo, targetStatus, actorName) {
  const status = String(targetStatus || "").trim();
  if (["완료", "취소", "예정"].indexOf(status) < 0) {
    throw new Error(`지원하지 않는 상태값: ${targetStatus}`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("원장DB");
  const rowNum = findRowByApprovalNo_(sheet, approvalNo);
  if (!rowNum) {
    throw new Error(`결재번호를 찾지 못했습니다: ${approvalNo}`);
  }

  const ts = normalizeTs_(sheet.getRange(rowNum, 17).getValue()); // Q열 thread ts
  const displayName = String(actorName || "수동처리").trim();

  // 스레드 ts가 있으면 기존 동기 처리 로직 재사용(DB+메시지 동시 변경)
  if (ts) {
    return handleStatusUpdate(ts, status, "", displayName, null, 2, true, String(approvalNo || "").trim());
  }

  // 스레드가 없는 경우 DB 상태만 갱신
  const now = new Date();
  const currentStatus = String(sheet.getRange(rowNum, 12).getValue() || "").trim();
  if (status === "완료") {
    sheet.getRange(rowNum, 12).setValue("완료");
    sheet.getRange(rowNum, 13).setValue(now);
    sheet.getRange(rowNum, 14).setValue(displayName);
  } else {
    sheet.getRange(rowNum, 12).setValue(status);
    sheet.getRange(rowNum, 18).setValue(displayName);
    sheet.getRange(rowNum, 19).setValue(now);
    if (status === "예정" && currentStatus === "완료") {
      sheet.getRange(rowNum, 13).clearContent();
      sheet.getRange(rowNum, 14).clearContent();
    }
  }
  const updatedRow = sheet.getRange(rowNum, 1, 1, 22).getValues()[0];
  notifyRequesterStatusChange_(updatedRow, status, displayName, currentStatus);
  return true;
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
  const requester = String(row[20] || '-').trim();
  const prefix = index ? `${index}. ` : "";
  const dueInfo = getDueInfo(row[6]);
  const amountRaw = Number(row[4] || 0);
  const amount = `${amountRaw.toLocaleString()}원`;
  const detailCodeBlock = `\`\`\`\n - 기안자명: ${requester}\n - 집행예정: ${dueInfo.dateLabel} | ${dueInfo.ddayLabel}\n - 집행대상: ${target}\n - 결제금액: ${amount}\n - 입금계좌: ${bank} | ${account} | ${owner}\n - 담당메모: ${memo}\n\`\`\``;
  
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
    const pendingTitleLine = `*${prefix}${itemTitle}* | *\`${approvalNo}\`*`;
    mainText = `\n${pendingTitleLine}\n${detailCodeBlock}`;
  }

  // 버튼 섹션 추가 (상태별 노출 제어)
  const actionButtons = getActionButtonsForStatus_(currentStatus, index);
  const blocks = [
    { "type": "section", "text": { "type": "mrkdwn", "text": mainText } },
    {
      "type": "actions",
      "elements": actionButtons
    }
  ];
  return blocks;
}

function getActionButtonsForStatus_(status, index) {
  const value = String(index || "");
  const normalized = String(status || "").trim();
  if (normalized === "완료") {
    return [
      { "type": "button", "text": { "type": "plain_text", "text": "취소" }, "style": "danger", "action_id": "status_cancel", "value": value },
      { "type": "button", "text": { "type": "plain_text", "text": "대기전환" }, "action_id": "status_pending", "value": value }
    ];
  }
  if (normalized === "취소") {
    return [
      { "type": "button", "text": { "type": "plain_text", "text": "완료" }, "style": "primary", "action_id": "status_done", "value": value },
      { "type": "button", "text": { "type": "plain_text", "text": "대기전환" }, "action_id": "status_pending", "value": value }
    ];
  }
  return [
    { "type": "button", "text": { "type": "plain_text", "text": "완료" }, "style": "primary", "action_id": "status_done", "value": value },
    { "type": "button", "text": { "type": "plain_text", "text": "취소" }, "style": "danger", "action_id": "status_cancel", "value": value }
  ];
}

/**
 * 3. 슬랙 인터랙션 수신 (doPost)
 */
function doPost(e) {
  try {
    if (!e || !e.postData) return ContentService.createTextOutput("");

    // Slash Command 처리 (application/x-www-form-urlencoded)
    if (e.parameter && e.parameter.command) {
      return handleSlashCommand_(e.parameter);
    }

    const payload = parseSlackPayload_(e);
    if (!payload) return ContentService.createTextOutput("");
    
    // URL 검증용
    if (payload.type === "url_verification") return ContentService.createTextOutput(payload.challenge);

    // 버튼 클릭(block_actions) 처리
    if (payload.type === "block_actions") {
      const action = payload.actions && payload.actions.length ? payload.actions[0] : null;
      if (!action) return ContentService.createTextOutput("");

      const actionId = action.action_id;
      const ts = payload.container && payload.container.message_ts ? payload.container.message_ts : "";
      const userId = payload.user && payload.user.id ? payload.user.id : "";
      const clickedUserName = payload.user && payload.user.name ? payload.user.name : userId;
      const displayIndex = extractDisplayIndex(action, payload.message && payload.message.blocks);
      const fallbackApprovalNo = extractApprovalNoFromBlocks_(payload.message && payload.message.blocks);

      let targetStatus = "";
      if (actionId === "status_done") targetStatus = "완료";
      else if (actionId === "status_cancel") targetStatus = "취소";
      else if (actionId === "status_pending") targetStatus = "예정";

      if (targetStatus && ts) {
        // 롤백: 구조 분리 이전처럼 버튼 클릭 시 즉시 동기 처리
        let handled = false;
        try {
          handled = handleStatusUpdate(ts, targetStatus, userId, clickedUserName, displayIndex, 2, true, fallbackApprovalNo);
        } catch (innerErr) {
          Logger.log(`[doPost Immediate Failed] ts=${ts}, status=${targetStatus}, error=${innerErr}`);
        }
        if (!handled) {
          Logger.log(`[doPost NotHandled] ts=${ts}, status=${targetStatus}, approvalNo=${fallbackApprovalNo}`);
        }
      }
      return ContentService.createTextOutput("");
    }
    
    // 기존 리액션 방식도 백업으로 유지
    if (payload.event && payload.event.type === "reaction_added") {
      const event = payload.event;
      if (event.reaction === "white_check_mark") {
        handleStatusUpdate(event.item.ts, "완료", event.user, event.user, null, 2, true);
      }
    }

    return ContentService.createTextOutput("ok");
  } catch (err) {
    // Slack 인터랙션에서 비-200 응답을 방지하기 위해 빈 200 응답 반환
    Logger.log(`[doPost Failed] ${err}`);
    return ContentService.createTextOutput("");
  }
}

function parseSlackPayload_(e) {
  const parameterPayload = e && e.parameter && e.parameter.payload ? String(e.parameter.payload) : "";
  if (parameterPayload) {
    try {
      return JSON.parse(parameterPayload);
    } catch (err) {
      Logger.log(`[parseSlackPayload_] parameter payload parse failed: ${err}`);
    }
  }

  const body = e && e.postData && e.postData.contents ? String(e.postData.contents) : "";
  if (!body) return null;

  // Slack 인터랙션 표준 포맷: payload={url-encoded-json}
  if (body.indexOf("payload=") === 0) {
    const encoded = body.substring("payload=".length);
    try {
      return JSON.parse(decodeURIComponent(encoded.replace(/\+/g, "%20")));
    } catch (err) {
      Logger.log(`[parseSlackPayload_] urlencoded payload parse failed: ${err}`);
      return null;
    }
  }

  // JSON body 포맷 fallback
  try {
    return JSON.parse(body);
  } catch (err) {
    Logger.log(`[parseSlackPayload_] raw body parse failed: ${err}`);
    return null;
  }
}

function handleSlashCommand_(params) {
  const text = String(params.text || "").trim();
  const userName = String(params.user_name || params.user_id || "슬랙사용자").trim();
  const parsed = parseSlashCommandText_(text);

  if (!parsed.ok) {
    return jsonResponse_({
      response_type: "ephemeral",
      text: buildSlashHelpText_(`명령어 해석 실패: ${parsed.error}`)
    });
  }

  try {
    if (parsed.type === "pending") {
      return handleListCommandHybrid_("pending", parsed.period, userName, String(params.response_url || "").trim());
    }

    if (parsed.type === "completed") {
      return handleListCommandHybrid_("completed", parsed.period, userName, String(params.response_url || "").trim());
    }

    if (parsed.type === "cancelled") {
      return handleListCommandHybrid_("cancelled", parsed.period, userName, String(params.response_url || "").trim());
    }

    if (parsed.type === "clear_queue") {
      clearSlackActionQueue();
      return jsonResponse_({ response_type: "ephemeral", text: "액션 큐를 비웠습니다." });
    }

    return jsonResponse_({
      response_type: "ephemeral",
      text: buildSlashHelpText_("지원하지 않는 명령입니다.")
    });
  } catch (err) {
    Logger.log(`[SlashCommand Failed] text=${text}, error=${err}`);
    return jsonResponse_({
      response_type: "ephemeral",
      text: `명령 처리 중 오류가 발생했습니다: ${err}`
    });
  }
}

function handleListCommandHybrid_(mode, period, userName, responseUrl) {
  const count = getReminderCandidateCount_(mode, period);
  const modeLabel = mode === "pending" ? "예정" : (mode === "completed" ? "완료" : "취소");
  const inlineThreshold = 5;

  if (count <= inlineThreshold) {
    const result = sendDailyReminders(mode, { period: period });
    return jsonResponse_({
      response_type: "ephemeral",
      text: `${modeLabel} 스레드 발송 완료 (대상 ${count}건 / 성공 ${Number(result.sentCount || 0)}건, 실패 ${Number(result.failedCount || 0)}건)`
    });
  }

  enqueueSlashCommandJob_(mode, {
        requestedBy: userName,
        responseUrl: responseUrl,
        period: period
      });
  return jsonResponse_({
    response_type: "ephemeral",
    text: `${modeLabel} 스레드 발송을 접수했습니다. (대상 ${count}건) 완료 후 결과를 안내합니다.`
  });
}

function parseSlashCommandText_(text) {
  const raw = String(text || "").trim();
  if (!raw || raw === "help" || raw === "도움말") {
    return { ok: false, error: "help" };
  }

  const parts = raw.split(/\s+/).filter(v => !!v);
  const command = String(parts[0] || "").toLowerCase();
  const periodRaw = parts.length > 1 ? parts.slice(1).join(" ") : "";
  const period = parsePeriodToken_(periodRaw);
  if (!period.ok) return { ok: false, error: period.error };

  if (["예정", "pending", "p"].indexOf(command) >= 0) {
    return { ok: true, type: "pending", period: period.value };
  }
  if (["완료", "completed", "c"].indexOf(command) >= 0) {
    return { ok: true, type: "completed", period: period.value };
  }
  if (["취소", "cancelled", "cancel", "x"].indexOf(command) >= 0) {
    return { ok: true, type: "cancelled", period: period.value };
  }
  if (["init", "initialize"].indexOf(command) >= 0) {
    return { ok: true, type: "clear_queue" };
  }

  return { ok: false, error: "지원하지 않는 형식입니다." };
}

function buildSlashHelpText_(prefix) {
  const help = [
    prefix || "명령어를 입력하세요.",
    "",
    "[사용 예시]",
    "- 예정",
    "- 완료",
    "- 취소",
    "- 예정 7",
    "- 완료 30",
    "- 취소 14",
    "- 예정 전체",
    "- 완료 2026-04-01~2026-04-30",
    "- init"
  ];
  return help.join("\n");
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function enqueueSlashCommandJob_(type, payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const queue = getSlashCommandQueue_();
    queue.push({
      id: Utilities.getUuid(),
      type: String(type || "").trim(),
      payload: payload || {},
      enqueuedAt: new Date().toISOString()
    });
    saveSlashCommandQueue_(queue);
  } finally {
    lock.releaseLock();
  }
  scheduleSlashCommandWorkerSoon_();
}

function processSlashCommandQueue() {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  let batch = [];
  try {
    const queue = getSlashCommandQueue_();
    if (!queue.length) return;
    batch = queue.slice(0, 5);
    saveSlashCommandQueue_(queue.slice(5));
  } finally {
    lock.releaseLock();
  }

  batch.forEach(job => {
    try {
      if (job.type === "pending" || job.type === "completed" || job.type === "cancelled") {
        const mode = job.type;
        const p = job.payload || {};
        const result = sendDailyReminders(mode, { period: p.period || { type: "default" } });
        notifySlashJobResult_(job, result);
      }
      else if (job.type === "status") {
        const p = job.payload || {};
        changeStatusByApprovalNo(p.approvalNo, p.targetStatus, p.actorName || p.requestedBy || "수동처리");
      }
    } catch (e) {
      Logger.log(`[SlashJob Failed] type=${job.type}, error=${e}`);
      notifySlashJobFailure_(job, e);
    }
  });
}

function processSlashCommandQueueOnce() {
  processSlashCommandQueue();
  cleanupSlashCommandTriggers_();
}

function scheduleSlashCommandWorkerSoon_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const exists = ScriptApp.getProjectTriggers()
      .some(t => t.getHandlerFunction() === "processSlashCommandQueueOnce");
    if (exists) return;
    ScriptApp.newTrigger("processSlashCommandQueueOnce")
      .timeBased()
      .after(1000)
      .create();
  } finally {
    lock.releaseLock();
  }
}

function cleanupSlashCommandTriggers_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "processSlashCommandQueueOnce") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function getSlashCommandQueue_() {
  const raw = SCRIPT_PROPS.getProperty("SLASH_COMMAND_QUEUE") || "[]";
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    return [];
  }
}

function saveSlashCommandQueue_(queue) {
  SCRIPT_PROPS.setProperty("SLASH_COMMAND_QUEUE", JSON.stringify(queue || []));
}

function notifySlashJobResult_(job, result) {
  const payload = job && job.payload ? job.payload : {};
  const responseUrl = String(payload.responseUrl || "").trim();
  if (!responseUrl) return;

  const modeLabel = result && result.mode === "completed" ? "완료" : (result && result.mode === "cancelled" ? "취소" : "예정");
  const lines = [
    `${modeLabel} 스레드 발송 결과`,
    `- 대상건수: ${Number(result && result.candidateCount || 0)}건`,
    `- 성공: ${Number(result && result.sentCount || 0)}건`,
    `- 실패: ${Number(result && result.failedCount || 0)}건`,
    `- 상태: ${String(result && result.reason || "unknown")}`
  ];
  if (result && result.quotaHit) lines.push(`- 비고: 상한/쿼터 도달로 일부 중단`);

  postToResponseUrl_(responseUrl, lines.join("\n"));
}

function notifySlashJobFailure_(job, error) {
  const payload = job && job.payload ? job.payload : {};
  const responseUrl = String(payload.responseUrl || "").trim();
  if (!responseUrl) return;
  postToResponseUrl_(responseUrl, `작업 실행 중 오류가 발생했습니다: ${String(error || "")}`);
}

function postToResponseUrl_(responseUrl, text) {
  try {
    UrlFetchApp.fetch(responseUrl, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({
        response_type: "ephemeral",
        text: String(text || "")
      }),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log(`[ResponseUrl Notify Failed] ${e}`);
  }
}

function parsePeriodToken_(rawPeriod) {
  const token = String(rawPeriod || "").trim();
  if (!token) return { ok: true, value: { type: "default" } };
  const normalized = token.toLowerCase();
  if (normalized === "전체" || normalized === "all") return { ok: true, value: { type: "all" } };
  // 숫자 단독 입력(예: 7, 30)은 최근 n일로 해석
  if (/^\d{1,3}$/.test(normalized)) {
    const days = Number(normalized);
    if (days <= 0) return { ok: false, error: "일수는 1 이상의 숫자로 입력하세요." };
    return { ok: true, value: { type: "last_days", days: days } };
  }
  // 숫자 + '일' 입력(예: 7일, 30일)도 지원
  const dayMatch = normalized.match(/^(\d{1,3})일$/);
  if (dayMatch) {
    const days = Number(dayMatch[1]);
    if (days <= 0) return { ok: false, error: "일수는 1 이상의 숫자로 입력하세요." };
    return { ok: true, value: { type: "last_days", days: days } };
  }
  if (normalized === "최근7일" || normalized === "7d") return { ok: true, value: { type: "last_days", days: 7 } };
  if (normalized === "최근30일" || normalized === "최근한달" || normalized === "30d") return { ok: true, value: { type: "last_days", days: 30 } };

  const rangeMatch = token.match(/^(\d{4}-\d{2}-\d{2})\s*~\s*(\d{4}-\d{2}-\d{2})$/);
  if (rangeMatch) {
    const from = new Date(rangeMatch[1]);
    const to = new Date(rangeMatch[2]);
    if (isNaN(from.getTime()) || isNaN(to.getTime())) return { ok: false, error: "기간 형식이 잘못되었습니다." };
    from.setHours(0, 0, 0, 0);
    to.setHours(0, 0, 0, 0);
    if (from > to) return { ok: false, error: "기간 시작일이 종료일보다 늦습니다." };
    return { ok: true, value: { type: "range", from: from.getTime(), to: to.getTime() } };
  }

  return { ok: false, error: "기간은 숫자(n일), 전체, YYYY-MM-DD~YYYY-MM-DD 형식으로 입력하세요." };
}

function getReferenceDateForMode_(row, mode, dueDate) {
  if (mode === "completed" && row[12]) {
    const d = new Date(row[12]);
    d.setHours(0, 0, 0, 0);
    return d;
  }
  if (mode === "cancelled" && row[18]) {
    const d = new Date(row[18]);
    d.setHours(0, 0, 0, 0);
    return d;
  }
  const base = new Date(dueDate);
  base.setHours(0, 0, 0, 0);
  return base;
}

function shouldIncludeByPeriod_(mode, diffDays, referenceDate, period, targetDays, today) {
  const p = period || { type: "default" };
  if (p.type === "default") {
    // 기본값 미입력 시 모든 모드 전체 조회
    return true;
  }
  if (p.type === "all") return true;
  if (p.type === "last_days") {
    const days = Number(p.days || 0);
    if (days <= 0) return false;
    if (mode === "pending") return diffDays >= 0 && diffDays <= days;
    const start = new Date(today);
    start.setDate(start.getDate() - days);
    return referenceDate >= start && referenceDate <= today;
  }
  if (p.type === "range") {
    const from = Number(p.from || 0);
    const to = Number(p.to || 0);
    if (!from || !to) return false;
    const t = referenceDate.getTime();
    return t >= from && t <= to;
  }
  return false;
}

function getReminderCandidateCount_(mode, period) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('원장DB');
  const lastRow = sheet.getLastRow();
  if (lastRow < 10) return 0;
  const data = sheet.getRange(1, 1, lastRow, 22).getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const targetDays = [10, 7, 5, 3, 1];
  let count = 0;

  for (let i = 9; i < data.length; i++) {
    const row = data[i];
    const status = String(row[11] || "").trim();
    if (mode === "pending") {
      if (['완료', '보류', '취소'].indexOf(status) >= 0) continue;
    } else if (mode === "completed") {
      if (status !== "완료") continue;
    } else if (mode === "cancelled") {
      if (status !== "취소") continue;
    } else continue;

    const dueDateRaw = row[6];
    if (!dueDateRaw) continue;
    const dueDate = new Date(dueDateRaw);
    dueDate.setHours(0, 0, 0, 0);
    const diffDays = Math.round((dueDate - today) / (1000 * 60 * 60 * 24));
    const referenceDate = getReferenceDateForMode_(row, mode, dueDate);
    if (shouldIncludeByPeriod_(mode, diffDays, referenceDate, period, targetDays, today)) count += 1;
  }

  return count;
}

/**
 * 4. 상태 업데이트 로직
 */
function handleStatusUpdate(ts, targetStatus, slackUserId, clickedUserName, displayIndex, maxApiAttempts, fastMode, fallbackApprovalNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('원장DB');
  const lastRow = sheet.getLastRow();
  if (lastRow < 10) return false;

  const data = sheet.getRange(1, 1, lastRow, 22).getValues();
  const searchTs = normalizeTs_(ts);
  const approvalNo = String(fallbackApprovalNo || "").trim();
  const userName = clickedUserName || getSlackUserName(slackUserId);
  const now = new Date();

  let targetRowIndex = -1;
  for (let i = 9; i < data.length; i++) {
    const rowTs = normalizeTs_(data[i][16]);
    if (rowTs && rowTs === searchTs) {
      targetRowIndex = i;
      break;
    }
  }

  if (targetRowIndex < 0 && approvalNo) {
    for (let i = 9; i < data.length; i++) {
      if (String(data[i][2] || "").trim() === approvalNo) {
        targetRowIndex = i;
        break;
      }
    }
  }

  if (targetRowIndex < 0) return false;

  const rowNum = targetRowIndex + 1;
  const currentStatus = String(data[targetRowIndex][11] || "").trim();

  if (targetStatus === "완료") {
    sheet.getRange(rowNum, 12).setValue("완료");
    sheet.getRange(rowNum, 13).setValue(now);
    sheet.getRange(rowNum, 14).setValue(userName);
  } else {
    sheet.getRange(rowNum, 12).setValue(targetStatus);
    sheet.getRange(rowNum, 18).setValue(userName);
    sheet.getRange(rowNum, 19).setValue(now);
    if (targetStatus === "예정" && currentStatus === "완료") {
      sheet.getRange(rowNum, 13).clearContent();
      sheet.getRange(rowNum, 14).clearContent();
    }
  }

  const updatedRow = sheet.getRange(rowNum, 1, 1, 22).getValues()[0];
  const normalizedIndex = displayIndex || deriveDisplayIndexFromRow_(targetRowIndex);
  const updatedBlocks = buildDetailBlocks(updatedRow, (targetStatus === "완료"), userName, normalizedIndex, targetStatus);
  updateSlackMessageBlocks(ts, updatedBlocks, maxApiAttempts || 2, !!fastMode);

  // 드물게 발생하는 "다른 목록 스레드의 상태 미동기화" 보완
  const approvalNoToSync = String(updatedRow[2] || "").trim();
  if (approvalNoToSync) {
    registerApprovalThreadTs_(approvalNoToSync, ts);
    syncApprovalRelatedThreads_(approvalNoToSync, ts, updatedBlocks, maxApiAttempts || 2);
  }

  // 상태 변경 시 기안자에게 개인 DM 안내
  notifyRequesterStatusChange_(updatedRow, targetStatus, userName, currentStatus);
  return true;
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
  if (!result.ok) {
    if (result.error === "bandwidth_quota_exceeded_apps_script") return null;
    throw new Error(`Slack post 실패: ${result.error || 'unknown_error'}`);
  }
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

function postToSlack(message, threadTs) {
  const payload = { channel: SLACK_CHANNEL_ID, text: message };
  if (threadTs) payload.thread_ts = threadTs;
  const result = slackApiCall("https://slack.com/api/chat.postMessage", payload);
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
      const statusLabel = buildSummaryStatusLabel_(mode, rows);
      lines.push(`- ${date} | ${statusLabel} | ${desc} | ${total.toLocaleString()}원 | 총 ${rows.length}건`);
    });
  });

  if (!lines.length) return "요약 데이터 없음";

  const separator = "--------------------------------";
  const totalLabel = mode === 'pending' ? "[지출 집행 예정 총계]" : (mode === "completed" ? "✅ [최근 완료 총계]" : "⛔ [최근 취소 총계]");
  return `${lines.join('\n')}\n\n${separator}\n${totalLabel} : ${Number(grandTotal || 0).toLocaleString()}원`;
}

function buildSummaryStatusLabel_(mode, rows) {
  if (mode === "completed") return "완료";
  if (mode === "cancelled") return "취소";
  if (!rows || !rows.length) return "➖ 상태없음";

  const minDiff = rows.reduce((minVal, item) => Math.min(minVal, Number(item.diffDays || 99999)), 99999);
  if (!isFinite(minDiff) || minDiff === 99999) return "D-?";
  if (minDiff >= 0) return `D-${minDiff}`;
  return `D+${Math.abs(minDiff)}`;
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

function buildDetailFallbackText_(row, index, statusLabel) {
  const amount = Number(row[4] || 0).toLocaleString();
  const itemTitle = String(row[5] || '내역 없음').trim();
  const target = String(row[7] || '-').trim();
  const approvalNo = String(row[2] || '-').trim();
  const dueInfo = getDueInfo(row[6]);
  const requester = String(row[20] || '-').trim();
  const bank = String(row[8] || '').trim();
  const account = String(row[9] || '').trim();
  const owner = String(row[10] || '').trim();
  const memo = String(row[15] || '-').trim();
  const prefix = index ? `${index}. ` : "";

  if (statusLabel === "완료" || statusLabel === "취소") {
    return `${prefix}\`${itemTitle}_${target}_${amount}원_${approvalNo}\`\n> ${statusLabel}`;
  }

  return `${prefix}*${itemTitle}* | *\`${approvalNo}\`*\n\`\`\`\n - 기안자명: ${requester}\n - 집행예정: ${dueInfo.dateLabel} | ${dueInfo.ddayLabel}\n - 집행대상: ${target}\n - 결제금액: ${amount}원\n - 입금계좌: ${bank} | ${account} | ${owner}\n - 담당메모: ${memo}\n\`\`\``;
}

function notifyRequesterStatusChange_(row, targetStatus, actorName, previousStatus) {
  try {
    const statusLabel = String(targetStatus || "").trim();
    const prevStatus = String(previousStatus || "").trim();
    // 최초 상태 대비 실변경이 있을 때만 DM 발송
    if (!statusLabel || prevStatus === statusLabel) return false;

    const requesterName = String(row[20] || "").trim();
    if (!requesterName) return false;

    // 1순위: 원장DB V열(슬랙 ID), 2순위: Script Properties 매핑
    const rowSlackUserId = sanitizeSlackUserId_(row[21]);
    const requesterUserId = rowSlackUserId || getRequesterSlackUserId_(requesterName);
    if (!requesterUserId) {
      Logger.log(`[Requester DM Skip] 매핑 없음: requester=${requesterName}, vCol=${String(row[21] || "").trim()}`);
      return false;
    }

    const itemTitle = String(row[5] || "내역 없음").trim();
    const target = String(row[7] || "-").trim();
    const approvalNo = String(row[2] || "-").trim();
    const amount = `${Number(row[4] || 0).toLocaleString()}원`;
    const dueInfo = getDueInfo(row[6]);
    const bank = String(row[8] || "").trim();
    const account = String(row[9] || "").trim();
    const owner = String(row[10] || "").trim();
    const memo = String(row[15] || "-").trim();
    const handler = extractDisplayName(actorName || "시스템");
    const nowLabel = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm");

    const dmText = [
      `*[지출결의 상태 변경 안내]*`,
      "",
      "```",
      "1. 처리정보",
      ` - 처리상태: ${prevStatus || "-"} > ${statusLabel}`,
      ` - 처리건명: ${itemTitle} | ${target}`,
      ` - 처리자명: ${handler}`,
      ` - 처리일시: ${nowLabel}`,
      "-----------------------------------------",
      "2. 결제정보",
      ` - 결재번호: ${approvalNo}`,
      ` - 결제금액: ${amount}`,
      ` - 입금정보: ${bank} | ${account} | ${owner}`,
      ` - 담당메모: ${memo}`,
      "```"
    ].join("\n");

    postToSlackDm_(requesterUserId, dmText);
    Logger.log(`[Requester DM Sent] requester=${requesterName}, userId=${requesterUserId}, status=${prevStatus}->${statusLabel}`);
    return true;
  } catch (err) {
    Logger.log(`[Requester DM Failed] error=${err}`);
    return false;
  }
}

function getRequesterSlackUserId_(requesterName) {
  const normalized = normalizeRequesterKey_(requesterName);
  if (!normalized) return "";

  // 예: {"홍길동":"U0123ABCD","김코덱":"U0456EFGH"}
  const raw = SCRIPT_PROPS.getProperty("REQUESTER_SLACK_MAP") || "{}";
  let map = {};
  try {
    map = JSON.parse(raw);
  } catch (e) {
    Logger.log(`[REQUESTER_SLACK_MAP Parse Failed] ${e}`);
    return "";
  }

  const direct = String(map[requesterName] || "").trim();
  if (direct) return direct;

  const keyMatched = Object.keys(map).find(name => normalizeRequesterKey_(name) === normalized);
  return keyMatched ? String(map[keyMatched] || "").trim() : "";
}

function normalizeRequesterKey_(name) {
  return String(name || "").trim().replace(/\s+/g, "").toLowerCase();
}

function sanitizeSlackUserId_(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const normalized = raw.replace(/[<@>]/g, "").toUpperCase();
  // Slack 사용자 ID는 보통 U... 또는 W... 형식을 사용
  return /^[UW][A-Z0-9]+$/.test(normalized) ? normalized : "";
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
    let response;
    try {
      response = UrlFetchApp.fetch(url, {
        method: "post",
        headers: { "Authorization": "Bearer " + SLACK_BOT_TOKEN },
        contentType: "application/json",
        muteHttpExceptions: true,
        payload: JSON.stringify(payload)
      });
    } catch (fetchErr) {
      const msg = String(fetchErr || "");
      if (msg.indexOf("Bandwidth quota exceeded") >= 0) {
        return { ok: false, error: "bandwidth_quota_exceeded_apps_script" };
      }
      if (i === attempts) return { ok: false, error: `fetch_exception:${msg}` };
      Utilities.sleep(waitMs);
      waitMs = Math.min(waitMs * 2, 12000);
      continue;
    }

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

function findRowByThreadTs_(sheet, threadTs) {
  if (!sheet || !threadTs) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 10) return null;

  const normalizedTarget = normalizeTs_(threadTs);
  const values = sheet.getRange(10, 17, lastRow - 9, 1).getValues(); // Q열(스레드 ts)
  for (let i = 0; i < values.length; i++) {
    const cellTs = normalizeTs_(values[i][0]);
    if (cellTs && cellTs === normalizedTarget) {
      return i + 10;
    }
  }
  return null;
}

function findRowByApprovalNo_(sheet, approvalNo) {
  const target = String(approvalNo || "").trim();
  if (!sheet || !target) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 10) return null;

  const values = sheet.getRange(10, 3, lastRow - 9, 1).getValues(); // C열 결재번호
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || "").trim() === target) {
      return i + 10;
    }
  }
  return null;
}

function deriveDisplayIndexFromRow_(_rowIndex) {
  // 신규 메시지는 버튼 value로 index를 유지하므로, 구형 메시지 fallback만 null 처리
  return null;
}

function normalizeTs_(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  return raw.replace(/^'+/, "");
}

function extractApprovalNoFromBlocks_(blocks) {
  try {
    if (!blocks || !blocks.length) return "";
    const first = blocks[0];
    const text = first && first.text ? String(first.text.text || "") : "";

    // 대기 포맷: "*1. 항목명* | *`PR2604...`*"
    let match = text.match(/`([A-Za-z0-9_-]+)`/);
    if (match && match[1]) return String(match[1]).trim();

    // 완료/취소 포맷: "*`1 항목_대상_금액_PR2604...`*"
    match = text.match(/_([A-Za-z0-9_-]+)`?\*/);
    if (match && match[1]) return String(match[1]).trim();
  } catch (e) {}
  return "";
}

function registerApprovalThreadTs_(approvalNo, threadTs) {
  const keyNo = String(approvalNo || "").trim();
  const keyTs = normalizeTs_(threadTs);
  if (!keyNo || !keyTs) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(3000);
  try {
    const raw = SCRIPT_PROPS.getProperty("APPROVAL_THREAD_TS_MAP") || "{}";
    let map = {};
    try { map = JSON.parse(raw); } catch (e) { map = {}; }

    const list = Array.isArray(map[keyNo]) ? map[keyNo] : [];
    const merged = [keyTs].concat(list.filter(v => normalizeTs_(v) !== keyTs)).slice(0, 10);
    map[keyNo] = merged;
    SCRIPT_PROPS.setProperty("APPROVAL_THREAD_TS_MAP", JSON.stringify(map));
  } finally {
    lock.releaseLock();
  }
}

function getApprovalThreadTsList_(approvalNo) {
  const keyNo = String(approvalNo || "").trim();
  if (!keyNo) return [];
  const raw = SCRIPT_PROPS.getProperty("APPROVAL_THREAD_TS_MAP") || "{}";
  let map = {};
  try { map = JSON.parse(raw); } catch (e) { map = {}; }
  const list = Array.isArray(map[keyNo]) ? map[keyNo] : [];
  return list.map(v => normalizeTs_(v)).filter(v => !!v);
}

function syncApprovalRelatedThreads_(approvalNo, currentTs, blocks, maxApiAttempts) {
  const current = normalizeTs_(currentTs);
  const list = getApprovalThreadTsList_(approvalNo).filter(ts => ts !== current);
  list.forEach(ts => {
    try {
      updateSlackMessageBlocks(ts, blocks, maxApiAttempts || 2, true);
      Utilities.sleep(80);
    } catch (e) {
      Logger.log(`[Sync Related Thread Failed] approvalNo=${approvalNo}, ts=${ts}, error=${e}`);
    }
  });
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

function clearSlackActionQueue() {
  SCRIPT_PROPS.setProperty("SLACK_ACTION_QUEUE", "[]");
  Logger.log("[Queue Cleared] SLACK_ACTION_QUEUE");
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
    const exists = ScriptApp.getProjectTriggers()
      .some(t => t.getHandlerFunction() === "processSlackActionQueueOnce");
    if (exists) return;

    // 단발 트리거 과생성 방지(짧게 유지)
    const cooldownMs = 3000;
    const lastScheduled = Number(SCRIPT_PROPS.getProperty("SLACK_QUEUE_ONESHOT_SCHEDULED_AT") || 0);
    if (lastScheduled && now - lastScheduled < cooldownMs) return;

    if (!exists) {
      ScriptApp.newTrigger("processSlackActionQueueOnce")
        .timeBased()
        .after(1000)
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

function buildQuickStatusBlocksFromPayload_(payload, targetStatus, clickedUserName, displayIndex) {
  const sectionText = ((payload || {}).message || {}).blocks && ((payload || {}).message || {}).blocks[0] && ((payload || {}).message || {}).blocks[0].text
    ? String(((payload || {}).message || {}).blocks[0].text.text || "")
    : "";
  const titleLine = extractTitleLineFromSectionText_(sectionText);
  const nowLabel = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd HH:mm");
  const name = extractDisplayName(clickedUserName || "시스템");
  const statusLabel = targetStatus === "완료" ? "✅ 완료" : "⛔ 취소";

  const mainText = `\n${titleLine}\n> ${statusLabel} | ${name} | ${nowLabel}`;
  const buttonValue = String(displayIndex || "");

  return [
    { type: "section", text: { type: "mrkdwn", text: mainText } },
    {
      type: "actions",
      elements: [
        { type: "button", text: { type: "plain_text", text: "완료" }, style: "primary", action_id: "status_done", value: buttonValue },
        { type: "button", text: { type: "plain_text", text: "취소" }, style: "danger", action_id: "status_cancel", value: buttonValue },
        { type: "button", text: { type: "plain_text", text: "대기전환" }, action_id: "status_pending", value: buttonValue }
      ]
    }
  ];
}

function extractTitleLineFromSectionText_(sectionText) {
  const lines = String(sectionText || "").split("\n").map(v => String(v || "").trim()).filter(v => !!v);
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.indexOf(">") === 0) continue;
    if (line.indexOf("```") === 0) continue;
    if (line.indexOf("- 집행") === 0) continue;
    return line;
  }
  return "항목 정보";
}

/**
 * 매일 오전 9시(스크립트 타임존 기준) 큐 상태 자동 리포트
 * - 같은 날짜에 중복 실행되더라도 1회만 발송
 * - 현재 운영에서는 비활성화됨(슬랙 발송 없음)
 */
function dailyHealthCheck() {
  // 운영 정책 변경: 9시 자동 DM은 dailyDigestJob으로 통합하고,
  // 큐 상태 리포트는 수동 점검 함수(inspectSlackActionQueueStatus)로만 사용.
  return { ok: true, disabled: true };
}

/**
 * 운영 트리거 일괄 설정
 * - 큐 워커(1분)
 * - 일일 상태 리포트(오전 9시)
 */
function setupOperationsTriggers() {
  setupSlackActionQueueTrigger();
  removeDailyHealthCheckTrigger_();
  setupDailyDigestTrigger();
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

/**
 * 하루 1회 DM 요약 발송 잡
 * - Script Properties의 DAILY_DM_USER_IDS(쉼표 구분 Slack User ID 목록) 대상으로 발송
 * - 중복 발송 방지 키: DAILY_DIGEST_SENT_DATE
 */
function dailyDigestJob() {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const todayKey = Utilities.formatDate(new Date(), "GMT+9", "yyyy-MM-dd");
    const sentKey = SCRIPT_PROPS.getProperty("DAILY_DIGEST_SENT_DATE");
    if (sentKey === todayKey) return;

    // /cdv 예정과 동일 로직으로 최신 데이터 기반 parent+thread 자동 발송
    const runResult = sendDailyReminders("pending", { period: { type: "default" } });
    const digestText = buildDailyDigestTextFromRunResult_(runResult);
    const recipients = getDailyDigestRecipients_();
    if (recipients.length) {
      recipients.forEach(userId => {
        try {
          postToSlackDm_(userId, digestText);
        } catch (e) {
          Logger.log(`[DailyDigest Failed] user=${userId}, error=${e}`);
        }
      });
    } else {
      // DAILY_DM_USER_IDS 미설정 시 기본 채널로 자동 발송
      postToSlack(digestText);
    }

    SCRIPT_PROPS.setProperty("DAILY_DIGEST_SENT_DATE", todayKey);
  } finally {
    lock.releaseLock();
  }
}

function setupDailyDigestTrigger() {
  const handler = "dailyDigestJob";
  const exists = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === handler);
  if (exists) return;

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .nearMinute(0)
    .create();
}

function removeDailyHealthCheckTrigger_() {
  const handler = "dailyHealthCheck";
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function buildDailyDigestText_(mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("원장DB");
  const lastRow = sheet.getLastRow();
  if (lastRow < 10) return "*[지출 집행 예정 요약]*\n```요약 데이터 없음```";

  const data = sheet.getRange(1, 1, lastRow, 22).getValues();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const targetDays = [10, 7, 5, 3, 1];

  let grandTotal = 0;
  const groups = {};

  for (let i = 9; i < data.length; i++) {
    const row = data[i];
    const status = String(row[11] || "").trim();
    if (mode === "pending") {
      if (["완료", "보류", "취소"].includes(status)) continue;
    } else if (status !== "완료") {
      continue;
    }

    const dueDateRaw = row[6];
    if (!dueDateRaw) continue;
    const dueDate = new Date(dueDateRaw);
    dueDate.setHours(0, 0, 0, 0);
    const diffDays = Math.round((dueDate - today) / (1000 * 60 * 60 * 24));

    if (mode === "completed" || targetDays.includes(diffDays)) {
      const dayLabels = ["일", "월", "화", "수", "목", "금", "토"];
      const dateStr = `${Utilities.formatDate(dueDate, "GMT+9", "yyyy-MM-dd")}(${dayLabels[dueDate.getDay()]})`;
      const desc = String(row[5] || "내역 없음").trim();
      if (!groups[dateStr]) groups[dateStr] = {};
      if (!groups[dateStr][desc]) groups[dateStr][desc] = [];
      groups[dateStr][desc].push({ data: row, diffDays: diffDays });
      grandTotal += Number(row[4] || 0);
    }
  }

  const lines = buildParentSummaryLines(
    Object.keys(groups).reduce((acc, date) => {
      acc[date] = { items: groups[date] };
      return acc;
    }, {}),
    mode,
    grandTotal
  );

  const title = mode === "pending" ? "*[지출 집행 예정 요약]*" : "*[최근 완료 내역]*";
  return `${title}\n\`\`\`\n${lines}\n\`\`\`\n※ 상세내역은 아래 스래드를 확인하세요`;
}

function buildDailyDigestTextFromRunResult_(result) {
  const modeLabel = result && result.mode === "completed" ? "완료" : (result && result.mode === "cancelled" ? "취소" : "예정");
  const candidateCount = Number(result && result.candidateCount || 0);
  const sentCount = Number(result && result.sentCount || 0);
  const failedCount = Number(result && result.failedCount || 0);
  const reason = String(result && result.reason || "unknown");

  if (candidateCount === 0) {
    return `*[일일 ${modeLabel} 요약]*\n\`\`\`\n오늘 기준 ${modeLabel} 대상 건이 없습니다.\n\`\`\``;
  }

  const lines = [
    `[일일 ${modeLabel} 자동 발송 완료]`,
    `- 대상건수: ${candidateCount}건`,
    `- 스레드 발송: ${sentCount}건`,
    `- 실패: ${failedCount}건`,
    `- 상태: ${reason}`,
    "",
    "상세 내역은 채널 스레드에서 확인해 주세요."
  ];
  return lines.join("\n");
}

function getDailyDigestRecipients_() {
  const raw = SCRIPT_PROPS.getProperty("DAILY_DM_USER_IDS") || "";
  return raw
    .split(",")
    .map(v => v.trim())
    .filter(v => !!v);
}

function postToSlackDm_(userId, text) {
  const openResult = slackApiCall("https://slack.com/api/conversations.open", { users: userId }, 3, { fastMode: true });
  if (!openResult.ok || !openResult.channel || !openResult.channel.id) {
    throw new Error(`DM 채널 오픈 실패: ${openResult.error || "unknown_error"}`);
  }
  const dmChannel = openResult.channel.id;
  const result = slackApiCall("https://slack.com/api/chat.postMessage", { channel: dmChannel, text: text }, 3, { fastMode: true });
  if (!result.ok) throw new Error(`DM 발송 실패: ${result.error || "unknown_error"}`);
  return result.ts;
}