/**
 * Kepty English 予約バックエンド
 *
 * 上流はスプレッドシート「calendar」。
 * 2行で1日。時刻の下のチェックがONなら、その時刻〜右隣の時刻の1時間が稼働。
 * 予約枠はその1時間を20分×3本に割る（例: 18時チェック → 18:00 / 18:20 / 18:40）。
 * 予約ローデータは同ブックの「data」「raw」シート A〜F 列へ書く
 * （A=LINE_UserID, B=date, C=time, D=remind message,
 *  E=lesson type, F=その行までの同一ユーザーのレッスン累計）。
 * 書き込み先は A 列が空の行の上から。
 * 予約を取り消した行の F は消す。同じ LINE userID の残った行は上から 1, 2, 3… と振り直し、
 * いちばん下の行に、その時点の受講回数を入れる。
 * data シートでは、レッスン終了後に I 列へコメントがあり J 列がチェック済みなら、
 * チェックから5分後にそのコメントを A 列の LINE userID へ送る。
 * チェック時刻は K 列、送信済みは L 列。チェックを外すと K・L は消える。
 * 送れなかったときは L 列のメモに LINE の応答を残し、次の実行で送り直す。
 * 既存のカレンダー予約を流し込むときは、エディタで backfillExistingBookingsToRaw を実行する。
 *
 * デプロイ:
 * 1. このコードを、カレンダーとスプシの両方に権限がある Google アカウントの Apps Script に貼る
 * 2. プロジェクトの設定 → スクリプト プロパティ
 *    - LINE_CHANNEL_ACCESS_TOKEN … Messaging API のチャネルアクセストークン
 *    - CALENDAR_ID … 省略時はデフォルトカレンダー
 *    - SPREADSHEET_ID … 省略時は下の定数
 * 3. デプロイ → 新しいデプロイ → ウェブアプリ
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員
 * 4. トリガー:
 *    - checkAndSendReminders … 10分おき
 *      （予約した本人へは前日21時のLINE案内のみ。1時間前は本人へ送らない）
 *      （STAFF_LINE_USER_IDS へは、前日21時と開始およそ1時間前も送る）
 *      （data シートのフィードバック送信も、この実行で確認する）
 *    - syncAvailabilityToCalendar … 1分おき（チェックの有無を空き枠へ反映。外した日も消す）
 *      （フィードバック送信も、カレンダー更新の前に確認する）
 */

var TZ = 'Asia/Tokyo';
var SLOT_MINUTES = 20;
var DURATION_MINUTES = 20;
var LEAD_MINUTES = 20;
var SPREADSHEET_ID = '1OLiHIs7HjtlSxE9j3ETGhlE8efjiayomMpmYb-376U0';
var SHEET_NAME = 'calendar';
var DATA_SHEET_NAME = 'data';
var RAW_SHEET_NAME = 'raw';
var DATA_HEADERS = ['LINE_UserID', 'date', 'time', 'remind message', 'lesson type', 'lesson count'];
var AVAILABILITY_TITLE = 'Online Lesson Booking Slot';
var AVAILABILITY_TAG = 'SHEET_AVAILABILITY:true';
var AVAILABILITY_COLOR = CalendarApp.EventColor.PALE_BLUE;
var BOOKING_TITLE_PREFIX = 'Lesson Booking';
var STAFF_LINE_USER_IDS = [
  'Ud0fefb32063023e68af8a6be50201e19',
  'U9de56864ced4187e6270d26ece292bcd'
];
var LESSON_TYPES = {
  guided: '専属講師が主導する授業形式のレッスン',
  freetalk: '専属講師と自由に会話するフリートークレッスン'
};
var STAFF_LESSON_LABELS = {
  guided: '[Lecture-style Lesson]',
  freetalk: '[Free Talk Lesson]'
};
var STAFF_MONTHS = ['Jan.', 'Feb.', 'Mar.', 'Apr.', 'May', 'Jun.', 'Jul.', 'Aug.', 'Sep.', 'Oct.', 'Nov.', 'Dec.'];
var ZOOM_URL = 'https://us05web.zoom.us/j/82772210913?pwd=ZuUZhN09KfLOz2oyZuafZZEH5GdaJI.1';
var WEEKDAY_JA = ['日', '月', '火', '水', '木', '金', '土'];
var FB_COMMENT_COL = 9;
var FB_CHECK_COL = 10;
var FB_CHECKED_AT_COL = 11;
var FB_SENT_COL = 12;
var FB_DELAY_MS = 5 * 60 * 1000;

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'slots';
    if (action === 'bookings') {
      var userId = String((e.parameter && e.parameter.userId) || '').trim();
      if (!userId || userId === 'GUEST') return json_({ ok: true, bookings: [] });
      return json_({ ok: true, bookings: listBookings_(userId) });
    }
    var date = (e && e.parameter && e.parameter.date) || tokyoDateString_(new Date());
    return json_({ ok: true, slots: buildSlots_(date) });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var data = JSON.parse((e.postData && e.postData.contents) || '{}');
    var action = String(data.action || 'book').trim();
    var userId = String(data.userId || '').trim();
    var userName = String(data.userName || 'Guest').trim();

    if (!userId || userId === 'GUEST') {
      return json_({ success: false, message: '公式LINEの予約メニューから開いてください。' });
    }

    if (action === 'cancel') {
      return json_(cancelBooking_(userId, String(data.eventId || '').trim()));
    }

    var date = String(data.date || '').trim();
    var time = String(data.time || '').trim();
    var lessonType = String(data.lessonType || '').trim();
    if (!date || !time) {
      return json_({ success: false, message: '日時を選択してください。' });
    }
    if (!LESSON_TYPES[lessonType]) {
      return json_({ success: false, message: 'レッスン形式を選択してください。' });
    }

    var slot = findSlot_(date, time);
    if (!slot || !slot.available) {
      return json_({ success: false, message: 'その枠はすでに埋まりました。別の日時を選んでください。' });
    }

    var start = slotStart_(date, time);
    var end = new Date(start.getTime() + DURATION_MINUTES * 60 * 1000);
    var calendar = getCalendar_();
    var event = calendar.createEvent(bookingEventTitle_(userName), start, end, {
      description: bookingDescription_(userId, userName, lessonType),
      location: ZOOM_URL
    });
    event.setColor(CalendarApp.EventColor.ORANGE);

    var confirmText = bookingLineMessage_(start);
    sendLine_(userId, confirmText);
    appendBookingRaw_(userId, date, time, eveLineMessage_(start), lessonType);
    var lessonNo = String(lessonNoFromRaw_(userId, date, time));
    setDescFlag_(event, 'LESSON_NO', lessonNo);
    sendLineToStaff_(staffBookingMessage_(calendarStudentName_(userName), lessonNo, start, lessonType));

    return json_({
      success: true,
      message: confirmText,
      eventId: event.getId()
    });
  } catch (err) {
    return json_({ success: false, message: '予約処理でエラーが発生しました。' });
  } finally {
    lock.releaseLock();
  }
}

function sendReminders() {
  sendEveReminders_();
  sendStaffHourReminders_();
}

function checkAndSendReminders() {
  sendFbMessages_();
  sendEveReminders_();
  sendStaffHourReminders_();
}

function sendEveReminders_() {
  var now = new Date();
  if (Number(Utilities.formatDate(now, TZ, 'H')) !== 21) return;

  var tomorrow = tokyoPlusDays_(1);
  var events = getCalendar_().getEvents(parseTokyoDate_(tomorrow), endOfDay_(parseTokyoDate_(tomorrow)));

  events.forEach(function (event) {
    if (!isOurBooking_(event)) return;
    var desc = event.getDescription() || '';
    var userId = valueOf_(desc, 'LINE_USER_ID');
    var sent = valueOf_(desc, 'PREV_DAY_REMINDER_SENT');
    if (!userId) return;
    if (sent !== 'true') {
      sendLine_(userId, eveLineMessage_(event.getStartTime()));
      if (desc.indexOf('PREV_DAY_REMINDER_SENT:false') !== -1) {
        event.setDescription(desc.replace('PREV_DAY_REMINDER_SENT:false', 'PREV_DAY_REMINDER_SENT:true'));
      } else if (desc.indexOf('PREV_DAY_REMINDER_SENT:') === -1) {
        event.setDescription(desc + '\nPREV_DAY_REMINDER_SENT:true');
      }
      desc = event.getDescription() || '';
    }
    notifyStaffEve_(event, desc);
  });
}

function notifyStaffEve_(event, desc) {
  if (valueOf_(desc, 'STAFF_EVE_SENT') === 'true') return;
  var userId = valueOf_(desc, 'LINE_USER_ID');
  if (!userId) return;
  var start = event.getStartTime();
  var name = calendarStudentName_(valueOf_(desc, 'LINE_DISPLAY_NAME'));
  var lessonNo = lessonNumber_(event);
  sendLineToStaff_(staffEveMessage_(name, lessonNo, start, valueOf_(desc, 'LESSON_TYPE')));
  setDescFlag_(event, 'STAFF_EVE_SENT', 'true');
}

function sendStaffHourReminders_() {
  var now = new Date();
  // getEvents の終端は含まないので、開始ちょうど60分前の予定も拾う。
  var horizon = new Date(now.getTime() + 61 * 60 * 1000);
  getCalendar_().getEvents(now, horizon).forEach(function (event) {
    if (!isOurBooking_(event)) return;
    var start = event.getStartTime();
    var remaining = start.getTime() - now.getTime();
    if (remaining <= 0 || remaining > 60 * 60 * 1000) return;
    var desc = event.getDescription() || '';
    if (valueOf_(desc, 'STAFF_HOUR_SENT') === 'true') return;
    if (!valueOf_(desc, 'LINE_USER_ID')) return;
    var name = calendarStudentName_(valueOf_(desc, 'LINE_DISPLAY_NAME'));
    var lessonNo = lessonNumber_(event);
    sendLineToStaff_(staffHourMessage_(name, lessonNo, start, valueOf_(desc, 'LESSON_TYPE')));
    setDescFlag_(event, 'STAFF_HOUR_SENT', 'true');
  });
}

function buildSlots_(date) {
  if (date < tokyoDateString_(new Date())) return [];

  var windows = windowsForDate_(date);
  if (!windows.length) return [];

  var busy = [];
  getCalendar_().getEvents(parseTokyoDate_(date), endOfDay_(parseTokyoDate_(date))).forEach(function (event) {
    if (isSheetAvailability_(event)) return;
    if (isOurBooking_(event)) {
      busy.push({ start: event.getStartTime(), end: event.getEndTime() });
    }
  });

  return slotsFromWindows_(windows, busy);
}

function windowsForDate_(date) {
  var days = loadAvailabilityDays_();
  return days[date] || [];
}

function loadAvailabilityDays_() {
  var values = getCalendarSheet_().getDataRange().getValues();
  var days = {};

  for (var i = 0; i < values.length - 1; i += 2) {
    var dateStr = sheetDateToYmd_(values[i][0]);
    if (!dateStr) continue;

    var timeRow = values[i];
    var checkRow = values[i + 1] || [];
    var windows = [];

    for (var c = 1; c < timeRow.length; c++) {
      if (!isChecked_(checkRow[c])) continue;
      var startHm = sheetTimeToHm_(timeRow[c]);
      var endHm = sheetTimeToHm_(timeRow[c + 1]);
      if (!startHm || !endHm) continue;
      windows.push({
        start: slotStart_(dateStr, startHm),
        end: slotStart_(dateStr, endHm)
      });
    }

    days[dateStr] = mergeWindows_(windows);
  }

  return days;
}

function mergeWindows_(windows) {
  if (!windows.length) return [];
  var sorted = windows.slice().sort(function (a, b) {
    return a.start.getTime() - b.start.getTime();
  });
  var merged = [{ start: sorted[0].start, end: sorted[0].end }];
  for (var i = 1; i < sorted.length; i++) {
    var last = merged[merged.length - 1];
    if (sorted[i].start.getTime() <= last.end.getTime() + 1000) {
      if (sorted[i].end.getTime() > last.end.getTime()) last.end = sorted[i].end;
    } else {
      merged.push({ start: sorted[i].start, end: sorted[i].end });
    }
  }
  return merged;
}

function slotsFromWindows_(windows, busy) {
  var now = new Date();
  var latest = now.getTime() + LEAD_MINUTES * 60 * 1000;
  var slots = [];
  var seen = {};
  var durationMs = DURATION_MINUTES * 60 * 1000;
  var stepMs = SLOT_MINUTES * 60 * 1000;

  windows.forEach(function (win) {
    var cursor = win.start;
    var endMs = win.end.getTime();
    while (cursor.getTime() + durationMs <= endMs + 1000) {
      var time = Utilities.formatDate(cursor, TZ, 'HH:mm');
      if (!seen[time]) {
        seen[time] = true;
        var slotEnd = new Date(cursor.getTime() + durationMs);
        if (cursor.getTime() >= latest && !overlapsBusy_(cursor, slotEnd, busy)) {
          slots.push({ time: time, available: true });
        }
      }
      cursor = new Date(cursor.getTime() + stepMs);
    }
  });

  slots.sort(function (a, b) {
    return a.time < b.time ? -1 : 1;
  });
  return slots;
}

function syncAvailabilityToCalendar() {
  sendFbMessages_();
  runAvailabilitySync_();
}

function runAvailabilitySync_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    var days = loadAvailabilityDays_();
    var today = tokyoDateString_(new Date());
    var desired = {};
    // チェックが無い日も読み取る。最後のチェック日で範囲を切ると、その後の空き枠が残る。
    var latestDate = tokyoPlusDays_(180);

    Object.keys(days).forEach(function (date) {
      if (date < today) return;
      if (date > latestDate) latestDate = date;
      (days[date] || []).forEach(function (win) {
        var key = date + '|' + Utilities.formatDate(win.start, TZ, 'HH:mm') + '|' + Utilities.formatDate(win.end, TZ, 'HH:mm');
        desired[key] = win;
      });
    });

    var calendar = getCalendar_();
    var rangeStart = parseTokyoDate_(today);
    var rangeEnd = endOfDay_(parseTokyoDate_(latestDate));
    var existing = calendar.getEvents(rangeStart, rangeEnd);
    var kept = {};

    existing.forEach(function (event) {
      if (!isSheetAvailability_(event)) return;
      var key = tokyoDateString_(event.getStartTime()) + '|' +
        Utilities.formatDate(event.getStartTime(), TZ, 'HH:mm') + '|' +
        Utilities.formatDate(event.getEndTime(), TZ, 'HH:mm');
      if (desired[key]) {
        kept[key] = true;
        event.setColor(AVAILABILITY_COLOR);
      } else {
        event.deleteEvent();
      }
    });

    Object.keys(desired).forEach(function (key) {
      if (kept[key]) return;
      var win = desired[key];
      var event = calendar.createEvent(AVAILABILITY_TITLE, win.start, win.end, {
        description: AVAILABILITY_TAG
      });
      event.setColor(AVAILABILITY_COLOR);
    });
  } finally {
    lock.releaseLock();
  }
}

function onEdit(e) {
  try {
    if (!e || !e.range) {
      runAvailabilitySync_();
      return;
    }
    recordFbCheckTime_(e);
    var sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAME) return;
    runAvailabilitySync_();
  } catch (err) {}
}

function recordFbCheckTime_(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  if (sheet.getName() !== DATA_SHEET_NAME) return;
  var startCol = e.range.getColumn();
  var endCol = startCol + e.range.getNumColumns() - 1;
  if (startCol > FB_CHECK_COL || endCol < FB_CHECK_COL) return;

  var values = e.range.getValues();
  var offset = FB_CHECK_COL - startCol;
  var startRow = e.range.getRow();
  var now = new Date();
  for (var r = 0; r < values.length; r++) {
    var row = startRow + r;
    if (row < 2) continue;
    if (isChecked_(values[r][offset])) {
      if (!sheet.getRange(row, FB_CHECKED_AT_COL).getValue()) {
        sheet.getRange(row, FB_CHECKED_AT_COL).setNumberFormat('yyyy-mm-dd hh:mm:ss').setValue(now);
      }
    } else {
      sheet.getRange(row, FB_CHECKED_AT_COL, 1, 2).clearContent();
    }
  }
}

function sendFbMessages_() {
  var lock = null;
  try {
    lock = LockService.getDocumentLock();
    if (!lock.tryLock(5000)) return;
  } catch (err) {
    lock = null;
  }
  try {
    sendFbRows_();
  } catch (err) {
    Logger.log(err);
  } finally {
    if (lock) lock.releaseLock();
  }
}

function sendFbRows_() {
  var sheet = getDataSheet_();
  ensureFbHeaders_(sheet);
  var last = sheet.getLastRow();
  if (last < 2) return;

  var values = sheet.getRange(2, 1, last - 1, FB_SENT_COL).getValues();
  var now = new Date();
  for (var i = 0; i < values.length; i++) {
    try {
      sendFbRow_(sheet, values[i], i + 2, now);
    } catch (err) {
      sheet.getRange(i + 2, FB_SENT_COL).setNote(String(err).slice(0, 200));
    }
  }
}

function sendFbRow_(sheet, rowValues, row, now) {
  var checked = isChecked_(rowValues[FB_CHECK_COL - 1]);
  var checkedAt = sheetDateTime_(rowValues[FB_CHECKED_AT_COL - 1]);
  var sent = isFbSent_(rowValues[FB_SENT_COL - 1]);
  if (!checked) {
    if (checkedAt || sent || rowValues[FB_SENT_COL - 1] !== '') {
      sheet.getRange(row, FB_CHECKED_AT_COL, 1, 2).clearContent();
      sheet.getRange(row, FB_SENT_COL).clearNote();
    }
    return;
  }

  var userId = String(rowValues[0] != null ? rowValues[0] : '').trim();
  var comment = String(rowValues[FB_COMMENT_COL - 1] != null ? rowValues[FB_COMMENT_COL - 1] : '').trim();
  if (!userId || !comment || sent) return;
  if (!checkedAt) {
    sheet.getRange(row, FB_CHECKED_AT_COL).setNumberFormat('yyyy-mm-dd hh:mm:ss').setValue(now);
    return;
  }

  var date = sheetWall_(rowValues[1], 'yyyy-MM-dd') || sheetDateToYmd_(rowValues[1]);
  var time = sheetWall_(rowValues[2], 'HH:mm') || sheetTimeToHm_(rowValues[2]);
  var sentCell = sheet.getRange(row, FB_SENT_COL);
  if (!date || !time) {
    sentCell.setNote('日付か時刻を読めないため送れません');
    return;
  }
  var end = slotStart_(date, time).getTime() + DURATION_MINUTES * 60 * 1000;
  if (isNaN(end)) {
    sentCell.setNote('日付か時刻を読めないため送れません');
    return;
  }
  if (now.getTime() < end) return;
  if (now.getTime() < checkedAt.getTime() + FB_DELAY_MS) return;

  var text = comment.length > 5000 ? comment.substring(0, 5000) : comment;
  var result = pushLine_(userId, text);
  if (!result.ok) {
    var detail = 'LINE ' + result.code;
    if (result.body) detail += ' ' + String(result.body).replace(/\s+/g, ' ').slice(0, 160);
    sentCell.setNote(detail);
    return;
  }
  sentCell.clearNote();
  sentCell.setValue(true);
}

function ensureFbHeaders_(sheet) {
  var headers = sheet.getRange(1, FB_CHECKED_AT_COL, 1, 2).getValues()[0];
  var next = [headers[0], headers[1]];
  var changed = false;
  if (!String(next[0] != null ? next[0] : '').trim()) {
    next[0] = 'FB checked at';
    changed = true;
  }
  if (!String(next[1] != null ? next[1] : '').trim()) {
    next[1] = 'FB sent';
    changed = true;
  }
  if (changed) sheet.getRange(1, FB_CHECKED_AT_COL, 1, 2).setValues([next]);
}

function isFbSent_(value) {
  return isChecked_(value) || String(value != null ? value : '').trim().toLowerCase() === 'sent';
}

function sheetDateTime_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return tokyoWallTime_(value);
  }
  if (value === '' || value === null || value === undefined) return null;
  var text = String(value).trim();
  var dmy = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if (dmy) {
    return new Date(
      dmy[3] + '-' + pad2_(Number(dmy[2])) + '-' + pad2_(Number(dmy[1])) + 'T' +
      pad2_(dmy[4] ? Number(dmy[4]) : 0) + ':' + pad2_(dmy[5] ? Number(dmy[5]) : 0) + ':' +
      pad2_(dmy[6] ? Number(dmy[6]) : 0) + '+09:00'
    );
  }
  var parsed = new Date(text);
  if (isNaN(parsed.getTime())) return null;
  return parsed;
}

function sheetWall_(value, pattern) {
  if (Object.prototype.toString.call(value) !== '[object Date]' || isNaN(value.getTime())) return '';
  var zone = Session.getScriptTimeZone() || TZ;
  return Utilities.formatDate(value, zone, pattern);
}

function tokyoWallTime_(value) {
  var wall = sheetWall_(value, "yyyy-MM-dd'T'HH:mm:ss");
  if (!wall) return value;
  var instant = new Date(wall + '+09:00');
  if (isNaN(instant.getTime())) return value;
  return instant;
}

function listBookings_(userId) {
    var now = new Date();
    var from = new Date(now.getTime() - DURATION_MINUTES * 60 * 1000);
    var to = parseTokyoDate_(tokyoPlusDays_(120));
    var bookings = [];

    getCalendar_().getEvents(from, to).forEach(function (event) {
      if (!isOurBooking_(event)) return;
      if (event.getEndTime().getTime() <= now.getTime()) return;
      var desc = event.getDescription() || '';
      if (valueOf_(desc, 'LINE_USER_ID') !== userId) return;
      bookings.push({
        eventId: event.getId(),
        when: formatLessonWhen_(event.getStartTime()),
        date: tokyoDateString_(event.getStartTime()),
        time: Utilities.formatDate(event.getStartTime(), TZ, 'HH:mm'),
        startMs: event.getStartTime().getTime(),
        endMs: event.getEndTime().getTime(),
        zoom: ZOOM_URL
      });
    });

  bookings.sort(function (a, b) {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    return a.time < b.time ? -1 : 1;
  });
  return bookings;
}

function cancelBooking_(userId, eventId) {
  if (!eventId) return { success: false, message: '予約が見つかりません。' };

  var event = findOwnedBooking_(userId, eventId);
  if (!event) return { success: false, message: '予約が見つかりません。' };
  if (event.getStartTime().getTime() < new Date().getTime()) {
    return { success: false, message: '開始済みの予約は取り消せません。' };
  }

  var desc = event.getDescription() || '';
  var start = event.getStartTime();
  var when = formatLessonWhen_(start);
  var date = tokyoDateString_(start);
  var time = Utilities.formatDate(start, TZ, 'HH:mm');
  var name = calendarStudentName_(valueOf_(desc, 'LINE_DISPLAY_NAME'));
  var lessonNo = lessonNumber_(event);
  var lessonType = valueOf_(desc, 'LESSON_TYPE');
  event.deleteEvent();
  removeBookingRaw_(userId, date, time);
  sendLine_(userId, cancelLineMessage_(when));
  sendLineToStaff_(staffCancelMessage_(name, lessonNo, start, lessonType), userId);
  return { success: true, message: when + ' の予約を取り消しました。' };
}

function findOwnedBooking_(userId, eventId) {
  try {
    var event = getCalendar_().getEventById(eventId);
    if (event && valueOf_(event.getDescription() || '', 'LINE_USER_ID') === userId) return event;
  } catch (err) {}

  var bookings = listBookings_(userId);
  for (var i = 0; i < bookings.length; i++) {
    if (bookings[i].eventId === eventId) {
      try {
        var found = getCalendar_().getEventById(bookings[i].eventId);
        if (found) return found;
      } catch (err2) {}
    }
  }
  return null;
}

function cancelLineMessage_(when) {
  return [
    'ご予約を取り消しました。',
    when,
    '',
    'またのご予約をお待ちしています。'
  ].join('\n');
}

function findSlot_(date, time) {
  var slots = buildSlots_(date);
  for (var i = 0; i < slots.length; i++) {
    if (slots[i].time === time) return slots[i];
  }
  return null;
}

function getSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || SPREADSHEET_ID;
  return SpreadsheetApp.openById(id);
}

function getCalendarSheet_() {
  var sheet = getSpreadsheet_().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('calendarシートが見つかりません');
  return sheet;
}

function getOrCreateSheet_(name) {
  var spreadsheet = getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    sheet.getRange(1, 1, 1, DATA_HEADERS.length).setValues([DATA_HEADERS]);
  }
  return sheet;
}

function getDataSheet_() {
  return getOrCreateSheet_(DATA_SHEET_NAME);
}

function getRawSheet_() {
  return getOrCreateSheet_(RAW_SHEET_NAME);
}

function appendBookingRaw_(userId, date, time, remindMessage, lessonType) {
  var row = [userId, date, time, remindMessage || '', lessonTypeSheetLabel_(lessonType)];
  appendBookingRawRows_(getDataSheet_(), [row]);
  appendBookingRawRows_(getRawSheet_(), [row]);
}

function lessonTypeSheetLabel_(lessonType) {
  if (lessonType === 'freetalk') return 'Free talk lesson';
  if (lessonType === 'guided') return 'lecture style lesson';
  return '';
}

/**
 * エディタから手動実行する。
 * Googleカレンダーの今後のレッスン予約を raw（なければ作成）と data の A〜F へ流し込む。
 * 既に同じ LINE_UserID + date + time がある行は重複追加しない。
 */
function backfillExistingBookingsToRaw() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var rows = collectUpcomingBookingRawRows_();
    var added = appendBookingRawRows_(getRawSheet_(), rows);
    appendBookingRawRows_(getDataSheet_(), rows);
    Logger.log('backfillExistingBookingsToRaw: added ' + added + ' row(s)');
    return { ok: true, added: added };
  } finally {
    lock.releaseLock();
  }
}

function collectUpcomingBookingRawRows_() {
  var now = new Date();
  var from = new Date(now.getTime() - DURATION_MINUTES * 60 * 1000);
  var to = parseTokyoDate_(tokyoPlusDays_(120));
  var rows = [];
  var seen = {};

  getCalendar_().getEvents(from, to).forEach(function (event) {
    if (!isOurBooking_(event)) return;
    if (isSheetAvailability_(event)) return;
    if (event.getEndTime().getTime() <= now.getTime()) return;
    var start = event.getStartTime();
    var desc = event.getDescription() || '';
    var userId = valueOf_(desc, 'LINE_USER_ID');
    if (!userId) return;
    var date = tokyoDateString_(start);
    var time = Utilities.formatDate(start, TZ, 'HH:mm');
    var key = bookingRawKey_(userId, date, time);
    if (!key || seen[key]) return;
    seen[key] = true;
    rows.push([userId, date, time, eveLineMessage_(start), lessonTypeSheetLabel_(valueOf_(desc, 'LESSON_TYPE'))]);
  });

  return rows;
}

function existingBookingRawKeys_(sheet) {
  var existing = {};
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var key = bookingRawKey_(
      String(values[i][0] != null ? values[i][0] : '').trim(),
      sheetDateToYmd_(values[i][1]),
      sheetTimeToHm_(values[i][2])
    );
    if (key) existing[key] = true;
  }
  return existing;
}

function appendBookingRawRows_(sheet, rows) {
  var existing = existingBookingRawKeys_(sheet);
  var fresh = [];
  for (var i = 0; i < rows.length; i++) {
    var key = bookingRawKey_(rows[i][0], rows[i][1], rows[i][2]);
    if (!key || existing[key]) continue;
    existing[key] = true;
    fresh.push(rows[i]);
  }
  if (!fresh.length) return 0;

  var targets = rowsWithEmptyColumnA_(sheet, fresh.length);
  var counts = lessonCountsThroughRows_(sheet, fresh, targets);
  var runStart = 0;
  for (var j = 1; j <= targets.length; j++) {
    if (j < targets.length && targets[j] === targets[j - 1] + 1) continue;
    var run = [];
    for (var k = runStart; k < j; k++) {
      run.push([
        fresh[k][0],
        fresh[k][1],
        fresh[k][2],
        fresh[k][3] || '',
        fresh[k][4] || '',
        counts[k]
      ]);
    }
    sheet.getRange(targets[runStart], 1, run.length, 6).setValues(run);
    if (sheet.getName() === DATA_SHEET_NAME) {
      sheet.getRange(targets[runStart], FB_COMMENT_COL, run.length, 4).clearContent();
    }
    runStart = j;
  }
  var seenUsers = {};
  for (var u = 0; u < fresh.length; u++) {
    var id = String(fresh[u][0] != null ? fresh[u][0] : '').trim();
    if (!id || seenUsers[id]) continue;
    seenUsers[id] = true;
    renumberLessonCounts_(sheet, id);
  }
  return fresh.length;
}

function lessonCountsThroughRows_(sheet, fresh, targets) {
  var last = Math.max(sheet.getLastRow(), 1);
  var colA = last <= 1 ? [] : sheet.getRange(2, 1, last - 1, 1).getValues();
  var counts = [];
  for (var i = 0; i < fresh.length; i++) {
    var userId = String(fresh[i][0] != null ? fresh[i][0] : '').trim();
    var through = targets[i] - 1;
    var total = 0;
    var limit = Math.min(Math.max(through - 1, 0), colA.length);
    for (var r = 0; r < limit; r++) {
      if (String(colA[r][0] != null ? colA[r][0] : '').trim() === userId) total++;
    }
    for (var p = 0; p < i; p++) {
      if (targets[p] <= through && String(fresh[p][0] != null ? fresh[p][0] : '').trim() === userId) total++;
    }
    counts.push(total + 1);
  }
  return counts;
}

function rowsWithEmptyColumnA_(sheet, count) {
  var last = Math.max(sheet.getLastRow(), 1);
  var colA = sheet.getRange(2, 1, Math.max(last - 1, 1), 1).getValues();
  var rows = [];
  var limit = last <= 1 ? 0 : colA.length;
  for (var i = 0; i < limit && rows.length < count; i++) {
    if (String(colA[i][0] != null ? colA[i][0] : '').trim() === '') rows.push(i + 2);
  }
  var next = last + 1;
  while (rows.length < count) {
    rows.push(next);
    next++;
  }
  return rows;
}

function renumberLessonCounts_(sheet, userId) {
  userId = String(userId || '').trim();
  if (!userId) return;
  var last = sheet.getLastRow();
  if (last < 2) return;
  var values = sheet.getRange(2, 1, last - 1, 6).getValues();
  var count = 0;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] != null ? values[i][0] : '').trim() !== userId) continue;
    count++;
    if (Number(values[i][5]) === count) continue;
    sheet.getRange(i + 2, 6).setValue(count);
  }
}

function bookingRawKey_(userId, date, time) {
  if (!userId || !date || !time) return '';
  return userId + '|' + date + '|' + time;
}

function removeBookingRaw_(userId, date, time) {
  removeBookingRawFromSheet_(getDataSheet_(), userId, date, time);
  removeBookingRawFromSheet_(getRawSheet_(), userId, date, time);
}

function removeBookingRawFromSheet_(sheet, userId, date, time) {
  var values = sheet.getDataRange().getValues();
  for (var i = values.length - 1; i >= 1; i--) {
    var rowUserId = String(values[i][0] != null ? values[i][0] : '').trim();
    var rowDate = sheetDateToYmd_(values[i][1]);
    var rowTime = sheetTimeToHm_(values[i][2]);
    if (rowUserId === userId && rowDate === date && rowTime === time) {
      sheet.getRange(i + 1, 1, 1, 6).clearContent();
      if (sheet.getName() === DATA_SHEET_NAME) {
        sheet.getRange(i + 1, FB_COMMENT_COL, 1, 4).clearContent();
      }
      renumberLessonCounts_(sheet, userId);
      return;
    }
  }
}

function getCalendar_() {
  var id = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID');
  var calendar = id ? CalendarApp.getCalendarById(id) : CalendarApp.getDefaultCalendar();
  if (!calendar) throw new Error('カレンダーが見つかりません');
  return calendar;
}

function isChecked_(value) {
  return value === true || value === 'TRUE' || value === 'true' || value === 1 || value === '1';
}

function isSheetAvailability_(event) {
  var desc = event.getDescription() || '';
  return desc.indexOf(AVAILABILITY_TAG) !== -1;
}

function isOurBooking_(event) {
  var title = String(event.getTitle() || '');
  var desc = event.getDescription() || '';
  return title.indexOf(BOOKING_TITLE_PREFIX) !== -1 ||
    title.indexOf('レッスン予約') !== -1 ||
    desc.indexOf('LINE_USER_ID:') !== -1;
}

function overlapsBusy_(start, end, busy) {
  return busy.some(function (block) {
    return block.start.getTime() < end.getTime() && block.end.getTime() > start.getTime();
  });
}

function sheetDateToYmd_(value) {
  if (value === '' || value === null || value === undefined) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }
  var text = String(value).trim();
  var dmy = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (dmy) return dmy[3] + '-' + pad2_(Number(dmy[2])) + '-' + pad2_(Number(dmy[1]));
  var ymd = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (ymd) return ymd[1] + '-' + ymd[2] + '-' + ymd[3];
  return '';
}

function sheetTimeToHm_(value) {
  if (value === '' || value === null || value === undefined) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'HH:mm');
  }
  var text = String(value).trim();
  var match = text.match(/^(\d{1,2}):(\d{2})/);
  if (!match) return '';
  return pad2_(Number(match[1])) + ':' + match[2];
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

function sendLineToStaff_(text, excludeUserId) {
  STAFF_LINE_USER_IDS.forEach(function (id) {
    if (excludeUserId && id === excludeUserId) return;
    sendLine_(id, text);
  });
}

function lessonNoFromRaw_(userId, date, time) {
  var values = getRawSheet_().getDataRange().getValues();
  var index = 0;
  var found = 0;
  for (var i = 1; i < values.length; i++) {
    var rowUserId = String(values[i][0] != null ? values[i][0] : '').trim();
    if (rowUserId !== userId) continue;
    index++;
    var rowDate = sheetDateToYmd_(values[i][1]);
    var rowTime = sheetTimeToHm_(values[i][2]);
    if (rowDate === date && rowTime === time) found = index;
  }
  if (found) return found;
  return index + 1;
}

function lessonNumber_(event) {
  var desc = event.getDescription() || '';
  var stored = valueOf_(desc, 'LESSON_NO');
  if (stored && stored !== '0') return stored;
  var userId = valueOf_(desc, 'LINE_USER_ID');
  var start = event.getStartTime();
  var no = String(lessonNoFromRaw_(userId, tokyoDateString_(start), Utilities.formatDate(start, TZ, 'HH:mm')));
  setDescFlag_(event, 'LESSON_NO', no);
  return no;
}

function setDescFlag_(event, key, value) {
  var desc = event.getDescription() || '';
  var pattern = new RegExp(key + ':([^\\n]*)');
  if (pattern.test(desc)) {
    event.setDescription(desc.replace(pattern, key + ':' + value));
  } else {
    event.setDescription(desc + '\n' + key + ':' + value);
  }
}

function formatStaffDate_(date) {
  var day = Number(Utilities.formatDate(date, TZ, 'd'));
  var month = Number(Utilities.formatDate(date, TZ, 'M'));
  return day + ordinalSuffix_(day) + ' ' + STAFF_MONTHS[month - 1];
}

function ordinalSuffix_(day) {
  var mod100 = day % 100;
  if (mod100 >= 11 && mod100 <= 13) return 'th';
  if (day % 10 === 1) return 'st';
  if (day % 10 === 2) return 'nd';
  if (day % 10 === 3) return 'rd';
  return 'th';
}

function formatStaffTime_(date) {
  return Utilities.formatDate(date, TZ, 'HH:mm');
}

function withLessonType_(text, lessonType) {
  var label = STAFF_LESSON_LABELS[lessonType] || '';
  if (!label) return text;
  return text + '\n' + label;
}

function staffBookingMessage_(name, lessonNo, start, lessonType) {
  return withLessonType_(
    'New Booking: ' + name + ' booked lesson #' + lessonNo + ' for ' + formatStaffDate_(start) + ' at ' + formatStaffTime_(start) + '.',
    lessonType
  );
}

function staffCancelMessage_(name, lessonNo, start, lessonType) {
  return withLessonType_(
    'Booking Cancelled: ' + name + ' cancelled lesson #' + lessonNo + ' for ' + formatStaffDate_(start) + ' at ' + formatStaffTime_(start) + '.',
    lessonType
  );
}

function staffEveMessage_(name, lessonNo, start, lessonType) {
  return withLessonType_(
    'Reminder: ' + name + ' has lesson #' + lessonNo + ' tomorrow (' + formatStaffDate_(start) + ') at ' + formatStaffTime_(start) + '.',
    lessonType
  );
}

function staffHourMessage_(name, lessonNo, start, lessonType) {
  return withLessonType_(
    "Reminder: " + name + "'s lesson #" + lessonNo + ' starts in 1 hour (' + formatStaffDate_(start) + ' at ' + formatStaffTime_(start) + ').',
    lessonType
  );
}

function sendLine_(userId, text) {
  return pushLine_(userId, text).ok;
}

function pushLine_(userId, text) {
  var token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token || !userId || !text) return { ok: false, code: 0, body: 'missing token, user, or text' };
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      to: userId,
      messages: [{ type: 'text', text: text }]
    }),
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  return {
    ok: code >= 200 && code < 300,
    code: code,
    body: res.getContentText()
  };
}

function slotStart_(date, time) {
  return new Date(date + 'T' + time + ':00+09:00');
}

function parseTokyoDate_(date) {
  return new Date(date + 'T00:00:00+09:00');
}

function endOfDay_(startDay) {
  var ymd = tokyoDateString_(startDay);
  return new Date(parseTokyoDate_(ymd).getTime() + 24 * 60 * 60 * 1000);
}

function tokyoDateString_(date) {
  return Utilities.formatDate(date, TZ, 'yyyy-MM-dd');
}

function tokyoPlusDays_(days) {
  var parts = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd').split('-');
  var utc = new Date(Date.UTC(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]) + days));
  return utc.toISOString().slice(0, 10);
}

function weekdayJa_(date) {
  var weekday = Number(Utilities.formatDate(date, TZ, 'u'));
  return WEEKDAY_JA[weekday === 7 ? 0 : weekday];
}

function formatLessonWhen_(date) {
  return Utilities.formatDate(date, TZ, 'M月d日') + '（' + weekdayJa_(date) + '）' +
    Utilities.formatDate(date, TZ, 'HH:mm') + '〜';
}

function calendarStudentName_(userName) {
  var name = String(userName || '').trim();
  if (!name || name === 'ゲスト') return 'Guest';
  return name;
}

function bookingEventTitle_(userName) {
  return BOOKING_TITLE_PREFIX + ' (' + calendarStudentName_(userName) + ')';
}

function lessonLineMessage_(start, intro, enterLine) {
  return [
    intro,
    '・' + formatLessonWhen_(start),
    '・レッスン時間は約20分',
    '',
    enterLine,
    ZOOM_URL,
    '',
    'Enjoy your lesson✈️'
  ].join('\n');
}

function bookingLineMessage_(start) {
  return lessonLineMessage_(
    start,
    'ご予約ありがとうございます。',
    '当日は、下記リンクよりご入室ください。'
  );
}

function eveLineMessage_(start) {
  return lessonLineMessage_(
    start,
    'レッスン前日のリマインドです。',
    '明日は、下記リンクよりご入室ください。'
  );
}

function bookingDescription_(userId, userName, lessonType) {
  return [
    'Online lesson booking',
    'Student: ' + calendarStudentName_(userName),
    'LINE_USER_ID:' + userId,
    'LINE_DISPLAY_NAME:' + userName,
    'LESSON_TYPE:' + (lessonType || ''),
    'PREV_DAY_REMINDER_SENT:false',
    'STAFF_EVE_SENT:false',
    'STAFF_HOUR_SENT:false',
    'Meeting link:',
    ZOOM_URL
  ].join('\n');
}

function valueOf_(desc, key) {
  var match = desc.match(new RegExp(key + ':([^\\n]+)'));
  return match ? match[1].trim() : '';
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
