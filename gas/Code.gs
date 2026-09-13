/**
 * Kepty English 予約バックエンド
 *
 * 上流はスプレッドシート「calendar」。
 * 2行で1日。時刻の下のチェックがONなら、その時刻〜右隣の時刻の1時間が稼働。
 * 予約枠はその1時間を20分×3本に割る（例: 18時チェック → 18:00 / 18:20 / 18:40）。
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
 * 4. トリガー: syncAvailabilityToCalendar を1分おき
 *    （チェックを入れたら、カレンダーの稼働ブロックが追従する）
 */

var TZ = 'Asia/Tokyo';
var SLOT_MINUTES = 20;
var DURATION_MINUTES = 20;
var LEAD_MINUTES = 60;
var SPREADSHEET_ID = '1OLiHIs7HjtlSxE9j3ETGhlE8efjiayomMpmYb-376U0';
var SHEET_NAME = 'calendar';
var AVAILABILITY_TITLE = 'Online Lesson Booking Slot';
var AVAILABILITY_TAG = 'SHEET_AVAILABILITY:true';

function doGet(e) {
  try {
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
    var userId = String(data.userId || '').trim();
    var userName = String(data.userName || 'ゲスト').trim();
    var date = String(data.date || '').trim();
    var time = String(data.time || '').trim();

    if (!userId || userId === 'GUEST') {
      return json_({ success: false, message: '公式LINEの予約メニューから開いてください。' });
    }
    if (!date || !time) {
      return json_({ success: false, message: '日時を選択してください。' });
    }

    var slot = findSlot_(date, time);
    if (!slot || !slot.available) {
      return json_({ success: false, message: 'その枠はすでに埋まりました。別の日時を選んでください。' });
    }

    var start = slotStart_(date, time);
    var end = new Date(start.getTime() + DURATION_MINUTES * 60 * 1000);
    var calendar = getCalendar_();
    var event = calendar.createEvent('レッスン予約（' + userName + '）', start, end, {
      description: [
        'LINE_USER_ID:' + userId,
        'LINE_DISPLAY_NAME:' + userName,
        'REMINDER_SENT:false'
      ].join('\n')
    });
    event.setColor(CalendarApp.EventColor.ORANGE);

    var whenText = formatWhen_(start);
    sendLine_(userId, 'ご予約を受け付けました。\n' + whenText + '\n開始の1時間前に、こちらへリマインドをお送りします。');

    return json_({
      success: true,
      message: whenText + ' で予約しました。レッスン1時間前にLINEでお知らせします。',
      eventId: event.getId()
    });
  } catch (err) {
    return json_({ success: false, message: '予約処理でエラーが発生しました。' });
  } finally {
    lock.releaseLock();
  }
}

function sendReminders() {
  var calendar = getCalendar_();
  var now = new Date();
  var from = new Date(now.getTime() + 50 * 60 * 1000);
  var to = new Date(now.getTime() + 70 * 60 * 1000);
  var events = calendar.getEvents(from, to);

  events.forEach(function (event) {
    var desc = event.getDescription() || '';
    var userId = valueOf_(desc, 'LINE_USER_ID');
    var sent = valueOf_(desc, 'REMINDER_SENT');
    if (!userId || sent === 'true') return;

    sendLine_(userId, 'まもなくレッスンです。\n' + formatWhen_(event.getStartTime()) + '\n準備ができたらお待ちしています。');
    event.setDescription(desc.replace('REMINDER_SENT:false', 'REMINDER_SENT:true'));
  });
}

function checkAndSendReminders() {
  sendReminders();
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

    days[dateStr] = windows;
  }

  return days;
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
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    var days = loadAvailabilityDays_();
    var today = tokyoDateString_(new Date());
    var desired = {};
    var minStart = null;
    var maxEnd = null;

    Object.keys(days).forEach(function (date) {
      if (date < today) return;
      days[date].forEach(function (win) {
        var key = date + '|' + Utilities.formatDate(win.start, TZ, 'HH:mm') + '|' + Utilities.formatDate(win.end, TZ, 'HH:mm');
        desired[key] = win;
        if (!minStart || win.start.getTime() < minStart.getTime()) minStart = win.start;
        if (!maxEnd || win.end.getTime() > maxEnd.getTime()) maxEnd = win.end;
      });
    });

    var calendar = getCalendar_();
    var rangeStart = parseTokyoDate_(today);
    var rangeEnd = maxEnd ? endOfDay_(maxEnd) : endOfDay_(rangeStart);
    var existing = calendar.getEvents(rangeStart, rangeEnd);
    var kept = {};

    existing.forEach(function (event) {
      if (!isSheetAvailability_(event)) return;
      var key = tokyoDateString_(event.getStartTime()) + '|' +
        Utilities.formatDate(event.getStartTime(), TZ, 'HH:mm') + '|' +
        Utilities.formatDate(event.getEndTime(), TZ, 'HH:mm');
      if (desired[key]) {
        kept[key] = true;
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
      event.setColor(CalendarApp.EventColor.PALE_BLUE);
    });
  } finally {
    lock.releaseLock();
  }
}

function onEdit(e) {
  try {
    if (!e || !e.range) {
      syncAvailabilityToCalendar();
      return;
    }
    var sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAME) return;
    syncAvailabilityToCalendar();
  } catch (err) {}
}

function findSlot_(date, time) {
  var slots = buildSlots_(date);
  for (var i = 0; i < slots.length; i++) {
    if (slots[i].time === time) return slots[i];
  }
  return null;
}

function getCalendarSheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || SPREADSHEET_ID;
  var spreadsheet = SpreadsheetApp.openById(id);
  var sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('calendarシートが見つかりません');
  return sheet;
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
  return title.indexOf('レッスン予約') !== -1 || desc.indexOf('LINE_USER_ID:') !== -1;
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

function sendLine_(userId, text) {
  var token = PropertiesService.getScriptProperties().getProperty('LINE_CHANNEL_ACCESS_TOKEN');
  if (!token || !userId) return;
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      to: userId,
      messages: [{ type: 'text', text: text }]
    }),
    muteHttpExceptions: true
  });
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

function formatWhen_(date) {
  return Utilities.formatDate(date, TZ, 'M月d日 HH:mm') + '〜';
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
