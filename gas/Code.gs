/**
 * Kepty English 予約バックエンド
 *
 * デプロイ:
 * 1. このコードを、予約カレンダーを持っている Google アカウントの Apps Script に貼る
 * 2. プロジェクトの設定 → スクリプト プロパティ
 *    - LINE_CHANNEL_ACCESS_TOKEN … Messaging API のチャネルアクセストークン
 *    - CALENDAR_ID … 省略時はデフォルトカレンダー
 * 3. デプロイ → 新しいデプロイ → 種類: ウェブアプリ
 *    - 次のユーザーとして実行: 自分
 *    - アクセスできるユーザー: 全員
 * 4. 発行される URL は必ず
 *    https://script.google.com/macros/s/..../exec
 *    （/a/macros/kepty.co/ だと LINE 内からログインを要求され、空き枠取得に失敗します）
 * 空き枠の出どころ（重要）:
 * Googleの予約スケジュール「Online Lesson Booking Slot」の未予約枠は
 * カレンダー予定として存在しない（公式ヘルプ: 未予約の予定はカレンダーに出ない）。
 * CalendarApp.getEvents では稼働時間は取れない。
 * そのため稼働時間は予約ページと同じ定義（平日 16:30開始〜19:30開始 / 20分）で切り、
 * カレンダーは LINE から入った「レッスン予約」の埋め済み判定にだけ使う。
 * 予約確定時は同じカレンダーに「レッスン予約（名前）」を作成する。
 */

var TZ = 'Asia/Tokyo';
var SLOT_MINUTES = 20;
var DURATION_MINUTES = 20;
var LEAD_MINUTES = 60;
var FIRST_SLOT = '16:30';
var LAST_SLOT = '19:30';
var WEEKDAYS = [1, 2, 3, 4, 5];
var SCHEDULE_NAME = 'Online Lesson Booking Slot';

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

/**
 * 空き枠 = 稼働時間（平日 16:30〜19:30開始）から、レッスン予約を除いたもの。
 * 未予約の Appointment Schedule を getEvents で探してはいけない（カレンダーに無い）。
 */
function buildSlots_(date) {
  if (date <= tokyoDateString_(new Date())) return [];

  var startDay = parseTokyoDate_(date);
  var weekday = Number(Utilities.formatDate(startDay, TZ, 'u'));
  if (WEEKDAYS.indexOf(weekday) === -1) return [];

  var window = {
    start: slotStart_(date, FIRST_SLOT),
    end: new Date(slotStart_(date, LAST_SLOT).getTime() + DURATION_MINUTES * 60 * 1000)
  };

  var busy = [];
  getCalendar_().getEvents(startDay, endOfDay_(startDay)).forEach(function (event) {
    if (isScheduleNamed_(event)) return;
    if (isOurBooking_(event)) {
      busy.push({ start: event.getStartTime(), end: event.getEndTime() });
    }
  });

  return slotsFromWindows_([window], busy);
}

function slotsFromWindows_(windows, busy) {
  var now = new Date();
  var latest = now.getTime() + LEAD_MINUTES * 60 * 1000;
  var slots = [];
  var seen = {};
  var maxOpeningMs = (DURATION_MINUTES + 5) * 60 * 1000;
  var durationMs = DURATION_MINUTES * 60 * 1000;
  var stepMs = SLOT_MINUTES * 60 * 1000;

  windows.forEach(function (win) {
    var start = win.start;
    var endMs = win.end.getTime();
    if (endMs - start.getTime() <= maxOpeningMs) {
      addAvailableSlot_(start, latest, busy, seen, slots);
      return;
    }
    var cursor = new Date(Math.floor(start.getTime() / 60000) * 60000);
    while (cursor.getTime() + durationMs <= endMs + 2000) {
      addAvailableSlot_(cursor, latest, busy, seen, slots);
      cursor = new Date(cursor.getTime() + stepMs);
    }
  });

  slots.sort(function (a, b) {
    return a.time < b.time ? -1 : 1;
  });
  return slots;
}

function addAvailableSlot_(start, latest, busy, seen, slots) {
  var time = Utilities.formatDate(start, TZ, 'HH:mm');
  if (seen[time]) return;
  seen[time] = true;
  if (start.getTime() < latest) return;
  var slotEnd = new Date(start.getTime() + DURATION_MINUTES * 60 * 1000);
  if (overlapsBusy_(start, slotEnd, busy)) return;
  slots.push({ time: time, available: true });
}

function isScheduleNamed_(event) {
  var title = String(event.getTitle() || '').toLowerCase();
  return title.indexOf('online lesson booking') !== -1 || title.indexOf('booking slot') !== -1;
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

function findSlot_(date, time) {
  var slots = buildSlots_(date);
  for (var i = 0; i < slots.length; i++) {
    if (slots[i].time === time) return slots[i];
  }
  return null;
}

function getCalendar_() {
  var id = PropertiesService.getScriptProperties().getProperty('CALENDAR_ID');
  var calendar = id ? CalendarApp.getCalendarById(id) : CalendarApp.getDefaultCalendar();
  if (!calendar) throw new Error('カレンダーが見つかりません');
  return calendar;
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
  return new Date(startDay.getTime() + 24 * 60 * 60 * 1000);
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
