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
 * 空き枠は「Online Lesson Booking Slot」の予約可能ブロックを20分刻みで返します。
 * ブロックが無い／全部潰れる日は、平日 16:30〜19:30 を予備の枠として出します。
 * 埋め済みは「レッスン予約」だけを見ます（Googleの空き枠イベントは塞ぎません）。
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

function buildSlots_(date) {
  if (date <= tokyoDateString_(new Date())) return [];

  var startDay = parseTokyoDate_(date);
  var weekday = Number(Utilities.formatDate(startDay, TZ, 'u'));
  var events = getCalendar_().getEvents(startDay, endOfDay_(startDay));
  var windows = [];
  var busy = [];
  var hasScheduleDay = false;

  events.forEach(function (event) {
    if (isOurBooking_(event)) {
      busy.push({ start: event.getStartTime(), end: event.getEndTime() });
      return;
    }
    if (isScheduleNamed_(event)) {
      hasScheduleDay = true;
      if (!event.isAllDayEvent()) {
        windows.push({ start: event.getStartTime(), end: event.getEndTime() });
      }
    }
  });

  windows = mergeWindows_(windows);
  if (!windows.length && (hasScheduleDay || WEEKDAYS.indexOf(weekday) !== -1)) {
    windows = [defaultWindow_(date)];
  }
  if (!windows.length) return [];

  var slots = slotsFromWindows_(windows, busy);
  if (!slots.length && WEEKDAYS.indexOf(weekday) !== -1) {
    slots = slotsFromWindows_([defaultWindow_(date)], busy);
  }
  return slots;
}

function defaultWindow_(date) {
  var start = slotStart_(date, FIRST_SLOT);
  var lastStart = slotStart_(date, LAST_SLOT);
  return {
    start: start,
    end: new Date(lastStart.getTime() + DURATION_MINUTES * 60 * 1000)
  };
}

function slotsFromWindows_(windows, busy) {
  var now = new Date();
  var slots = [];
  var seen = {};

  windows.forEach(function (win) {
    var cursor = ceilToSlot_(win.start);
    var winEnd = win.end.getTime();
    while (cursor.getTime() + DURATION_MINUTES * 60 * 1000 <= winEnd + 1000) {
      var slotEnd = new Date(cursor.getTime() + DURATION_MINUTES * 60 * 1000);
      var time = Utilities.formatDate(cursor, TZ, 'HH:mm');
      if (!seen[time]) {
        var tooSoon = cursor.getTime() < now.getTime() + LEAD_MINUTES * 60 * 1000;
        var blocked = overlapsBusy_(cursor, slotEnd, busy);
        if (!tooSoon && !blocked) {
          slots.push({ time: time, available: true });
        }
        seen[time] = true;
      }
      cursor = new Date(cursor.getTime() + SLOT_MINUTES * 60 * 1000);
    }
  });

  slots.sort(function (a, b) {
    return a.time < b.time ? -1 : 1;
  });
  return slots;
}

function mergeWindows_(windows) {
  if (!windows.length) return [];
  windows.sort(function (a, b) {
    return a.start.getTime() - b.start.getTime();
  });
  var merged = [windows[0]];
  for (var i = 1; i < windows.length; i++) {
    var last = merged[merged.length - 1];
    if (windows[i].start.getTime() <= last.end.getTime() + 1000) {
      if (windows[i].end.getTime() > last.end.getTime()) last.end = windows[i].end;
    } else {
      merged.push(windows[i]);
    }
  }
  return merged;
}

function overlapsBusy_(start, end, busy) {
  return busy.some(function (block) {
    return block.start.getTime() < end.getTime() && block.end.getTime() > start.getTime();
  });
}

function isOurBooking_(event) {
  var title = String(event.getTitle() || '');
  var desc = event.getDescription() || '';
  return title.indexOf('レッスン予約') !== -1 || desc.indexOf('LINE_USER_ID:') !== -1;
}

function isScheduleNamed_(event) {
  var title = String(event.getTitle() || '').toLowerCase();
  return title.indexOf(SCHEDULE_NAME.toLowerCase()) !== -1;
}

function ceilToSlot_(date) {
  var hour = Number(Utilities.formatDate(date, TZ, 'H'));
  var minute = Number(Utilities.formatDate(date, TZ, 'm'));
  var ymd = Utilities.formatDate(date, TZ, 'yyyy-MM-dd');
  var extra = minute % SLOT_MINUTES;
  if (extra !== 0) {
    minute += SLOT_MINUTES - extra;
    if (minute >= 60) {
      hour += 1;
      minute -= 60;
    }
  }
  return slotStart_(ymd, pad2_(hour) + ':' + pad2_(minute));
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
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
