/** @OnlyCurrentDoc */

const SPREADSHEET_ID = '1dT30c_Im0UuOMaLgYmyc4t4yLGduhug6T4abpFs5CAQ';
const SHEET_NAME = '가족달력';
const TZ = 'Asia/Seoul';
const OWNERS = ['가족', '아빠', '엄마', '현준', '현아'];

function doGet(e) {
  ensureSheet_();

  const p = (e && e.parameter) ? e.parameter : {};
  const action = String(p.action || '');

  try {
    let result;

    if (action === 'calendarState') {
      result = getCalendarState_(Number(p.year), Number(p.month));
    } else if (action === 'calendarAdd') {
      result = addCalendarEvent_(p.date, p.owner, p.text);
    } else if (action === 'calendarUpdate') {
      result = updateCalendarEvent_(p.id, p.date, p.owner, p.text);
    } else if (action === 'calendarDelete') {
      result = deleteCalendarEvent_(p.id);
    } else {
      result = { ok: true, message: '가족달력 API' };
    }

    return jsonp_(p.callback, result);

  } catch (err) {
    return jsonp_(p.callback, {
      ok: false,
      message: err && err.message ? err.message : '오류가 발생했습니다.'
    });
  }
}

function ensureSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 5).setValues([[
      'ID', '날짜', '대상', '내용', '등록시각'
    ]]);
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('yyyy-mm-dd');
    sh.getRange('E:E').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }

  return sh;
}

function jsonp_(callback, data) {
  let cb = String(callback || 'callback');

  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(cb)) {
    cb = 'callback';
  }

  return ContentService
    .createTextOutput(cb + '(' + JSON.stringify(data) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function normalizeDate_(value) {
  if (!value) return '';

  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TZ, 'yyyy-MM-dd');
  }

  const text = String(value).trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return text;
  }

  const parsed = new Date(value);

  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, TZ, 'yyyy-MM-dd');
  }

  return text;
}

function getCalendarState_(year, month) {
  const now = new Date();

  if (!year || !month || month < 1 || month > 12) {
    year = Number(Utilities.formatDate(now, TZ, 'yyyy'));
    month = Number(Utilities.formatDate(now, TZ, 'M'));
  }

  const first = new Date(year, month - 1, 1);
  const start = new Date(first);
  start.setDate(start.getDate() - start.getDay());

  const end = new Date(start);
  end.setDate(end.getDate() + 42);

  const startKey = Utilities.formatDate(start, TZ, 'yyyy-MM-dd');
  const endKey = Utilities.formatDate(end, TZ, 'yyyy-MM-dd');

  const sh = ensureSheet_();
  const values = sh.getDataRange().getValues();
  const events = [];

  for (let i = 1; i < values.length; i++) {
    const [id, date, owner, text] = values[i];
    const dateKey = normalizeDate_(date);

    if (!dateKey || dateKey < startKey || dateKey >= endKey) continue;

    events.push({
      id: String(id || ''),
      date: dateKey,
      owner: String(owner || '가족'),
      text: String(text || '')
    });
  }

  return {
    ok: true,
    events: events,
    holidays: getKoreanHolidays_(start, end),
    serverTime: new Date().getTime()
  };
}

function addCalendarEvent_(date, owner, text) {
  validateInput_(date, owner, text);

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const sh = ensureSheet_();
    sh.appendRow([
      Utilities.getUuid(),
      String(date),
      String(owner),
      String(text).trim(),
      new Date()
    ]);

    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function updateCalendarEvent_(id, date, owner, text) {
  validateInput_(date, owner, text);

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const sh = ensureSheet_();
    const values = sh.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0] || '') === String(id || '')) {
        sh.getRange(i + 1, 2, 1, 3).setValues([[
          String(date),
          String(owner),
          String(text).trim()
        ]]);

        return { ok: true };
      }
    }

    return { ok: false, message: '일정을 찾지 못했습니다.' };
  } finally {
    lock.releaseLock();
  }
}

function deleteCalendarEvent_(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const sh = ensureSheet_();
    const values = sh.getDataRange().getValues();

    for (let i = values.length - 1; i >= 1; i--) {
      if (String(values[i][0] || '') === String(id || '')) {
        sh.deleteRow(i + 1);
        return { ok: true };
      }
    }

    return { ok: false, message: '일정을 찾지 못했습니다.' };
  } finally {
    lock.releaseLock();
  }
}

function validateInput_(date, owner, text) {
  date = String(date || '').trim();
  owner = String(owner || '').trim();
  text = String(text || '').trim();

  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
    throw new Error('날짜가 올바르지 않습니다.');
  }

  if (OWNERS.indexOf(owner) === -1) {
    throw new Error('가족 구분을 확인해 주세요.');
  }

  if (!text) {
    throw new Error('일정 내용을 입력해 주세요.');
  }

  if (text.length > 40) {
    throw new Error('일정은 40자 이내로 입력해 주세요.');
  }
}

/*
  대한민국 공휴일:
  개인 Google Calendar 권한을 요구하지 않도록
  공개된 대한민국 공휴일 ICS 파일만 읽습니다.
*/
function getKoreanHolidays_(start, end) {
  const result = {};

  const url =
    'https://calendar.google.com/calendar/ical/' +
    'ko.south_korea%23holiday%40group.v.calendar.google.com/' +
    'public/basic.ics';

  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      let ics = response.getContentText('UTF-8');

      // ICS 줄바꿈(folding) 해제
      ics = ics.replace(/\r?\n[ \t]/g, '');

      const blocks = ics.split('BEGIN:VEVENT');
      const startKey = Utilities.formatDate(start, TZ, 'yyyy-MM-dd');
      const endKey = Utilities.formatDate(end, TZ, 'yyyy-MM-dd');

      for (let i = 1; i < blocks.length; i++) {
        const block = blocks[i];

        const dateMatch =
          block.match(/DTSTART(?:;VALUE=DATE)?:([0-9]{8})/);

        const titleMatch =
          block.match(/SUMMARY(?:;[^:]*)?:(.*)/);

        if (!dateMatch || !titleMatch) continue;

        const raw = dateMatch[1];
        const dateKey =
          raw.slice(0, 4) + '-' +
          raw.slice(4, 6) + '-' +
          raw.slice(6, 8);

        if (dateKey < startKey || dateKey >= endKey) continue;

        let title = titleMatch[1]
          .split(/\r?\n/)[0]
          .replace(/\\,/g, ',')
          .replace(/\\;/g, ';')
          .replace(/\\n/gi, ' ')
          .trim();

        if (title) {
          result[dateKey] = title;
        }
      }
    }
  } catch (err) {
    // 아래 고정 공휴일 fallback 사용
  }

  // 공개 공휴일 데이터를 못 불러온 경우 최소한의 고정 공휴일 표시
  if (Object.keys(result).length === 0) {
    const years = new Set();
    const d = new Date(start);

    while (d < end) {
      years.add(d.getFullYear());
      d.setFullYear(d.getFullYear() + 1);
    }

    years.forEach(y => {
      result[y + '-01-01'] = '신정';
      result[y + '-03-01'] = '삼일절';
      result[y + '-05-05'] = '어린이날';
      result[y + '-06-06'] = '현충일';
      result[y + '-08-15'] = '광복절';
      result[y + '-10-03'] = '개천절';
      result[y + '-10-09'] = '한글날';
      result[y + '-12-25'] = '성탄절';
    });
  }

  return result;
}
