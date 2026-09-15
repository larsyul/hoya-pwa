/* =========================================================
   가족달력 API 추가 코드
   사용법:
   1) 기존 Code.gs의 doGet() 함수를 아래 doGet(e)로 교체
   2) 이 파일의 나머지 코드를 Code.gs 맨 아래에 붙여넣기
   3) 저장 후 기존 웹 앱 배포를 "새 버전"으로 업데이트
   ========================================================= */

const FAMILY_CAL_SHEET = '가족달력';
const FAMILY_OWNERS = ['가족','아빠','엄마','현준','현아'];

/* ★ 기존 doGet() 대신 이 함수 사용 */
function doGet(e) {
  ensureSheet_();

  const action = e && e.parameter ? String(e.parameter.action || '') : '';

  if (action.indexOf('calendar') === 0) {
    return handleFamilyCalendarApi_(e);
  }

  return HtmlService
    .createTemplateFromFile('index')
    .evaluate()
    .setTitle('호야 밥 체크')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/* =========================
   가족달력 시트
========================= */

function ensureFamilyCalendarSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(FAMILY_CAL_SHEET);

  if (!sh) {
    sh = ss.insertSheet(FAMILY_CAL_SHEET);
    sh.getRange(1, 1, 1, 5).setValues([[
      'ID',
      '날짜',
      '대상',
      '내용',
      '등록시각'
    ]]);
    sh.setFrozenRows(1);
    sh.getRange('B:B').setNumberFormat('yyyy-mm-dd');
    sh.getRange('E:E').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }

  return sh;
}


/* =========================
   가족달력 JSONP API
========================= */

function handleFamilyCalendarApi_(e) {
  const p = e.parameter || {};
  const action = String(p.action || '');

  try {
    let result;

    if (action === 'calendarState') {
      result = familyCalendarState_(
        Number(p.year),
        Number(p.month)
      );
    } else if (action === 'calendarAdd') {
      result = familyCalendarAdd_(
        p.date,
        p.owner,
        p.text
      );
    } else if (action === 'calendarUpdate') {
      result = familyCalendarUpdate_(
        p.id,
        p.date,
        p.owner,
        p.text
      );
    } else if (action === 'calendarDelete') {
      result = familyCalendarDelete_(p.id);
    } else {
      result = {
        ok: false,
        message: '지원하지 않는 요청입니다.'
      };
    }

    return jsonpResponse_(p.callback, result);

  } catch (err) {
    return jsonpResponse_(p.callback, {
      ok: false,
      message: err && err.message
        ? err.message
        : '오류가 발생했습니다.'
    });
  }
}


function jsonpResponse_(callback, data) {
  let cb = String(callback || 'callback');

  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(cb)) {
    cb = 'callback';
  }

  return ContentService
    .createTextOutput(
      cb + '(' + JSON.stringify(data) + ');'
    )
    .setMimeType(
      ContentService.MimeType.JAVASCRIPT
    );
}


/* =========================
   월간 상태
========================= */

function familyCalendarState_(year, month) {
  const now = new Date();

  if (!year || !month || month < 1 || month > 12) {
    year = Number(
      Utilities.formatDate(now, TZ, 'yyyy')
    );
    month = Number(
      Utilities.formatDate(now, TZ, 'M')
    );
  }

  const first =
    new Date(year, month - 1, 1);

  const start =
    new Date(first);

  start.setDate(
    start.getDate() - start.getDay()
  );

  const end =
    new Date(start);

  end.setDate(
    end.getDate() + 42
  );

  const startKey =
    Utilities.formatDate(
      start,
      TZ,
      'yyyy-MM-dd'
    );

  const endKey =
    Utilities.formatDate(
      end,
      TZ,
      'yyyy-MM-dd'
    );

  const sh =
    ensureFamilyCalendarSheet_();

  const values =
    sh.getDataRange().getValues();

  const events = [];

  for (let i = 1; i < values.length; i++) {
    const [
      id,
      date,
      owner,
      text
    ] = values[i];

    const dateKey =
      normalizeDateKey_(date);

    if (
      !dateKey ||
      dateKey < startKey ||
      dateKey >= endKey
    ) {
      continue;
    }

    events.push({
      id: String(id || ''),
      date: dateKey,
      owner: String(owner || '가족'),
      text: String(text || '')
    });
  }

  events.sort((a, b) => {
    if (a.date !== b.date) {
      return a.date.localeCompare(b.date);
    }

    return (
      FAMILY_OWNERS.indexOf(a.owner) -
      FAMILY_OWNERS.indexOf(b.owner)
    );
  });

  return {
    ok: true,
    events: events,
    holidays:
      koreanHolidayMap_(start, end),
    serverTime:
      new Date().getTime()
  };
}


/* =========================
   일정 추가 / 수정 / 삭제
========================= */

function familyCalendarAdd_(date, owner, text) {
  validateFamilyCalendarInput_(
    date,
    owner,
    text
  );

  const lock =
    LockService.getScriptLock();

  lock.waitLock(5000);

  try {
    const sh =
      ensureFamilyCalendarSheet_();

    sh.appendRow([
      Utilities.getUuid(),
      String(date),
      String(owner),
      String(text).trim(),
      new Date()
    ]);

    return {
      ok: true,
      message: '일정을 저장했습니다.'
    };

  } finally {
    lock.releaseLock();
  }
}


function familyCalendarUpdate_(
  id,
  date,
  owner,
  text
) {
  id = String(id || '').trim();

  if (!id) {
    throw new Error(
      '수정할 일정이 없습니다.'
    );
  }

  validateFamilyCalendarInput_(
    date,
    owner,
    text
  );

  const lock =
    LockService.getScriptLock();

  lock.waitLock(5000);

  try {
    const sh =
      ensureFamilyCalendarSheet_();

    const values =
      sh.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0] || '') === id) {
        sh.getRange(i + 1, 2, 1, 3)
          .setValues([[
            String(date),
            String(owner),
            String(text).trim()
          ]]);

        return {
          ok: true,
          message: '일정을 수정했습니다.'
        };
      }
    }

    return {
      ok: false,
      message: '일정을 찾지 못했습니다.'
    };

  } finally {
    lock.releaseLock();
  }
}


function familyCalendarDelete_(id) {
  id = String(id || '').trim();

  if (!id) {
    throw new Error(
      '삭제할 일정이 없습니다.'
    );
  }

  const lock =
    LockService.getScriptLock();

  lock.waitLock(5000);

  try {
    const sh =
      ensureFamilyCalendarSheet_();

    const values =
      sh.getDataRange().getValues();

    for (
      let i = values.length - 1;
      i >= 1;
      i--
    ) {
      if (String(values[i][0] || '') === id) {
        sh.deleteRow(i + 1);

        return {
          ok: true,
          message: '일정을 삭제했습니다.'
        };
      }
    }

    return {
      ok: false,
      message: '일정을 찾지 못했습니다.'
    };

  } finally {
    lock.releaseLock();
  }
}


function validateFamilyCalendarInput_(
  date,
  owner,
  text
) {
  date = String(date || '').trim();
  owner = String(owner || '').trim();
  text = String(text || '').trim();

  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(date)
  ) {
    throw new Error(
      '날짜가 올바르지 않습니다.'
    );
  }

  if (
    FAMILY_OWNERS.indexOf(owner) === -1
  ) {
    throw new Error(
      '가족 구성원을 확인해 주세요.'
    );
  }

  if (!text) {
    throw new Error(
      '일정 내용을 입력해 주세요.'
    );
  }

  if (text.length > 40) {
    throw new Error(
      '일정은 40자 이내로 입력해 주세요.'
    );
  }
}


/* =========================
   대한민국 공휴일
   Google 공휴일 캘린더 사용
========================= */

function koreanHolidayMap_(start, end) {
  const result = {};

  const calendarIds = [
    'ko.south_korea#holiday@group.v.calendar.google.com',
    'en.south_korea#holiday@group.v.calendar.google.com'
  ];

  for (
    let c = 0;
    c < calendarIds.length;
    c++
  ) {
    try {
      const cal =
        CalendarApp.getCalendarById(
          calendarIds[c]
        );

      if (!cal) {
        continue;
      }

      const list =
        cal.getEvents(start, end);

      list.forEach(event => {
        const dateKey =
          Utilities.formatDate(
            event.getStartTime(),
            TZ,
            'yyyy-MM-dd'
          );

        result[dateKey] =
          String(event.getTitle() || '');
      });

      if (
        Object.keys(result).length > 0
      ) {
        break;
      }

    } catch (err) {
      // 다음 캘린더 ID 시도
    }
  }

  // Google 공휴일 캘린더를 불러오지 못했을 때
  // 고정 날짜 공휴일만 최소 표시
  if (
    Object.keys(result).length === 0
  ) {
    const years = {};

    let d = new Date(start);

    while (d < end) {
      years[d.getFullYear()] = true;
      d.setFullYear(
        d.getFullYear() + 1
      );
    }

    Object.keys(years).forEach(y => {
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
