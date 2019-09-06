// https://tc39.github.io/ecma262/#sec-array.prototype.includes
if (!Array.prototype.includes) {
  Object.defineProperty(Array.prototype, 'includes', {
    value: function(searchElement, fromIndex) {
      if (this == null) {
        throw new TypeError('"this" is null or not defined');
      }

      // 1. Let O be ? ToObject(this value).
      var o = Object(this);

      // 2. Let len be ? ToLength(? Get(O, "length")).
      var len = o.length >>> 0;

      // 3. If len is 0, return false.
      if (len === 0) {
        return false;
      }

      // 4. Let n be ? ToInteger(fromIndex).
      //    (If fromIndex is undefined, this step produces the value 0.)
      var n = fromIndex | 0;

      // 5. If n ≥ 0, then
      //  a. Let k be n.
      // 6. Else n < 0,
      //  a. Let k be len + n.
      //  b. If k < 0, let k be 0.
      var k = Math.max(n >= 0 ? n : len - Math.abs(n), 0);

      function sameValueZero(x, y) {
        return (
          x === y ||
          (typeof x === 'number' &&
            typeof y === 'number' &&
            isNaN(x) &&
            isNaN(y))
        );
      }

      // 7. Repeat, while k < len
      while (k < len) {
        // a. Let elementK be the result of ? Get(O, ! ToString(k)).
        // b. If SameValueZero(searchElement, elementK) is true, return true.
        if (sameValueZero(o[k], searchElement)) {
          return true;
        }
        // c. Increase k by 1.
        k++;
      }

      // 8. Return false
      return false;
    },
  });
}
const titleColumn = 'A';
const calendarColumn = 'B';
const startDateColumn = 'C';
const endDateColumn = 'D';
const idColumn = 'E';
type Column = {
  titleColumn: string;
  calendarColumn: string;
  startDateColumn: string;
  endDateColumn: string;
  idColumn: string;
  allDayColumn: string;
};
const columnTitles: Column = {
  titleColumn: '名前',
  calendarColumn: '種類',
  startDateColumn: '開始日時',
  endDateColumn: '終了日時',
  idColumn: 'ID',
  allDayColumn: '終日',
};
let column: Column = {
  titleColumn,
  calendarColumn,
  startDateColumn,
  endDateColumn,
  idColumn,
  allDayColumn: 'F',
};
function getColumnFromIndex(index: number) {
  var alphabet = [
    'A',
    'B',
    'C',
    'D',
    'E',
    'F',
    'G',
    'H',
    'I',
    'J',
    'K',
    'L',
    'M',
    'N',
    'O',
    'P',
    'Q',
    'R',
    'S',
    'T',
    'U',
    'V',
    'W',
    'X',
    'Y',
    'Z',
  ];
  var column = '';
  while (index > 0) {
    var digit = (index - 1) % alphabet.length;
    column = alphabet[digit] + column;
    index = (index - 1 - digit) / alphabet.length;
  }
  return column;
}

function updateColumn() {
  const columnNames = getScheduleSheet()
    .getRange('1:1')
    .getValues()[0];
  const updatedColumnName = Object.keys(column).map(v => {
    for (let i = 0; i < columnNames.length; i++) {
      if (columnNames[i] === v) {
        return getColumnFromIndex(i);
      }
    }
    return null;
  });
  if (updatedColumnName.every(v => v != null)) {
    const updatedColumn: any = {};
    Object.keys(column).forEach(
      (k, i) => (updatedColumn[k] = updatedColumnName[i])
    );

    column = updatedColumn;
  }
}

let calendarsCache:
  | GoogleAppsScript.Calendar.Schema.CalendarListEntry[]
  | null = null;

function getScheduleSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
}

function getLastRowNumber() {
  const sheet = getScheduleSheet();
  const lastRow = sheet.getLastRow();
  const titles: string[][] = sheet
    .getRange(`${titleColumn}2:${titleColumn}`)
    .getValues();
  for (let i = lastRow; i >= 0; i--) {
    if (typeof titles[i][0] === 'string' && titles[i][0] !== '') {
      console.log(i);
      console.log(titles);
      console.log(titles[i][0]);
      console.log(titles.length);
      return i + 2;
    }
  }
  return null;
}

function getSpreadsheetSchedules(): Schedule[] {
  const sheet = getScheduleSheet();
  const rowNumber = getLastRowNumber();
  if (rowNumber === null) {
    return [];
  }
  const titles = sheet
    .getRange(`${titleColumn}2:${titleColumn}${rowNumber}`)
    .getValues();
  const calendars = sheet
    .getRange(`${calendarColumn}2:${calendarColumn}${rowNumber}`)
    .getValues();
  const startDates = sheet
    .getRange(`${startDateColumn}2:${startDateColumn}${rowNumber}`)
    .getValues();
  const endDates = sheet
    .getRange(`${endDateColumn}2:${endDateColumn}${rowNumber}`)
    .getValues();
  const ids = sheet
    .getRange(`${idColumn}2:${idColumn}${rowNumber}`)
    .getValues();
  return titles
    .map((title, i) => ({
      line: i + 2,
      title: title[0],
      calendar: calendars[i][0],
      startDate: startDates[i][0],
      endDate: endDates[i][0],
      id: ids[i][0],
    }))
    .filter(
      v =>
        Object.keys(v).every(
          k => k === 'id' || (v[k] != null && v[k] !== '')
        ) &&
        (Moment.moment(v.endDate).isSame(Moment.moment()) ||
          Moment.moment(v.endDate).isAfter(Moment.moment()))
    );
}

type Schedule = {
  line: string | number;
  title: string;
  calendar: string;
  startDate: string;
  endDate: string;
  id: string;
};
function getCalendars() {
  if (calendarsCache !== null) {
    return calendarsCache;
  }
  const calendars = Calendar.CalendarList.list({ maxResults: 100 });
  const ownerCalendars = calendars.items.filter(
    ({ accessRole }) => accessRole === 'owner'
  );
  calendarsCache = ownerCalendars;
  return ownerCalendars;
}

function getCalendarNames() {
  const calendars = getCalendars();
  return calendars.map(v => v.summary);
}

function addCalendarValidate() {
  const calendarNames = getCalendarNames();
  const sheet = getScheduleSheet().getRange(
    `${calendarColumn}2:${calendarColumn}`
  );
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(calendarNames)
    .build();
  sheet.setDataValidation(rule);
}

function addDateValidate() {
  const sheet = getScheduleSheet();
  const rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .build();
  sheet
    .getRange(`${startDateColumn}2:${endDateColumn}`)
    .setDataValidation(rule);
}

export function update() {
  updateColumn();
  addCalendarValidate();
  addDateValidate();
  const schedules = getSpreadsheetSchedules();
  const unregistered = schedules.filter(s => s.id === '' || s.id == null);
  register(unregistered);
  const registrated = schedules.filter(
    s => s.id !== '' && typeof s.id === 'string'
  );
  updateSchedule(registrated);
}
function updateSchedule(schedules: Schedule[]) {
  schedules.forEach(s => {
    const event = CalendarApp.getEventById(s.id);
    if (event.getTitle() !== s.title) {
      event.setTitle(s.title);
    }
    if (
      Moment.moment(event.getStartTime()).isSame(Moment.moment(s.startDate)) ||
      Moment.moment(event.getEndTime()).isSame(Moment.moment(s.endDate))
    ) {
      event.setTime(
        Moment.moment(s.startDate).toDate(),
        Moment.moment(s.endDate).toDate()
      );
    }
  });
}
function register(schedules: Schedule[]) {
  const calendars = getCalendars();
  const calendarNames = getCalendarNames();
  schedules.forEach(s => {
    for (let i = 0; i < calendarNames.length; i++) {
      if (s.calendar === calendarNames[i]) {
        const event = CalendarApp.getCalendarById(calendars[i].id).createEvent(
          s.title,
          Moment.moment(s.startDate).toDate(),
          Moment.moment(s.endDate).toDate()
        );
        const sheet = getScheduleSheet();
        sheet.getRange(`${idColumn}${s.line}`).setValue(event.getId());
      }
    }
  });
}

export function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('カレンダー連携', [
    { name: '更新', functionName: 'update' },
  ]);
}
