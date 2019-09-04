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
const calendarColumn = 'B';

const startDateColumn = 'C';
const endDateColumn = 'D';

let calendarsCache:
  | GoogleAppsScript.Calendar.Schema.CalendarListEntry[]
  | null = null;
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()[0]
    .getRange(`${calendarColumn}2:${calendarColumn}`);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(calendarNames)
    .build();
  sheet.setDataValidation(rule);
}

function addDateValidate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .build();
  sheet
    .getRange(`${startDateColumn}2:${endDateColumn}`)
    .setDataValidation(rule);
}

function update() {
  addCalendarValidate();
  addDateValidate();
}

function getSchedulesTester() {
  const calendars = getCalendars();
  getSchedules(calendars.map(c => c.id));
}

function getSchedules(calendarIds: string[]) {
  const today = new Date();
  const nextYear = new Date();
  nextYear.setFullYear(today.getFullYear() + 1);
  const events = CalendarApp.getEvents(today, nextYear).filter(
    v =>
      calendarIds.includes(v.getOriginalCalendarId()) && !v.isRecurringEvent()
  );
  console.log(events.map(v => v.getTitle()));
  return events;
}
