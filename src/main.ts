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
    .getRange('B2:B');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(calendarNames)
    .build();
  sheet.setDataValidation(rule);
}

function update() {
  addCalendarValidate();
}

function getSchedulesTester() {
  getSchedules;
}

function getSchedules(
  calendar: GoogleAppsScript.Calendar.Schema.CalendarListEntry
) {}
