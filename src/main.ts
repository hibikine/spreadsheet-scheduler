function getCalenadar() {
  const calendars = Calendar.CalendarList.list({maxResults: 100});
  const calendarNames = calendars.items.map(v => v.summary)
  console.log(calendarNames)
  console.log(calendars.items.map(v=>v.accessRole))
  return calendars;
}
