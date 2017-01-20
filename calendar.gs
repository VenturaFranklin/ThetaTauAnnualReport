function event_add_calendar(event_name, event_date, event_type, event_desc){
  // Adds an event to the calendar
  var event = CalendarApp.getDefaultCalendar()
    .createAllDayEvent(event_name,
                       new Date(event_date),
                       {description: 'Type: ' + event_type + '\n' + 
                        'Description: ' + event_desc});
  //OPTIONS:
  //description: the description of the event
  //location: the location of the event
  //guests: a comma-separated list of email addresses that should be added as guests
  //sendInvites: whether to send invitation emails (default: false)
}

function calendar_add_event(){
  // Gets a list of events from the calendar and adds it to the Events sheet
}