function event_add_calendar(event_name, event_date, event_type, event_desc){
//  var event_name = "First Event";
//  var event_date = "01/01/2017";
  event_date = new Date(event_date);
  var calendar = CalendarApp.getDefaultCalendar();
  var event = find_event(calendar, event_name, event_date);
  if (!event){
  // Adds an event to the calendar
    var event = calendar
    .createAllDayEvent(event_name,
                       new Date(event_date),
                       {description: 'Type: ' + event_type + '\n' + 
                        'Description: ' + event_desc});
  } else {
    event.setDescription('Type: ' + event_type + '\n' + 
                         'Description: ' + event_desc)
  }
  //OPTIONS:
  //description: the description of the event
  //location: the location of the event
  //guests: a comma-separated list of email addresses that should be added as guests
  //sendInvites: whether to send invitation emails (default: false)
}

function find_event(calendar, event_name, event_date){
  var events = calendar.getEventsForDay(event_date);
  for (var i in events){
    var event = events[i];
    var name = event.getTitle();
    if (name == event_name){
      return event;
    }
  }
}

function calendar_add_event(){
  // Gets a list of events from the calendar and adds it to the Events sheet
}