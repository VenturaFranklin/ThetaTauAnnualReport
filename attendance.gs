function show_att_sheet_alert(){
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'ERROR',
     'Please edit the events or members on the Events or Membership Sheet',
      ui.ButtonSet.OK);
}

function att_name(name){
  return name;
//Used to undo vertical name, not needed
//  var new_string = "";
//  for (var j = 0; j < name.length; j++){
//    var char = name[j];
//    if (j % 2 == 0){
//      new_string = new_string.concat(char);
//    }
//  }
//  return new_string
}

function check_duplicates(event_name, event_date){
  // Just added event need to make sure no duplicate
//  var event_date = "Sun Jan 01 2017 00:00:00 GMT-0700 (MST)";
//  var event_name = "First Event";
  var AttendanceObject = get_sheet_data("Attendance");
  var last = AttendanceObject.name_date.lastIndexOf(event_name+event_date);
  var first = AttendanceObject.name_date.indexOf(event_name+event_date);
  if (first != last){
    Logger.log("(" + arguments.callee.name + ") " + "There's a duplicate!");
    var sheet = AttendanceObject.sheet;
    sheet.deleteRow(last+1+1); // Extra 1 is for the header row which is not included in name_date
    if ((last-first) > 1){
      // When there are > 1 extra events
      Logger.log("(" + arguments.callee.name + ") " + "More than one duplicate!");
      check_duplicates(event_name, event_date);
    }
  }
}

function attendance_add_event(event_name, event_date){
  //align_attendance_events(myObject["Event Name"][0], myObject.Date[0])
//  var event_name = "Test";
//  var event_date = "Mon Aug 01 2016 00:00:00 GMT-0700 (MST)";
  if (!event_name || !event_date){
    return;
//    var event_data = get_sheet_data("Events");
//    var event_values = event_data.range.getValues();
  }
  sleep(2000); // Sometimes the function gets called twice, need to sleep
  var att_data = get_sheet_data("Attendance");
  if (att_data.name_date.indexOf(event_name+event_date) > -1){
    return;
  }
  var sheet = att_data.sheet;
//  var att_values = att_data.range.getValues();
  var attendance_rows = att_data.max_row;
  var attendance_cols = att_data.max_column;
  Logger.log("(" + arguments.callee.name + ") " +attendance_rows);
  sheet.insertRowBefore(attendance_rows+1);
  var att_row = sheet.getRange(attendance_rows+1, 1, 1, 2);
  att_row.setValues([[event_name, event_date]]);
  var att_row_full = sheet.getRange(attendance_rows+1, 3, 1, attendance_cols-2);
  var default_values =
      Array.apply(null, Array(attendance_cols-2)).map(function() { return 'U' });;
  att_row_full.setValues([default_values]);
//  att_row_full.setBackground("white");
  var attendance = range_object(sheet, attendance_rows+1);
  update_attendance(attendance);
  main_range_object("Attendance");
  check_duplicates(event_name, event_date);
}

function update_attendance(attendance){
  // Function to update the events sheet with the attendance counts
  // input: attendance object
  // example:
  // var attendance = range_object("Attendance", 26);
  // needs the membership sheet, and event sheet 

  var MemberObject = main_range_object("Membership");
//  Logger.log("(" + arguments.callee.name + ") " +attendance);
  var event_name_att = attendance["Event Name"][0];
  var event_date_att = attendance["Date"][0];
  Logger.log("(" + arguments.callee.name + ") " +event_name_att);
  if (event_name_att == ""){
    return;
  }
  var counts = {};
  counts["Student"] = {};
  counts["Pledge"] = {};
  counts["Shiny"] = {}
  counts["Away"] = {};
  counts["Alumn"] = {};
  var test_len = attendance.object_count;
  for(var i = 2; i< attendance.object_count; i++) {
    var member_name_att = attendance.object_header[i];
    var member_name_short = att_name(attendance.object_header[i]);
    var member_object = find_member_shortname(MemberObject, member_name_short);
    if (typeof member_object == 'undefined') {
      // Member may no longer exists on Membership sheet
      continue;
    }
    var event_status = attendance[member_name_att][0];
    event_status = event_status.toUpperCase();
    var member_status = member_object["Chapter Status"][0];
    switch (member_status) {
      case "Away":
        var start = member_object["Status Start"][0];
        var end = member_object["Status End"][0];
        if ((event_date_att > end) || (event_date_att < start)){
          member_status = "Student";
        }
        break;
      case "Alumn":
        var start = member_object["Status Start"][0];
        if (event_date_att < start){
          member_status = "Student";
        }
        break;
      case "Shiny":
        var start = member_object["Status Start"][0];
        if (event_date_att < start){
          member_status = "Pledge";
        } else {
          member_status = "Student";
        }
        break;
    }
//    Logger.log("(" + arguments.callee.name + ") " +[member_name_short, member_object, event_status, member_status]);
    counts[member_status][event_status] = counts[member_status][event_status] ? counts[member_status][event_status] + 1 : 1;
  }
  Logger.log("(" + arguments.callee.name + ") " +counts)
  var event_info = att_event_exists("Events", attendance)
  Logger.log("(" + arguments.callee.name + ") " +"ROW: " + event_info.event_row +
             " Active: " + event_info.active_col + " Pledge: " + event_info.pledge_col)
  if (typeof event_info.event_row == 'undefined'){
    Logger.log("(" + arguments.callee.name + ") Event may have been deleted?");
    return}; // This might mean that the event has been deleted
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Events");
  var active_range = sheet.getRange(event_info.event_row, event_info.active_col)
  var pledge_range = sheet.getRange(event_info.event_row, event_info.pledge_col)
  var num_actives = counts["Student"]["P"] ? counts["Student"]["P"]:0;
  var num_pledges = counts["Pledge"]["P"] ? counts["Pledge"]["P"]:0;
  active_range.setValue(num_actives)
  pledge_range.setValue(num_pledges)
}