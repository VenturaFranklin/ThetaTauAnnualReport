function get_event_list() {
  var event_object = main_range_object("Events");
  var event_list = [];
  for(var i = 0; i< event_object.object_count; i++) {
    var event_name = event_object.original_names[i];
    event_list.push(event_name);
    }
  return event_list
}

function att_event_exists(sheet_name, myObject) {
  // find if event exists on attendance or event sheet
//  var sheet_name = "Events";
//  var myObject = range_object("Attendance", 3);
  Logger.log("(" + arguments.callee.name + ") " + sheet_name);
  Logger.log(myObject);
  if (sheet_name == "Attendance") {
    var check_1 = 'Event Name';
    var check_2 = 'Date';
  } else {
    var check_1 = '# Members';
    var check_2 = '# Pledges';
  }
  var name_check = myObject["Event Name"][0];
  var date_check = myObject["Date"][0];
  var Object = main_range_object(sheet_name);
  for (var i = 0; i < Object.object_count; i++){
    var event_name = Object.object_header[i];
//    var event_date = Object[event_name]["Date"][0];
    if (event_name == name_check+date_check){
      Logger.log("(" + arguments.callee.name + ") " +[event_name, name_check+date_check]);
      var active_col = Object[event_name][check_1][1];
      var pledge_col = Object[event_name][check_2][1];
      var event_row = Object[event_name].object_row;
      break;
    }
  }
  return {active_col: active_col,
          pledge_col: pledge_col,
          event_row: event_row
         }
}

function get_needed_fields(event_type, ScoringObject){
  if (!ScoringObject){
    var ScoringObject = main_range_object("Scoring");
  }
  var score_object = ScoringObject[event_type];
  var needed_fields = score_object["Event Fields"][0];
  needed_fields = needed_fields.split(', ');
  var score_description = score_object["Long Description"][0];
  return {needed_fields: needed_fields,
          score_description: score_description
         }
}

function event_fields_set(myObject){
  var score_info = get_needed_fields(myObject["Type"][0]);
  var needed_fields = score_info.needed_fields;
  var score_description = score_info.score_description;
  var event_row = myObject["object_row"];
  var sheet = myObject["sheet"];
  var new_range = sheet.getRange(event_row, 3);
  new_range.setNote(score_description);
  var field_range = sheet.getRange(event_row, 10, 1, 4);
  field_range.setBackground("black")
             .setNote("Do not edit");
  // No needed fields
  if (needed_fields[0] == ""){return true;};
  var needed_field_values = [];
  Logger.log("(" + arguments.callee.name + ") " +needed_fields);
  var yes_no_fields = ['STEM?', 'HOST'];
  var optional_fields = yes_no_fields.slice(0);
  optional_fields.push('# Non- Members', 'MILES');
  for (var i in needed_fields){
    var needed_field = needed_fields[i];
    var needed_value = myObject[needed_field][0];
    var needed_col = myObject[needed_field][1];
    if (optional_fields.indexOf(needed_field) > -1) {
      var needed_range = sheet.getRange(event_row, needed_col);
      needed_range.setBackground("white")
      .clearNote();
      if (yes_no_fields.indexOf(needed_field) > -1){
        needed_range.setNote('Yes or No');
        if (needed_value==""){
          needed_range.setValue('No');
          needed_value = 'No';
        }
      } else {
        if (needed_value==""){
          needed_range.setValue(0);
          needed_value = 0;
        }
      }
    }
    needed_field_values.push(needed_value);
  }
  Logger.log("(" + arguments.callee.name + ") " +needed_field_values);
  if (needed_field_values.indexOf("") > -1){
    return false;
  }
  return true;
}

function events_to_att(ss, attendance_object, EventObject){
  try{
    progress_update("EVENTS TO ATTENDANCE");
    if (!ss){
      var ss = get_active_spreadsheet();
      var attendance_object = main_range_object("Attendance", undefined, ss);
      var EventObject = main_range_object("Events", undefined, ss);
    }
    var new_events = []
    var att_sheet = attendance_object.sheet;
    var max_rows = attendance_object.object_count + 1
    for (var j in EventObject.object_header){
      var event_name = EventObject.object_header[j];
      var event = EventObject[event_name];
      if (attendance_object.object_header.indexOf(event_name) < 0){
        new_events.push([event["Event Name"][0], event["Date"][0]]);
        att_sheet.insertRowAfter(max_rows);
      }
    }
    Logger.log(new_events);
    var att_sheet = attendance_object.sheet;
    var start_row = attendance_object.object_count ? attendance_object.object_count + 1:2;
    if (new_events.length == 0){
      progress_update("NO NEW EVENTS");
      return};
    var att_range = att_sheet.getRange(start_row, 1, new_events.length, 2);
    var attendance_cols = attendance_object.header_values.length;
    att_range.setValues(new_events);
    var default_values =
        Array.apply(null, Array(attendance_cols-2)).map(function() { return 'U' });
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['P', 'E', 'U', 'p', 'e', 'u'], false)
      .setHelpText('P-Present; E-Excused; U-Unexcused')
      .setAllowInvalid(false).build();
    for (var row = start_row;row < start_row + new_events.length;row++){
      var att_row_full = att_sheet.getRange(row, 3, 1, attendance_cols-2);
      att_row_full.setValues([default_values]);
      att_row_full.setDataValidation(rule);
  }
    var format_range = ss.getRangeByName("FORMAT");
    var max_column = att_sheet.getLastColumn();
    var max_row = att_sheet.getLastRow();
    format_range.copyFormatToRange(att_sheet, 3, max_column, 2, 1000);
    main_range_object("Attendance");
    progress_update("EVENTS TO ATTENDANCE FINISHED");
  } catch (e) {
    var message = Utilities.formatString('This error has automatically been sent to the developers. %s: %s (line %s, file "%s"). Stack: "%s" . While processing %s.',
                                         e.name||'', e.message||'', e.lineNumber||'', e.fileName||'',
                                         e.stack||'', arguments.callee.name||'');
    Logger = startBetterLog();
    Logger.severe(message);
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
     'ERROR',
      message,
      ui.ButtonSet.OK);
    return "";
  }
}

function refresh_events() {
  // This function should adjust the black bg add to calendar
  try{
    progress_update("REFRESH EVENTS");
    var ss = get_active_spreadsheet();
    var EventObject = main_range_object("Events", undefined, ss);
    var ScoringObject = main_range_object("Scoring");
    var all_infos = [];
    var bg_colors = [];
    var need_indx = {"# Non- Members": 0,
                     "STEM?": 1,
                     "HOST": 2,
                     "MILES": 3
                    }
    for (var j in EventObject.object_header){
      var event_name = EventObject.object_header[j];
      var event = EventObject[event_name];
      var score_info = get_needed_fields(event["Type"][0], ScoringObject);
      var needed_fields = score_info.needed_fields;
      var score_description = score_info.score_description;
      all_infos.push([score_description]);
      var color_array = ["black", "black", "black", "black"];
      if (needed_fields > 0){
        for (var i in needed_fields){
          var needed_field = needed_fields[i];
          color_array[need_indx[needed_field]] = "white";
        }
      }
      bg_colors.push(color_array);
    }
    var event_sheet = EventObject.sheet;
    var event_col = EventObject.header_values.indexOf("Type") + 1;
    var event_range = event_sheet.getRange(2, event_col, EventObject.object_count, 1);
    event_range.setNotes(all_infos);
    var field_col = EventObject.header_values.indexOf("# Non- Members") + 1;
    var field_range = event_sheet.getRange(2, field_col, EventObject.object_count, 4);
    field_range.setBackgrounds(bg_colors);
    setup_dataval();
    progress_update("REFRESH EVENTS FINISHED");
  } catch (e) {
    var message = Utilities.formatString('This error has automatically been sent to the developers. %s: %s (line %s, file "%s"). Stack: "%s" . While processing %s.',
                                         e.name||'', e.message||'', e.lineNumber||'', e.fileName||'',
                                         e.stack||'', arguments.callee.name||'');
    Logger = startBetterLog();
    Logger.severe(message);
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
     'ERROR',
      message,
      ui.ButtonSet.OK);
    return "";
  }
}