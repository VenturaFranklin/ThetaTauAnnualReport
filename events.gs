function get_event_list() {
  var event_object = main_range_object("Events");
  var event_list = [];
  for(var i = 0; i< event_object.object_count; i++) {
    var event_name = event_object.original_names[i];
    event_list.push(event_name);
    }
  return event_list
}

function add_event_sheet(){
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('What do you want to name new event sheet?\n'+
                         '(Events is added automatically,\neg. Events-yourname)\n'+
                         'Sheet added at end, you will have to move it if you want.',
      ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  var name = result.getResponseText();
  if (button == ui.Button.OK) {
    var ss = get_active_spreadsheet();
    var sheets = find_all_event_sheets(ss);
//   for (var sheet_name in sheets){
//     var sheet = sheets[sheet_name];
//     break;
//   }
    var sheet = sheets[Object.keys(sheets)[0]];
    var new_sheet = sheet.copyTo(ss);
//   SpreadsheetApp.flush();
    i = 0;
    var new_sheet_name = "Events-" + name;
    while (new_sheet_name in sheets){
      new_sheet_name = "Events-" + name + str(i);
      i+=1;
    }
    new_sheet.setName(new_sheet_name);
    new_sheet.deleteRows(2, new_sheet.getLastRow());
  }
}

function find_all_event_sheets(ss){
  var event_sheets = new Array();
  if (!ss){
    var ss = get_active_spreadsheet();
  }
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++){
    var sheet = sheets[i];
    var sheet_name = sheet.getName();
    if (sheet_name.indexOf('Event') >= 0){
      event_sheets[sheet_name] = sheet;
    }
  }
  return event_sheets;
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

function get_needed_fields(event_type, ScoringObject, notify){
  if (!ScoringObject){
    var ScoringObject = main_range_object("Scoring");
  }
  var score_object = ScoringObject[event_type];
  if (!score_object){
    var message = Utilities.formatString('This event does not exist!\nEvent name: %s?\n'+
                                         'Please make sure the event is from the drop down list.\nUse '+
                                         '"Refresh Events Background Stuff" if the drop downs are missing.',
                                           event_type||'');
      Logger = startBetterLog();
      Logger.severe(message);
    if(notify){
      var ui = SpreadsheetApp.getUi();
      var result = ui.alert(
        'ERROR',
        message,
        ui.ButtonSet.OK);
    }
    return {needed_fields: [],
            score_description: "Event does not exist",
            not_set: true
           }
  }
  var needed_fields = score_object["Event Fields"][0];
  needed_fields = needed_fields.split(', ');
  var score_description = score_object["Long Description"][0];
  return {needed_fields: needed_fields,
          score_description: score_description,
          not_set: false
         }
}

function event_fields_set(myObject){
  var score_info = get_needed_fields(myObject["Type"][0], undefined, true);
  var needed_fields = score_info.needed_fields;
  var score_description = score_info.score_description;
  var event_row = myObject["object_row"];
  var sheet = myObject["sheet"];
  var new_range = sheet.getRange(event_row, 3);
  new_range.setNote(score_description);
  new_range.setBackground('white');
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

//function events_to_att(ss, attendance_object, EventObject){
//  try{
//    progress_update("EVENTS TO ATTENDANCE");
//    if (!ss){
//      var ss = get_active_spreadsheet();
//      var attendance_object = main_range_object("Attendance", undefined, ss);
//      var EventObject = main_range_object("Events", undefined, ss);
//    }
//    var new_events = []
//    var att_sheet = attendance_object.sheet;
//    var max_rows = attendance_object.object_count + 2;
//    for (var j in EventObject.object_header){
//      var event_name = EventObject.object_header[j];
//      var event = EventObject[event_name];
//      if (attendance_object.object_header.indexOf(event_name) < 0){
//        new_events.push([event["Event Name"][0], event["Date"][0]]);
//      }
//    }
//    Logger.log(new_events);
//    var att_sheet = attendance_object.sheet;
//    var start_row = attendance_object.object_count ? attendance_object.object_count + 1:2;
//    if (new_events.length == 0){
//      progress_update("NO NEW EVENTS");
//      return};
//    att_sheet.insertRows(max_rows, new_events.length);
//    var att_range = att_sheet.getRange(start_row, 1, new_events.length, 2);
//    var attendance_cols = attendance_object.header_values.length;
//    att_range.setValues(new_events);
//    var default_values =
//        Array.apply(null, Array(attendance_cols-2)).map(function() { return 'U' });
//    var rule = SpreadsheetApp.newDataValidation()
//      .requireValueInList(['P', 'E', 'U', 'p', 'e', 'u'], false)
//      .setHelpText('P-Present; E-Excused; U-Unexcused')
//      .setAllowInvalid(false).build();
//    for (var row = start_row;row < start_row + new_events.length;row++){
//      var att_row_full = att_sheet.getRange(row, 3, 1, attendance_cols-2);
//      att_row_full.setValues([default_values]);
//      att_row_full.setDataValidation(rule);
//  }
//    var format_range = ss.getRangeByName("FORMAT");
//    var max_column = att_sheet.getLastColumn();
//    var max_row = att_sheet.getLastRow();
//    format_range.copyFormatToRange(att_sheet, 3, max_column, 2, 1000);
//    main_range_object("Attendance");
//    check_duplicate_missing_att_events(ss, EventObject);
//    progress_update("EVENTS TO ATTENDANCE FINISHED");
//  } catch (e) {
//    var message = Utilities.formatString('This error has automatically been sent to the developers. %s: %s (line %s, file "%s"). Stack: "%s" . While processing %s.',
//                                         e.name||'', e.message||'', e.lineNumber||'', e.fileName||'',
//                                         e.stack||'', arguments.callee.name||'');
//    Logger = startBetterLog();
//    Logger.severe(message);
//    var ui = SpreadsheetApp.getUi();
//    var result = ui.alert(
//     'ERROR',
//      message,
//      ui.ButtonSet.OK);
//    return "";
//  }
//}

function refresh_events_silent(){
  SILENT = true;
  refresh_events();
  SILENT = false;
}

function refresh_events() {
  // This function should adjust the black bg add to calendar
  try{
    progress_update("REFRESH EVENTS");
    var ss = get_active_spreadsheet();
    var EventObject = main_range_object("Events", undefined, ss);
    var ScoringObject = main_range_object("Scoring");
    var all_infos = {};
    var bg_colors = {};
    var bg_events = {};
    var bg_dates = {};
    var date_notes = {};
    var sheet_names = {};
    var need_indx = {"# Non- Members": 0,
                     "STEM?": 1,
                     "HOST": 2,
                     "MILES": 3
                    }
    var year_semesters = get_year_semesters();
    year_semesters = Object.keys(year_semesters);
    for (var j in EventObject.object_header){
      var event_name = EventObject.object_header[j];
      var event = EventObject[event_name];
      var event_sheet_name = event.sheet_name;
      sheet_names[event_sheet_name] = event.sheet;
      var event_color = 'white';
      all_infos[event_sheet_name] = all_infos[event_sheet_name] ? all_infos[event_sheet_name]:[];
      bg_colors[event_sheet_name] = bg_colors[event_sheet_name] ? bg_colors[event_sheet_name]:[];
      bg_events[event_sheet_name] = bg_events[event_sheet_name] ? bg_events[event_sheet_name]:[];
      bg_dates[event_sheet_name] = bg_dates[event_sheet_name] ? bg_dates[event_sheet_name]:[];
      date_notes[event_sheet_name] = date_notes[event_sheet_name] ? date_notes[event_sheet_name]:[];
      if (!check_date_year_semester(event.Date[0])){
        bg_dates[event_sheet_name].push(['red'])
        date_notes[event_sheet_name].push(["Date should be within year/semesters of Annual report.\n" + year_semesters.join(", ")]);
      } else {
        bg_dates[event_sheet_name].push(['white'])
        date_notes[event_sheet_name].push(['']);
      }
      var score_info = get_needed_fields(event["Type"][0], ScoringObject, false);
      var needed_fields = score_info.needed_fields;
      var score_description = score_info.score_description;
      all_infos[event_sheet_name].push([score_description]);
      if (score_info.not_set){
        event_color = 'red';
      }
      bg_events[event_sheet_name].push([event_color]);
      var color_array = ["black", "black", "black", "black"];
      if (needed_fields.length > 0){
        for (var i in needed_fields){
          var needed_field = needed_fields[i];
          color_array[need_indx[needed_field]] = "white";
        }
      }
      bg_colors[event_sheet_name].push(color_array);
    }
  var date_col = event.Date[1];
  var event_col = EventObject.header_values.indexOf("Type") + 1;
  var field_col = EventObject.header_values.indexOf("# Non- Members") + 1;
  for (event_sheet_name in sheet_names){
    var event_sheet = sheet_names[event_sheet_name];
    var rows = all_infos[event_sheet_name].length;
    var event_range = event_sheet.getRange(2, event_col, rows, 1);
    event_range.setNotes(all_infos[event_sheet_name]);
    event_range.setBackgrounds(bg_events[event_sheet_name]);
    var date_range = event_sheet.getRange(2, date_col, rows, 1);
    date_range.setNotes(date_notes[event_sheet_name]);
    date_range.setBackgrounds(bg_dates[event_sheet_name]);
    var field_range = event_sheet.getRange(2, field_col, rows, 4);
    field_range.setBackgrounds(bg_colors[event_sheet_name]);
  }
    setup_dataval();
    progress_update("REFRESH EVENTS FINISHED");
  } catch (e) {
    var message = Utilities.formatString('This error has automatically been sent to the developers. %s: %s (line %s, file "%s"). Stack: "%s" . While processing %s.',
                                         e.name||'', e.message||'', e.lineNumber||'', e.fileName||'',
                                         e.stack||'', arguments.callee.name||'');
    Logger = startBetterLog();
    Logger.severe(message);
//    var ui = SpreadsheetApp.getUi();
//    var result = ui.alert(
//     'ERROR',
//      message,
//      ui.ButtonSet.OK);
    return "";
  }
}