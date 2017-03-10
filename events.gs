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

function get_needed_fields(event_type){
  var ScoringObject = main_range_object("Scoring");
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