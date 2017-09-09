/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
var WORKING = false;
var SCRIPT_PROP = PropertiesService.getDocumentProperties();
var betterLogStarted = false;
logging_check();

function logging_check(){
  try {
    var logger = SCRIPT_PROP.getProperty("logger") == 'true';
    Logger.log("Checking logging: " + logger);
    if (logger){
      var current = new Date();
      var stop = new Date(SCRIPT_PROP.getProperty("logger_stop"));
      if (current < stop) {
        startBetterLog();
      } else {
        SCRIPT_PROP.setProperty("logger", false);
      }
    }
  } catch (e) {
    Logger.log("(" + arguments.callee.name + ") " +e);
  }
}

function start_logging() {
  var newDateObj = new Date();
  newDateObj.setTime(newDateObj.getTime() + (30 * 60 * 1000));
  SCRIPT_PROP.setProperty("logger_stop", newDateObj);
  SCRIPT_PROP.setProperty("logger", true);
  startBetterLog();
}

function startBetterLog() {
  if (!betterLogStarted) {
    var chapter_name = get_chapter_name();
    Logger.log("Starting Better Logger for chapter: " + chapter_name);
    Logger = BetterLog.useSpreadsheet('1mo5t1Uu7zmP9t7w2hL1mWrdba4CtgD_Q9ImbAKjGZyM',
                                      chapter_name);
    betterLogStarted = true;
  }
  return Logger;
}

function clientLog() {
  var Logger = startBetterLog();
  var args = Array.slice(arguments);    // Convert arguments to array
  var func = args.shift();              // Remove first argument, Logger method
//  if (!Logger.hasOwnProperty(func))     // Validate Logger method
//    throw new Error( "Unknown Logger method: " + func );
  args[0] = "CLIENT "+args[0];          // Prepend CLIENT tag
  Logger[func].apply(null,args);        // Pass all arguments to Logger method
}

function get_active_spreadsheet() {
  var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
//  var doc = SpreadsheetApp.getActiveSpreadsheet();
  return doc
}

function onOpen(e) {
  SCRIPT_PROP.setProperty("password", "FALSE");
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addSubMenu(SpreadsheetApp.getUi().createMenu("Refresh")
//                  .addItem('Refresh Attendance on Events', 'refresh_attendance')
                  .addItem('Refresh All Scores', 'refresh_scores')
                  .addItem('Refresh Events Background Stuff', 'refresh_events')
//                  .addItem('Refresh Events to Attendance', 'events_to_att')
                  .addItem('Refresh Members', 'refresh_members')
  );
  menu.addSubMenu(SpreadsheetApp.getUi().createMenu("Submit")
                  .addItem('Pledge Forms', 'side_pledge')
                  .addItem('Status Change', 'side_member')
                  .addItem('Submit Item', 'side_submit')
                  .addItem('Update Officers', 'side_officers')
                  .addItem('SUBMIT ANNUAL REPORT', 'submit_report')
  );
  menu.addItem('Send Survey', 'send_survey');
  menu.addItem('SYNC', 'sync');
  menu.addSeparator();
  menu.addSubMenu(SpreadsheetApp.getUi().createMenu("Debugging")
                  .addItem('Create Triggers', 'run_createTriggers')
                  .addItem('Add Missing Member', 'missing_form')
                  .addItem("RESET", 'RESET')
                  .addItem('SETUP', 'run_install')
                  .addItem('Start Logging', 'start_logging')
                  .addItem('Unlock', 'unlock')
  );
  menu.addToUi();
//  check_sheets();
}

function TEST(){
//  SCRIPT_PROP.setProperty("logger", false)
//  SCRIPT_PROP.setProperty('key', '1wBICuD_CvSm3BonA_OZg-sOTRJylVWVbLYi9nr8vn8Q');
//  SCRIPT_PROP.setProperty('chapter', 'Chi');
//  SCRIPT_PROP.setProperty('director', 'werd@thetatau.org');
//  SCRIPT_PROP.setProperty('email', 'venturafranklin@gmail.com');
//  SCRIPT_PROP.setProperty("region", "Western");
//  SCRIPT_PROP.setProperty("dash", "10ebwK7tTKgveVCEOpRle2S17d4UjwmsoXXCPFvC9A-A");
//  var dash_id = SCRIPT_PROP.getProperty("dash");
//  var dash_file = SpreadsheetApp.openById(dash_id);
//  Logger.log("(" + arguments.callee.name + ") " +SCRIPT_PROP.getProperty('key'));
//  Logger.log("(" + arguments.callee.name + ") " +SCRIPT_PROP.getProperty('chapter'));
//  Logger.log("(" + arguments.callee.name + ") " +SCRIPT_PROP.getProperty('director'));
//  Logger.log("(" + arguments.callee.name + ") " +SCRIPT_PROP.getProperty('email'));
//  Logger.log("(" + arguments.callee.name + ") " +SCRIPT_PROP.getProperty("region"));
//  Logger.log("(" + arguments.callee.name + ") " +SCRIPT_PROP.getProperty("folder"));
//  var ui = SpreadsheetApp.getUi();
//  ui.alert('SETUP COMPLETE!\n'+
//           'Next steps:\n'+
//           '- Fill out Chapter Sheet\n'+
//           '- Verify Membership\n'+
//           '- Add Events & Attendance\n\n'+
//           'Do not edit gray or black cells\n'+
//           'Submit forms in menu "Add-ons-->ThetaTauReports"');
//  var ss = get_active_spreadsheet();
//  Logger.log("(" + arguments.callee.name + ") " +range.getValues());
}

function unlock() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('What is the password?',
      ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    SCRIPT_PROP.setProperty("password", text);
  }  
}

function getAllIndexes(arr, val) {
    var indexes = [], i;
    for(i = 0; i < arr.length; i++)
        if (arr[i] === val)
            indexes.push(i);
    return indexes;
}

function get_member_list(status){
//  var status = "Student";
//  var status = "Pledge";
  var MemberObject = main_range_object("Membership");
  var member_list = [];
  for(var i = 0; i< MemberObject.object_count; i++) {
    var member_name = MemberObject.object_header[i];
    var member_status = MemberObject[member_name]["Chapter Status"][0];
    if (member_status == "Away" || member_status == "Shiny" || member_status == "Alumn"){
      member_status = "Student";
    }
    if (member_status == status){
      member_name = shorten(member_name, 15);
      member_list.push(member_name);
    }
  }
  return member_list
}

function format_date(date) {
  //"YYYY-MM-DD" to DD/MM/YYYY 
  try{
    var raw = date.split("-");
    return raw[1] + "/" + raw[2] + "/" + raw[0]
  } catch (error) {
    Logger.log("(" + arguments.callee.name + ") " +error);
    return "";
  }
}

function find_member_shortname(MemberObject, member_name_raw){
  try {
    var member_name = member_name_raw.split("...")[0]
  } catch (error) {
    Logger.log("(" + arguments.callee.name + ") ");
    Logger.log(error);
    return;
  }
  for (var full_name in MemberObject){
    if (~full_name.indexOf(member_name)){
      return MemberObject[full_name]
    }
  }
}

function shorten(long_string, max_len, ellipse){
  if (ellipse !== undefined){
    ellipse="";
  } else {
    ellipse="...";
  }
  return (long_string.length > max_len) ? long_string.substr(0,max_len-1)+ellipse : long_string.substr(0,long_string.length);
}


function cleanArray(actual, short_length) {
  short_length = short_length || 15;
  var newArray = new Array();
  for (var i = 0; i < actual.length; i++) {
    if (actual[i]) {
      var newactual = actual[i].toString();
      newactual = shorten(newactual, short_length);
      newArray.push(newactual);
    }
  }
  return newArray;
}

function get_type_list(score_type, desc){
//  var score_type = "Submit";
//  var score_type = "Events";
  var ScoringObject = main_range_object("Scoring");
  var newArray = new Array();
  var descArray = {};
  for (var type_ind = 0;  type_ind < parseInt(ScoringObject.object_count); type_ind++){
    var type_name = ScoringObject.object_header[type_ind];
    var description = ScoringObject[type_name]["Long Description"][0];
    var thistype = ScoringObject[type_name]["Score Type"][0];
    if (~thistype.indexOf(score_type)){
      newArray.push(type_name);
      descArray[type_name] = description;
    }
  }
//  newArray.sort();
  Logger.log("(" + arguments.callee.name + ") ");
  Logger.log(newArray);
  if (!desc){
    return newArray;
  } else {
    return {
      type_list: newArray,
      type_desc: descArray
    }
  }
}

function get_ind_list(type){
//  var type = "Brotherhood";
//  var type = "Operate";
//  var type = "ProDev";
//  var type = "Service";
  Logger.log("(" + arguments.callee.name + ") " +type);
  var ScoringObject = main_range_object("Scoring");
  var newArray = new Array();
  for (var type_ind = 0;  type_ind < parseInt(ScoringObject.object_count); type_ind++){
    var type_name = ScoringObject.object_header[type_ind];
    var thistype = ScoringObject[type_name]["Type"][0];
    var thisind = ScoringObject[type_name].object_row;
    if (~thistype.indexOf(type)){
      newArray.push(thisind);
    }
  }
  newArray.sort();
  Logger.log("(" + arguments.callee.name + ") ");
  Logger.log(newArray);
  return newArray;
}

function getList(RangeName) {
  //' MemberNamesOnly
//  var RangeName = 'MemberNamesOnly'
//  var RangeName = 'EventTypes';submit_col_max
  Logger.log("(" + arguments.callee.name + ") " +'Called getList, RangeName: ' + RangeName);
  var ss = get_active_spreadsheet();
  var events = ss
      .getRangeByName(RangeName)
      .getValues()
  var event_list = [].concat.apply([], events);
  var event_list = cleanArray(event_list);
  event_list.sort();
  Logger.log("(" + arguments.callee.name + ") " +event_list);
  return event_list;
}

//function onChange(e){
////  show_att_sheet_alert();
//  Logger.log("(" + arguments.callee.name + ") " +arguments.callee.name + "onChange");
//  Logger.log("(" + arguments.callee.name + ") " +e);
//  _onEdit(e);
////  var ss = get_active_spreadsheet();
////  var sheet = ss.getActiveSheet();
////  var sheet_name = sheet.getName();
////  if (sheet_name == "Events"){
////    if (e.changeType == "INSERT_ROW"){
////      Logger.log("(" + arguments.callee.name + ") " +"EVENTS ROW ADDED");
////    }else if (e.changeType == "REMOVE_ROW"){
////      Logger.log("(" + arguments.callee.name + ") " +"EVENTS ROW REMOVED");
////    }
////  } 
//}

function reset_range(range, user_old_value){
//  return;
  var this_password = SCRIPT_PROP.getProperty("password");
  if (this_password == password){
    return;
  }
  var user_old_value = (user_old_value != undefined) ? user_old_value:"";
  range.setValue(user_old_value);
}

function _onEdit(e){
  try{
  Logger.log("(" + arguments.callee.name + ") " +"onEDIT");
  Logger.log(e)
  Logger.log("(" + arguments.callee.name + ") " +"onEdit, authMode: " + e.authMode);
  Logger.log("(" + arguments.callee.name + ") " +"onEdit, user: " + e.user);
  Logger.log("(" + arguments.callee.name + ") " +"onEdit, source: " + e.source);
  Logger.log("(" + arguments.callee.name + ") " +"onEdit, range: " + e.range);
  Logger.log("(" + arguments.callee.name + ") " +"onEdit, value: " + e.value);
  var sheet = e.range.getSheet();
  var sheet_name = sheet.getName();
  var user_range = e.range
  var user_row = user_range.getRow();
  var user_col = user_range.getColumn();
  var user_old_value = e.oldValue
  Logger.log("(" + arguments.callee.name + ") " +"Row: " + user_row + " Col: " + user_col);
  var this_password = SCRIPT_PROP.getProperty("password");
  if (sheet_name == "Events"){
    Logger.log("(" + arguments.callee.name + ") " +"EVENTS CHANGED");
    if (user_row == 1 || user_col == 4 
//        || user_col == 5 || user_col == 6
       ){
      reset_range(user_range, user_old_value)
      if (this_password == password){
        return;
      }
      var ui = SpreadsheetApp.getUi();
      var result = ui.alert(
        'ERROR',
        'The Score is updated automatically',
        ui.ButtonSet.OK);
    } else {
    update_scores_event(user_row);
    }
//    show_event_sheet_alert();
//    align_event_attendance();
//  } else if (sheet_name == "Attendance"){
//    if (user_row == 1 || user_col < 3){
//      reset_range(user_range, user_old_value);
//      if (this_password == password){
//        return;
//      }
//      show_att_sheet_alert();
//    } else {
//      var attendance = range_object(sheet, user_row);
//      var header = attendance.object_header;
//      var clean_header = cleanArray(header, 50);
//      if (clean_header.length == header.length){
//        update_attendance(attendance);
//        update_scores_event(attendance);
//      } else {
//        return;
//      }
//    }
  } else if (sheet_name == "Scoring") {
    reset_range(user_range, user_old_value)
    if (this_password == password){
      return;
    }
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
     'ERROR',
     'Please do not edit the Scoring Sheet',
      ui.ButtonSet.OK);
  } else if (sheet_name == "Submissions") {
    reset_range(user_range, user_old_value)
    if (this_password == password){
      return;
    }
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
     'ERROR',
     'Please do not edit the Submissions Sheet\n'+
     'Please use the submissions sidebar',
      ui.ButtonSet.OK);
    side_submit();
  } else if (sheet_name == "Dashboard") {
    reset_range(user_range, user_old_value)
    if (this_password == password){
      return;
    }
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
     'ERROR',
     'Please do not edit the Dashboard Sheet',
      ui.ButtonSet.OK);
  }else if (sheet_name == "Membership") {
    Logger.log("(" + arguments.callee.name + ") " +"MEMBER CHANGED");
    if (user_col > 12){
      update_scores_org_gpa_serv();
    } else {
      reset_range(user_range, user_old_value)
      if (this_password == password){
        return;
      }
//      var ui = SpreadsheetApp.getUi();
//      var result = ui.alert(
//        'ERROR',
//        'Please do not edit member information here\n'+
//        'Member information is changed by notifying the central office',
//        ui.ButtonSet.OK);
    }
  } else if (sheet_name == "Chapter") {
    Logger.log("(" + arguments.callee.name + ") " +"CHAPTER CHANGED");
    update_score_member_pledge()
  }
  } catch (e) {
    var message = Utilities.formatString('This error has automatically been sent to the developers. %s: %s (line %s, file "%s"). Stack: "%s" . While processing %s.',
                                         e.name||'', e.message||'', e.lineNumber||'', e.fileName||'', e.stack||'', sheet_name||'');
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

function get_membership_ranges(){
  var ss = get_active_spreadsheet();
  var init_sp_range = ss.getRangeByName("INIT_SP");
  var init_fa_range = ss.getRangeByName("INIT_FA");
  var pledge_sp_range = ss.getRangeByName("PLEDGE_SP");
  var pledge_fa_range = ss.getRangeByName("PLEDGE_FA");
  var grad_sp_range = ss.getRangeByName("GRAD_SP");
  var grad_fa_range = ss.getRangeByName("GRAD_FA");
  var act_sp_range = ss.getRangeByName("ACT_SP");
  var act_fa_range = ss.getRangeByName("ACT_FA");
  return {init_sp_range: init_sp_range,
          init_fa_range: init_fa_range,
          pledge_sp_range: pledge_sp_range,
          pledge_fa_range: pledge_fa_range,
          grad_sp_range: grad_sp_range,
          grad_fa_range: grad_fa_range,
          act_sp_range: act_sp_range,
          act_fa_range: act_fa_range
  }
}

function get_column_values(col, range_values){
	var newArray = new Array();
	for(var i=0; i<range_values.length; i++){
		newArray.push(range_values[i][col]);
     }
	return newArray;
}

function check_sheets(){
  try {
  var sheet_names = ["Events", "Chapter", "Scoring",
                     "Membership", "Submissions", "Dashboard"];
  var ss = get_active_spreadsheet();
  for (var i in sheet_names){
    var sheetName = sheet_names[i];
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet){
      var message = Utilities.formatString('You are missing a sheet!\nWhere is sheet name: %s?\nPlease rename the sheet back to its original name:\n%s',
                                           sheetName||'', sheetName||'');
      Logger = startBetterLog();
      Logger.severe(message);
      var ui = SpreadsheetApp.getUi();
      var result = ui.alert(
        'ERROR',
        message,
        ui.ButtonSet.OK);
    }
  }
  } catch (e) {
    Logger.log("(" + arguments.callee.name + ") " +e);
  }
}

function refresh_members(){
  get_chapter_members();
}

function check_refresh(refresh, property_name){
  // if property_name does not exist will be caught by end_time return true
  if (refresh == true){
    return true;
  }
  var end_time = new Date(SCRIPT_PROP.getProperty(property_name+"_refresh"));
  var current = new Date();
  if (current > end_time){
    return true; 
  }
  return SCRIPT_PROP.getProperty(property_name);
}

function set_refresh(property_name, property_value){
  SCRIPT_PROP.setProperty(property_name, property_value);
  var newDateObj = new Date();
  newDateObj.setTime(newDateObj.getTime() + (30 * 60 * 1000)); //30 minute delay
  SCRIPT_PROP.setProperty(property_name+"_refresh", newDateObj);
}

function get_dues_years(){
  var today = new Date();
  var month = today.getMonth() + 1;
  if (month > 8){
    // if current date is after 8 month then 11/1 this year 3/15 next year
    var fall_year = today.getFullYear();
    var spring_year = fall_year + 1;
  } else {
    // else 11/1 last year and 3/15 this year
    var spring_year = today.getFullYear();
    var fall_year = spring_year - 1;
  }
  return {
      fall_date: new Date(fall_year, 11-1, 1),
      spring_date: new Date(spring_year, 3-1, 15)
  };
}

function get_total_members(refresh){
  
  var refresh = check_refresh(refresh, "get_total_members");
  if (refresh == true) {
    var MemberObject = main_range_object("Membership");
    var counts = {};
    counts["FALL"] = {};
    counts["SPRING"] = {};
    for(var i = 0; i< MemberObject.object_count; i++) {
      var member_name = MemberObject.object_header[i];
      var member_object = MemberObject[member_name];
      // Need to take into account semesters
      var due_dates = get_dues_years();
      var fall_status = member_status_semester(member_object, due_dates.fall_date);
      var spring_status = member_status_semester(member_object, due_dates.spring_date);
      counts["FALL"][fall_status] = counts["FALL"][fall_status] ? counts["FALL"][fall_status] + 1 : 1;
      counts["SPRING"][spring_status] = counts["SPRING"][spring_status] ? counts["SPRING"][spring_status] + 1 : 1;
    }
    Logger.log("(" + arguments.callee.name + ") ");
    set_refresh("get_total_members", JSON.stringify(counts));
  } else {
    var counts = JSON.parse(refresh);
  }
  Logger.log(counts);
  return counts;
}

function get_ind_from_string(str, range_values){
  var range_values = range_values[0]
  for (val_ind in range_values){
    if (range_values[val_ind] == str){
      return +val_ind+1
    }
  }
}

function main_range_object(sheetName, short_header, ss){
//  var sheetName = "Membership"
//  var sheetName = "Scoring"
//  var sheetName = "Events"
//  var sheetName = "Submissions";
//  var sheetName = "Attendance";
  if (!ss){
    var ss = get_active_spreadsheet();
  }
  var sheet = ss.getSheetByName(sheetName);
  switch (sheetName){
    case "Membership":
    case "REGIONS":
    case "Jewelry":
    case "MAIN":
      if (short_header == undefined){
      var short_header = "Member Name";
      }
      var sort_val = short_header;
      break;
    case "Scoring":
      var short_header = "Short Name";
      var sort_val = short_header;
      break;
    case "Events":
      var short_header = "Event Name";
      var sort_val = "Date";
      break;
    case "Attendance":
      var short_header = "Event Name";
      var sort_val = "Date";
      break;
    case "Submissions":
      var short_header = "File Name";
      var sort_val = "Date";
      break;
  }
  var max_row = sheet.getLastRow()-1;
  Logger.log("(" + arguments.callee.name + ") " +"MAX_"+sheetName+": "+max_row);
  var max_column = sheet.getLastColumn();
  var header_range = sheet.getRange(1, 1, 1, max_column);
  var header_values = header_range.getValues();
  var short_names_ind = get_ind_from_string(short_header, header_values);
  var sort_ind = get_ind_from_string(sort_val, header_values);
  if (max_row > 0){
    var full_data_range = sheet.getRange(2, 1, max_row, max_column);
    var sorted_range = full_data_range.sort({column: sort_ind, ascending: true});
    var full_data_values = sorted_range.getValues();
    var short_names_range = sheet.getRange(2, short_names_ind, max_row, 1);
    var short_names = short_names_range.getValues();
    short_names = [].concat.apply([], short_names);
    short_names = cleanArray(short_names, 100);
  } else {
    short_names = [];
  }
  var myObject = new Array();
  myObject["object_header"] = new Array();
  myObject["original_names"] = new Array();
  myObject["header_values"] = header_values[0];
  myObject["sheet"] = sheet;
  myObject["object_count"] = 0;
  for (var val in short_names){
//    short_names.forEach(function (item) {
//      var test = item;
//      console.log(item);
//     Logger.log("(" + arguments.callee.name + ") " +item);
//    });
    var short_name_ind = parseInt(val);
    var short_name = short_names[short_name_ind];
    myObject["original_names"].push(short_name);
    var range_values = full_data_values[short_name_ind]
    var temp = range_object_fromValues(header_values[0], range_values, short_name_ind + 2);
    if (sheetName == "Events" || sheetName == "Attendance" || sheetName == "Submissions"){
      // This prevents event duplicates
      short_name = short_name+temp["Date"][0];
    }
    myObject[short_name] = temp;
    myObject["object_header"].push(short_name);
    myObject["object_count"] = myObject["object_count"] ? myObject["object_count"]+1 : 1;
  }
  return myObject
}

function range_object(sheet, range_row){
//  var sheet = "Attendance";
//  var range_row = 3;
  var ss = get_active_spreadsheet();
  if (typeof sheet === "string"){
    var sheet = ss.getSheetByName(sheet);
  }
  var max_column = sheet.getLastColumn()
  var range = sheet.getRange(range_row, 1, 1, max_column);
  var range_values = range.getValues()
  Logger.log("(" + arguments.callee.name + ") " +range_values)
  var header_range = sheet.getRange(1, 1, 1, max_column);
  var header_values = header_range.getValues();
  Logger.log("(" + arguments.callee.name + ") " +header_values)
  var myObject = new Array();
//  myObject["range"] = range; TODO
  myObject["sheet"] = sheet;
  myObject["object_header"] = new Array();
  myObject["object_row"] = range_row;
  for (header in header_values[0]){
    var header_ind = parseInt(header)
    var header_name = header_values[0][header_ind]
    myObject[header_name] = [range_values[0][header_ind], header_ind + 1];
    myObject["object_count"] = myObject["object_count"] ? myObject["object_count"]+1 : 1;
    myObject["object_header"].push(header_name);
  }
  return myObject
}

function range_object_fromValues(header_values, range_values, range_row){
  var myObject = new Array();
  myObject["object_header"] = new Array();
  myObject["object_row"] = range_row;
  for (header in header_values){
    var header_ind = parseInt(header)
    var header_name = header_values[header_ind]
    myObject[header_name] = [range_values[header_ind], header_ind + 1];
    myObject["object_count"] = myObject["object_count"] ? myObject["object_count"]+1 : 1;
    myObject["object_header"].push(header_name);
  }
  return myObject
}

function test_onEdit() {
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Attendance");
  var range = sheet.getRange(2, 3, 1, 1);
  var value = range.getValue();
  _onEdit({
    user : Session.getActiveUser().getEmail(),
    source : ss,
    range : range, //ss.getActiveCell(),
    value : value, //ss.getActiveCell().getValue(),
    authMode : "LIMITED"
  });
//  var ui = SpreadsheetApp.getUi();
//  var result = ui.alert(
//     'ERROR',
//     'Value: '+
//      value,
//      ui.ButtonSet.OK);
}

function get_sheet_data(SheetName) {
//  var SheetName="Events"
//  var SheetName="Attendance"
  Logger.log("(" + arguments.callee.name + ") " +SheetName)
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName(SheetName);
  if (sheet == null) { return; }
  var max_row = sheet.getLastRow();
  var max_row = (max_row != 0) ? max_row:1;
  var max_column = sheet.getLastColumn()
  var range = sheet.getRange(2, 1, max_row, max_column);
  var header_range = sheet.getRange(1, 1, 1, max_column);
  var header_values = header_range.getValues();
  //    Logger.log("(" + arguments.callee.name + ") " +header_values);
  var date_index = header_values[0].indexOf("Date");
  Logger.log("(" + arguments.callee.name + ") " +"date index: " + date_index);
  var name_index =header_values[0].indexOf("Event Name");
  Logger.log("(" + arguments.callee.name + ") " +"name index: " + name_index);
  var range_values = range.getValues();
  var name_date = [];
  var names = [];
  var dates = [];
  for (var value in range_values){
    var name = range_values[value][name_index];
    var date = range_values[value][date_index];
    names.push(name);
    dates.push(date);
    name_date.push(name+date);
  }
  return {sheet: sheet,
          range: range,
          range_values: range_values,
          max_row: max_row,
          header: header_values,
          date_index: date_index,
          name_index: name_index,
          max_column: max_column,
          names: names,
          dates: dates,
          name_date: name_date
         }
}

function sleep(milliseconds) {
  var start = new Date().getTime();
  for (var i = 0; i < 1e7; i++) {
    if ((new Date().getTime() - start) > milliseconds){
      break;
    }
  }
}
