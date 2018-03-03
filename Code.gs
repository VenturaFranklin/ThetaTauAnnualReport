/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
var WORKING = false;
var SILENT = false;
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

function update_property(key, val) {
  var ui = SpreadsheetApp.getUi();
  var keys = SCRIPT_PROP.getKeys();
  keys = keys.map(function(x){ return x.replace('password',"") });
  var result = ui.prompt('What key do you want to update?\nProperties are: ' + keys,
    ui.ButtonSet.OK_CANCEL);
  var key = result.getResponseText();
  var old_val = SCRIPT_PROP.getProperty(key);
  var result = ui.prompt('What value do you want to set key: ' + key + 
                         '\nCurrent value is: ' + old_val,
    ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  var val = result.getResponseText();
  if (button == ui.Button.OK) {
    SCRIPT_PROP.setProperty(key, val);
  }
}

function onOpen(e) {
  SCRIPT_PROP.setProperty("password", "FALSE");
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addSubMenu(SpreadsheetApp.getUi().createMenu("Refresh")
//                  .addItem('Refresh Attendance on Events', 'refresh_attendance')
                  .addItem('Refresh All Scores', 'refresh_scores')
                  .addItem('Refresh All Scores Silent', 'refresh_scores_silent')
                  .addItem('Refresh Events Background Stuff', 'refresh_events')
                  .addItem('Refresh Events Background Stuff Silent', 'refresh_events_silent')
//                  .addItem('Refresh Events to Attendance', 'events_to_att')
                  .addItem('Refresh Members', 'refresh_members')
                  .addItem('Refresh Members Silent', 'refresh_members_silent')
  );
  menu.addSubMenu(SpreadsheetApp.getUi().createMenu("Submit")
                  .addItem('Pledge Forms', 'side_pledge')
                  .addItem('Status Change', 'side_member')
                  .addItem('Submit Item', 'side_submit')
                  .addItem('Update Officers', 'side_officers')
  );
  menu.addItem('Send Survey', 'side_survey');
  menu.addItem('Add event sheet', 'add_event_sheet');
  menu.addItem('SUBMIT ANNUAL REPORT', 'submit_report');
  menu.addItem('Update', 'update');
  menu.addSeparator();
  menu.addSubMenu(SpreadsheetApp.getUi().createMenu("Debugging")
                  .addItem('Create Triggers', 'run_createTriggers')
                  .addItem('Add Missing Member', 'missing_form')
                  .addItem('Update Chapter Name', 'chapter_name')
                  .addItem("RESET", 'RESET')
                  .addItem('Start Logging', 'start_logging')
                  .addItem('SETUP', 'run_install')
                  .addItem('Unlock', 'unlock')
                  .addItem('Update Property', 'update_property')
  );
  menu.addItem('Version', 'version');
  menu.addToUi();
}

function version(){
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert('Chapter Management Tool\n'+
                        'version: 2.2.6 = 43 (google version) = 40 (published version)\n'+
                        'Maintaned and developed by Franklin Ventura Frank.Ventura@thetatau.org\n'+
                        'https://github.com/VenturaFranklin/ThetaTauAnnualReport',
                        ui.ButtonSet.OK);
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
    } else if (status == "All"){
      member_name = shorten(member_name, 15);
      member_list.push(member_name);
    }
  }
  return member_list
}

function date_check(start, end){
//  var start = "2017-01-01";
//  var end = "2016-01-01";
  try{
    var date_start = new Date(start);
    var date_end = new Date(end);
    if (date_end >= date_start){
      return false;
    }
    return true;
  } catch (error) {
    Logger.log("(" + arguments.callee.name + ") " +error);
    return false;
  }
}

function format_date_first(date) {
  //"YYYY-MM-DD" to DD/MM/YYYY with day first of month
//  var date = "2016-07-31";
//  var date = "2017-12-31";
//  var date = "2016-08-20";
//  var date = "2017-06-01";
  try{
    date = new Date(date);
    date.setDate(date.getDate() + 1);
    date.setDate(1);
    var new_date = (date.getMonth()+1) + "/" + date.getDate() + "/" + date.getFullYear()
    return new_date
  } catch (error) {
    Logger.log("(" + arguments.callee.name + ") " +error);
    return "";
  }
}

function format_date(date) {
  //"YYYY-MM-DD" to DD/MM/YYYY 
//  var date = "2016-12-31";
  try{
    var raw = date.split("-");
    var new_date = raw[1] + "/" + raw[2] + "/" + raw[0];
    return new_date
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
    update_20171015();
    if(!check_sheets()){
      Logger.log("(" + arguments.callee.name + ") " +"Check sheets false");
      return;
    }
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
  if (sheet_name.indexOf('Event') >= 0){
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
        'The Score is updated automatically; And Do not edit header.',
        ui.ButtonSet.OK);
    } else {
    update_scores_event(sheet_name, user_row);
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
      var ui = SpreadsheetApp.getUi();
      var result = ui.alert(
        'ERROR',
        'Please do not edit member information here\n'+
        'Member information is changed by notifying the central office',
        ui.ButtonSet.OK);
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

function update(){
  Logger.log("(" + arguments.callee.name + ") " +"update");
  update_20171015_main();
  check_sheets();
}

function update_20171015_main(){
  var MemberObject = main_range_object("Membership");
  var member_sheet = MemberObject.sheet;
  var cols_to_del = [];
  if (MemberObject.header_values.indexOf("Service Hours Fall") >= 0){
    cols_to_del.push(MemberObject.header_values.indexOf("Service Hours Fall"));
  }
  if (MemberObject.header_values.indexOf("Service Hours Spring") >= 0){
    cols_to_del.push(MemberObject.header_values.indexOf("Service Hours Spring"));
  }
  if (MemberObject.header_values.indexOf("Self Service Hrs FA") >= 0){
    col = MemberObject.header_values.indexOf("Self Service Hrs FA");
    member_sheet.getRange(1, +col+1).setValue("2016 FALL Service");
  }
  if (MemberObject.header_values.indexOf("Service Hrs FA") >= 0){
    col = MemberObject.header_values.indexOf("Service Hrs FA");
    member_sheet.getRange(1, +col+1).setValue("2016 FALL Service");
  }
  if (MemberObject.header_values.indexOf("Self Service Hrs SP") >= 0){
    col = MemberObject.header_values.indexOf("Self Service Hrs SP");
    member_sheet.getRange(1, +col+1).setValue("2017 SPRING Service");
  }
  if (MemberObject.header_values.indexOf("Service Hrs SP") >= 0){
    col = MemberObject.header_values.indexOf("Service Hrs SP");
    member_sheet.getRange(1, +col+1).setValue("2017 SPRING Service");
  }
  if (MemberObject.header_values.indexOf("Fall GPA") >= 0){
    col = MemberObject.header_values.indexOf("Fall GPA");
    member_sheet.getRange(1, +col+1).setValue("2016 FALL GPA");
  }
  if (MemberObject.header_values.indexOf("Spring GPA") >= 0){
    col = MemberObject.header_values.indexOf("Spring GPA");
    member_sheet.getRange(1, +col+1).setValue("2017 SPRING GPA");
  }
  for (var i in cols_to_del){
    var col = cols_to_del[i];
    member_sheet.deleteColumn(+col+1);
  }
  var ScoringObject = main_range_object("Scoring");
  var scoring_sheet = ScoringObject.sheet;
  if (ScoringObject.header_values.indexOf("FALL SCORE") >= 0){
    col = ScoringObject.header_values.indexOf("FALL SCORE");
    scoring_sheet.getRange(1, +col+1).setValue("2016 FALL");
  }
  if (ScoringObject.header_values.indexOf("SPRING SCORE") >= 0){
    col = ScoringObject.header_values.indexOf("SPRING SCORE");
    scoring_sheet.getRange(1, +col+1).setValue("2017 SPRING");
  }
  var EventObject = main_range_object("Events");
  var event_sheet = EventObject.sheet;
  if (EventObject.header_values.indexOf("PLEDGE Focus") >= 0){
    col = EventObject.header_values.indexOf("PLEDGE Focus");
    event_sheet.deleteColumn(+col+1);;
  }
  var chapter_info = get_chapter_info();
  var chapter_sheet = chapter_info.sheet;
  if ("Total Pledges" in chapter_info){
    var row = chapter_info["Total Pledges"].row;
    chapter_sheet.getRange(row + 1, 1).setNote("This includes depledged pledges; All pledges given and accepted bid.");
  }
  if ("Graduated Members" in chapter_info){
    var row = chapter_info["Graduated Members"].row;
    chapter_sheet.getRange(row + 1, 1).setNote("This includes pre alumn members.");
  }
  if ("Time of Year" in chapter_info){
    chapter_sheet.deleteRow(4);
  }
  if (!("Physical Address for mailing things:" in chapter_info)){
    chapter_sheet.insertRows(4);
    chapter_sheet.getRange(4, 1).setValue("Physical Address for mailing things:").setWrap(true);
  }
  if (!("Years" in chapter_info)){
    chapter_sheet.insertRows(5);
    chapter_sheet.getRange(5, 1, 1, 5).setValues([["Years", '2016', '2017', '2017', '2018']]);
  }
  if (!("Semesters" in chapter_info)){
    chapter_sheet.insertRows(6);
    chapter_sheet.getRange(6, 1, 1, 5).setValues([["Semesters", 'FALL', 'SPRING',  'FALL', 'SPRING']]);
  }
  if (!("Regent" in chapter_info)){
    chapter_sheet.insertRows(18);
    chapter_sheet.getRange(18, 1).setValue("Regent");
    chapter_sheet.getRange(18, 2).setValue("Emails used when chapter has a single email for each officer role. Set this before submitting officers to the central office.");
  }
  if (!("Vice Regent" in chapter_info)){
    chapter_sheet.insertRows(19);
    chapter_sheet.getRange(19, 1).setValue("Vice Regent");
  }
  if (!("Treasurer" in chapter_info)){
    chapter_sheet.insertRows(20);
    chapter_sheet.getRange(20, 1).setValue("Treasurer");
  }
  if (!("Scribe" in chapter_info)){
    chapter_sheet.insertRows(21);
    chapter_sheet.getRange(21, 1).setValue("Scribe");
  }
  if (!("Corresponding Secretary" in chapter_info)){
    chapter_sheet.insertRows(22);
    chapter_sheet.getRange(22, 1).setValue("Corresponding Secretary");
    chapter_sheet.getRange(18, 2, 5, 2).merge().setWrap(true);
  }
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName('Attendance');
  sheet.insertRows(2);
  sheet.getRange(2, 2, 1, 40).setDataValidation(null).merge().setBackground('red').setValue("This sheet is only for your chapter's use." + 
    "The attendance will no longer be automatically sent to the events sheet. You do not have to use it if you do not want to.").setFontWeight("bold").setHorizontalAlignment('left');
  
}

function update_20171015(){
  var update_test = SCRIPT_PROP.getProperty('20171015');
  if (!update_test){
    update_20171015_main()
    SCRIPT_PROP.setProperty('20171015', true);
  }
}

function get_start_year(){
  var chapter_info = get_chapter_info();
  var start_year = chapter_info['Years'].values[0];
  return start_year
}

function get_stop_year(){
  var chapter_info = get_chapter_info();
  var stop_year = chapter_info['Years'].values[3];
  return stop_year
}

function get_chapter_info(){
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName('Chapter');
  var chapter_info = {};
//   get_column_values(col, range_values)
  var max_row = sheet.getLastRow();
  var max_column = sheet.getLastColumn();
  var range = sheet.getRange(1, 1, max_row, max_column);
  var range_values = range.getValues();
  chapter_info.range = range;
  chapter_info.values = range_values;
  chapter_info.sheet = sheet;
  for (var row in range_values){
    row = parseInt(row);
    var range = sheet.getRange(row+1, 2, 1, max_column-1);
    var row_name = range_values[row][0];
    chapter_info[row_name] = chapter_info[row_name] ? chapter_info[row_name]:{};
    chapter_info[row_name].row = row;
    chapter_info[row_name].range = range;
//     chapter_info[row_name].values = range.getValues()[0];
    chapter_info[row_name].values = range_values[row].slice(1, max_column);
  }
  return chapter_info;
}

function get_year_semesters(){
//   var update_test = SCRIPT_PROP.getProperty('year_semesters');
//   if (!update_test){
    var year_semesters = {};
    var membership_ranges = get_membership_ranges();
    for (var year_semester in membership_ranges){
      year_semesters[year_semester] = null;
    }
//     SCRIPT_PROP.setProperty('year_semesters', JSON.stringify(year_semesters))
//   } else {
//      var year_semesters = JSON.parse(update_test);
//   }
  return year_semesters;
}

function get_membership_ranges(){
  /*
  membership_ranges[4]
    2016 Fall[4]
    2017 Spring[4]
    2017 Fall[4]
    2018 Spring[4]
      Initiated Pledges[2]
      Total Pledges[2]
      Graduated Members[2]
      Active Members[2]
        range
        value[1]
  */
  var chapter_info = get_chapter_info();
  var sheet = chapter_info.sheet;
  var membership_ranges = {};
  try{
  var years = chapter_info['Years'].values;
  } catch (e) {
    update_20171015_main();
    var years = chapter_info['Years'].values;
  }
  var semesters = chapter_info['Semesters'].values;
  var rows = ["Initiated Pledges", "Total Pledges",
              "Graduated Members", "Active Members"];
  for (var i in years){
    var year = years[i];
    var semester = semesters[i].toUpperCase();
    var sm_yr = year + " " + semester;
    membership_ranges[sm_yr] = {};
    for (var j in rows){
      j = parseInt(j);
      var name_row = rows[j];
      membership_ranges[sm_yr][name_row] = {};
      var row = chapter_info[name_row].row;
      var range = sheet.getRange(row+1, +i+2, 1, 1);
      membership_ranges[sm_yr][name_row].range = range;
      var val = chapter_info[name_row].values.slice(+i, +i+1);
      membership_ranges[sm_yr][name_row].value = val;
      if (val < 1){
        range.setNote("Scribe should set this value")
        .setBackground('red');
      } else {
        range.clearNote()
        .setBackground('white');
      }
    }
  }
  return membership_ranges
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
  Logger.log("(" + arguments.callee.name + ") " +"check_sheets");
  var sheet_names = ["Chapter", "Scoring",
                     "Membership", "Submissions", "Dashboard"];
  var chapter_info = get_chapter_info();
  var years = chapter_info['Years'].values;
  var semesters = chapter_info['Semesters'].values;
  var ss = get_active_spreadsheet();
  var event_sheets = find_all_event_sheets(ss);
  for (var sheet_name in event_sheets){
    sheet_names.push(sheet_name);
  }
  for (var i in sheet_names){
    var sheet_name = sheet_names[i];
    var sheet = ss.getSheetByName(sheet_name);
    if (!sheet){
      var message = Utilities.formatString('You are missing a sheet!\nWhere is sheet name: %s?\nPlease rename the sheet back to its original name:\n%s',
                                           sheet_name||'', sheet_name||'');
      Logger = startBetterLog();
      Logger.severe(message);
      var ui = SpreadsheetApp.getUi();
      var result = ui.alert(
        'ERROR',
        message,
        ui.ButtonSet.OK);
      set_refresh("check", true);
      return false;
    }
    var col_names = [];
    switch (sheet_name){
      case "Scoring":
        col_names = ["ACTIVITY", "Long Description", "Type", "Points"];
        for (var i in years){
          col_names.push(years[i] + " " + semesters[i]);
        }
        var col_names2 = ["CHAPTER TOTAL", "Top 10 Chapters", "EVENTS/ YEAR",
                          "Max/ Semester", "How points are calculated", "Short Name",
                          "Score Type", "Base Points", "Attendance Multiplier",
                          "Member Add", "Special", "Event Fields"];
        col_names.push.apply(col_names, col_names2);
        break;
      case "Membership":
        col_names = ["Member Name", "Last Update", "First Name", "Last Name", "Badge Number",
                     "Chapter Status", "Status Start", "Status End", "Chapter Role",
                     "Current Major", "School Status", "Phone Number", "Email Address"];
        for (var i in years){
          col_names.push(years[i]  + " " + semesters[i] + " Service");
        }
        for (var i in years){
          col_names.push(years[i]  + " " + semesters[i] + " GPA");
        }
        var col_names2 = ["Professional/ Technical Orgs", "Officer (Pro/Tech)", "Honor Orgs",
                          "Officer (Honor)", "Other Orgs", "Officer (Other)"];
        col_names.push.apply(col_names, col_names2);
        break;
      case "Submissions":
        col_names = ["Date", "File Name", "Type", "Score", "Location of Upload"];
        break;
    }
    if (sheet_name.indexOf("Event") >= 0){
      col_names = ["Event Name", "Date", "Type", "Score", "# Members", "# Pledges",
                   "# Alumni", "Description", "Event Hours", "# Non- Members",
                   "STEM?", "HOST", "MILES"];
    }
    if (!check_cols(sheet, col_names)){
      return false;
    };
    }
  return true;
  } catch (e) {
    var message = Utilities.formatString('This error has automatically been sent to the developers. %s: %s (line %s, file "%s"). Stack: "%s" . While processing %s.',
                                         e.name||'', e.message||'', e.lineNumber||'', e.fileName||'',
                                         e.stack||'', arguments.callee.name||'');
    Logger = startBetterLog();
    Logger.severe(message);
  }
}

function check_cols(sheet, col_names){
  Logger.log("(" + arguments.callee.name + ") " +"check_cols");
  var max_column = sheet.getLastColumn();
  var header_range = sheet.getRange(1, 1, 1, max_column);
  var header_values = header_range.getValues();
  for (var ind in col_names){
    var col_name = col_names[ind];
    if (header_values[0].indexOf(col_name) < 0){
      sheet.insertColumns(+ind+1);
      var new_range = sheet.getRange(1, +ind + 1);
      new_range.setValue(col_name);
      var sheet_name = sheet.getSheetName();
      var message = Utilities.formatString('You are missing a column name!\nWhere is column name: %s in sheet %s?\nPlease rename the column back to its original name.',
                                           col_name||'', sheet_name||'');
      Logger = startBetterLog();
      Logger.severe(message);
//       var ui = SpreadsheetApp.getUi();
//       var result = ui.alert(
//         'ERROR',
//         message,
//         ui.ButtonSet.OK);
//       return false;
    }
  }
  return true;
}

function refresh_members_silent(){
  SILENT = true;
  refresh_members();
  SILENT = false;
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

function check_date_year_semester(date){
  try{
    var year_semesters = get_year_semesters();
    var semester = get_semester(date);
    var year = date.getFullYear();
    var year_semester = year + " " + semester;
    if (!(year_semester in year_semesters)){
      return false;
    }
    return true;
  } catch (e) {
    return false;
  }
}

function check_date(date){
  try{
    var cur_date = new Date();
    var cur_year = cur_date.getFullYear();
    var event_year = date.getFullYear();
    var delta = cur_year - event_year;
    if (delta > 2 || delta < -2){
      return false;
    }
    return true;
  } catch (e) {
    return false;
  }
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
  var sheets = new Array(); 
  switch (sheetName){
    case "Membership":
    case "REGIONS":
    case "Jewelry":
    case "MAIN":
      if (short_header == undefined){
      var short_header = "Member Name";
      }
      var sort_val = short_header;
      sheets[sheetName] = ss.getSheetByName(sheetName);
      break;
    case "ScoreInfo":
    case "Scoring":
      var short_header = "Short Name";
      var sort_val = short_header;
      sheets[sheetName] = ss.getSheetByName(sheetName);
      break;
    case "Events":
      var short_header = "Event Name";
      var sort_val = "Date";
      sheets = find_all_event_sheets(ss);
      break;
    case "Attendance":
      var short_header = "Event Name";
      var sort_val = "Date";
      sheets[sheetName] = ss.getSheetByName(sheetName);
      break;
    case "Submissions":
      var short_header = "File Name";
      var sort_val = "Date";
      sheets[sheetName] = ss.getSheetByName(sheetName);
      break;
  }
 var myObject = new Array();
 myObject["object_header"] = new Array();
 myObject["original_names"] = new Array();
 for (var sheetName in sheets){
  var sheet = sheets[sheetName];
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
   temp.sheet = sheet;
   temp.sheet_name = sheetName;
    myObject[short_name] = temp;
    myObject["object_header"].push(short_name);
    myObject["object_count"] = myObject["object_count"] ? myObject["object_count"]+1 : 1;
  }
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
  var sheet = ss.getSheetByName("Events");
  var range = sheet.getRange(2, 3, 1, 1);
  var value = range.getValue();
  Logger.log("TEST");
  Logger.log(value);
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
