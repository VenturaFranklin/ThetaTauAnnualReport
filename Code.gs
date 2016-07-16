/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Update Members', 'getChapterMembers');
  menu.addItem('Event Functions', 'showaddEvent');
  menu.addItem('Update Officers', 'officerSidebar');
  menu.addItem('Submit Item', 'submitSidebar');
  menu.addItem('Create Triggers', 'createTriggers');
  menu.addToUi();
}

function createTriggers() {
  var sheet = SpreadsheetApp.getActive();
  ScriptApp.newTrigger("onChange")
  .forSpreadsheet(sheet)
  .onChange()
  .create();
}

function testEvents() {
  Logger.log(HtmlService
      .createTemplateFromFile('Officers')
      .getCodeWithComments());
}

function officerSidebar() {
  var template = HtmlService
      .createTemplateFromFile('Officers');

  var htmlOutput = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Update Officers')
      .setWidth(500);

  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(htmlOutput);
}

function submitSidebar() {
   var template = HtmlService
      .createTemplateFromFile('SubmitForm');
  
  template.submissions = get_type_list('Submit');
  
  var htmlOutput = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Submit Item')
      .setWidth(500);
  
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(htmlOutput);
}

function showaddEvent() {
  var html = addEvent()
  html.setTitle('Event Functions')
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);
}

function addEvent() {
  Logger.log('Called addEvent');
  var t = HtmlService.createTemplateFromFile('Events');
  t.events = get_type_list("Events");
  return t.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function uploadFiles(form) {
  
  try {
//    var folder_name = "Student Files";
    var folder = DriveApp.getFolderById('0BwvK5gYQ6D4nWVhUVlo4dUhYV0E');
//    var folder, folders = DriveApp.getFoldersByName(folder_name);
    
//    if (folders.hasNext()) {
//      folder = folders.next();
//    } else {
//      folder = DriveApp.createFolder(folder_name);
//    }
    
    var blob = form.myFile;
    Logger.log("fileBlob Name: " + blob.getName())
    Logger.log("fileBlob type: " + blob.getContentType())
    Logger.log('fileBlob: ' + blob);
    
    var file = folder.createFile(blob);    
//    file.setDescription("Uploaded by " + form.myName);
    var template = HtmlService.createTemplateFromFile('SubmitFormResponse');
    var file_url = template.fileUrl = file.getUrl();
    var submission_date = template.date = new Date();
    var submission_type = template.type = form.submissions;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Submissions");
    SpreadsheetApp.setActiveSheet(sheet);
    var max_column = sheet.getLastColumn();
    var max_row = sheet.getLastRow();
    var submit_range = sheet.getRange(max_row + 1, 1, 1, max_column);
    var file_name = template.name = file.getName();
    submit_range.setValues([[submission_date, file_name, submission_type, 10, file_url]])
    
    return template.evaluate().getContent();
  } catch (error) {
    return error.toString();
  }
}

function shorten(long_string, max_len){
  return (long_string.length > max_len) ? long_string.substr(0,max_len-1)+'...' : long_string.substr(0,long_string.length);
}


function cleanArray(actual, short_length) {
  short_length = short_length || 15;
  var newArray = new Array();
  for (var i = 0; i < actual.length; i++) {
    if (actual[i]) {
      var newactual = actual[i]
      newactual = shorten(newactual, short_length)
      newArray.push(newactual);
    }
  }
  return newArray;
}

function get_type_list(score_type){
//  var score_type = "Submit";
//  var score_type = "Events";
  var ScoringObject = main_range_object("Scoring");
  var newArray = new Array();
  for (var type_ind = 0;  type_ind < parseInt(ScoringObject.object_count); type_ind++){
    var type_name = ScoringObject.object_header[type_ind];
    var thistype = ScoringObject[type_name]["Score Type"][0];
    if (~thistype.indexOf(score_type)){
      newArray.push(type_name);
    }
  }
  newArray.sort();
  Logger.log(newArray);
  return newArray;
}

function getList(RangeName) {
  //' MemberNamesOnly
//  var RangeName = 'MemberNamesOnly'
//  var RangeName = 'EventTypes';
  Logger.log('Called getList, RangeName: ' + RangeName);
  var events = SpreadsheetApp
      .getActiveSpreadsheet()//      .openById('10avD_q_RiDwUDuJ8Nw36vWixeJzmPATjW73DnnHRNJo')
      .getRangeByName(RangeName)
      .getValues()
  var event_list = [].concat.apply([], events);
  var event_list = cleanArray(event_list);
  event_list.sort();
  Logger.log(event_list);
  return event_list;
}

//function onChange(e){
//  Logger.log("onChange")
//  Logger.log(e)
//  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
//  var sheet_name = sheet.getName();
//  if (sheet_name == "Events"){
//    if (e.changeType == "INSERT_ROW"){
//      Logger.log("EVENTS ROW ADDED");
//    }else if (e.changeType == "REMOVE_ROW"){
//      Logger.log("EVENTS ROW REMOVED");
//    }
//  } 
//}

function onEdit(e){
  Logger.log("onEDIT" + e)
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheet_name = sheet.getName();
  var user_event_range = e.range
  var user_row = user_event_range.getRow();
  var user_col = user_event_range.getColumn();
  var user_old_value = e.oldValue
  Logger.log("Row: " + user_row + " Col: " + user_col);
  if (sheet_name == "Events"){
    Logger.log("EVENTS CHANGED");
    update_scores(user_row);
//    show_event_sheet_alert();
//    align_event_attendance();
  } else if (sheet_name == "Attendance"){
    if (user_row == 1 || user_col < 3){
      var user_old_value = (user_old_value != undefined) ? user_old_value:"";
//      user_event_range.setValue(user_old_value);
      show_att_sheet_alert();
    } else {
      var attendance = range_object(sheet, user_row)
      update_attendance(attendance);
    }
  }
}

function update_attendance(attendance){
  var MemberObject = main_range_object("Membership");
  Logger.log(attendance);
  var counts = {};
  counts["Active"] = {};
  counts["Pledge"] = {};
  var test_len = attendance.object_count;
  for(var i = 2; i< attendance.object_count; i++) {
    var member_name = attendance.object_header[i];
    var event_status = attendance[member_name][0];
    var member_status = MemberObject[member_name]["Chapter Status"][0]
    counts[member_status][event_status] = counts[member_status][event_status] ? counts[member_status][event_status]+1 : 1;
  }
  Logger.log(counts)
  var event_name = attendance["Event Name"];
  Logger.log(event_name);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Events");
  var max_column = sheet.getLastColumn();
  var event_headers = sheet.getRange(1, 1, 1, max_column);
  var header_values = event_headers.getValues()
  var active_col = get_ind_from_string("# Members", header_values)
  var pledge_col = get_ind_from_string("# Pledges", header_values)
  var event_row = attendance.object_row;
  Logger.log("ROW: " + event_row + " Active: " + active_col + " Pledge: " + pledge_col)
  var active_range = sheet.getRange(event_row, active_col)
  var pledge_range = sheet.getRange(event_row, pledge_col)
  active_range.setValue(counts["Active"]["PR"])
  pledge_range.setValue(counts["Pledge"]["PR"])
}

function update_scores(user_row){
//  var user_row = 17;
  var other_type_rows = update_score(user_row);
  Logger.log(other_type_rows);
  for (i in other_type_rows){
    if (parseInt(other_type_rows[i])!=parseInt(user_row)){
      update_score(other_type_rows[i]);
    }
  }
}

function update_score(event_row){
//  var event_row = 4
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Events");
  var myEvent = range_object(sheet, event_row);
  var score_data = get_score(myEvent);
  var score_ind = myEvent["Score"][1];
  var event_date = myEvent["Event Date"][0];
  var event_type = myEvent["Event Type"][0];
  var score_range = sheet.getRange(event_row, score_ind);
  score_range.setValue(0); // To prevent the current score from affecting max
  Logger.log("Event Date: " + event_date + " Type:" + event_type)
  var total_scores = get_event_scores();
  Logger.log(total_scores)
  var month = event_date.getMonth();
  var semester = "FALL";
  if (month<5){
	var semester = "SPRING";
	}
  var type_score = total_scores[semester][event_type][0];
  var other_type_rows = total_scores[semester][event_type][1];
  Logger.log("Type Score: " + type_score);
  var score = score_data.score;
  var total = parseFloat(type_score) + parseFloat(score);
  Logger.log(total)
  if (total > parseFloat(score_data.score_max_semester)){
    score = score_data.score_max_semester - type_score;
    score = score > 0 ? score:0;
  }
  Logger.log("FINAL SCORE: " + score)
  score_range.setValue(score);
  score_range.setNote(score_data.score_method_note);
  return other_type_rows;
}

function get_event_scores(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Events");
  var max_column = sheet.getLastColumn();
  var max_row = sheet.getLastRow();
  var full_data_range = sheet.getRange(1, 1, max_row, max_column);
  var full_data_values = full_data_range.getValues();
  var score_ind = get_ind_from_string("Score", full_data_values);
  var date_ind = get_ind_from_string("Event Date", full_data_values);
  var type_ind = get_ind_from_string("Event Type", full_data_values);
  var score_values = get_column_values(score_ind-1, full_data_values);
  var date_values = get_column_values(date_ind-1, full_data_values);
  var type_values = get_column_values(type_ind-1, full_data_values);
  var date_types = new Array();
  date_types["SPRING"] = {};
  date_types["FALL"] = {};
  for(var i = 1; i< date_values.length; i++) {
		var date = date_values[i];
        var month = date.getMonth();
		var type_name = type_values[i];
		var score = score_values[i];
		var semester = "FALL";
		if (month<5){
			var semester = "SPRING";
		}
        var old_score = date_types[semester][type_name] ? 
				date_types[semester][type_name][0] : 0;
        var new_score = parseFloat(old_score) + parseFloat(score);
        var old_rows = date_types[semester][type_name] ? 
				date_types[semester][type_name][1] : [];
        old_rows.push(parseInt(i) + 1);
		date_types[semester][type_name] = [new_score, old_rows]
	  }
  return date_types;
}

function get_column_values(col, range_values){
	var newArray = new Array();
	for(var i=0; i<range_values.length; i++){
		newArray.push(range_values[i][col]);
     }
	return newArray;
}

function get_score(myEvent){
  var event_type = myEvent["Event Type"][0]
  var score_data = get_score_method(event_type);
  Logger.log(score_data);
  var score_method_edit = edit_score_method(myEvent, score_data.score_method);
  var score = eval(score_method_edit);
  score = score.toFixed(1);
  score_data.score = score;
  Logger.log("SCORE RAW: " + score);
  return score_data
}

function edit_score_method(myEvent, score_method){
  var attend = myEvent["# Members"][0];
  var attend = (attend != "") ? attend:0;
  if (~score_method.indexOf("memberATT")){
      var total_members = 30;
      var percent_attend = attend / total_members;
      score_method = score_method.replace("memberATT", percent_attend);
          }
  if (~score_method.indexOf("memberADD")){
      score_method = score_method.replace("memberADD", attend);
          }
  if (~score_method.indexOf("NON-MEMBER") || ~score_method.indexOf("ALUMNI")){
      var non_members = myEvent["# Non- Members"][0];
      var non_members = (non_members != "") ? non_members:0;
      score_method = score_method.replace("NON-MEMBER", non_members);
          }
  if (~score_method.indexOf("ALUMNI")){
      var alumni_members = myEvent["# Alumni"][0];
      var alumni_members = (alumni_members != "") ? alumni_members:0;
      score_method = score_method.replace("ALUMNI", alumni_members);
          }
  if (~score_method.indexOf("STEM")){
      var stem = myEvent["STEM?"][0];
      var stem = (stem == "Yes") ? 1:0;
      score_method = score_method.replace("STEM", stem);
          }
  if (~score_method.indexOf("P_FOCUS")){
      var focus = myEvent["PLEDGE Focus"][0];
      var focus = (focus == "Yes") ? 1:0;
      score_method = score_method.replace("P_FOCUS", focus);
          }
  if (~score_method.indexOf("HOURS")){
      score_method = "0";
          }
  if (~score_method.indexOf("MEMBERSHIP")){
      score_method = "0";
          }
  if (~score_method.indexOf("PROPERTY")){
      score_method = "0";
          }
  if (~score_method.indexOf("MEETINGS")){
      score_method = "0";
          }
  if (~score_method.indexOf("GPA")){
      score_method = score_method.replace("GPA", 0);
          }
  if (~score_method.indexOf("INIT")){
      score_method = score_method.replace("INIT", 0);
          }
  if (~score_method.indexOf("PLEDGE")){
      score_method = score_method.replace("PLEDGE", 0);
          }
  if (~score_method.indexOf("SOCIETY")){
      score_method = score_method.replace("SOCIETY", 0);
          }
  if (~score_method.indexOf("OFFICER")){
      score_method = score_method.replace("OFFICER", 0);
          }
  Logger.log("Score Method Raw: " + score_method)
  return score_method
}

function get_score_method(event_type){
  var ScoringObject = main_range_object("Scoring");
  var score_object = ScoringObject[event_type];
  var score_type = score_object["Score Type"][0];
  var score_method_note = score_object["How points are calculated"][0];
  var att =  score_object["Attendence Multiplier"][0];
  var att = (att != "") ? att:0;
  var add = score_object["Member Add"][0];
  var add = (add != "") ? add:0;
  var base =  score_object["Base Points"][0];
  var special = score_object["Special"][0];
  if (score_type == "Events"){
   var score_method = "memberATT*" + att + "+memberADD*" + add;
  }
  if (score_type == "Submit"){
   var score_method = base;
  }
  if (score_type == "Events/Submit"){
   var score_method =  "memberATT*" + att + "+memberADD*" + add + "+" + base;
  }
  if (score_type == "Events/Special" || score_type == "Special"){
   var score_method =  special;
  }
  return {score_method: score_method,
          score_method_note: score_method_note,
          score_max_semester: score_object["Max/ Semester"][0]
         }
}

function get_ind_from_string(str, range_values){
  var range_values = range_values[0]
  for (val_ind in range_values){
    if (range_values[val_ind] == str){
      return +val_ind+1
    }
  }
}

function main_range_object(sheetName){
//  var sheetName = "Membership"
//  var sheetName = "Scoring"
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheetName=="Membership"){
    var short_header = "Member Name"
  } else if (sheetName=="Scoring"){
    var short_header = "Short Name"
  }
  var max_row = sheet.getLastRow() - 1;
  var max_column = sheet.getLastColumn();
  var full_data_range = sheet.getRange(2, 1, max_row, max_column);
  var full_data_values = full_data_range.getValues();
  var header_range = sheet.getRange(1, 1, 1, max_column);
  var header_values = header_range.getValues();
  var short_names_ind = get_ind_from_string(short_header, header_values);
  var short_names_range = sheet.getRange(2, short_names_ind, max_row, 1);
  var short_names = short_names_range.getValues();
  short_names = [].concat.apply([], short_names);
  short_names = cleanArray(short_names, 100);
  var myObject = new Array();
  myObject["object_header"] = new Array();
  for (val in short_names){
    var short_name_ind = parseInt(val);
    var short_name = short_names[short_name_ind];
    var range_values = full_data_values[short_name_ind]
    var temp = range_object_fromValues(header_values[0], range_values, short_name_ind + 2);
    myObject[short_name] = temp;
    myObject["object_count"] = myObject["object_count"] ? myObject["object_count"]+1 : 1;
    myObject["object_header"].push(short_name);
  }
  return myObject
}

function range_object(sheet, range_row){
  if (typeof sheet === "string"){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet);
  }
  var max_column = sheet.getLastColumn()
  var range = sheet.getRange(range_row, 1, 1, max_column);
  var range_values = range.getValues()
  Logger.log(range_values)
  var header_range = sheet.getRange(1, 1, 1, max_column);
  var header_values = header_range.getValues();
  Logger.log(header_values)
  var myObject = new Array();
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
  onEdit({
    user : Session.getActiveUser().getEmail(),
    source : SpreadsheetApp.getActiveSpreadsheet(),
    range : SpreadsheetApp.getActiveSpreadsheet().getActiveCell(),
    value : SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getValue(),
    authMode : "LIMITED"
  });
}

function show_att_sheet_alert(){
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'ERROR',
     'Please edit the events or members on the Events or Membership Sheet',
      ui.ButtonSet.OK);
}

function show_event_sheet_alert() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'ERROR',
     'Please add/edit/remove events using the form',
      ui.ButtonSet.OK);
  showaddEvent();
}

function getChapterMembers(){
  var chapterName = "Chi";
  var ChapterMembers = getChapterMembers_(chapterName)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  for (var i = 0; i < ChapterMembers.length; i++) {
    sheet.getRange(i+1, 1, 1, ChapterMembers[i].length).setValues(new Array(ChapterMembers[i]));
  }
}

function get_event_data(SheetName) {
//  var SheetName="Events"
//  var SheetName="Attendance"
  Logger.log(SheetName)
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SheetName);
  if (sheet != null) {
    var max_row = sheet.getLastRow() - 1
    var max_row = (max_row != 0) ? max_row:1;
    var max_column = sheet.getLastColumn()
    var range = sheet.getRange(2, 1, max_row, max_column);
    var header_range = sheet.getRange(1, 1, 1, max_column);
    var header_values = header_range.getValues();
    Logger.log(header_values);
    for (i in header_values[0]){
      if (header_values[0][i] == "Event Date") {
        var date_index = parseInt(i);
        Logger.log("date index: " + date_index);
      } else if (header_values[0][i] == "Event Name") {
        var name_index = parseInt(i);
        Logger.log("name index: " + name_index);
      }
    }
//    var sorted_range = range.sort({column: +date_index+1, ascending: true});
//    sheet.setFrozenRows(1);
  }
  return {range: range,
          header: header_values,
          date_index: date_index,
          name_index: name_index
         }
}

function align_event_attendance(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance");
  var event_data = get_event_data("Events");
  var att_data = get_event_data("Attendance");
  var event_values = event_data.range.getValues();
  var att_values = att_data.range.getValues();
  var attendance_rows = att_values.length;
  Logger.log(attendance_rows);
  for (row in event_values){
    var this_row = parseInt(row) + 1
    var event_name = event_values[row][event_data.name_index];
    var event_date = event_values[row][event_data.date_index];
    if (this_row - 1 < attendance_rows){
      var att_event_name = att_values[row][att_data.name_index];
      var att_event_date = att_values[row][att_data.date_index];
    }
    if (event_name != att_event_name){
      sheet.insertRowAfter(this_row);
      var name_range = sheet.getRange(this_row+1, att_data.name_index+1);
      name_range.setValue(event_name);
      var date_range = sheet.getRange(this_row+1, +att_data.date_index+1);
      date_range.setValue(event_date);
      sheet.setRowHeight(this_row+1, 10);
    }
    Logger.log(event_name);
  }
}