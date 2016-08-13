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
  menu.addItem('Status Change', 'member_update_sidebar');
//  menu.addItem('Grad Change', 'form_gradDialog');
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

function form_statusDialog() {
  var html = HtmlService.createHtmlOutputFromFile('FORM_STATUS')
      .setWidth(800)
      .setHeight(400);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'STATUS FORM');
}

function form_gradDialog() {
  var html = HtmlService.createHtmlOutputFromFile('FORM_GRAD')
      .setWidth(800)
      .setHeight(400);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'GRAD FORM');
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

function member_update_sidebar() {
  var template = HtmlService
      .createTemplateFromFile('member_select');
  var htmlOutput = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Update Members');
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
  Logger.log('Called addEvent');
  var html = HtmlService.createTemplateFromFile('Events');
  html.events = get_type_list("Events");
  var htmlOutput  = html.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Event Functions')
    .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(htmlOutput);
}

function member_update(form) {
//  Logger.log(form);
//  var select_members = ["Daniel Tranfag...", "Jacob Landsied...", "Jessyca Thomas", "Louis Bertani", "Mark Silvern"];
  var select_members = form.memberlist
  var MemberObject = main_range_object("Membership");
  var members = [];
  for (var i = 0; i < MemberObject.object_count; i++) {
    var member_name = MemberObject.object_header[i];
    for (var j = 0; j < select_members.length; j++) {
      var member_name_select = select_members[j];
      member_name_select = member_name_select.replace("...","")
      if (~member_name.indexOf(member_name_select)){
        members.push(MemberObject[member_name]);
      }
    }
  }
  Logger.log(members);
//  var update_type = "Degree received";
  if (form.update_type == "Degree received"){
    Logger.log("DEGREE");
    var html = HtmlService.createTemplateFromFile('FORM_GRAD');
    html.members = members;
//    var html = HtmlService.createTemplateFromFile('FORM_STATUS');
    var htmlOutput = html.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(700)
      .setHeight(400);
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(htmlOutput, 'GRAD FORM');
  }
}

function format_date(date) {
  var raw = date.split("-");
  return raw[1] + "/" + raw[2] + "/" + raw[0]
}

function find_member_shortname(MemberObject, member_name_raw){
  var member_name = member_name_raw.split("...")[0]
  for (var full_name in MemberObject){
    if (~full_name.indexOf(member_name)){
      return MemberObject[full_name]
    }
  }
}

function save_form(csvFile, form_type){
  try {
    var folder = DriveApp.getFolderById('0BwvK5gYQ6D4nWVhUVlo4dUhYV0E');
    var chapterName = SpreadsheetApp
                      .getActiveSpreadsheet()
                      .getRangeByName("ChapterName").getValue();
    var date = new Date();
    var currentMonth = date.getMonth() + 1;
    if (currentMonth < 10) { currentMonth = '0' + currentMonth; }
    var fileName = date.getFullYear().toString()+
                   currentMonth.toString()+
                   date.getDate().toString()+"_"+
                   chapterName+"_"+
                   form_type+"_"+
                   date.getTime().toString()+
                   ".csv";
    var file = folder.createFile(fileName, csvFile);
    Logger.log("fileBlob Name: " + file.getName())
    Logger.log('fileBlob: ' + file);
    
    var template = HtmlService.createTemplateFromFile('SubmitFormResponse');
    var file_url = template.fileUrl = file.getUrl();
    var submission_date = template.date = date;
    var submission_type = template.type = form_type;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Submissions");
    SpreadsheetApp.setActiveSheet(sheet);
    var max_column = sheet.getLastColumn();
    var max_row = sheet.getLastRow();
    var submit_range = sheet.getRange(max_row + 1, 1, 1, max_column);
    var file_name = template.name = file.getName();
    submit_range.setValues([[submission_date, file_name, submission_type, 0, file_url]])
    return template.evaluate().getContent();
  } catch (error) {
    Logger.log(error);
    return error.toString();
  }
}

function process_oer(form) {
//  var form = {"Scribe": "Eugene Balaguer", "Service Chair": "Kyle Wilson", 
//              "Treasurer": "Jeremy Faber", "Fundraising Chair": "N/A", 
//              "Risk Management Chair": "N/A", "Recruitment Chair": "Hannah Rowe", 
//              "Website/Social Media Chair": "N/A", "Pledge/New Member Educator": "Adam Schilpero...", 
//              "officer_end": "2016-12-31", "Corresponding Secretary": "Kyle Wilson",
//              "TCS_start": "2016-08-01", "TCS_end": "2017-06-01",
//              "Social/Brotherhood Chair": "N/A", "officer_start": "2016-08-01", 
//              "Scholarship Chair": "N/A", "Vice Regent": "David Montgome...", "PD Chair": "N/A", 
//              "Regent": "Adam Schilpero...", "Project Chair": "N/A"};
  Logger.log(form);
  var MemberObject = main_range_object("Membership");
  if (form.officer_start == "" || form.officer_end == "" || 
      form.TCS_start == "" || form.TCS_end == ""){
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
     'ERROR',
     'You must set all of the dates',
      ui.ButtonSet.OK);
    return false;
  }
  var officer_start = format_date(form.officer_start);
  var officer_end = format_date(form.officer_end);
  var TCS_start = format_date(form.TCS_start);
  var TCS_end = format_date(form.TCS_end);
  delete form.officer_start
  delete form.officer_end
  delete form.TCS_start
  delete form.TCS_end
  var header = ["Submitted by", "Date Submitted", "Chapter Name", "Office",
            "Term Begins (M/D/YYYY)", "Term Ends (M/D/YYYY)", "*ChapRoll",
            "First Name", "Last Name", "Mobile Phone", "Campus Email", "Other Email"];
  var start_date = [];
  var data = [];
  data.push(header);
  data.push(["test", "date", "chapter","","","","","","","","",""]);
  for (var key in form){
    var start = officer_start;
    var end = officer_end
    if (form[key] == "N/A"){
      continue;
    }
    if (~key.indexOf("Treasurer") || ~key.indexOf("Corresponding")){
      start = TCS_start;
      end = TCS_end;
    }
    var member_object = find_member_shortname(MemberObject, form[key]);
    var row = ["", "", "", key, start, end, member_object["Badge Number"][0],
              member_object["First Name"][0], member_object["Last Name"][0],
              member_object["Phone Number"][0], member_object["Email Address"][0], ""];
    Logger.log(row);
    data.push(row);
  }
  Logger.log(data);
  var csvFile = create_csv(data);
  Logger.log(csvFile);
  return save_form(csvFile, "OER");
}

function process_grad(form) {
//  var form = {'owed': [10, 50, 120], 
//              'new_location': ['New York', 'LA', 'TEST'], 
//              'student': ['NO', 'YES', 'NO'], 
//              'name': ['Cole Mobberley', 'Austin Mutschler', 'Adam Schilperoort'], 
//              'degree': ['Cole Mobberley MAJOR', 'Austin Mutschler MAJOR', 'Adam Schilperoort MAJOR'], 
//              'email': ['ColeMobberley@email.com', 'AustinMutschler@email.com', 'AdamSchilperoort@email.com']};
  Logger.log(form);
  var header = ["name", "owed", "new_location", "student", "degree", "email"];
  var data = [];
  data.push(header);
  for (var i = 0; i < form["name"].length; i++){
    var row = [];
    header.forEach(function (item) {
      row.push(form[item][i]);
//      Logger.log(form[item][i]);
    })
    data.push(row);
  }
  Logger.log(data);
       }

function out_OER(){

}

function out_MSCR(){
//  Submitted by
//  Date Submitted
//  School Name

//  ChapRoll
//  First Name
//  Last Name
//  Mobile Phone
//  Email  Address
//  Reason for Status Change
  //Graduated from school
  //Withdrawing from school
  //Transferring to another school
  //Wishes to REQUEST Premature Alum Status
//  Degree Received
//  Graduation Date (M/D/YYYY)
//  Employer
//  Work Email
//  Attending Graduate School where ?
//  Withdrawing from school?
//  Date withdrawn (M/D/YYYY)
//  Transferring to what school ?
//  Date of transfer (M/D/YYYY)
//  REQUESTING what type of Premature Alum Status?
  //Undergrad Premature <4 years
  //Undergrad Premature > 4 years
  //Grad Student Premature
}

function out_COOP(){
//Submitted by
//Date Submitted
//Chapter Name

//*ChapRoll
//First Name
//Last Name
//Reason Away
  //Co-Op/Internship
  //Study Abroad
  //Called to Active/Reserve Military Duty
//Start Date (M/D/YYYY)
//End Date (M/D/YYYY)
//Miles from Campus**
}

function out_INIT(){
//Submitted by
//Date Submitted
//Initiation Date
//Chapter Name

//Graduation Year
//Roll Number
//First Name
//Middle Name
//Last Name
//Overall GPA
  //A Pledge Test Scores
  //B Pledge Test Scores
//Initiation Fee
//Late Fee
//Badge Style
  //109 ($20)
  //106 ($67)
  //107 ($117)
  //102 ($165)
  //103 ($209)
//Guard Type
  //Chose Gold gloss or 10k Gold, 
  //and one of: 
    //Plain
    //Chased/Engraved
    //Close Set Pearl
    //Crown Set Pearl
//Badge Cost
//Guard Cost
//Sum for member
}

function out_DEPL(){
//Submitted by
//Date Submitted
//Chapter Name
//
//First Name
//Last Name
//Reason Depledged
  //Voluntarily decided not to continue
  //Too much time required
  //Poor grades
  //Lost interest
  //Negative Chapter Vote
  //Withdrew from Engineering/University
  //Transferring to another school
  //Other
//Date Depledged (M/D/YYYY)
}

function create_csv(data){
  try {
    var csvFile = undefined;

    // Loop through the data in the range and build a string with the CSV data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // Join each row's columns
        // Add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
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
    submit_range.setValues([[submission_date, file_name, submission_type, 0, file_url]])
    update_scores_submit(max_row + 1);
    
    return template.evaluate().getContent();
  } catch (error) {
    return error.toString();
  }
}

function get_score_submit(myScore){
  var event_type = myScore["Type"][0]
  var score_data = get_score_method(event_type);
  Logger.log(score_data);
  var score = eval(score_data.score_method);
  score = score.toFixed(1);
  score_data.score = score;
  Logger.log("SCORE RAW: " + score);
  return score_data
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

function get_ind_list(type){
//  var type = "Brotherhood";
//  var type = "Operate";
//  var type = "ProDev";
//  var type = "Service";
  Logger.log(type);
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
    update_scores_event(user_row);
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
      update_scores_event(user_row);
    }
  } else if (sheet_name == "Membership") {
    Logger.log("MEMBER CHANGED");
    if (user_col > 8){
      update_scores_org_gpa_serv();
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
    counts[member_status][event_status] = counts[member_status][event_status] ? counts[member_status][event_status] + 1 : 1;
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

function update_scores_event(user_row){
//  var user_row = 17;
  var myObject = range_object("Events", user_row);
  var score_data = get_score_event(myObject);
  var other_type_rows = update_score(user_row, "Events", score_data, myObject);
  Logger.log("OTHER ROWS" + other_type_rows);
  for (i in other_type_rows){
    if (parseInt(other_type_rows[i])!=parseInt(user_row)){
      var myObject = range_object("Events", other_type_rows[i]);
      var score_data = get_score_event(myObject);
      update_score(other_type_rows[i], "Events", score_data, myObject);
    }
  }
}

function update_score_att(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Scoring");
  var EventObject = main_range_object("Events");
  var ScoringObject = main_range_object("Scoring");
  var total_members = get_total_members().Active;
  var date_types = [];
  var counts = [];
  for (var i = 0; i < EventObject.object_count; i++){
    var event_name = EventObject.object_header[i];
    var event_type = EventObject[event_name]["Type"][0];
    if (event_type == "Meetings"){
      var object_date = EventObject[event_name]["Date"][0];
      var meeting_att = EventObject[event_name]["# Members"][0];
      meeting_att = parseFloat(meeting_att / total_members);
      var month = object_date.getMonth();
      var semester = "FALL";
      if (month<5){
      var semester = "SPRING";
      }
      date_types[semester] = date_types[semester] ? 
        date_types[semester] + meeting_att:meeting_att;
      counts[semester] = counts[semester] ? 
        counts[semester] + 1:1;
    }
  }
  var fall_avg = date_types["FALL"]/counts["FALL"];
  var spring_avg = date_types["SPRING"]/counts["SPRING"];
  Logger.log("FALL ATT: " + fall_avg + " SPRING ATT: " + spring_avg);
  var score_method_raw = ScoringObject["Meetings"]["Special"][0];
  var score_max = ScoringObject["Meetings"]["Max/ Semester"][0];
  var score_method_fa = score_method_raw.replace("MEETINGS", fall_avg);
  var score_row = ScoringObject["Meetings"].object_row;
  var total_col = ScoringObject["Meetings"]["CHAPTER TOTAL"][1];
  var score_range_fa = sheet.getRange(score_row, ScoringObject["Meetings"]["FALL SCORE"][1]);
  var score_range_sp = sheet.getRange(score_row, ScoringObject["Meetings"]["SPRING SCORE"][1]);
  var score_range_tot = sheet.getRange(score_row, total_col);
  var score_method_sp = score_method_raw.replace("MEETINGS", spring_avg);
  var score_fa = eval_score(score_method_fa, score_max);
  var score_sp = eval_score(score_method_sp, score_max);
  score_range_fa.setValue(score_fa);
  score_range_sp.setValue(score_sp);
  score_range_tot.setValue(+score_fa + score_sp);
  update_dash_score("Operate", total_col);
}

function update_score_member_pledge(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Scoring");
  var member_value_obj = get_membership_ranges();
  var init_sp_value = member_value_obj.init_sp_range.getValue();
  var init_fa_value = member_value_obj.init_fa_range.getValue();
  var pledge_sp_value = member_value_obj.pledge_sp_range.getValue();
  var pledge_fa_value = member_value_obj.pledge_fa_range.getValue();
  var grad_sp_value = member_value_obj.grad_sp_range.getValue();
  var grad_fa_value = member_value_obj.grad_fa_range.getValue();
  var act_sp_value = member_value_obj.act_sp_range.getValue();
  var act_fa_value = member_value_obj.act_fa_range.getValue();
  var ScoringObject = main_range_object("Scoring");
  var score_method_pledge_raw = ScoringObject["Pledge Ratio"]["Special"][0];
  var score_pledge_max = ScoringObject["Pledge Ratio"]["Max/ Semester"][0];
  var score_method_pledge_fa = score_method_pledge_raw.replace("INIT", init_fa_value);
  score_method_pledge_fa = score_method_pledge_fa.replace("PLEDGE", pledge_fa_value);
  var score_pledge_fa = eval_score(score_method_pledge_fa, score_pledge_max);
  var score_method_pledge_sp = score_method_pledge_raw.replace("INIT", init_sp_value);
  score_method_pledge_sp = score_method_pledge_sp.replace("PLEDGE", pledge_sp_value);
  var score_pledge_sp = eval_score(score_method_pledge_sp, score_pledge_max);
  var score_method_raw = ScoringObject["Membership"]["Special"][0];
  var score_max = ScoringObject["Membership"]["Max/ Semester"][0];
  var score_method_fa = score_method_raw.replace("OUT", grad_fa_value);
  score_method_fa = score_method_fa.replace("IN", init_fa_value);
  score_method_fa = score_method_fa.replace("MEMBERS", act_fa_value);
  var score_fa = eval_score(score_method_fa, score_max);
  var score_method_sp = score_method_raw.replace("OUT", grad_sp_value);
  score_method_sp = score_method_sp.replace("IN", init_sp_value);
  score_method_sp = score_method_sp.replace("MEMBERS", act_sp_value);
  var score_sp = eval_score(score_method_sp, score_max);
  var score_row = ScoringObject["Membership"].object_row;
  var score_fa_range = sheet.getRange(score_row,
                                      ScoringObject["Membership"]["FALL SCORE"][1]);
  var score_sp_range = sheet.getRange(score_row,
                                      ScoringObject["Membership"]["SPRING SCORE"][1]);
  var total_col = ScoringObject["Membership"]["CHAPTER TOTAL"][1];
  var score_tot_range = sheet.getRange(score_row,total_col);
  var score_pledge_row = ScoringObject["Pledge Ratio"].object_row;
  var score_pledge_fa_range = sheet.getRange(score_pledge_row,
                                      ScoringObject["Pledge Ratio"]["FALL SCORE"][1]);
  var score_pledge_sp_range = sheet.getRange(score_pledge_row,
                                      ScoringObject["Pledge Ratio"]["SPRING SCORE"][1]);
  var score_pledge_tot_range = sheet.getRange(score_pledge_row,total_col);
  score_fa_range.setValue(score_fa);
  score_sp_range.setValue(score_sp);
  score_tot_range.setValue(score_fa + score_sp);
  update_dash_score("Operate", total_col);
  score_pledge_fa_range.setValue(score_pledge_fa);
  score_pledge_sp_range.setValue(score_pledge_sp);
  score_pledge_tot_range.setValue(score_pledge_fa + score_pledge_sp);
  update_dash_score("Brotherhood", total_col);
}

function get_membership_ranges(){
  var init_sp_range = SpreadsheetApp
                      .getActiveSpreadsheet()
                      .getRangeByName("INIT_SP");
  var init_fa_range = SpreadsheetApp
                      .getActiveSpreadsheet()
                      .getRangeByName("INIT_FA");
  var pledge_sp_range = SpreadsheetApp
                        .getActiveSpreadsheet()
                        .getRangeByName("PLEDGE_SP");
  var pledge_fa_range = SpreadsheetApp
                        .getActiveSpreadsheet()
                        .getRangeByName("PLEDGE_FA");
  var grad_sp_range = SpreadsheetApp
                      .getActiveSpreadsheet()
                      .getRangeByName("GRAD_SP");
  var grad_fa_range = SpreadsheetApp
                      .getActiveSpreadsheet()
                      .getRangeByName("GRAD_FA");
  var act_sp_range = SpreadsheetApp
                     .getActiveSpreadsheet()
                     .getRangeByName("ACT_SP");
  var act_fa_range = SpreadsheetApp
                     .getActiveSpreadsheet()
                     .getRangeByName("ACT_FA");
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

function update_scores_org_gpa_serv(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Scoring");
  var score_data = get_scores_org_gpa_serv();
  var ScoringObject = main_range_object("Scoring");
  var total_col = ScoringObject["Societies"]["CHAPTER TOTAL"][1];
  var fall_col = ScoringObject["Societies"]["FALL SCORE"][1];
  var spring_col = ScoringObject["Societies"]["SPRING SCORE"][1];
  var societies_range = sheet.getRange(ScoringObject["Societies"].object_row, total_col);
  var societies_method = ScoringObject["Societies"]["Special"][0];
  societies_method = societies_method.replace("ORG", score_data.percent_org);
  societies_method = societies_method.replace("OFFICER", score_data.officer_count);
  var societies_max = ScoringObject["Societies"]["Max/ Semester"][0];
  var socieities_score = eval_score(societies_method, societies_max);
  var gpa_fall_range = sheet.getRange(ScoringObject["GPA"].object_row, fall_col);
  var gpa_spring_range = sheet.getRange(ScoringObject["GPA"].object_row, spring_col);
  var gpa_range = sheet.getRange(ScoringObject["GPA"].object_row, total_col);
  var gpa_method_raw = ScoringObject["GPA"]["Special"][0];
  var gpa_fall_method = gpa_method_raw.replace("GPA", score_data.gpa_avg_fall);
  var gpa_spring_method = gpa_method_raw.replace("GPA", score_data.gpa_avg_spring);
  var gpa_max = ScoringObject["GPA"]["Max/ Semester"][0];
  var gpa_fall_score = eval_score(gpa_fall_method, gpa_max);
  var gpa_spring_score = eval_score(gpa_spring_method, gpa_max);
  var service_range = sheet.getRange(ScoringObject["Service Hours"].object_row, total_col);
  var service_method = ScoringObject["Service Hours"]["Special"][0];
  service_method = service_method.replace("HOURS", score_data.percent_service);
  var service_max = ScoringObject["Service Hours"]["Max/ Semester"][0];
  var service_score = eval_score(service_method, service_max);
  Logger.log("SOC: " + societies_method + ", SCORE: " + socieities_score);
  Logger.log("GPA_FALL: " + gpa_fall_method + ", SCORE: " + gpa_fall_score);
  Logger.log("GPA_SPRING: " + gpa_spring_method + ", SCORE: " + gpa_spring_score);
  Logger.log("SERV: " + service_method + ", SCORE: " + service_score);
  societies_range.setValue(socieities_score);
  gpa_fall_range.setValue(gpa_fall_score);
  gpa_spring_range.setValue(gpa_spring_score);
  gpa_range.setValue(gpa_fall_score + gpa_spring_score);
  service_range.setValue(service_score);
  update_dash_score("ProDev", total_col);
  update_dash_score("Service", total_col);
}

function eval_score(score_method, score_max){
  var score = eval(score_method);
  score = parseFloat(score.toFixed(1));
  score = score > parseFloat(score_max) ? score_max: score;
  return score;
}

function get_scores_org_gpa_serv(){
  var gpa_counts = {};
  var officer_counts = {};
  var org_counts = {};
  var service_count = 0;
  var officer_count = 0;
  var org_count = 0;
  var officers = ["Officer (Pro/Tech)", "Officer (Honor)", "Officer (Other)"];
  var orgs = ["Professional/ Technical Orgs", "Honor Orgs", "Other Orgs"];
  var gpas = ["Fall GPA", "Service Hours", "Spring GPA"];
  var MemberObject = main_range_object("Membership");
  var gpa = 0;
  for (var i = 0; i < MemberObject.object_count; i++){
    var member_name = MemberObject.object_header[i];
    var org_true = false;
    var officer_true = false;
    for (var j = 0; j <= 2; j++){
      var gpa = parseInt(MemberObject[member_name][gpas[j]][0]);
      gpa_counts[gpas[j]] = gpa_counts[gpas[j]] ? gpa_counts[gpas[j]]+gpa:gpa;
      var this_org = MemberObject[member_name][orgs[j]][0];
      org_counts[orgs[j]] = org_counts[orgs[j]] ? org_counts[orgs[j]]:0;
      org_counts[orgs[j]] = this_org!="None" ? org_counts[orgs[j]]+1:org_counts[orgs[j]];
      org_true = this_org!="None" ? true:org_true;
      var officer = MemberObject[member_name][officers[j]][0];
      officer_counts[officers[j]] = officer_counts[officers[j]] ? officer_counts[officers[j]]:0;
      officer_counts[officers[j]] = officer=="YES" ? officer_counts[officers[j]]+1:officer_counts[officers[j]];
      officer_true = officer=="YES" ? true:officer_true;
      Logger.log("GPA: " + gpa + " ORG: " + org + " OFFICER: " + officer);
    }
    var service_hours = MemberObject[member_name]["Service Hours"][0];
    var service_hours_self = MemberObject[member_name]["Self Service Hours"][0];
    service_hours = +service_hours + service_hours_self
    service_count = service_hours >= 16 ? service_count + 1:service_count;
    officer_count = officer_true ? officer_count + 1:officer_count;
    org_count = org_true ? org_count + 1:org_count;
  }
  var percent_service = service_count / MemberObject.object_count;
  var percent_org = org_count / MemberObject.object_count;
  var gpa_avg_fall = gpa_counts["Fall GPA"] / MemberObject.object_count;
  var gpa_avg_spring = gpa_counts["Spring GPA"] / MemberObject.object_count;
  return {percent_service: percent_service,
          percent_org: percent_org,
          officer_count: officer_count,
          gpa_avg_fall: gpa_avg_fall,
          gpa_avg_spring: gpa_avg_spring
          }
}

function update_scores_submit(user_row){
//  var user_row = 2;
  Logger.log("ROW: " + user_row);
  var myObject = range_object("Submissions", parseInt(user_row));
  var score_data = get_score_submit(myObject);
  var other_type_rows = update_score(user_row, "Submissions", score_data, myObject);
  Logger.log(other_type_rows);
}

function update_score(row, sheetName, score_data, myObject){
//  var row = 4
//  var shetName = "Events";
  Logger.log("SHEET: " + sheetName + " ROW: " + row)
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var score_ind = myObject["Score"][1];
  var object_date = myObject["Date"][0];
  var object_type = myObject["Type"][0];
  var score_range = sheet.getRange(row, score_ind);
  score_range.setValue(0); // To protect the current score from affecting max
  Logger.log("Date: " + object_date + " Type:" + object_type)
  var total_scores = get_current_scores(sheetName);
  Logger.log(total_scores)
  var month = object_date.getMonth();
  var semester = "FALL";
  if (month<5){
  var semester = "SPRING";
  }
  score_data.semester = semester;
  var type_score = total_scores[semester][object_type][0];
  var other_type_rows = total_scores[semester][object_type][1];
  Logger.log("Type Score: " + type_score);
  var score = score_data.score;
  var total = parseFloat(type_score) + parseFloat(score);
  Logger.log(total)
  if (total > parseFloat(score_data.score_max_semester)){
    score = score_data.score_max_semester - type_score;
    score = score > 0 ? score:0;
  }
  Logger.log("FINAL SCORE: " + score);
  score_data.final_score = score;
  score_data.type_score = type_score;
  update_main_score(score_data);
  score_range.setValue(score);
  score_range.setNote(score_data.score_method_note);
  return other_type_rows;
}

function update_main_score(score_data){
  Logger.log(score_data);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Scoring");
  var score_row = score_data.score_ids.score_row
  var semester_range = sheet.getRange(score_row, score_data.score_ids[score_data.semester]);
  var other_semester = score_data.semester=="FALL" ? "SPRING":"FALL";
  var other_semester_range = sheet.getRange(score_row, score_data.score_ids[other_semester]);
  var other_semester_value = other_semester_range.getValue();
  var other_semester_value = (other_semester_value != "") ? other_semester_value:0;
  var total_range = sheet.getRange(score_row, score_data.score_ids.chapter);
  var total_sem_score = parseFloat(score_data.final_score) + score_data.type_score;
  var total_score = parseFloat(other_semester_value) + total_sem_score;
  semester_range.setValue(total_sem_score);
  total_range.setValue(total_score);
  update_dash_score(score_data.score_type, score_data.score_ids.chapter);
}

function update_dash_score(score_type, score_column){
  Logger.log(score_type);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Scoring");
  if (score_type != undefined){
    var type_inds = get_ind_list(score_type);
    var type_count = type_inds.length;
  } else {
    var type_count = sheet.getLastRow();
    var type_inds = [];
    for (var i = 0; i <= type_count; i++) {
        type_inds.push(i);
    }
  }
  var total = 0;
  for (var j = 0; j < type_count; j++){
    var row = type_inds[j];
    var row_total = sheet.getRange(row, score_column).getValue();
    total = +total + row_total;
  }
  Logger.log(type_inds);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  var RangeName = "SCORE" + "_" + score_type.toUpperCase();
  var dash_score_range = SpreadsheetApp
                         .getActiveSpreadsheet()
                         .getRangeByName(RangeName);
  dash_score_range.setValue(total);
}

function get_current_scores(sheetName){
//  var sheetName = "Events";
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var max_column = sheet.getLastColumn();
  var max_row = sheet.getLastRow();
  var full_data_range = sheet.getRange(1, 1, max_row, max_column);
  var full_data_values = full_data_range.getValues();
  var score_ind = get_ind_from_string("Score", full_data_values);
  var date_ind = get_ind_from_string("Date", full_data_values);
  var type_ind = get_ind_from_string("Type", full_data_values);
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

function get_score_event(myEvent){
  var event_type = myEvent["Type"][0]
  var score_data = get_score_method(event_type);
  Logger.log(score_data);
  var score_method_edit = edit_score_method_event(myEvent, score_data.score_method);
  var score = eval(score_method_edit);
  score = score.toFixed(1);
  score_data.score = score;
  Logger.log("SCORE RAW: " + score);
  return score_data
}

function get_total_members(){
  var MemberObject = main_range_object("Membership");
  var counts = {};
  for(var i = 0; i< MemberObject.object_count; i++) {
    var member_name = MemberObject.object_header[i];
    var member_status = MemberObject[member_name]["Chapter Status"][0]
    counts[member_status] = counts[member_status] ? counts[member_status] + 1 : 1;
  }
  Logger.log(counts);
  return counts;
}

function edit_score_method_event(myEvent, score_method){
  var attend = myEvent["# Members"][0];
  var attend = (attend != "") ? attend:0;
  if (~score_method.indexOf("memberATT")){
      var totals = get_total_members();
      var percent_attend = attend / totals.Active;
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
  if (~score_method.indexOf("MEETINGS")){
      score_method = "MEETINGS";
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
  if (score_type == "Events/Special" || score_type == "Special"){
   var score_method =  special;
  }
  var score_ids = {
      score_row: score_object.object_row,
      FALL: score_object["FALL SCORE"][1],
      SPRING: score_object["SPRING SCORE"][1],
      chapter: score_object["CHAPTER TOTAL"][1]
  }
  return {score_method: score_method,
          score_method_note: score_method_note,
          score_max_semester: score_object["Max/ Semester"][0],
          score_ids: score_ids,
          score_type: score_object["Type"][0]
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
//  var sheetName = "Events"
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheetName=="Membership"){
    var short_header = "Member Name"
  } else if (sheetName=="Scoring"){
    var short_header = "Short Name"
  } else if (sheetName=="Events"){
    var short_header = "Event Name"
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
//    short_names.forEach(function (item) {
//      var test = item;
//      console.log(item);
//     Logger.log(item);
//    });
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
      if (header_values[0][i] == "Date") {
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