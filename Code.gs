/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
var betterLogStarted = false;
var SCRIPT_PROP = PropertiesService.getDocumentProperties();
//Logger.log(SCRIPT_PROP);
//startBetterLog();

function startBetterLog() {
  if (!betterLogStarted) {
    Logger = BetterLog.useSpreadsheet('1mo5t1Uu7zmP9t7w2hL1mWrdba4CtgD_Q9ImbAKjGZyM');
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
//  menu.addItem('Create Triggers', 'createTriggers');
//  menu.addItem('Event Functions', 'showaddEvent');
  menu.addItem('Pledge Forms', 'pledge_sidebar');
  menu.addItem("RESET", 'RESET');
  menu.addItem('Refresh', 'refresh')
  menu.addItem('SETUP', 'onInstall');
  menu.addItem('Status Change', 'member_update_sidebar');
  menu.addItem('Submit Item', 'submitSidebar');
//  menu.addItem("TEST", 'TEST');//test_onEdit
  menu.addItem('Unlock', 'unlock');
//  menu.addItem('Update Members', 'get_chapter_members');
  menu.addItem('Update Officers', 'officerSidebar');
  menu.addToUi();
}

function TEST(){
//  SCRIPT_PROP.setProperty('key', '1wBICuD_CvSm3BonA_OZg-sOTRJylVWVbLYi9nr8vn8Q');
//  SCRIPT_PROP.setProperty('chapter', 'Chi');
//  SCRIPT_PROP.setProperty('director', 'werd@thetatau.org');
//  SCRIPT_PROP.setProperty('email', 'venturafranklin@gmail.com');
//  SCRIPT_PROP.setProperty("region", "Western");
//  SCRIPT_PROP.setProperty("folder", "0BwvK5gYQ6D4nOHNPekh5cmF5RFk");
  Logger.log(SCRIPT_PROP.getProperty('key'));
  Logger.log(SCRIPT_PROP.getProperty('chapter'));
  Logger.log(SCRIPT_PROP.getProperty('director'));
  Logger.log(SCRIPT_PROP.getProperty('email'));
  Logger.log(SCRIPT_PROP.getProperty("region"));
  Logger.log(SCRIPT_PROP.getProperty("folder"));
  var ui = SpreadsheetApp.getUi();
  ui.alert('SETUP COMPLETE!\n'+
           'Next steps:\n'+
           '- Fill out Chapter Sheet\n'+
           '- Verify Membership\n'+
           '- Add Events & Attendance\n\n'+
           'Do not edit gray or black cells\n'+
           'Submit forms in menu "Add-ons-->ThetaTauReports"');
//  var ss = get_active_spreadsheet();
//  Logger.log(range.getValues());
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

function form_gradDialog() {
  var html = HtmlService.createHtmlOutputFromFile('FORM_GRAD')
      .setWidth(800)
      .setHeight(400);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'GRAD FORM');
}

function sync() {
  var dash_id = "10ebwK7tTKgveVCEOpRle2S17d4UjwmsoXXCPFvC9A-A";//SCRIPT_PROP.getProperty("dash");
  var dash_file = SpreadsheetApp.openById(dash_id);
  var chapter = SCRIPT_PROP.getProperty("chapter");
  var main_sheet = dash_file.getSheetByName("MAIN");
  var main_max = main_sheet.getMaxRows();
  var main_values = main_sheet.getRange(1, 1, main_max).getValues();
  var main_arr = get_column_values(0, main_values);
  var main_row = main_arr.indexOf(chapter);
  var submit_sheet = dash_file.getSheetByName("SUBMISSIONS");
  var submit_row_max = submit_sheet.getMaxRows();
  var submit_col_max = submit_sheet.getMaxColumns();
  var submit_values = submit_sheet.getRange(1, 1, submit_row_max, submit_col_max).getValues();
  var submit_arr = submit_values[0];
  var submit_col = submit_arr.indexOf(chapter);
  var officer_sheet = dash_file.getSheetByName("OFFICERS");
  var officer_row_max = officer_sheet.getMaxRows();
  var officer_col_max = officer_sheet.getMaxColumns();
  var officer_values = officer_sheet.getRange(1, 1, officer_row_max, officer_col_max).getValues();
  var officer_arr = get_column_values(0, officer_values);
  var officer_rows = getAllIndexes(officer_arr, chapter);
  Logger.log("TEST");
}

function getAllIndexes(arr, val) {
    var indexes = [], i;
    for(i = 0; i < arr.length; i++)
        if (arr[i] === val)
            indexes.push(i);
    return indexes;
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

function get_member_list(status){
//  var status = "Student";
//  var status = "Pledge";
  var MemberObject = main_range_object("Membership");
  var member_list = [];
  for(var i = 0; i< MemberObject.object_count; i++) {
    var member_name = MemberObject.object_header[i];
    var member_status = MemberObject[member_name]["Chapter Status"][0];
    if (member_status == status){
      member_name = shorten(member_name, 15);
      member_list.push(member_name);
    }
  }
  return member_list
}

function pledge_sidebar(){
  var template = HtmlService
      .createTemplateFromFile('pledge_select');
  template.pledge = get_member_list("Pledge");
  var htmlOutput = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Update Pledges');
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
   .createTemplateFromFile('SubmitForm')
//  .createHtmlOutputFromFile('SubmitForm');
  template.submissions = get_type_list('Submit');
  template.folder_id = get_folder_id();
  var htmlOutput = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Submit Item')
      .setWidth(500);
//  Logger.log(htmlOutput.getContent());
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(htmlOutput);
//      .showModalDialog(template, "SUBMIT");
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

function pledge_update(form) {
  Logger.log(form);
  var html = HtmlService.createTemplateFromFile('FORM_INIT');
  var INIT = []
  var DEPL = []
  if (typeof form["name"] === 'string'){
    for (var obj in form){
      form[obj] = [form[obj]];
    }
  }
  for (var k in form["name"]){
    var status = form["status"][k];
    var name = form["name"][k];
    if (status == "Initiated"){
      INIT.push(name);
    } else {
      DEPL.push(name);
    }
  }
  html.init = INIT
  html.depl = DEPL
  var htmlOutput = html.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(700)
    .setHeight(400);
  Logger.log(htmlOutput.getContent());
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showModalDialog(htmlOutput, 'PLEDGE FORM');
}

function member_update(form) {
//  var form = {"update_type": "Transfer", "memberlist": "Adam Schilpero...",
//              "Degree": ["Adam Schilpero...", "Austin Mutschl...", "Cole Mobberley"],//};
//              "Abroad": "Adam Schilpero...", "Transfer": "Adam Schilpero...",
//              "PreAlumn": ["Derek Hogue", "Esgar Moreno"], "Military": "Adam Schilpero...",
//              "CoOp": ["Adam Schilpero...", "Austin Mutschl...", "Cole Mobberley"]};
  Logger.log(form);
  var MemberObject = main_range_object("Membership");
  var html = HtmlService.createTemplateFromFile('FORM_STATUS');
  var CSMTA = []
  for (var k in form){
    var type = k;
    if (type == "update_type" || type == "memberlist"){
      continue;
    }
    Logger.log(k);
    var select_members = form[k];
    if (typeof select_members === 'string'){
      select_members = [select_members];
    }
    Logger.log(select_members);
    var members = [];
    for (var i = 0; i < MemberObject.object_count; i++) {
      var member_name = MemberObject.object_header[i];
      for (var j = 0; j < select_members.length; j++) {
        var member_name_select = select_members[j];
        member_name_select = member_name_select.replace("...","")
        if (~member_name.indexOf(member_name_select)){
          var this_obj = {}
          this_obj["Member Name"] = MemberObject[member_name]["Member Name"][0];
          this_obj["Email Address"] = MemberObject[member_name]["Email Address"][0];
          this_obj["Current Major"] = MemberObject[member_name]["Current Major"][0];
          this_obj["Phone Number"] = MemberObject[member_name]["Phone Number"][0];
          this_obj["New Status"] = type;
          members.push(this_obj);
        }
      }
    }
    html[type] = members;
    if (type != "Degree"){
      var CSMTA = CSMTA.concat(members)
    }
  }
    html.CSMTA = CSMTA;
    var htmlOutput = html.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(700)
      .setHeight(400);
    Logger.log(htmlOutput.getContent());
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(htmlOutput, 'STATUS FORM');
}

function format_date(date) {
  try{
    var raw = date.split("-");
    return raw[1] + "/" + raw[2] + "/" + raw[0]
  } catch (error) {
    Logger.log(error);
    return "";
  }
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
    var folder_id = get_folder_id();
    var folder = DriveApp.getFolderById(folder_id);
    var ss = get_active_spreadsheet();
    var chapterName = get_chapter_name();
    var date = new Date();
    var currentMonth = date.getMonth() + 1;
    if (currentMonth < 10) { currentMonth = '0' + currentMonth; }
    var currentDay = date.getDate().toString();
    if (currentDay < 10) { currentDay = '0' + currentDay; }
    var fileName = date.getFullYear().toString()+
                   currentMonth.toString()+
                   currentDay.toString()+"_"+
                   chapterName+"_"+
                   form_type+"_"+
                   date.getTime().toString()+
                   ".csv";
    var file = folder.createFile(fileName, csvFile);
    Logger.log("fileBlob Name: " + file.getName())
    Logger.log('fileBlob: ' + file);
    
    var template = HtmlService.createTemplateFromFile('SubmitFormResponse');
    var submission = {};
    submission.file = file;
    submission.folder_id = folder_id;
    submission.id = file.getId();
    var file_url = submission.alternateLink = template.fileUrl = file.getUrl();
    var submission_date = template.date = date;
    var submission_type = template.type = form_type;
    var sheet = ss.getSheetByName("Submissions");
    var max_column = sheet.getLastColumn();
    var max_row = sheet.getLastRow();
    var submit_range = sheet.getRange(max_row + 1, 1, 1, max_column);
    var file_name = submission.title = template.name = file.getName();
    submit_range.setValues([[submission_date, file_name, submission_type, 0, file_url]])
    sendemail_submission(submission_type, submission);
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
  var arr = [form.officer_start, form.officer_end, form.TCS_start, form.TCS_end];
  if (arr.indexOf("") > -1){
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
  var ss = get_active_spreadsheet();
  var chapterName = get_chapter_name();
  var date = new Date();
  var formatted = (date.getMonth() + 1) + '-' + date.getDate() + '-' +
                  date.getFullYear() + ' ' + date.getHours() + ':' +
                  date.getMinutes() + ':' + date.getSeconds();
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
    var row = ["N/A", formatted, chapterName, key, start, end, member_object["Badge Number"][0],
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

function process_init(form) {
//  var form = {"badge": ["109 ($20)", "109 ($20)", "106 ($67)", "102 ($165)", "109 ($20)", "102 ($165)", "102 ($165)"],
//              "reason": ["Lost interest", "Too much time required", "Voluntarily decided not to continue"],
//              "name_init": ["Nicholas Larson", "David Montgome...", "Ryan Richard", "Justine Saugen",
//                            "Mark Silvern", "Monica Sproul", "Daniel Tranfag..."],
//              "testB": ["1", "2", "3", "4", "5", "6", "7"],
//              "date_init": "2016-08-01", "name_depl": ["Esgar Moreno", "Adam Schilpero...", "Jessyca Thomas"],
//              "guard": ["None", "Goldgloss & Plain", "Goldgloss & Chased/Engraved",
//                        "10k Gold & Chased/Engraved", "10k Gold & Crown Set Pearl",
//                        "10k Gold & Close Set Pearl", "Goldgloss & Plain"],
//              "roll": ["1", "2", "3", "4", "5", "6", "7"],
//              "GPA": ["1", "2", "3", "4", "5", "6", "7"],
//              "date_grad": ["2016-08-01", "2016-08-01", "2015-08-01", "2016-08-01",
//                            "2016-08-01", "2016-08-01", "2016-08-01"],
//              "testA": ["1", "2", "3", "4", "5", "6", "7"],
//              "date_depl": ["2015-08-01", "2016-08-02", "2016-08-03"]}
  Logger.log(form);
//  return;
  var MemberObject = main_range_object("Membership");
  var INIT = [header_INIT()];
  var DEPL = [header_DEPL()];
  var date = new Date();
  var date_init = form["date_init"];
  var ss = get_active_spreadsheet();
  if (date_init == ""){
      ss.toast('You must set the initiation date!', 'ERROR', 5);
      return [false, "date_init"];
    }
  date_init = format_date(date_init);
  var chapterName = get_chapter_name();
  var formatted = (date.getMonth() + 1) + '-' + date.getDate() + '-' +
                  date.getFullYear() + ' ' + date.getHours() + ':' +
                  date.getMinutes() + ':' + date.getSeconds();
  var init_count = 0;
  var depl_count = 0;
  if (typeof form["name_depl"] === 'string'){
    for (var obj in form){
      form[obj] = [form[obj]];
    }
  }
  if (typeof form["name_init"] === 'string'){
    for (var obj in form){
      form[obj] = [form[obj]];
    }
  }
  Logger.log(form);
  if (form["name_init"] !== undefined){
    for (var i = 0; i < form["name_init"].length; i++){
      var name = form["name_init"][i];
      var member_object = find_member_shortname(MemberObject, name);
      var first = member_object["First Name"][0];
      var last = member_object["Last Name"][0];
      var date_grad = form["date_grad"][i];
      var roll = form["roll"][i];
      var GPA = form["GPA"][i];
      var testA = form["testA"][i];
      var testB = form["testB"][i];
      var badge = form["badge"][i];
      var guard = form["guard"][i];
      var arr = [date_grad, roll, GPA, testA, testB];
      if (arr.indexOf("") > -1){
        ss.toast('You must set all of the fields!\nMissing information for:\n'
                 +name, 'ERROR', 5);
        return [false, name];
      }
      date_grad = format_date(date_grad);
      INIT.push(["N/A", formatted, date_init, chapterName,
                 date_grad, roll, first, "",
                 last, GPA, testA,
                 testB, "Initiation Fee", "Late Fee",
                 "Badge Style", "Guard Type", "Badge Cost", "Guard Cost", "Sum for member"]);
    }
  }
  if (form["name_depl"] !== undefined){
    for (var i = 0; i < form["name_depl"].length; i++){
      var name = form["name_depl"][i];
      var member_object = find_member_shortname(MemberObject, name);
      var first = member_object["First Name"][0];
      var last = member_object["Last Name"][0];
      var date_depl = form["date_depl"][i];
      var arr = [date_depl];
      if (arr.indexOf("") > -1){
        ss.toast('You must set all of the fields!\nMissing information for:\n'
                 +name, 'ERROR', 5);
        return [false, name];
      }
      date_depl = format_date(date_depl);
      var reason = form["reason"][i];
      DEPL.push(["N/A", formatted, chapterName, first, last, reason, date_depl]);
    }
  }
  Logger.log("INIT");
  Logger.log(INIT);
  var csvFile = create_csv(INIT);
  Logger.log(csvFile);
  var init_out = "";
  if (INIT.length > 1){
    init_out = save_form(csvFile, "INIT");
  }
  Logger.log("DEPL");
  Logger.log(DEPL);
  var csvFile = create_csv(DEPL);
  Logger.log(csvFile);
  var depl_out = ""
  if (DEPL.length > 1){
    depl_out = save_form(csvFile, "DEPL");
  }
    return [init_out+depl_out, null];
}

function process_grad(form) {
//  var form = {"date_start": ["2016-08-01", "2016-08-01", "2016-08-01",
//                "2016-08-01", "2016-08-01", "2016-08-01", "2016-08-01",
//                "2016-08-01", "2016-08-01", "2016-08-01", "2016-08-01"],
//              "new_location": ["Test Cole 1", "Test Austin 1", "Test Adam 1", "Test Adam 2",
//                 "Test Adam 3", "Test Derek", "Test Esgar", "Test Adam 4",
//                 "Test Cole 2", "Test Austin 2", "Test Adam 5"],
//              "phone": ["520-664-5654", "520-664-5654", "520-664-5654"],
//              "prealumn": ["Undergrad > 4 yrs", "Undergrad < 4 yrs"],
//              "name": ["Cole Mobberley", "Austin Mutschler", "Adam Schilperoort", "Adam Schilperoort",
//                       "Adam Schilperoort", "Derek Hogue", "Esgar Moreno", "Adam Schilperoort",
//                       "Cole Mobberley", "Austin Mutschler", "Adam Schilperoort"],
//              "degree": ["Cole Mobberley MAJOR", "Austin Mutschler MAJOR", "Adam Schilperoort MAJOR"],
//              "dist": ["> 60 mi", "> 60 mi", "< 60 mi", "> 60 mi", "> 60 mi"],
//              "date_end": ["2016-08-03", "2016-08-03", "2016-08-03", "2016-08-03", "2016-08-03"],
//              "type": ["Degree received", "Degree received", "Degree received", "Abroad",
//                       "Transfer", "PreAlumn", "PreAlumn", "Military", "CoOp", "CoOp", "CoOp"],
//              "email": ["ColeMobberley@email.com", "AustinMutschler@email.com", "AdamSchilperoort@email.com"]}
  Logger.log(form);
//  return;
  var MemberObject = main_range_object("Membership");
  var MSCR = [header_MSCR()];
//  var MSCR_type = ["Degree received", "Transfer", "PreAlumn"];
  var COOP = [header_COOP()];
  var date = new Date();
  var formatted = (date.getMonth() + 1) + '-' + date.getDate() + '-' +
                  date.getFullYear() + ' ' + date.getHours() + ':' +
                  date.getMinutes() + ':' + date.getSeconds();
//  var COOP_type = ["Abroad", "Military", "CoOp"]
  var degree_count = 0;
  var alum_count = 0;
  var nonalum_count = 0;
  if (typeof form["type"] === 'string'){
    for (var obj in form){
      form[obj] = [form[obj]];
    }
  }
  for (var i = 0; i < form["type"].length; i++){
    var type = form["type"][i];
    Logger.log(type);
    var name = form["name"][i];
    var member_object = find_member_shortname(MemberObject, name);
    var badge = member_object["Badge Number"][0];
    var first = member_object["First Name"][0];
    var last = member_object["Last Name"][0];
    var loc = form["new_location"][i]
    var date_start = form["date_start"][i];
    if (type == "Degree received"){
      var email = form["email"][degree_count];
      var phone = form["phone"][degree_count];
      var degree = form["degree"][degree_count];
      if (typeof email === 'string'){
          email = form["email"];
          phone = form["phone"];
          degree = form["degree"];
        }
      var arr = [loc, date_start, email, phone, degree];
      degree_count++
    } else if (type != "PreAlumn"){
      if (type != "Withdrawn" && type != "Transfer"){
        var date_end = form["date_end"][nonalum_count];
        var dist = form["dist"][nonalum_count];
        if (typeof date_end === 'string'){
          date_end =form["date_end"];
          dist = form["dist"];
        }
        var arr = [loc, date_start, date_end, dist];
        date_end = format_date(date_end);
        nonalum_count++
      }
    } else {
      var prealumn = form["prealumn"][alum_count];
      if (prealumn == "Undergrad < 4 yrs"){
        prealumn = "Undergrad Premature <4 years";
      } else if (prealumn == "Undergrad > 4 yrs"){
        prealumn = "Undergrad Premature > 4 years";
      } else if (prealumn == "Grad Premature"){
        prealumn = "Grad Student Premature";
      }
      alum_count++
    }
    date_start = format_date(date_start);
    if (arr.indexOf("") > -1){
      var ss = get_active_spreadsheet();
      ss.toast('You must set all of the fields!\nMissing information for:\n'
             +name, 'ERROR', 5);
      return [false, name];
    }
    switch (type) {
      case "Degree received":
        MSCR.push(["N/A", formatted, badge, first, last, phone, email,
                   "Graduated from school", degree, date_start, loc, "",
                   "", "", "", "", "", ""]);
        break;
      case "Transfer":
        MSCR.push(["N/A", formatted, badge, first, last, "", "",
                   "Transferring to another school", "",
                   "", "", "", "", "", "",
                   loc, date_start, ""]);
        break;
      case "Withdrawn":
        MSCR.push(["N/A", formatted, badge, first, last, "", "",
                   "Withdrawing from school", "", "", "", "", "",
                   "Yes", date_start, "", "", ""]);
        break;
      case "PreAlumn":
        MSCR.push(["N/A", formatted, badge, first, last, "", "",
                   "Wishes to REQUEST Premature Alum Status", "",
                   "", "", "", "", "", "", "", "", prealumn]);
        break;
      case "Abroad":
        COOP.push(["N/A", formatted, badge, first, last,
                   "Study Abroad", date_start,
                   date_end, dist]);
        break;
      case "Military":
        COOP.push(["N/A", formatted, badge, first, last,
                   "Called to Active/Reserve Military Duty",
                   date_start, date_end, dist]);
        break;
      case "CoOp":
        COOP.push(["N/A", formatted, badge, first, last,
                   "Co-Op/Internship",
                   date_start, date_end, dist]);
        break;
    }
  }
  Logger.log("COOP");
  Logger.log(COOP);
  var csvFile = create_csv(COOP);
  Logger.log(csvFile);
  var coop_out = "";
  if (COOP.length > 1){
    coop_out = save_form(csvFile, "COOP");
  }
  Logger.log("MSCR");
  Logger.log(MSCR);
  var csvFile = create_csv(MSCR);
  Logger.log(csvFile);
  var mscr_out = ""
  if (MSCR.length > 1){
    mscr_out = save_form(csvFile, "MSCR");
  }
    return [coop_out+mscr_out, null];
       }

function header_MSCR(){
  return ["Submitted by", "Date Submitted", "ChapRoll",
   "First Name", "Last Name", "Mobile Phone", "EmailAddress",
   "Reason for Status Change", "Degree Received",
   "Graduation Date (M/D/YYYY)", "Employer", "Work Email",
   "Attending Graduate School where ?", "Withdrawing from school?",
   "Date withdrawn (M/D/YYYY)", "Transferring to what school ?",
   "Date of transfer (M/D/YYYY)"];
}

function header_COOP(){
  return ["Submitted by", "Date Submitted", "*ChapRoll",
    "First Name", "Last Name", "Reason Away", "Start Date (M/D/YYYY)",
    "End Date (M/D/YYYY)", "Miles from Campus**"]
}

function header_INIT(){
  return ["Submitted by", "Date Submitted", "Initiation Date", "Chapter Name",
          "Graduation Year", "Roll Number", "First Name", "Middle Name",
          "Last Name", "Overall GPA", "A Pledge Test Scores",
          "B Pledge Test Scores", "Initiation Fee", "Late Fee",
          "Badge Style", "Guard Type", "Badge Cost", "Guard Cost", "Sum for member"];
}

function header_DEPL(){
  return ["Submitted by", "Date Submitted", "Chapter Name",
          "First Name", "Last Name", "Reason Depledged",
          "Date Depledged (M/D/YYYY)"];
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
  Logger.log(form);
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
    var ss = get_active_spreadsheet();
    var sheet = ss.getSheetByName("Submissions");
    var max_column = sheet.getLastColumn();
    var max_row = sheet.getLastRow();
    var submit_range = sheet.getRange(max_row + 1, 1, 1, max_column);
    var file_name = template.name = file.getName();
    submit_range.setValues([[submission_date, file_name, submission_type, 0, file_url]])
    update_scores_submit(max_row + 1);
    return template.evaluate().getContent();
  } catch (error) {
    var this_error = error.toString();
    Logger.log(this_error);
    return this_error;
  }
}

function post_submit(file_object, submission_type) {
  var template = HtmlService.createTemplateFromFile('SubmitFormResponse');
  var file_url = template.fileUrl = file_object.alternateLink;
  var submission_date = template.date = new Date();
  var submission_type = template.type = submission_type;
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Submissions");
  var max_column = sheet.getLastColumn();
  var max_row = sheet.getLastRow();
  var submit_range = sheet.getRange(max_row + 1, 1, 1, max_column);
  var file_name = template.name = file_object.title;
  submit_range.setValues([[submission_date, file_name, submission_type, 0, file_url]])
  update_scores_submit(max_row + 1);
  var output = template.evaluate().getContent();
  Logger.log(output);
  sendemail_submission(submission_type, file_object)
  return output;
}

function sendemail_submission(submission_type, submission) {
  var email_director = SCRIPT_PROP.getProperty("director");
  var email_chapter = SCRIPT_PROP.getProperty("email");
  var chapter = SCRIPT_PROP.getProperty("chapter");
  var subject = "Chapter Submission: "+chapter;
  if (submission.file){
    var file_obj = submission.file;
    var folder_id = submission.folder_id;
  } else {
    var file_id = submission.id;
    var folder_id = submission.parents[0].id;
    var file_obj = DriveApp.getFileById(file_id).getBlob();
  }
  var file_url = submission.alternateLink;
  var file_name = submission.title;
  var folder_url = DriveApp.getFolderById(folder_id).getUrl();
  Logger.log([email_director, email_chapter, chapter, subject, file_id,
             file_url, file_name, folder_id, folder_url]);

  var emailBody = "Chapter Submission: "+chapter+
    "\nSubmission Type: "+submission_type+
      "\nSubmission Location: " + folder_url +
        "\nFile Location: " + file_url +
          "\nSubmission Title: " + file_name +
            "\nPlease see attached.";

  var htmlBody = "Chapter Submission: "+chapter+
    "<br/>Submission Type: "+submission_type+
      "<br/>Submission Location: " + folder_url +
        "<br/>File Location: " + file_url +
          "<br/>Submission Title: " + file_name +
            "<br/>Please see attached.";

  var optAdvancedArgs = {name: chapter +" Chapter", htmlBody: htmlBody,
                         replyTo: email_chapter, attachments: [file_obj]};
  MailApp.sendEmail(email_director, subject, emailBody, optAdvancedArgs);
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
//  newArray.sort();
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

function get_event_list() {
  var event_object = main_range_object("Events");
  var event_list = [];
  for(var i = 0; i< event_object.object_count; i++) {
    var event_name = event_object.object_header[i];
    event_list.push(event_name);
    }
  return event_list
}

function getList(RangeName) {
  //' MemberNamesOnly
//  var RangeName = 'MemberNamesOnly'
//  var RangeName = 'EventTypes';submit_col_max
  Logger.log('Called getList, RangeName: ' + RangeName);
  var ss = get_active_spreadsheet();
  var events = ss
      .getRangeByName(RangeName)
      .getValues()
  var event_list = [].concat.apply([], events);
  var event_list = cleanArray(event_list);
  event_list.sort();
  Logger.log(event_list);
  return event_list;
}

//function onChange(e){
////  show_att_sheet_alert();
//  Logger.log("onChange");
//  Logger.log(e);
//  _onEdit(e);
////  var ss = get_active_spreadsheet();
////  var sheet = ss.getActiveSheet();
////  var sheet_name = sheet.getName();
////  if (sheet_name == "Events"){
////    if (e.changeType == "INSERT_ROW"){
////      Logger.log("EVENTS ROW ADDED");
////    }else if (e.changeType == "REMOVE_ROW"){
////      Logger.log("EVENTS ROW REMOVED");
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
//  var Logger = startBetterLog();
  try{
  Logger.log("onEDIT" + e);
  Logger.log(e);
  Logger.log("onEdit, authMode: " + e.authMode);
  Logger.log("onEdit, user: " + e.user);
  Logger.log("onEdit, source: " + e.source);
  Logger.log("onEdit, range: " + e.range);
  Logger.log("onEdit, value: " + e.value);
  var sheet = e.range.getSheet();
  var sheet_name = sheet.getName();
  var user_range = e.range
  var user_row = user_range.getRow();
  var user_col = user_range.getColumn();
  var user_old_value = e.oldValue
  Logger.log("Row: " + user_row + " Col: " + user_col);
  var this_password = SCRIPT_PROP.getProperty("password");
  if (sheet_name == "Events"){
    Logger.log("EVENTS CHANGED");
    if (user_row == 1 || user_col == 4 ||
        user_col == 5 || user_col == 6){
      reset_range(user_range, user_old_value)
      if (this_password == password){
        return;
      }
      var ui = SpreadsheetApp.getUi();
      var result = ui.alert(
        'ERROR',
        'Score, #Members, & #Pledges are are updated automatically',
        ui.ButtonSet.OK);
    } else {
    update_scores_event(user_row);
    }
//    show_event_sheet_alert();
//    align_event_attendance();
  } else if (sheet_name == "Attendance"){
    if (user_row == 1 || user_col < 3){
      reset_range(user_range, user_old_value);
      if (this_password == password){
        return;
      }
      show_att_sheet_alert();
    } else {
      var attendance = range_object(sheet, user_row)
      var header = attendance.object_header;
      var clean_header = cleanArray(header, 50);
      if (clean_header.length == header.length){
        update_attendance(attendance);
        update_scores_event(user_row);
      } else {
        return;
      }
    }
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
    submitSidebar();
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
    Logger.log("MEMBER CHANGED");
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
  }
  } catch (error) {
    Logger.log(error);
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
     'ERROR',
      error,
      ui.ButtonSet.OK);
    return "";
  }
}

function refresh() {
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Events");
  var max_rows = sheet.getLastRow();
  for (var user_row = 2; user_row < max_rows; user_row++){
    update_scores_event(user_row);
  }
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

function update_attendance(attendance){
  var MemberObject = main_range_object("Membership");
//  Logger.log(attendance);
  var counts = {};
  counts["Student"] = {};
  counts["Pledge"] = {};
  var test_len = attendance.object_count;
  for(var i = 2; i< attendance.object_count; i++) {
    var member_name_att = attendance.object_header[i];
    var member_name_short = att_name(attendance.object_header[i]);
    var member_object = find_member_shortname(MemberObject, member_name_short);
    var event_status = attendance[member_name_att][0];
    event_status = event_status.toUpperCase();
    var member_status = member_object["Chapter Status"][0]
//    Logger.log([member_name_short, member_object, event_status, member_status]);
    counts[member_status][event_status] = counts[member_status][event_status] ? counts[member_status][event_status] + 1 : 1;
  }
  Logger.log(counts)
  var event_name_att = attendance["Event Name"][0];
  var event_date_att = attendance["Event Date"][0];
  Logger.log(event_name_att);
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Events");
  var EventObject = main_range_object("Events");
  for (var i = 0; i < EventObject.object_count; i++){
    var event_name = EventObject.object_header[i];
    var event_date = EventObject[event_name]["Date"][0];
    Logger.log([event_name+event_date, event_name_att+event_date_att]);
    if (event_name+event_date == event_name_att+event_date_att){
      var active_col = EventObject[event_name]['# Members'][1];
      var pledge_col = EventObject[event_name]['# Pledges'][1];
      var event_row = EventObject[event_name].object_row;
      break;
    }
  }
//  var max_column = sheet.getLastColumn();
//  var event_headers = sheet.getRange(1, 1, 1, max_column);
//  var header_values = event_headers.getValues()
//  var active_col = get_ind_from_string("# Members", header_values)
//  var pledge_col = get_ind_from_string("# Pledges", header_values)
//  var event_row = attendance.object_row;
  Logger.log("ROW: " + event_row + " Active: " + active_col + " Pledge: " + pledge_col)
  var active_range = sheet.getRange(event_row, active_col)
  var pledge_range = sheet.getRange(event_row, pledge_col)
  var num_actives = counts["Student"]["P"] ? counts["Student"]["P"]:0;
  var num_pledges = counts["Pledge"]["P"] ? counts["Pledge"]["P"]:0;
  active_range.setValue(num_actives)
  pledge_range.setValue(num_pledges)
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
  var field_range = sheet.getRange(event_row, 10, 1, 5);
  field_range.setBackground("black")
             .setNote("Do not edit");
  var needed_field_values = [];
  Logger.log(needed_fields);
  var yes_no_fields = ['STEM?', 'PLEDGE Focus', 'HOST'];
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
  Logger.log(needed_field_values);
  if (needed_field_values.indexOf("") > -1){
    return false;
  }
  return true;
}

function update_scores_event(user_row){
//  var user_row = 2;
  var myObject = range_object("Events", user_row);
  if (myObject.Type[0] == "" || myObject.Date[0] == "" ||
      myObject["Event Name"][0] == ""){
    return;
  } else if (myObject["# Members"][0] === ""){
    attendance_add_event(myObject["Event Name"][0], myObject.Date[0]);
  }
  if (!event_fields_set(myObject)){
    return;
  }
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

function update_service_hours(){
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Membership");
  var EventObject = main_range_object("Events");
  var MemberObject = main_range_object("Membership");
  var AttendanceObject = main_range_object("Attendance");
  var score_obj = {};
  for (var i = 0; i < EventObject.object_count; i++){
    var event_name = EventObject.object_header[i];
    var event_type = EventObject[event_name]["Type"][0];
    if (event_type == "Service Hours"){
      var event_hours = EventObject[event_name]["Event Hours"][0];
      var event_date = EventObject[event_name]["Date"][0];
      var month = event_date.getMonth();
      var semester = "FALL";
      if (month<5){
	    var semester = "SPRING";
      }
      var att_obj = AttendanceObject[event_name];
      for (var j = 2; j < att_obj.object_count; j++){
        var member_name_raw = AttendanceObject.header_values[j];
        var member_name_short = att_name(member_name_raw);
        var member_object = find_member_shortname(MemberObject, member_name_short);
        var member_name = member_object["Member Name"][0];
        var att = att_obj[member_name_raw][0];
        if (att == "P"){
          score_obj[member_name] = score_obj[member_name] ? score_obj[member_name]:{};
          score_obj[member_name][semester] = score_obj[member_name][semester] ?
            score_obj[member_name][semester]+event_hours:event_hours;
        }
//        Logger.log(score_obj);
      }
    }
  }
  for (var member_name in score_obj){
//    Logger.log(member_name);
    var member_obj = MemberObject[member_name];
    var member_row = member_obj.object_row;
    var fall_col = member_obj["Service Hours Fall"][1];
    var spring_col = member_obj["Service Hours Spring"][1];
    var member_fall_range = sheet.getRange(member_row, fall_col);
    var member_spring_range = sheet.getRange(member_row, spring_col);
    var fall_score = score_obj[member_name]["FALL"] ? score_obj[member_name]["FALL"]:0;
    member_fall_range.setValue(fall_score);
    var spring_score = score_obj[member_name]["SPRING"] ? score_obj[member_name]["SPRING"]:0;
    member_spring_range.setValue(spring_score);
    Logger.log("FALL: "+fall_col+" SPRING: "+spring_col+" ROW: "+member_row);
  }
  update_scores_org_gpa_serv();
}

function update_score_att(){
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Scoring");
  var EventObject = main_range_object("Events");
  var ScoringObject = main_range_object("Scoring");
  var total_members = get_total_members()["Student"];
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
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Scoring");
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

function update_scores_org_gpa_serv(){
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Scoring");
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
  var service_fall_range = sheet.getRange(ScoringObject["Service Hours"].object_row, fall_col);
  var service_spring_range = sheet.getRange(ScoringObject["Service Hours"].object_row, spring_col);
  var service_range = sheet.getRange(ScoringObject["Service Hours"].object_row, total_col);
  var service_method_raw = ScoringObject["Service Hours"]["Special"][0];
  var service_fall_method = service_method_raw.replace("HOURS", score_data.percent_service_fa);
  var service_spring_method = service_method_raw.replace("HOURS", score_data.percent_service_sp);
  var service_max = ScoringObject["Service Hours"]["Max/ Semester"][0];
  var service_fall_score = eval_score(service_fall_method, service_max);
  var service_spring_score = eval_score(service_spring_method, service_max);
  Logger.log("SOC: " + societies_method + ", SCORE: " + socieities_score);
  Logger.log("GPA_FALL: " + gpa_fall_method + ", SCORE: " + gpa_fall_score);
  Logger.log("GPA_SPRING: " + gpa_spring_method + ", SCORE: " + gpa_spring_score);
  Logger.log("SERV_FALL: " + service_fall_method + ", SCORE: " + service_fall_score);
  Logger.log("SERV_SPRING: " + service_spring_method + ", SCORE: " + service_spring_score);
  societies_range.setValue(socieities_score);
  gpa_fall_range.setValue(gpa_fall_score);
  gpa_spring_range.setValue(gpa_spring_score);
  gpa_range.setValue(gpa_fall_score + gpa_spring_score);
  service_fall_range.setValue(service_fall_score);
  service_spring_range.setValue(service_spring_score);
  service_range.setValue(service_fall_score + service_spring_score);
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
  var service_count_fa = 0;
  var service_count_sp = 0;
  var officer_count = 0;
  var org_count = 0;
  var officers = ["Officer (Pro/Tech)", "Officer (Honor)", "Officer (Other)"];
  var orgs = ["Professional/ Technical Orgs", "Honor Orgs", "Other Orgs"];
  var gpas = ["Fall GPA", "Service Hours Fall", "Spring GPA"];
  var MemberObject = main_range_object("Membership");
  var gpa = 0;
  for (var i = 0; i < MemberObject.object_count; i++){
    var member_name = MemberObject.object_header[i];
    var org_true = false;
    var officer_true = false;
    for (var j = 0; j <= 2; j++){
      var gpa_raw = MemberObject[member_name][gpas[j]][0]
      gpa_raw = gpa_raw == "" ? 0:gpa_raw;
      var gpa = parseInt(gpa_raw);
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
    var service_hours_fa = MemberObject[member_name]["Service Hours Fall"][0];
    var service_hours_sp = MemberObject[member_name]["Service Hours Spring"][0];
    var service_hours_self_fa = MemberObject[member_name]["Self Service Hrs FA"][0];
    var service_hours_self_sp = MemberObject[member_name]["Self Service Hrs SP"][0];
    service_hours_fa = +service_hours_fa + service_hours_self_fa
    service_hours_sp = +service_hours_sp + service_hours_self_sp
    var service_count_fa = service_hours_fa >= 8 ? service_count_fa + 1:service_count_fa;
    var service_count_sp = service_hours_sp >= 8 ? service_count_sp + 1:service_count_sp;
    officer_count = officer_true ? officer_count + 1:officer_count;
    org_count = org_true ? org_count + 1:org_count;
  }
  var percent_service_fa = service_count_fa / MemberObject.object_count;
  var percent_service_sp = service_count_sp / MemberObject.object_count;
  var percent_org = org_count / MemberObject.object_count;
  var gpa_avg_fall = gpa_counts["Fall GPA"] / MemberObject.object_count;
  var gpa_avg_spring = gpa_counts["Spring GPA"] / MemberObject.object_count;
  return {percent_service_fa: percent_service_fa,
          percent_service_sp: percent_service_sp,
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
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName(sheetName);
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
  score_range.setNote(score_data.score_method_note);
  var score = score_data.score;
  if (score === null){
    score_range.setBackground("black");
    return [];
  } else {
    score_range.setBackground("dark gray 1");
  }
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
  return other_type_rows;
}

function update_main_score(score_data){
  Logger.log(score_data);
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Scoring");
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
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Scoring");
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
  var sheet = ss.getSheetByName("Dashboard");
  var RangeName = "SCORE" + "_" + score_type.toUpperCase();
  var dash_score_range = ss.getRangeByName(RangeName);
  dash_score_range.setValue(total);
}

function get_current_scores(sheetName){
//  var sheetName = "Events";
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName(sheetName);
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
  var score = null
  if (score_method_edit !== null){
    score = eval(score_method_edit);
    score = score.toFixed(1);
  }
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
      var percent_attend = attend / totals.Student;
      score_method = score_method.replace("memberATT", percent_attend);
          }
  if (~score_method.indexOf("memberADD")){
      score_method = score_method.replace("memberADD", attend);
          }
  if (~score_method.indexOf("MILES")){
    var miles = myEvent["MILES"][0];
    miles = (miles != "") ? miles:0;
    score_method = score_method.replace("MILES", miles);
          }
  if (~score_method.indexOf("HOST")){
    var host = myEvent["HOST"][0];
    host = (host == "Yes") ? 1:0;
    score_method = score_method.replace("HOST", host);
    score_method = score_method.replace("HOST", host);
    score_method = score_method.replace("HOST", host);
          }
  if (~score_method.indexOf("NON-MEMBER")){
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
    update_score_att();
    return null;
          }
  if (~score_method.indexOf("HOURS")){
    update_service_hours();
    return null;
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

function main_range_object(sheetName, short_header, ss){
//  var sheetName = "Membership"
//  var sheetName = "Scoring"
//  var sheetName = "Events"
  if (!ss){
    var ss = get_active_spreadsheet();
  }
  var sheet = ss.getSheetByName(sheetName);
  if (sheetName=="Membership" || sheetName=="MAIN"){
    if (short_header == undefined){
      var short_header = "Member Name";
      }
    var sort_val = short_header;
  } else if (sheetName=="Scoring"){
    var short_header = "Short Name";
    var sort_val = short_header;
  } else if (sheetName=="Events"){
    var short_header = "Event Name";
    var sort_val = "Date";
  } else if (sheetName=="Attendance"){
    var short_header = "Event Name";
    var sort_val = "Event Date";
  }
  var max_row = sheet.getLastRow()-1;
  Logger.log("MAX_"+sheetName+": "+max_row);
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
  myObject["header_values"] = header_values[0];
  for (var val in short_names){
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
  var ss = get_active_spreadsheet();
  if (typeof sheet === "string"){
    var sheet = ss.getSheetByName(sheet);
  }
  var max_column = sheet.getLastColumn()
  var range = sheet.getRange(range_row, 1, 1, max_column);
  var range_values = range.getValues()
  Logger.log(range_values)
  var header_range = sheet.getRange(1, 1, 1, max_column);
  var header_values = header_range.getValues();
  Logger.log(header_values)
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
  var range = sheet.getRange(3, 3, 1, 1);
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

function show_att_sheet_alert(){
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'ERROR',
     'Please edit the events or members on the Events or Membership Sheet',
      ui.ButtonSet.OK);
}

function get_sheet_data(SheetName) {
//  var SheetName="Events"
//  var SheetName="Attendance"
  Logger.log(SheetName)
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName(SheetName);
  if (sheet != null) {
    var max_row = sheet.getLastRow();
    var max_row = (max_row != 0) ? max_row:1;
    var max_column = sheet.getLastColumn()
    var range = sheet.getRange(2, 1, max_row, max_column);
    var header_range = sheet.getRange(1, 1, 1, max_column);
    var header_values = header_range.getValues();
//    Logger.log(header_values);
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
  return {sheet: sheet,
          range: range,
          max_row: max_row,
          header: header_values,
          date_index: date_index,
          name_index: name_index,
          max_column: max_column
         }
}

function attendance_add_event(event_name, event_date){
  //align_attendance_events(myObject["Event Name"][0], myObject.Date[0])
//  var event_name = "TEST";
//  var event_date = "01/01/2018";
  if (!event_name || !event_date){
    return;
    var event_data = get_sheet_data("Events");
    var event_values = event_data.range.getValues();
  }
  var att_data = get_sheet_data("Attendance");
  var sheet = att_data.sheet;
//  var att_values = att_data.range.getValues();
  var attendance_rows = att_data.max_row;
  var attendance_cols = att_data.max_column;
  Logger.log(attendance_rows);
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
//  for (row in event_values){
//    var this_row = parseInt(row) + 1
//    var event_name = event_values[row][event_data.name_index];
//    var event_date = event_values[row][event_data.date_index];
//    if (this_row - 1 < attendance_rows){
//      var att_event_name = att_values[row][att_data.name_index];
//      var att_event_date = att_values[row][att_data.date_index];
//    }
//    if (event_name != att_event_name){
//      sheet.insertRowAfter(this_row);
//      var name_range = sheet.getRange(this_row+1, att_data.name_index+1);
//      name_range.setValue(event_name);
//      var date_range = sheet.getRange(this_row+1, +att_data.date_index+1);
//      date_range.setValue(event_date);
//      sheet.setRowHeight(this_row+1, 10);
//    }
//    Logger.log(event_name);
//  }
}