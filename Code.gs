/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
var betterLogStarted = false;
var SCRIPT_PROP = PropertiesService.getScriptProperties();
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

function setup() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
    'Setup Sheet ID',
    'What is your sheet id?\n'+
    'Located in the URL at SHEETID, see example below:\n'+
    'https://docs.google.com/spreadsheets/d/SHEETID/edit#gid=0',
    ui.ButtonSet.OK_CANCEL);
  var button = result.getSelectedButton();
  var text = result.getResponseText();
//  var doc = SpreadsheetApp.getActiveSpreadsheet();
  if (button == ui.Button.OK) {
    ui.alert('Verify you sheet id is: ' + text);
  } else {
    // User clicked "Cancel".
    ui.alert('The scripts will not work without the Sheet ID.');
  }
  SCRIPT_PROP.setProperty("key", text);
}

function setup_sheets(chapter_name) {
//  var chapter_name = "Chi";
  var default_id = "19aWLtjJJ-Uh6XOqOuseLpQcNJYslQHe9Y9Gaj2vSjEw";
  var default_doc = SpreadsheetApp.openById(default_id);
  var target_doc = get_active_spreadsheet();
  var sheets = default_doc.getSheets();
  for (var i in sheets){
    var sheet = sheets[i];
    var sheet_name = sheet.getName();
    sheet.copyTo(target_doc).setName(sheet_name);
  }
  var sheet = target_doc.getSheetByName("Sheet1");
  target_doc.deleteSheet(sheet);
  var doc_name = default_doc.getName();
  doc_name = doc_name.replace("DEFAULT ", ""); //DEFAULT Theta Tau Chapter Report - Chapter
  doc_name = doc_name.replace("- Chapter", "- "+chapter_name);
  Logger.log(doc_name);
  target_doc.rename(doc_name);
  var sheet = target_doc.getSheetByName("Dashboard"); //A1
  //? CHAPTER ANNUAL REPORT
  var sheet = target_doc.getSheetByName("Chapter"); //B2
  var named_ranges = default_doc.getNamedRanges();
  for (var j in named_ranges){
    var named_range = named_ranges[j];
    var name = named_range.getName();
    var range = named_range.getRange();
    var sheet = range.getSheet().getSheetName();
    var old_range = range.getA1Notation();
    Logger.log(old_range);
    var new_sheet = target_doc.getSheetByName(sheet);
    var new_range = new_sheet.getRange(old_range);
    Logger.log(name);
    target_doc.setNamedRange(name, new_range);
  }
}

function RESET() {
  var target_doc = get_active_spreadsheet();
  var sheets = target_doc.getSheets();
  var sheet = target_doc.insertSheet();
  sheet.setName("Sheet1");
  for (var i in sheets){
    var sheet = sheets[i];
    target_doc.deleteSheet(sheet)
  }
  var named_ranges = target_doc.getNamedRanges();
  for (var j in named_ranges){
    var named_range = named_ranges[j];
    named_range.remove();
  }
  target_doc.rename("BLANK");
}

function get_active_spreadsheet() {
  var doc = SpreadsheetApp.openById(SCRIPT_PROP.getProperty("key"));
  return doc
}

function onInstall(e) {
  onOpen(e);
  setup();
  var chapter_name = "Chi";
  setup_sheets(chapter_name);
  get_chapter_members(chapter_name);
  createTriggers();
}

function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  menu.addItem('Update Members', 'get_chapter_members');
  menu.addItem('Event Functions', 'showaddEvent');
  menu.addItem('Update Officers', 'officerSidebar');
  menu.addItem('Submit Item', 'submitSidebar');
  menu.addItem('Status Change', 'member_update_sidebar');
  menu.addItem('Pledge Forms', 'pledge_sidebar');
  menu.addItem('Create Triggers', 'createTriggers');
  menu.addItem('SETUP', 'onInstall');
  menu.addItem("TEST", 'TEST');//test_onEdit
  menu.addItem("RESET", 'RESET');
  menu.addToUi();
}

function createTriggers() {
  var ss = get_active_spreadsheet();
  ScriptApp.newTrigger('_onEdit')
      .forSpreadsheet(ss)
      .onEdit()
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
    var folder = DriveApp.getFolderById('0BwvK5gYQ6D4nWVhUVlo4dUhYV0E');
    var ss = get_active_spreadsheet();
    var chapterName = ss.getRangeByName("ChapterName").getValue();
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
    var sheet = ss.getSheetByName("Submissions");
//    SpreadsheetApp.setActiveSheet(sheet);
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
  var chapterName = ss.getRangeByName("ChapterName").getValue();
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
  var chapterName = ss.getRangeByName("ChapterName").getValue();
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
//  var RangeName = 'EventTypes';
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

function _onEdit(e){
  var Logger = startBetterLog();
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

function update_attendance(attendance){
  var MemberObject = main_range_object("Membership");
  Logger.log(attendance);
  var counts = {};
  counts["Student"] = {};
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
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Events");
  var max_column = sheet.getLastColumn();
  var event_headers = sheet.getRange(1, 1, 1, max_column);
  var header_values = event_headers.getValues()
  var active_col = get_ind_from_string("# Members", header_values)
  var pledge_col = get_ind_from_string("# Pledges", header_values)
  var event_row = attendance.object_row;
  Logger.log("ROW: " + event_row + " Active: " + active_col + " Pledge: " + pledge_col)
  var active_range = sheet.getRange(event_row, active_col)
  var pledge_range = sheet.getRange(event_row, pledge_col)
  active_range.setValue(counts["Student"]["P"])
  pledge_range.setValue(counts["Pledge"]["P"])
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
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Scoring");
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

function main_range_object(sheetName, short_header){
//  var sheetName = "Membership"
//  var sheetName = "Scoring"
//  var sheetName = "Events"
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (sheetName=="Membership"){
    if (short_header == undefined){
      var short_header = "Member Name";
    }
  } else if (sheetName=="Scoring"){
    var short_header = "Short Name"
  } else if (sheetName=="Events"){
    var short_header = "Event Name"
  }
  var max_row = sheet.getLastRow() - 1;
  var max_column = sheet.getLastColumn();
  var header_range = sheet.getRange(1, 1, 1, max_column);
  var header_values = header_range.getValues();
  var short_names_ind = get_ind_from_string(short_header, header_values);
  if (max_row > 2){
    var full_data_range = sheet.getRange(2, 1, max_row, max_column);
    var full_data_values = full_data_range.getValues();
    var sorted_range = full_data_range.sort({column: short_names_ind, ascending: true});
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
  var range = sheet.getRange(1, 1, 1, 1);
  var value = range.getValue();
  onEdit({
    user : Session.getActiveUser().getEmail(),
    source : ss,
    range : range, //ss.getActiveCell(),
    value : value, //ss.getActiveCell().getValue(),
    authMode : "LIMITED"
  });
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
     'ERROR',
     'Value: '+
      value,
      ui.ButtonSet.OK);
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

function CSVToArray( strData, strDelimiter ){
  // Check to see if the delimiter is defined. If not,
  // then default to comma.
  strDelimiter = (strDelimiter || ",");
  // Create a regular expression to parse the CSV values.
  var objPattern = new RegExp(
    (
      // Delimiters.
      "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +
      // Quoted fields.
      "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +
      // Standard fields.
      "([^\"\\" + strDelimiter + "\\r\\n]*))"
    ),
    "gi"
  );

  // Create an array to hold our data. Give the array
  // a default empty first row.
  var arrData = [[]];

  // Create an array to hold our individual pattern
  // matching groups.
  var arrMatches = null;

  // Keep looping over the regular expression matches
  // until we can no longer find a match.
  while (arrMatches = objPattern.exec( strData )){

    // Get the delimiter that was found.
    var strMatchedDelimiter = arrMatches[ 1 ];

    // Check to see if the given delimiter has a length
    // (is not the start of string) and if it matches
    // field delimiter. If it does not, then we know
    // that this delimiter is a row delimiter.
    if (
      strMatchedDelimiter.length &&
      (strMatchedDelimiter != strDelimiter)
    ){
      // Since we have reached a new row of data,
      // add an empty row to our data array.
      arrData.push( [] );
    }


    // Now that we have our delimiter out of the way,
    // let's check to see which kind of value we
    // captured (quoted or unquoted).
    if (arrMatches[ 2 ]){

      // We found a quoted value. When we capture
      // this value, unescape any double quotes.
      var strMatchedValue = arrMatches[ 2 ].replace(
        new RegExp( "\"\"", "g" ),
        "\""
      );

    } else {

      // We found a non-quoted value.
      var strMatchedValue = arrMatches[ 3 ];
    }
    // Now that we have our value string, let's add
    // it to the data array.
    arrData[ arrData.length - 1 ].push( strMatchedValue );
  }
  // Return the parsed data.
  return( arrData );
}

function get_chapter_members(chapter_name){
//  var chapter_name = "Chi";
  var folder = DriveApp.getFolderById('0BwvK5gYQ6D4nOXB2UHFUV0w5WnM');
  var files = folder.getFiles();
  var old_date = new Date(2000, 01, 01);
  var new_file = null;
  while (files.hasNext()) {
    var file = files.next();
    var file_name = file.getName()
    Logger.log(file_name);
    var date_str = file_name.split("_")[0];
    var year = date_str.substring(0, 4);
    var month = date_str.substring(4, 6);
    var day = parseInt(date_str.substring(6, 8));
    if (day === NaN){
      continue;
    }
    var date = new Date(year, month, day);
    if (date > old_date){
      old_date = date;
      new_file = file;
    }
  }
  var csvFile = new_file.getBlob().getDataAsString();
  var csvData = CSVToArray(csvFile, ",");
//  Logger.log(csvData);
  var header = csvData[0];
  var chapter_index = header.indexOf("Constituent Specific Attributes Chapter Name Description");
  var status_index = header.indexOf("Constituency Code");
  var badge_index = header.indexOf("Constituent ID");
  var last_index = header.indexOf("Last Name");
  var first_index = header.indexOf("First Name");
  var email_index = header.indexOf("Email Address Number");
  var phone_index = header.indexOf("Mobile Phone Number");
//  var role_index = header.indexOf("Constituent Specific Attributes Chapter Name Description");
//  var major_index = header.indexOf("Constituent Specific Attributes Chapter Name Description");
//  var school_index = header.indexOf("Constituent Specific Attributes Chapter Name Description");
  var CentralMemberObject = {};
  CentralMemberObject['badge_numbers'] = [];
  for (var j in csvData){
    var row = csvData[j];
    var chapter_row = row[chapter_index];
    if (chapter_row == chapter_name){
      var member_object={};
      var badge_number = row[badge_index];
      member_object['First Name'] = row[first_index];
      member_object['Last Name'] = row[last_index];
      member_object['Badge Number'] = badge_number;
      member_object['Chapter Status'] = row[status_index];
      member_object['Chapter Role'] = row[last_index]+"_ROLE";//row[];
      member_object['Current Major'] = row[last_index]+"_MAJOR";//row[];
      member_object['School Status'] = row[last_index]+"_STATUS";//row[];
      member_object['Phone Number'] = row[phone_index];
      member_object['Email Address'] = row[email_index];
      CentralMemberObject['badge_numbers'].push(badge_number);
      CentralMemberObject[badge_number] = member_object;
    }
  }
  Logger.log(CentralMemberObject[badge_number]);
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Membership");
  var max_row = sheet.getLastRow() - 1;
  var max_row = (max_row != 0) ? max_row:1;
  var max_column = sheet.getLastColumn();
  var header_range = sheet.getRange(1, 1, 1, max_column);
  var header_values = header_range.getValues()[0];
  var badge_index_chapter = header_values.indexOf("Badge Number");
  var ChapterMemberObject = main_range_object("Membership", "Badge Number");
  Logger.log(ChapterMemberObject["object_header"]);
  var new_members = [];
  for (var k in CentralMemberObject['badge_numbers']){
    var badge_number = CentralMemberObject['badge_numbers'][k];
    if (ChapterMemberObject["object_header"].indexOf(badge_number) < 0){
      // Member is on Central list, not on chapter list
      Logger.log("NEW MEMBER!");
      Logger.log(CentralMemberObject[badge_number]['Last Name']);
      new_members.push(badge_number);
    }
  }
//  
  var old_members = [];
  for (var k in ChapterMemberObject["object_header"]){
    var badge_number = ChapterMemberObject["object_header"][k];
    if (CentralMemberObject['badge_numbers'].indexOf(badge_number) < 0){
      // Member is on chapter list, not on central list
      Logger.log("OLD MEMBER!");
      Logger.log(ChapterMemberObject[badge_number]['Last Name'][0]);
      old_members.push(badge_number)
    }
  }
  new_members.sort();
  new_members.reverse();
  for (var m in new_members){
    var this_row = 1;
    var previous_member = undefined;
    var new_badge = new_members[m];
    for (var k in ChapterMemberObject["object_header"]){
//      Logger.log(ChapterMemberObject["object_header"]);
      var badge_number = ChapterMemberObject["object_header"][k];
      var badge_next = ChapterMemberObject["object_header"][+k+1];
      badge_next = badge_next ? badge_next:new_badge+1;
//      Logger.log([badge_number, new_badge, badge_next]);
      if (new_badge > badge_number && new_badge < badge_next){
        this_row = ChapterMemberObject[badge_number]['object_row'];
        previous_member = ChapterMemberObject[badge_number]["Member Name"];
        break;
      }
    }
    sheet.insertRowAfter(this_row);
    var range = sheet.getRange(this_row+1, 1, 1, 10);
    var member_object = CentralMemberObject[new_badge];
    var new_values = [];
    for (var j in ChapterMemberObject["header_values"]){
      var header = ChapterMemberObject["header_values"][j];
      if (header == "Member Name"){
        var full_name = member_object["First Name"]+" "+member_object["Last Name"];
        new_values.push(full_name);
        continue;
      }
      var new_value = member_object[header];
      if ( new_value !== undefined){
        new_values.push(new_value);
      }
    }
    Logger.log(new_values);
    range.setValues([new_values]);
    align_attendance_members(previous_member, full_name);
  }
}

function get_event_data(SheetName) {
//  var SheetName="Events"
//  var SheetName="Attendance"
  Logger.log(SheetName)
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName(SheetName);
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

function align_attendance_members(previous_member, new_member){
//  var previous_member = "REALLYLONGNAMEFORTESTINGTHINGSLIKETHIS";
//  var new_member = "REALL3YLONGNAMEFORTESTINGTHINGSLIKETHIS";
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Attendance");
  var max_row = sheet.getLastRow() - 1;
  max_row = (max_row != 0) ? max_row:1;
  var max_column = sheet.getLastColumn();
//  var range = sheet.getRange(1, 3, max_row, max_column);
//  var range_values = range.getValues();
//  var sorted_range = sortHorizontal(range_values);
  var header_range = sheet.getRange(1, 1, 1, max_column);
  var header_values = header_range.getValues()[0];
  var previous_index = 2;
  if (previous_member !== undefined){
    previous_member = shorten(previous_member, 12, false);
    for (var i in header_values){
      var header_name = header_values[i];
      if (header_name == "Event Name" || header_name == "Event Date"){
        continue;
      }
      var new_string = "";
      for (var j = 0; j < header_name.length; j++){
        var char = header_name[j];
        if (j % 2 == 0){
          new_string = new_string.concat(char);
        }
      }
      if (new_string == previous_member){
        previous_index = parseInt(i)+1;
      }
    }
  }
  sheet.insertColumnAfter(previous_index);
  var new_range = sheet.getRange(1, +previous_index+1);
  new_member = shorten(new_member, 12, false);
  var formula = '=regexreplace("'+new_member+'", "(.)", "$1"&char(10))';
  Logger.log(formula);
  new_range.setFormula(formula);
  var val = new_range.getValue();
  val = val.substring(0, val.length - 1);
  new_range.setValue(val);
  sheet.setColumnWidth(+previous_index+1, 21)
  var format_range = ss.getRangeByName("FORMAT");
  format_range.copyFormatToRange(sheet, +previous_index+1, +previous_index+1, 2, 100);
  new_range = sheet.getRange(1, +previous_index+1, 100, 1);
  new_range.clearDataValidations();
}

function align_attendance_events(){
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Attendance");
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