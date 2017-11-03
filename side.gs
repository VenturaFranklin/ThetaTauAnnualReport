function side_officers() {
  var template = HtmlService
      .createTemplateFromFile('side_officers');
  var htmlOutput = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Update Officers')
      .setWidth(500);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(htmlOutput);
}

function side_pledge(){
  var template = HtmlService
      .createTemplateFromFile('side_pledge');
  template.pledge = get_member_list("Pledge");
  var htmlOutput = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Update Pledges');
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(htmlOutput);
}

function side_member() {
  var template = HtmlService
      .createTemplateFromFile('side_member');
  var htmlOutput = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Update Members');
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(htmlOutput);
}

function side_survey() {
  Logger.log("(" + arguments.callee.name + ") ");
  var template = HtmlService
      .createTemplateFromFile('side_survey');
  var year_semesters = get_year_semesters();
  year_semesters = Object.keys(year_semesters);
  Logger.log(year_semesters);
  template.year_semesters = year_semesters;
  var htmlOutput = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Survey Members');
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(htmlOutput);
}

function get_submit_folders(submit_types){
//   var update_test = SCRIPT_PROP.getProperty('submit_folders');
//   if (!update_test){
  var properties_id = "1vCVKh8MExPxg8eHTEGYx7k-KTu9QUypGwbtfliLm58A";
  var ss_prop = SpreadsheetApp.openById(properties_id);
  var ss = get_active_spreadsheet();
  var main_object = main_range_object("ScoreInfo", "Short Name", ss_prop);
  var submit_folders = [];
  for (var i in submit_types){
    var submit_type = submit_types[i];
    var submit_folder = main_object[submit_type]["Submit Folder"][0];
    submit_folders.push(submit_folder);
  }
//     SCRIPT_PROP.setProperty('submit_folders', JSON.stringify(submit_folders));
//   } else {
//     var submit_folders = JSON.parse(update_test);
//   }
  return submit_folders
}

function side_submit() {
   var template = HtmlService
   .createTemplateFromFile('side_submit')
//  .createHtmlOutputFromFile('SubmitForm');
   var list_info = get_type_list('Submit', true);
  template.submissions = list_info.type_list;
  var submit_folders = get_submit_folders(list_info.type_list);
  template.submissions_folders = submit_folders;
  template.descriptions = list_info.type_desc;
  template.folder_id = get_submit_id();
  var chapter = get_chapter_name();
  var date = new Date();
  var currentMonth = date.getMonth() + 1;
  if (currentMonth < 10) { currentMonth = '0' + currentMonth; }
  var currentDay = date.getDate().toString();
  if (currentDay < 10) { currentDay = '0' + currentDay; }
  var str_date = date.getFullYear().toString()+
                 currentMonth.toString()+
                 currentDay.toString();
  template.name = chapter + "_" + str_date  + "_";
  var htmlOutput = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Submit Item')
      .setWidth(500);
//  Logger.log("(" + arguments.callee.name + ") " +htmlOutput.getContent());
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(htmlOutput);
//      .showModalDialog(template, "SUBMIT");
}

function side_event() {
  Logger.log("(" + arguments.callee.name + ") " +'Called addEvent');
  var html = HtmlService.createTemplateFromFile('side_event');
  html.events = get_type_list("Events");
  var htmlOutput  = html.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Event Functions')
    .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(htmlOutput);
}