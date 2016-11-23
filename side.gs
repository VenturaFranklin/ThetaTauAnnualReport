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

function side_submit() {
   var template = HtmlService
   .createTemplateFromFile('side_submit')
//  .createHtmlOutputFromFile('SubmitForm');
   var list_info = get_type_list('Submit', true);
  template.submissions = list_info.type_list;
  template.descriptions = list_info.type_desc;
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

function side_event() {
  Logger.log('Called addEvent');
  var html = HtmlService.createTemplateFromFile('side_event');
  html.events = get_type_list("Events");
  var htmlOutput  = html.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setTitle('Event Functions')
    .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(htmlOutput);
}