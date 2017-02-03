function create_survey(){
  progress_update("Started Survey Creation");
  Logger.log("(" + arguments.callee.name + ") ");
  var ss = get_active_spreadsheet();
  var member_list = get_member_list("Student");
  var form = survey_create(member_list);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId())
  .setAllowResponseEdits(false)
  .setAcceptingResponses(true)
  .setRequireLogin(false);
  var survey_id = form.getId();
  SCRIPT_PROP.setProperty("survey", survey_id);
  var file = DriveApp.getFileById(survey_id);
  var folder = DriveApp.getFolderById(get_folder_id());
  folder.addFile(file);
  var root_folder = DriveApp.getRootFolder();
  root_folder.removeFile(file);
  try {
    ScriptApp.newTrigger('submit_survey')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  } catch (error) {
    Logger.log("(" + arguments.callee.name + ") ");
    Logger.log(error);
  }
//  update_survey_sheet();
}

//function update_survey_sheet(){
//  Does not work, list of sheets does not contain new Form sheet
//  var ss = get_active_spreadsheet();
//  var sheets = ss.getSheets();
//  for (var i in sheets){
//    var sheet = sheets[i];
//    var name = sheet.getName();
//    if (name.indexOf("Form") > -1){
//        sheet.setName("FORM");
//    }
//  }
//}

function delete_survey() {
  SCRIPT_PROP.deleteProperty("survey");
}

function get_survey(){
  return FormApp.openById(SCRIPT_PROP.getProperty("survey"));
}

function survey_create(member_list){
  Logger.log("(" + arguments.callee.name + ") ");
  Logger.log("(" + arguments.callee.name + ") ");
  var chapter_name = get_chapter_name();
  var form = FormApp.create(chapter_name + ' Survey');
  var gpaValidation = FormApp.createTextValidation()
   .setHelpText("GPA is on a 4.0 scale; should be number between 0.0 and 5.0")
   .requireNumberBetween(0, 5)
   .build();
  var servValidation = FormApp.createTextValidation()
   .setHelpText("Service should be a number greater than 0.")
   .requireNumberGreaterThan(0)
   .build();
  form.addListItem()
  .setTitle("Name")
  .setChoiceValues(member_list)
  .setRequired(true);
  form.addTextItem()
  .setTitle('Fall Service')
  .setHelpText('How many community service hours have you spent outside of Theta Tau events in the fall? (July-December)')
  .setValidation(servValidation)
  .setRequired(true);
  form.addTextItem()
  .setTitle('Spring Service')
  .setHelpText('How many community service hours have you spent outside of Theta Tau events in the spring? (January-June)')
  .setValidation(servValidation)
  .setRequired(true);
  form.addTextItem()
  .setTitle('Fall GPA')
  .setHelpText('What is your fall semester GPA? Set as 0 if not yet calculated')
  .setValidation(gpaValidation)
  .setRequired(true);
  form.addTextItem()
  .setTitle('Spring GPA')
  .setHelpText('What is your spring semester GPA? Set as 0 if not yet calculated')
  .setValidation(gpaValidation)
  .setRequired(true);
  form.addTextItem()
  .setTitle('Professional Orgs')
  .setHelpText('What Professional/ Technical Organizations are you a member? None, if none.')
  .setRequired(true);
  form.addMultipleChoiceItem()
  .setTitle('Professional Officer')
  .setHelpText('Are you an officer in a Professional/ Technical Organization?')
  .setChoiceValues(['Yes','No'])
  .setRequired(true);
  form.addTextItem()
  .setTitle('Honor Orgs')
  .setHelpText('What Honor Organizations are you a member? None, if none.')
  .setRequired(true);
  form.addMultipleChoiceItem()
  .setTitle('Honor Officer')
  .setHelpText('Are you an officer in an Honor Organization?')
  .setChoiceValues(['Yes','No'])
  .setRequired(true);
  form.addTextItem()
  .setTitle('Other Orgs')
  .setHelpText('What Other Organizations are you a member? None, if none.')
  .setRequired(true);
  form.addMultipleChoiceItem()
  .setTitle('Other Officer')
  .setHelpText('Are you an officer in any Other Organization?')
  .setChoiceValues(['Yes','No'])
  .setRequired(true);
  return form
}

function submit_survey(e) {
//  var e = {authMode:"FULL",
//           values:["1/21/2017 22:37:23", "AlexanderNEW G...",
//                   5, 1, 3, 3.5, "TEST", "Yes", "None", "No", "None", "No"],
//           namedValues:{"Spring GPA":[3.5], "Other Orgs":["None"], "Honor Officer":["No"],
//                        "Professional Officer":["Yes"], "Fall Service":[5], "Fall GPA":[3],
//                        "Professional Orgs":["TEST"], "Honor Orgs":["None"],
//                        "Timestamp":["1/21/2017 22:37:23"], "Other Officer":["No"],
//                        "Name":["AlexanderNEW G..."], "Spring Service":[1]},
//           triggerUid:"8395862820412211388"}
  try{
  Logger.log("(" + arguments.callee.name + ") ");
  Logger.log(e);
  var fields = {"Fall Service": "Self Service Hrs FA",
                "Spring Service": "Self Service Hrs SP", 
                "Fall GPA": "Fall GPA", 
                "Spring GPA": "Spring GPA",
                "Professional Orgs": "Professional/ Technical Orgs", 
                "Professional Officer": "Officer (Pro/Tech)", 
                "Honor Orgs": "Honor Orgs", 
                "Honor Officer": "Officer (Honor)",
                "Other Orgs": "Other Orgs", 
                "Other Officer": "Officer (Other)"};
  var MemberObject = main_range_object("Membership");
  var sheet = MemberObject["sheet"];
  var survey_name = e.namedValues["Name"][0];
  var member_object = find_member_shortname(MemberObject, survey_name);
  var member_row = member_object["object_row"];
  var response = {};
  for (var field in fields) {
    var survey_val = e.namedValues[field][0];
    var field_col = member_object[fields[field]][1];
    var field_range = sheet.getRange(member_row,
                                     field_col);
    field_range.setValue(survey_val);
  }
    } catch (e) {
    var message = Utilities.formatString('This error has automatically been sent to the developers. %s: %s (line %s, file "%s"). Stack: "%s" . While processing %s.',
                                         e.name||'', e.message||'', e.lineNumber||'', e.fileName||'',
                                         e.stack||'', arguments.callee.name||'');
    Logger = startBetterLog();
    Logger.severe(message);
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
     'ERROR',
      message,
      ui.ButtonSet.OK);
    return "";
  }
}

function send_survey() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Do you want to send a survey to all members?',
                           ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (response == ui.Button.NO) {
    return;
  }
  try{
  Logger.log("(" + arguments.callee.name + ") ");
  var form = get_survey();
  var MemberObject = main_range_object("Membership");
  for (var i in MemberObject["object_header"]){
    var member = MemberObject["object_header"][i];
    var email = MemberObject[member]["Email Address"][0];
    survey_email(form, email);
  }
    } catch (e) {
    var message = Utilities.formatString('This error has automatically been sent to the developers. %s: %s (line %s, file "%s"). Stack: "%s" . While processing %s.',
                                         e.name||'', e.message||'', e.lineNumber||'', e.fileName||'',
                                         e.stack||'', arguments.callee.name||'');
    Logger = startBetterLog();
    Logger.severe(message);
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
     'ERROR',
      message,
      ui.ButtonSet.OK);
    return "";
  }
}

function survey_email(form, email) {
//  var form = get_survey();
//  var email = "venturafranklin@gmail.com";
  Logger.log("(" + arguments.callee.name + ") ");
  Logger.log("(" + arguments.callee.name + ") " + email);
  var url = form.getPublishedUrl();
  Logger.log("(" + arguments.callee.name + ") " + url);
  // Fetch form's HTML
//  var response = UrlFetchApp.fetch(url);
//  var htmlBody = HtmlService.createHtmlOutput(response).getContent();
  var chapter_name = get_chapter_name();
  var subject = chapter_name + " Chapter Survey";
  var email_chapter = SCRIPT_PROP.getProperty("email");
  var emailBody = "Please fill out the following survey for the chapter's annual report:"+
    "\nSurvey("+url+")";

  var htmlBody = "Please fill out the following survey for the chapter's annual report:"+
    '<br/><a href="'+url+'"> Survey</a> ('+url+')';
  var optAdvancedArgs = {name: chapter_name +" Chapter Scribe", htmlBody: htmlBody,
                         replyTo: email_chapter};
  if (!WORKING){
    MailApp.sendEmail(email, subject,
                      emailBody,
                      optAdvancedArgs);
  }
}