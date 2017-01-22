function create_survey(){
  var ss = get_active_spreadsheet();
  var member_list = get_member_list("Student");
  var form = survey_create(member_list);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId())
  .setAllowResponseEdits(false)
  .setAcceptingResponses(true)
  .setRequireLogin(false);
  SCRIPT_PROP.setProperty("survey", form.getId());
  try {
    ScriptApp.newTrigger('submit_survey')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  } catch (error) {
    Logger.log("(" + arguments.callee.name + ") ");
    Logger.log(error);
  }
}

function delete_survey() {
  SCRIPT_PROP.deleteProperty("survey");
}

function get_survey(){
  return FormApp.openById(SCRIPT_PROP.getProperty("survey"));
}

function survey_create(member_list){
  var chapter_name = get_chapter_name();
  var form = FormApp.create(chapter_name + ' Survey');
  form.addListItem()
  .setTitle("Name")
  .setChoiceValues(member_list)
  .setRequired(true);
  form.addTextItem()
  .setTitle('Fall Service')
  .setHelpText('How many community service hours have you spent outside of Theta Tau events in the fall? (July-December)')
  .setRequired(true);
  form.addTextItem()
  .setTitle('Spring Service')
  .setHelpText('How many community service hours have you spent outside of Theta Tau events in the spring? (January-June)')
  .setRequired(true);
  form.addTextItem()
  .setTitle('Fall GPA')
  .setHelpText('What is your fall semester GPA?')
  .setRequired(true);
  form.addTextItem()
  .setTitle('Spring GPA')
  .setHelpText('What is your spring semester GPA?')
  .setRequired(true);
  form.addTextItem()
  .setTitle('Professional Orgs')
  .setHelpText('What Professional/ Technical Organizations are you a member?')
  .setRequired(true);
  form.addMultipleChoiceItem()
  .setTitle('Professional Officer')
  .setHelpText('Are you an officer in a Professional/ Technical Organization?')
  .setChoiceValues(['Yes','No'])
  .setRequired(true);
  form.addTextItem()
  .setTitle('Honor Orgs')
  .setHelpText('What Honor Organizations are you a member?')
  .setRequired(true);
  form.addMultipleChoiceItem()
  .setTitle('Honor Officer')
  .setHelpText('Are you an officer in an Honor Organization?')
  .setChoiceValues(['Yes','No'])
  .setRequired(true);
  form.addTextItem()
  .setTitle('Other Orgs')
  .setHelpText('What Other Organizations are you a member?')
  .setRequired(true);
  form.addMultipleChoiceItem()
  .setTitle('Other Officer')
  .setHelpText('Are you an officer in any Other Organization?')
  .setChoiceValues(['Yes','No'])
  .setRequired(true);
  return form
}

function submit_survey(e) {
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
  var member_object = MemberObject[survey_name];
  var member_row = member_object["object_row"];
  var response = {};
  for (var field in fields) {
    var survey_val = e.namedValues[field][0];
    var field_col = member_object[fields[field]][1];
    var field_range = sheet.getRange(member_row,
                                     field_col);
    field_range.setValue(survey_val);
  }
}

function send_survey() {
  var form = get_survey();
  var MemberObject = main_range_object("Membership");
  for (var i in MemberObject["object_header"]){
    var member = MemberObject["object_header"][i];
    var email = MemberObject[member]["Email Address"][0];
    survey_email(form, email);
  }
}

function survey_email(form, email) {
//  var form = get_survey();
//  var email = "venturafranklin@gmail.com";
  var url = form.getPublishedUrl();
  // Fetch form's HTML
  var response = UrlFetchApp.fetch(url);
  var htmlBody = HtmlService.createHtmlOutput(response).getContent();
  var chapter_name = get_chapter_name();
  var subject = chapter_name + " Chapter Survey";
  var email_chapter = SCRIPT_PROP.getProperty("email");
  var optAdvancedArgs = {name: chapter_name +" Chapter Scribe", htmlBody: htmlBody,
                         replyTo: email_chapter};
  if (!WORKING){
    MailApp.sendEmail(email, subject,
                      htmlBody,
                      optAdvancedArgs);
  }
}