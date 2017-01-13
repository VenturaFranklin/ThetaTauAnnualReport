//https://developers.google.com/apps-script/quickstart/forms
//https://developers.google.com/apps-script/quickstart/forms-add-on
function survey(){
  var ss = get_active_spreadsheet();
  var form = survey_create(member_info);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  try {
    ScriptApp.newTrigger('submit_survey')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  } catch (error) {
    Logger.log("(" + arguments.callee.name + ") " +error);
  }
}

function survey_create(member_info){
  var form = FormApp.create(semester + ' AR Survey');
  var item = form.addTextItem();
  form.addTextItem().setTitle('How many community service hours have you spent outside of Theta Tau events in the fall?');
  item.setTitle('How many community service hours have you spent outside of Theta Tau events in the spring?');
  item.setTitle('What is your fall semester GPA?').setRequired(true);
  item.setTitle('What is your spring semester GPA?');
//  Professional/ Technical Orgs
//  Officer (Pro/Tech)
//  Honor Orgs
//  Officer (Honor)
//  Other Orgs
//  Officer (Other)
  
  
//  First Name
//  Last Name
//  Badge Number
//  Chapter Status
//  Chapter Role
//  Current Major
//  School Status
//  Phone Number
//  Email Address
  var item = form.addCheckboxItem();
  item.setTitle('What condiments would you like on your hot dog?');
  item.setChoices([
    item.createChoice('Ketchup'),
    item.createChoice('Mustard'),
    item.createChoice('Relish')
  ]);
  form.addMultipleChoiceItem()
  .setTitle('Do you prefer cats or dogs?')
  .setChoiceValues(['Cats','Dogs'])
  .showOtherOption(true);
  form.addPageBreakItem()
  .setTitle('Getting to know you');
  form.addDateItem()
  .setTitle('When were you born?');
  form.addGridItem()
  .setTitle('Rate your interests')
  .setRows(['Cars', 'Computers', 'Celebrities'])
  .setColumns(['Boring', 'So-so', 'Interesting']);
  Logger.log("(" + arguments.callee.name + ") " +'Published URL: ' + form.getPublishedUrl());
  Logger.log("(" + arguments.callee.name + ") " +'Editor URL: ' + form.getEditUrl());
  return form
}

function submit_survey(e) {
  var user = {name: e.namedValues['Name'][0], email: e.namedValues['Email'][0]};
}

function survey_email(form, email) {
  var url = form.getPublishedUrl();
  form.setRequireLogin(false);
  // Fetch form's HTML
  var response = UrlFetchApp.fetch(url);
  var htmlBody = HtmlService.createHtmlOutput(response).getContent();
  var subject = form.getTitle();
  GmailApp.sendEmail(email,
                    subject,
                    'Please complete this form in order to help the chapter fill out our annual report.<br>',
                    {
                      name: chapter_name + ' Chapter Scribe',
                      htmlBody: htmlBody,
                      noReply: true
                    });
}