function pledge_update(form) {
  Logger.log("(" + arguments.callee.name + ") " +form);
  var html = HtmlService.createTemplateFromFile('form_init');
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
  Logger.log("(" + arguments.callee.name + ") " +htmlOutput.getContent());
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showModalDialog(htmlOutput, 'PLEDGE FORM');
}

function member_update(form) {
//  var form = {"update_type": "Transfer", "memberlist": "Adam Schilpero...",
//              "Degree": ["Adam Schilpero...", "Austin Mutschl...", "Cole Mobberley"],//};
//              "Abroad": "Adam Schilpero...", "Transfer": "Adam Schilpero...",
//              "PreAlumn": ["Derek Hogue", "Esgar Moreno"], "Military": "Adam Schilpero...",
//              "CoOp": ["Adam Schilpero...", "Austin Mutschl...", "Cole Mobberley"]};
  Logger.log("(" + arguments.callee.name + ") " +form);
  var MemberObject = main_range_object("Membership");
  var html = HtmlService.createTemplateFromFile('form_status');
  var CSMTA = []
  for (var k in form){
    var type = k;
    if (type == "update_type" || type == "memberlist"){
      continue;
    }
    Logger.log("(" + arguments.callee.name + ") " +k);
    var select_members = form[k];
    if (typeof select_members === 'string'){
      select_members = [select_members];
    }
    Logger.log("(" + arguments.callee.name + ") " +select_members);
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
    Logger.log("(" + arguments.callee.name + ") " +htmlOutput.getContent());
    SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(htmlOutput, 'STATUS FORM');
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
    Logger.log("(" + arguments.callee.name + ") " +"fileBlob Name: " + file.getName())
    Logger.log("(" + arguments.callee.name + ") " +'fileBlob: ' + file);
    
    var template = HtmlService.createTemplateFromFile('response');
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
    Logger.log("(" + arguments.callee.name + ") " +error);
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
  Logger.log("(" + arguments.callee.name + ") " +form);
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
  var dash = {};
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
    Logger.log("(" + arguments.callee.name + ") " +row);
    data.push(row);
    dash[key] = [chapterName, key, member_object["Member Name"][0],
                 member_object["Phone Number"][0], member_object["Email Address"][0],
                 start, end, date];
  }
  Logger.log("(" + arguments.callee.name + ") " +data);
  var csvFile = create_csv(data);
  Logger.log("(" + arguments.callee.name + ") " +csvFile);
  sync_officers(dash);
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
  Logger.log("(" + arguments.callee.name + ") " +form);
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
  Logger.log("(" + arguments.callee.name + ") " +form);
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
  Logger.log("(" + arguments.callee.name + ") " +"INIT");
  Logger.log("(" + arguments.callee.name + ") " +INIT);
  var csvFile = create_csv(INIT);
  Logger.log("(" + arguments.callee.name + ") " +csvFile);
  var init_out = "";
  if (INIT.length > 1){
    init_out = save_form(csvFile, "INIT");
  }
  Logger.log("(" + arguments.callee.name + ") " +"DEPL");
  Logger.log("(" + arguments.callee.name + ") " +DEPL);
  var csvFile = create_csv(DEPL);
  Logger.log("(" + arguments.callee.name + ") " +csvFile);
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
  Logger.log("(" + arguments.callee.name + ") " +form);
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
    Logger.log("(" + arguments.callee.name + ") " +type);
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
  Logger.log("(" + arguments.callee.name + ") " +"COOP");
  Logger.log("(" + arguments.callee.name + ") " +COOP);
  var csvFile = create_csv(COOP);
  Logger.log("(" + arguments.callee.name + ") " +csvFile);
  var coop_out = "";
  if (COOP.length > 1){
    coop_out = save_form(csvFile, "COOP");
  }
  Logger.log("(" + arguments.callee.name + ") " +"MSCR");
  Logger.log("(" + arguments.callee.name + ") " +MSCR);
  var csvFile = create_csv(MSCR);
  Logger.log("(" + arguments.callee.name + ") " +csvFile);
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
    Logger.log("(" + arguments.callee.name + ") " +err);
    Browser.msgBox(err);
  }
}


function uploadFiles(form) {
  Logger.log("(" + arguments.callee.name + ") " +form);
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
    Logger.log("(" + arguments.callee.name + ") " +"fileBlob Name: " + blob.getName())
    Logger.log("(" + arguments.callee.name + ") " +"fileBlob type: " + blob.getContentType())
    Logger.log("(" + arguments.callee.name + ") " +'fileBlob: ' + blob);
    var file = folder.createFile(blob);    
//    file.setDescription("Uploaded by " + form.myName);
    var template = HtmlService.createTemplateFromFile('response');
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
    Logger.log("(" + arguments.callee.name + ") " +this_error);
    return this_error;
  }
}

function post_submit(file_object, submission_type) {
  var template = HtmlService.createTemplateFromFile('response');
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
  Logger.log("(" + arguments.callee.name + ") " +output);
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
  Logger.log("(" + arguments.callee.name + ") " +[email_director, email_chapter, chapter, subject, file_id,
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