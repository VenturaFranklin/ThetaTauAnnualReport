function pledge_update(form) {
  Logger.log("(" + arguments.callee.name + ") ");
  try{
    Logger.log(form);
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
    } else if (status == "Depledged") {
      DEPL.push(name);
    }
  }
  html.init = INIT
  html.depl = DEPL
  var jewelry = get_guard_badge();
  html.badges = jewelry.badges;
  html.guards = jewelry.guards;
  var htmlOutput = html.evaluate()
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setWidth(700)
    .setHeight(400);
  Logger.log("(" + arguments.callee.name + ") " +htmlOutput.getContent());
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
  .showModalDialog(htmlOutput, 'PLEDGE FORM');
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
  }
}

function member_update(form) {
//  var form = {"update_type": "Transfer", "memberlist": "Adam Schilpero...",
//              "Degree": ["Adam Schilpero...", "Austin Mutschl...", "Cole Mobberley"],//};
//              "Abroad": "Adam Schilpero...", "Transfer": "Adam Schilpero...",
//              "PreAlumn": ["Derek Hogue", "Esgar Moreno"], "Military": "Adam Schilpero...",
//              "CoOp": ["Adam Schilpero...", "Austin Mutschl...", "Cole Mobberley"]};
  Logger.log("(" + arguments.callee.name + ") ");
  try{
    Logger.log(form);
  var MemberObject = main_range_object("Membership");
  var html = HtmlService.createTemplateFromFile('form_status');
  Logger.log(html);
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
  }
}

function save_form(csvFile, form_type){
  try {
    var folder_id = get_form_id();
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
    return e.toString();
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
  Logger.log("(" + arguments.callee.name + ") ");
  try{
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
    var row = ["", formatted, chapterName, key, start, end, member_object["Badge Number"][0],
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

function process_init(form) {
//  var form = {badge:["Crown pearls Gold-Filled/Layered", "Basic"],
//              reason:"Too much time required",
//              name_init:["Austin Mutschl...", "Avery Davidson"],
//              testB:["1", "10"],
//              date_init:"2017-01-01",
//              name_depl:"Benjamin Ambri",
//              guard:["Engraved 10kt gold", "None"],
//              roll:["1", "10"], "GPA":["1", "10"],
//              date_grad:["2017-05-01", "2018-05-10"],
//              testA:["1", "10"],
//              date_depl:"2017-01-20"};
  Logger.log("(" + arguments.callee.name + ") ");
  try{
  Logger.log(form);
//  return;
  var MemberObject = main_range_object("Membership");
  var sheet = MemberObject["sheet"];
  var INIT = [header_INIT()];
  var DEPL = [header_DEPL()];
  var this_date = new Date();
  var date_init = form["date_init"];
  var ss = get_active_spreadsheet();
  if (date_init == ""){
      ss.toast('You must set the initiation date!', 'ERROR', 5);
      return [false, "date_init"];
    }
  date_init = format_date(date_init);
  var chapterName = get_chapter_name();
  var formatted = (this_date.getMonth() + 1) + '-' + this_date.getDate() + '-' +
                  this_date.getFullYear() + ' ' + this_date.getHours() + ':' +
                  this_date.getMinutes() + ':' + this_date.getSeconds();
  var init_count = 0;
  var depl_count = 0;
  var depl_objs = ["reason", "date_depl", "name_depl"];
  if (typeof form["name_depl"] === 'string'){
    for (var obj in form){
      if (depl_objs.indexOf(obj) > -1){
        form[obj] = [form[obj]];
      }
    }
  }
  if (typeof form["name_init"] === 'string'){
    for (var obj in form){
      if (depl_objs.indexOf(obj) < 0){
        form[obj] = [form[obj]];
      }
    }
  }
  Logger.log(form);
  if (form["name_init"] !== undefined){
    for (var i = 0; i < form["name_init"].length; i++){
      var name = form["name_init"][i];
      var member_object = find_member_shortname(MemberObject, name);
      if (typeof member_object == 'undefined') {
        // Something bad happened
        ss.toast('An Error Occured processing: '
                 +name, 'ERROR', 5);
        return [false, name];
      }
      var status_range = sheet.getRange(member_object["object_row"],
                                        member_object["Chapter Status"][1]);
      var status_start_range = sheet.getRange(member_object["object_row"],
                                              member_object["Status Start"][1]);
      status_range.setValue("Shiny");
      status_start_range.setValue(date_init);
      var first = member_object["First Name"][0];
      var last = member_object["Last Name"][0];
      var date_grad = form["date_grad"][i];
      var roll = form["roll"][i];
      var GPA = form["GPA"][i];
      var testA = form["testA"][i];
      var testB = form["testB"][i];
      var badge = form["badge"][i];
      var guard = form["guard"][i];
      var init_date = new Date(date_init);
      var timeDiff = Math.abs(this_date.getTime() - init_date.getTime());
      var diffDays = Math.ceil(timeDiff / (1000 * 3600 * 24)); 
      var late = (diffDays > 14) ? 25:0;
      var init_fee = (chapterName.indexOf("Colony") > -1) ? 30:75;
      var arr = [date_grad, roll, GPA, testA, testB];
      if (arr.indexOf("") > -1){
        ss.toast('You must set all of the fields!\nMissing information for:\n'
                 +name, 'ERROR', 5);
        return [false, name];
      }
      var jewelry = get_guard_badge();
      var badge_cost = jewelry.badges[badge];
      var guard_cost = jewelry.guards[guard];
      var sum = +init_fee + late + badge_cost + guard_cost;
      date_grad = format_date(date_grad);
      INIT.push(["", formatted, date_init, chapterName,
                 date_grad, roll, first, "",
                 last, GPA, testA,
                 testB, init_fee, late,
                 badge, guard, badge_cost, guard_cost, sum]);
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
      DEPL.push(["", formatted, chapterName, first, last, reason, date_depl]);
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
    return ["ERROR", null];
  }
}

function process_grad(form) {
//  var form = {"date_start":["2017-01-01", "2017-01-02", "2017-01-03", "2017-01-04", "2017-01-05", "2017-01-06"],
//              "new_location":["Test", "Test1", "Test2", "TEst3", "Arizona State University"],
//              "phone":"(714) 656-5839",
//              "name":["Aimee Largier", "Albert Hu", "Alec Sonderman", "AlexanderNEW Gerwe", "Amelia Sylvester", "AmmarNEW Mustafa"],
//              "degree":"Mechanical Engineeering",
//              "dist":["100", "1000", "1000"], "date_end":["2017-02-02", "2017-02-03", "2017-02-04"],
//              "type":["Degree received", "CoOp", "Military", "Abroad", "Transfer", "Withdrawn"], "email":"allisonbeth@cox.net"}
//  form = {"date_start": "2017-01-01", "new_location": "Test", "name": "Adam Schilperoort",
//          "dist": "100", "date_end": "2017-01-30", "type": "Abroad"};
//  var form = {date_start:["2017-01-01", "2017-01-01"], new_location:["Test", "Test"],
//              phone:"(707) 779-9411", name:["Aimee Largier", "Albert Hu"],
//              degree:"Industrial Engineering", dist:"100", date_end:"2017-01-30",
//              type:["Degree received", "CoOp"], email:"miminoxolo@comcast.net"};
//  var form = {"date_start": ["2017-01-01", "2017-01-01", "2017-01-01"],
//              "name": ["Adam Schilperoort", "Austin Conry", "Austin Conry"],
//              "type": ["Withdrawn", "Withdrawn", "Withdrawn"]}
  Logger.log("(" + arguments.callee.name + ") ");
  try{
  Logger.log(form);
//  return;
  var MemberObject = main_range_object("Membership");
  var sheet = MemberObject["sheet"];
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
  for (var obj in form){
    if (typeof form[obj] === 'string'){
      form[obj] = [form[obj]];
    }
  }
  Logger.log(form);
  for (var i = 0; i < form["type"].length; i++){
    var type = form["type"][i];
    Logger.log("(" + arguments.callee.name + ") " +type);
    var name = form["name"][i];
    var member_object = find_member_shortname(MemberObject, name);
    var badge = member_object["Badge Number"][0];
    var first = member_object["First Name"][0];
    var last = member_object["Last Name"][0];
    if (type != "Returning" && type != "Withdrawn"){
      var loc = form["new_location"][i];
    } else {
      var loc = "None";
    }
    var date_start = form["date_start"][i];
    date_start = format_date(date_start);
    var status_range = sheet.getRange(member_object["object_row"],
                                      member_object["Chapter Status"][1]);
    var status_start_range = sheet.getRange(member_object["object_row"],
                                            member_object["Status Start"][1]);
    var status_end_range = sheet.getRange(member_object["object_row"],
                                          member_object["Status End"][1]);
    status_start_range.setValue(date_start);
    if (type == "Degree received"){
      var email = form["email"][degree_count];
      var phone = form["phone"][degree_count];
      var degree = form["degree"][degree_count];
      var arr = [loc, date_start, email, phone, degree];
      status_range.setValue("Alumn");
      degree_count++
    } else if (type != "PreAlumn"){
      if (type != "Withdrawn" && type != "Transfer"){
        var date_end = form["date_end"][nonalum_count];
        var dist = form["dist"][nonalum_count];
        var arr = [loc, date_start, date_end, dist];
        date_end = format_date(date_end);
        status_range.setValue("Away");
        status_end_range.setValue(date_end);
        nonalum_count++
      } else {
        status_range.setValue("Alumn");
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
      status_range.setValue("Alumn");
      alum_count++
    }
    status_start_range.setValue(date_start);
    if (arr.indexOf("") > -1){
      var ss = get_active_spreadsheet();
      ss.toast('You must set all of the fields!\nMissing information for:\n'
             +name, 'ERROR', 5);
      return [false, name];
    }
    switch (type) {
      case "Degree received":
        MSCR.push(["", formatted, badge, first, last, phone, email,
                   "Graduated from school", degree, date_start, loc, "",
                   "", "", "", "", "", ""]);
        break;
      case "Transfer":
        MSCR.push(["", formatted, badge, first, last, "", "",
                   "Transferring to another school", "",
                   "", "", "", "", "", "",
                   loc, date_start, ""]);
        break;
      case "Withdrawn":
        MSCR.push(["", formatted, badge, first, last, "", "",
                   "Withdrawing from school", "", "", "", "", "",
                   "Yes", date_start, "", "", ""]);
        break;
      case "PreAlumn":
        MSCR.push(["", formatted, badge, first, last, "", "",
                   "Wishes to REQUEST Premature Alum Status", "",
                   "", "", "", "", "", "", "", "", prealumn]);
        break;
      case "Abroad":
        COOP.push(["", formatted, badge, first, last,
                   "Study Abroad", date_start,
                   date_end, dist]);
        break;
      case "Military":
        COOP.push(["", formatted, badge, first, last,
                   "Called to Active/Reserve Military Duty",
                   date_start, date_end, dist]);
        break;
      case "CoOp":
        COOP.push(["", formatted, badge, first, last,
                   "Co-Op/Internship",
                   date_start, date_end, dist]);
        break;
      default:
        return ["Student Resuming Student Member Status", null];
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
    return ["ERROR", null];
  }
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

function post_submit(file_object, submission_type, program) {
  try{
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
  if (submission_type == "Pledge Program"){
    submission_type = submission_type + " - " + program;
  }
  submit_range.setValues([[submission_date, file_name, submission_type, 0, file_url]])
  update_scores_submit(max_row + 1);
  var output = template.evaluate().getContent();
  Logger.log("(" + arguments.callee.name + ") " +output);
  sendemail_submission(submission_type, file_object)
  return output;
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

function sendemail_submission(submission_type, submission) {
  try{
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
  if (!WORKING){
    MailApp.sendEmail(email_director, subject, emailBody, optAdvancedArgs);
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