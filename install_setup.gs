function run_install(e){
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert('Proceed with installing?',
      ui.ButtonSet.YES_NO);
//  var button = result.getSelectedButton();
  if (result == ui.Button.YES) {
    onInstall(e);
  }
}

function onInstall(e) {
  onOpen(e);
  try {
    setup();
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

message = ""

function progress_update(this_message){
  try {
    if (SILENT){
      return;}
    Logger.log(this_message);
    message += "<br>" + this_message;
    var htmlOutput = HtmlService
    .createHtmlOutput(message)
    .setWidth(400)
    .setHeight(300);
    SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, 'Progress');
  } catch (e) {
    var error_message = Utilities.formatString('This error has automatically been sent to the developers. %s: %s (line %s, file "%s"). Stack: "%s" . While processing %s.',
                                               e.name||'', e.message||'', e.lineNumber||'', e.fileName||'',
                                               e.stack||'', arguments.callee.name||'');
//    Logger = startBetterLog();
    Logger.severe(error_message);
  }
}

function run_createTriggers() {
  unlock();
  var this_password = SCRIPT_PROP.getProperty("password");
  if (this_password != password){
    var ui = SpreadsheetApp.getUi();
    ui.alert('Incorrect Password!');
    return;
  }
  createTriggers();
}

function createTriggers() {
  try {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
    progress_update("Creating Edit Trigger");
    var ss = get_active_spreadsheet();
    ScriptApp.newTrigger('_onEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  } catch (error) {
    Logger.log("(" + arguments.callee.name + ") ");
    Logger.log(error);
  } 
  try {
    progress_update("Creating Sync Trigger");
    ScriptApp.newTrigger("sync")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .create();
  } catch (error) {
    Logger.log("(" + arguments.callee.name + ") ");
    Logger.log(error);
  }
}

function get_chapter_name(){
  return SCRIPT_PROP.getProperty("chapter")
}

function get_chapter_fee(){
  return SCRIPT_PROP.getProperty("chapter_fee")
}

function chapter_name_process(form) {
  try{
  var cur_date = new Date();
  var start_year = get_start_year();
  var stop_year = get_stop_year();
  var cur_year = start_year + "-" + stop_year;
  Logger.log("(" + arguments.callee.name + ") ");
  Logger.log(form);
//  var form = {'chapterslist': 'Chi Gamma'}
  var chapter_name = form.chapterslist;
  SCRIPT_PROP.setProperty("chapter", chapter_name);
  start_logging();
  progress_update("Chapter Name set: " + chapter_name);
  var properties_id = "1vCVKh8MExPxg8eHTEGYx7k-KTu9QUypGwbtfliLm58A";
  var ss_prop = SpreadsheetApp.openById(properties_id);
  var ss = get_active_spreadsheet();
  var default_id = "19aWLtjJJ-Uh6XOqOuseLpQcNJYslQHe9Y9Gaj2vSjEw";
  var default_doc = SpreadsheetApp.openById(default_id);
  var doc_name = default_doc.getName();
  doc_name = doc_name.replace("DEFAULT ", "");
  doc_name = doc_name.replace("- Chapter", "- "+chapter_name + " " + cur_year);
  Logger.log("(" + arguments.callee.name + ") " +doc_name);
  ss.rename(doc_name);
  var sheet = ss.getSheetByName("Chapter");
  var range = sheet.getRange(2, 2);
  range.setValue(chapter_name);
  var test = range.getValue();
  var chapter_object = main_range_object("MAIN", "Organization Name", ss_prop);
  var chapter_info = chapter_object[chapter_name];
  progress_update("Chapter Information: " + chapter_info.toString());
  Logger.log("(" + arguments.callee.name + ") ");
  Logger.log(chapter_info);
  var region = chapter_info["Region Description"][0];
  range = sheet.getRange(2, 3);
  range.setValue(region);
  SCRIPT_PROP.setProperty("region", region);
  var director = chapter_info["Regional Director"][0];
  SCRIPT_PROP.setProperty("director", director);
  var balance = chapter_info["Balance Due"][0];
  var balance_date = chapter_info["Balance Updated"][0]
  var tax = chapter_info["Tax ID Number"][0];
  range = sheet.getRange(2, 5);
  range.setValue(tax);
  SCRIPT_PROP.setProperty("tax", tax);
  var email = chapter_info["Email"][0];
  range = sheet.getRange(2, 4);
  range.setValue(email);
  SCRIPT_PROP.setProperty("email", email);
  var sheet_dash = ss.getSheetByName("Dashboard");
  range = sheet_dash.getRange(2, 5);
  range.setValue(balance);
  range.setNote("Last Updated: "+balance_date);
  range = sheet_dash.getRange(1, 1);
  range.setValue(chapter_name + " CHAPTER ANNUAL REPORT " + cur_year);
  range.getValue();
  create_submit_folder(chapter_name, region);
  get_chapter_members();
  update();
  setup_dataval();
  createTriggers();
  create_survey();
  progress_update("Started Sync Main Info");
  sync_main();
  var ui = SpreadsheetApp.getUi();
  ui.alert('SETUP COMPLETE!\n'+
           'Next steps:\n'+
           '- Fill out Chapter Sheet\n'+
           '- Verify Membership\n'+
           '- Add Events\n\n'+
           'Do not edit gray or black cells\n'+
           'Submit forms in menu "Add-ons-->ThetaTauReports"');
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

function create_submit_folder(chapter_name, region) {
//  var chapter_name = "Epsilon Delta";
//  var region = "Western";
  progress_update("Started Submit Folder Creation");
  var folder_id = "0BwvK5gYQ6D4nTDRtY1prZG12UU0";
  var folder_submit = DriveApp.getFolderById(folder_id);
  var folder_region = folder_submit.getFoldersByName(region);
  if (folder_region.hasNext()) {
    folder_region = folder_region.next()
    progress_update("Found Region Folder: " + region);
  } else {
    folder_region = folder_submit.createFolder(region);
    progress_update("Created Region Folder: " + region);
  }
  var files = folder_region.getFiles();
  var file_dash = undefined;
  while (files.hasNext()) {
    var file = files.next();
    var file_name = file.getName();
    Logger.log("(" + arguments.callee.name + ") " +file_name);
    if (file_name.indexOf('Dashboard') > -1){
      file_dash = file;
      progress_update("Found Region Dashboard");
    }
  }
  if (!file_dash){
    var default_id = "1eqiek9iR1AtV7tw0WtrrTLX6BBZK7rOUAl2idD770hI";
//    var default_doc = SpreadsheetApp.openById(default_id);
    var default_doc = DriveApp.getFileById(default_id);
    var file_dash = default_doc.makeCopy(region + " Dashboard", folder_region);
    file_dash.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.EDIT);
//    var default_blob = default_doc.getBlob();
//    var file_dash = folder_region.createFile(default_blob);
    progress_update("Created Region Dashboard");
  }
  var dash_id = file_dash.getId();
  SCRIPT_PROP.setProperty("dash", dash_id);
  var dash = SpreadsheetApp.openById(dash_id);
  var main_sheet = dash.getSheetByName("MAIN");
  var main_values = main_sheet.getDataRange().getValues();
  var found_chapter = false;
  for (var i in main_values){
    var row = main_values[i];
    if (row.indexOf(chapter_name) > -1){
      found_chapter = true;
      break;
    }
  }
  if (!found_chapter){
    var values = [[chapter_name, SCRIPT_PROP.getProperty("email"),
                  SCRIPT_PROP.getProperty("tax"), SCRIPT_PROP.getProperty("balance")]];
    main_sheet.getRange(main_values.length+1, 1, 1, 4)
    .setValues(values);
  }
  var folder_chapter = folder_region.getFoldersByName(chapter_name);
  if (folder_chapter.hasNext()) {
    folder_chapter = folder_chapter.next();
    progress_update("Found Chapter Folder: " + chapter_name);
  } else {
    folder_chapter = folder_region.createFolder(chapter_name);
    var submit_folder = folder_chapter.createFolder("Submissions");
    var form_folder = folder_chapter.createFolder("Forms");
    progress_update("Created Chapter Folder: " + chapter_name);
  }
  var submit_folder = folder_chapter.getFoldersByName("Submissions");
  if (submit_folder.hasNext()) {
    submit_folder = submit_folder.next();
  } else {
    submit_folder = folder_chapter.createFolder("Submissions");
  }
  var form_folder = folder_chapter.getFoldersByName("Forms");
  if (form_folder.hasNext()) {
    form_folder = form_folder.next();
  } else {
    form_folder = folder_chapter.createFolder("Forms");
  }
  var folder_id = folder_chapter.getId();
  SCRIPT_PROP.setProperty("folder", folder_id);
  SCRIPT_PROP.setProperty("submit", submit_folder.getId());
  SCRIPT_PROP.setProperty("form", form_folder.getId());
  var file = DriveApp.getFileById(SCRIPT_PROP.getProperty("key"));
  DriveApp.addFolder(folder_chapter); // Adds the chapter folder to user drive
  folder_chapter.addFile(file); // Adds the ss to chapter folder
  DriveApp.removeFile(file); // Removes ss from user drive
}

function get_form_id() {
  return SCRIPT_PROP.getProperty("form");
}

function get_submit_id() {
  return SCRIPT_PROP.getProperty("submit");
}

function get_folder_id() {
  return SCRIPT_PROP.getProperty("folder");
}

function chapter_name() {
  var properties_id = "1vCVKh8MExPxg8eHTEGYx7k-KTu9QUypGwbtfliLm58A";
  var ss = SpreadsheetApp.openById(properties_id);
  var chapter_object = main_range_object("MAIN", "Organization Name", ss);
  var chapter_list = [];
  for(var i = 0; i< chapter_object.object_count; i++) {
    var chapter_name = chapter_object.object_header[i];
    chapter_list.push(chapter_name);
    }
  var template = HtmlService
      .createTemplateFromFile('chapter_name');
  template.chapters = chapter_list
  var htmlOutput = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(250)
      .setHeight(175);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(htmlOutput, "Chapter Name");
}
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function setup() {
  var template = HtmlService
      .createTemplateFromFile('ss_id');
  var htmlOutput = template.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(600)
      .setHeight(400);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(htmlOutput, "Spreadsheet ID?");
}

function sheet_id_process(form) {
  SCRIPT_PROP.setProperty("key", form.sheetid);
  progress_update("Spread Sheet ID set:" + form.sheetid);
  try {
  setup_sheets();
  chapter_name();
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

function setup_sheets() {
  var default_id = "19aWLtjJJ-Uh6XOqOuseLpQcNJYslQHe9Y9Gaj2vSjEw";
  var default_doc = SpreadsheetApp.openById(default_id);
  var target_doc = get_active_spreadsheet();
  var sheets = default_doc.getSheets();
  for (var i in sheets){
    var sheet = sheets[i];
    var sheet_name = sheet.getName();
    sheet.copyTo(target_doc).setName(sheet_name);
  }
  progress_update("Default Document Copied");
  var sheet = target_doc.getSheetByName("Sheet1");
  target_doc.deleteSheet(sheet);
  var named_ranges = default_doc.getNamedRanges();
  for (var j in named_ranges){
    var named_range = named_ranges[j];
    var name = named_range.getName();
    var range = named_range.getRange();
    var sheet = range.getSheet().getSheetName();
    var old_range = range.getA1Notation();
    Logger.log("(" + arguments.callee.name + ") " +old_range);
    var new_sheet = target_doc.getSheetByName(sheet);
    var new_range = new_sheet.getRange(old_range);
    Logger.log("(" + arguments.callee.name + ") " +name);
    target_doc.setNamedRange(name, new_range);
  }
  progress_update("Sheets and Ranges Setup");
}

function setup_dataval(){
  progress_update("Started Data Val Setup");
  var ss = get_active_spreadsheet();
  var event_sheets = find_all_event_sheets(ss);
  var EventObject = main_range_object("Events", undefined, ss);
  var yes_no_rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'])
    .setAllowInvalid(false).build();
  var events = get_type_list("Events");
  var type_rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(events)
    .setHelpText('Must be a valid event type.')
    .setAllowInvalid(false).build();
  var date_rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setHelpText('Enter a valid date MM/DD/YYYY')
    .setAllowInvalid(false).build();
  var type_col = EventObject.header_values.indexOf("Type") + 1;
  var member_col = EventObject.header_values.indexOf("# Members") + 1;
  var pledge_col = EventObject.header_values.indexOf("# Pledges") + 1;
  var date_col = EventObject.header_values.indexOf("Date") + 1;
  var stem_col = EventObject.header_values.indexOf("STEM?") + 1;
  var host_col = EventObject.header_values.indexOf("HOST") + 1;
  var hours_col = EventObject.header_values.indexOf("Event Hours") + 1;
  var name_col = EventObject.header_values.indexOf("Event Name") + 1;
  for (sheet_name in event_sheets){
    var event_sheet = event_sheets[sheet_name];
    var rows = event_sheet.getLastRow()-1;
    if (rows < 100){
      rows = 100;
    }
    var event_range = event_sheet.getRange(2, type_col, rows, 1);
    event_range.setDataValidation(type_rule);
    var date_range = event_sheet.getRange(2, date_col, rows, 1);
    date_range.setDataValidation(date_rule);
    var stem_range = event_sheet.getRange(2, stem_col, rows, 1);
    stem_range.setDataValidation(yes_no_rule);
    var host_range = event_sheet.getRange(2, host_col, rows, 1);
    host_range.setDataValidation(yes_no_rule);
    var hours_range = event_sheet.getRange(2, hours_col, 1, 1);
    hours_range.setNote("How long was the event? Number of hours?");
    var name_range = event_sheet.getRange(2, name_col, 1, 1);
    name_range.clearDataValidations();
    var member_range = event_sheet.getRange(2, member_col, rows, 1);
    member_range.setBackground("white");
    var pledge_range = event_sheet.getRange(2, pledge_col, rows, 1);
    pledge_range.setBackground("white");
  }
//  var range = ss.getRange("Attendance!1:149");
//  var rule = SpreadsheetApp.newDataValidation()
//    .requireValueInList(['P', 'E', 'U', 'p', 'e', 'u'], false)
//    .setHelpText('P-Present; E-Excused; U-Unexcused')
//    .setAllowInvalid(false).build();
//  range.setDataValidation(rule);
  var MemberObject = main_range_object("Membership", undefined, ss);
  var member_sheet = MemberObject.sheet;
  var name_col = MemberObject.header_values.indexOf("Member Name") + 1;
  var name_range = member_sheet.getRange(2, name_col, 1, 1);
  name_range.clearDataValidations();
  var start_year = get_start_year();
  var start_col_name = start_year + " FALL Service";
  var edit_col = MemberObject.header_values.indexOf(start_col_name) + 1;
  var max_row = member_sheet.getLastRow() - 1;
  var max_row = (max_row != 0) ? max_row:1;
  var max_column = member_sheet.getLastColumn();
  var edit_range = member_sheet.getRange(2, edit_col, max_row, max_column-edit_col+1);
  edit_range.setBackground("white");
  var ranges = ["Officer (Pro/Tech)", "Officer (Honor)", "Officer (Other)"];
  for (var i in ranges){
    var range_name = ranges[i];
    var col = MemberObject.header_values.indexOf(range_name) + 1;
    var range = member_sheet.getRange(2, col, max_row, 1);
    range.setDataValidation(yes_no_rule);
  }
  var rule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(0)
    .setHelpText('Set this value greater than 0.')
    .setAllowInvalid(false).build();
  var member_ranges = get_membership_ranges();
  for (var member_range_year in member_ranges){
    for (var member_range_type in member_ranges[member_range_year]){
      member_ranges[member_range_year][member_range_type].range.setDataValidation(rule);
    }
  }
  progress_update("Finished Data Val Setup");
}

function get_chapter_members(){
  try{
  progress_update("Started Get Chapter Members<br>"+
                  "This will take some time, please be patient...");
  var chapter_name = get_chapter_name();
  var folder = DriveApp.getFolderById('0BwvK5gYQ6D4nOXB2UHFUV0w5WnM');
  var files = folder.getFiles();
  var old_date = new Date(2000, 01, 01);
  var new_file = null;
  while (files.hasNext()) {
    var file = files.next();
    var file_name = file.getName();
    Logger.log("(" + arguments.callee.name + ") " +file_name);
    var date_str = file_name.split("_")[0];
    var year = date_str.substring(0, 4);
    var month = date_str.substring(4, 6);
    var day = parseInt(date_str.substring(6, 8), 10);
    if (isNaN(day)){
      continue;
    }
    month = parseInt(month, 10)-1;
    var date = new Date(year, month, day);
    if (date > old_date){
      old_date = date;
      new_file = file;
      var new_file_name = file.getName();
      var new_date = date;
    }
  }
  progress_update("Found Member list:" + new_file_name);
  var csvFile = new_file.getBlob().getDataAsString();
  var csvData = Utilities.parseCsv(csvFile, ",");
  var header = csvData[0];
  var chapter_index = header.indexOf("Constituent Specific Attributes Chapter Name Description");
  var CentralMemberObject = {};
  CentralMemberObject['badge_numbers'] = [];
  progress_update("Finding chapter members...");
  var indx = {
    "First Name": header.indexOf("First Name"),
    "Last Name": header.indexOf("Last Name"),
    "Badge Number": header.indexOf("Constituent ID"),
    "Chapter Status": header.indexOf("Constituency Code"),
    "Chapter Role": header.indexOf("Organization Relation Relationship"),
    "Current Major": header.indexOf("Constituent Specific Attributes y_Major Description"),
    "School Status": header.indexOf("Primary Education Class of"),
    "Phone Number": header.indexOf("Mobile Phone Number"),
    "Email Address": header.indexOf("Email Address Number")
  };
  var role_start_col = header.indexOf("Organization Relation From Date");
  var role_end_col = header.indexOf("Organization Relation To Date");
  var date_today = new Date();
  for (var j in csvData){
    var row = csvData[j];
    var chapter_row = row[chapter_index];
    if (chapter_row == chapter_name){
      var member_object={};
      var badge_number = row[indx["Badge Number"]];
      for (var col_name in indx){
        if (col_name == "Chapter Status"){
          var member_status = row[indx["Chapter Status"]];
          member_status = member_status=="Prospective Pledge" ? "Pledge":member_status;
          member_status = member_status=="Student of Colony" ? "Student":member_status;
          member_object[col_name] = member_status;
          continue;
        } else if(col_name == "Chapter Role"){
          var role = row[indx[col_name]];
          if(role != ""){
            var role_start = new Date(row[role_start_col]);
            var role_end = new Date(row[role_end_col]);
            if (date_today >= role_start && date_today < role_end){
              member_object[col_name] = role;
              continue;
            } else {
              member_object[col_name] = "";
              continue;
            }
          } else {
            member_object[col_name] = "";
            continue;
          }
        }
        member_object[col_name] = row[indx[col_name]];
      }
      if (CentralMemberObject['badge_numbers'].indexOf(badge_number) > -1){
        var new_role = member_object["Chapter Role"];
        if (new_role == ""){continue;}
        var prev_role = CentralMemberObject[badge_number]["Chapter Role"];
        prev_role = (prev_role == "") ? new_role: prev_role+= ", " + new_role;
        CentralMemberObject[badge_number]["Chapter Role"] = new_role;
        continue;
      }
      CentralMemberObject['badge_numbers'].push(badge_number);
      CentralMemberObject[badge_number] = member_object;
    }
  }
  progress_update("Found "+ CentralMemberObject['badge_numbers'].length +" Chapter Members");
  Logger.log("(" + arguments.callee.name + ") " +CentralMemberObject[badge_number]);
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Membership");
  var max_row = sheet.getLastRow() - 1;
  var max_row = (max_row != 0) ? max_row:1;
  var max_column = sheet.getLastColumn();
  var header_range = sheet.getRange(1, 1, 1, max_column);
  var header_values = header_range.getValues()[0];
  var badge_index_chapter = header_values.indexOf("Badge Number");
  var ChapterMemberObject = main_range_object("Membership", "Badge Number");
  Logger.log("(" + arguments.callee.name + ") ");
  Logger.log(ChapterMemberObject["object_header"]);
  var new_members = [];
  var verify_members = [];
  for (var k in CentralMemberObject['badge_numbers']){
    var badge_number = CentralMemberObject['badge_numbers'][k];
    if (ChapterMemberObject["object_header"].indexOf(badge_number) < 0){
      // Member is on Central list, not on chapter list
      Logger.log("(" + arguments.callee.name + ") " +"NEW MEMBER TO SHEET!");
      Logger.log("(" + arguments.callee.name + ") " +CentralMemberObject[badge_number]['Last Name']);
      new_members.push(badge_number);
    } else {
      Logger.log("(" + arguments.callee.name + ") " +"VERIFY MEMBER!");
      Logger.log("(" + arguments.callee.name + ") " +CentralMemberObject[badge_number]['Last Name']);
      verify_members.push(badge_number);
      // Member is already on chapter list, need to check for update
    }
  }
  progress_update("Found "+ new_members.length +" NEW Chapter Members TO SHEET");
  progress_update("Found "+ verify_members.length +" Previous Chapter Members ON SHEET");
//  
  var old_members = [];
  for (var k in ChapterMemberObject["object_header"]){
    var badge_number = ChapterMemberObject["object_header"][k];
    if (CentralMemberObject['badge_numbers'].indexOf(badge_number) < 0){
      // Member is on chapter list, not on central list
      // Need to remove from Membership Sheet
      Logger.log("(" + arguments.callee.name + ") " +"OLD MEMBER!");
      Logger.log("(" + arguments.callee.name + ") " +ChapterMemberObject[badge_number]['Last Name'][0]);
      old_members.push(badge_number)
    }
  }
  progress_update("Found "+ old_members.length +" Chapter Members to Remove");
  verify_members.sort();
  verify_members.reverse();  
  for (var q in verify_members){
    var badge = verify_members[q];
    var this_row = ChapterMemberObject[badge]['object_row'];
    var member = CentralMemberObject[badge];
    var full_name = CentralMemberObject[badge]["First Name"]+" "+CentralMemberObject[badge]["Last Name"];
    var chapter_member_name = ChapterMemberObject[badge]["Member Name"][0];
    if (full_name != chapter_member_name){
      var col = ChapterMemberObject[badge]["Member Name"][1];
      sheet.getRange(this_row, col).setValue(full_name);
    }
    for (var col_name in indx){
      var col_val = ChapterMemberObject[badge][col_name][0];
      var member_val = CentralMemberObject[badge][col_name];
      if (col_val == "Away" || col_val == "Alumn" || col_val == "Shiny"){continue;}
      if (col_val != member_val){
        var col = ChapterMemberObject[badge][col_name][1];
        sheet.getRange(this_row, col).setValue(member_val);
      }
    }
    var col = ChapterMemberObject[badge]['Last Update'][1];
    sheet.getRange(this_row, col).setValue(new_date);
    var range_note = sheet.getRange(this_row+1, 1);
    range_note.clearNote();
  }
  old_members.sort();
  old_members.reverse();
  var delete_att = [];
//  var alumn = [];
  for (var p in old_members){
    var badge = old_members[p];
    var this_row = ChapterMemberObject[badge]['object_row'];
    if (ChapterMemberObject[badge]['Chapter Status'][0] == "Alumn"){
//      alumn.push(ChapterMemberObject[badge]['Member Name']);
      continue;}
    delete_att.push(ChapterMemberObject[badge]['Member Name']);
    var badge_ind = ChapterMemberObject["object_header"].indexOf(badge);
    ChapterMemberObject["object_header"].splice(badge_ind, 1);
    delete ChapterMemberObject[badge];
    sheet.deleteRow(this_row);
  }
  new_members.sort();
  new_members.reverse();
  for (var m in new_members){
    var this_row = 1;
    var new_badge = new_members[m];
    for (var k in ChapterMemberObject["object_header"]){
//      Logger.log("(" + arguments.callee.name + ") " +ChapterMemberObject["object_header"]);
      var badge_number = ChapterMemberObject["object_header"][k];
      var badge_next = ChapterMemberObject["object_header"][+k+1];
      badge_next = badge_next ? badge_next:new_badge+1;
//      Logger.log("(" + arguments.callee.name + ") " +[badge_number, new_badge, badge_next]);
      if (new_badge > badge_number && new_badge < badge_next){
        this_row = ChapterMemberObject[badge_number]['object_row'];
        break;
      }
    }
    sheet.insertRowAfter(this_row);
    var range = sheet.getRange(this_row+1, 1, 1, 13);
    var range_note = sheet.getRange(this_row+1, 1);
//    range_note.setNote("Member Info Updated: "+new_date);
    range_note.clearNote();
    var member_object = CentralMemberObject[new_badge];
    var new_values = [];
    for (var j in ChapterMemberObject["header_values"]){
      var header = ChapterMemberObject["header_values"][j];
      if (header == "Member Name"){
        var full_name = member_object["First Name"]+" "+member_object["Last Name"];
        new_values.push(full_name);
        continue;
      }
      if (header == "Status Start" || header == "Status End"){
        new_values.push("");
        continue;
      }
      if (header == "Last Update"){
        new_values.push(new_date);
        continue;
      }
      var new_value = member_object[header];
      if ( new_value !== undefined){
        new_values.push(new_value);
      }
    }
    Logger.log("(" + arguments.callee.name + ") " +new_values);
    range.setValues([new_values]);
  }
//  setup_attendance();
//  check_duplicate_names();
  progress_update("Finished Get Chapter Members");
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

function shorten_membership_list(object_header) {
  var short_list = [];
  for (var i in object_header){
    var short = shorten(object_header[i], 20, false);
    short_list.push(short);
  }
  return short_list;
}

//function setup_attendance(){
////  var delete_att = ["Jeremy Faber", "Eugene Balaguer", "Jacob Landsiedel"];
//  progress_update("Started Updating Attendance Sheet");
//  Logger.log("(" + arguments.callee.name + ") ");
//  var previous_member = undefined;
//  var ChapterMemberObject = main_range_object("Membership");
//  var ss = get_active_spreadsheet();
//  var sheet = ss.getSheetByName("Attendance");
//  var max_column = sheet.getLastColumn();
//  var header_range = sheet.getRange(1, 1, 1, max_column);
//  var header_values = header_range.getValues()[0];
//  ChapterMemberObject["short_names"] = shorten_membership_list(ChapterMemberObject["object_header"]);
//  progress_update("Removing From Att, not on Membership");
//  var AttendanceObject = main_range_object("Attendance", "None", ss);
//  var att_header_values = AttendanceObject["header_values"];
//  att_header_values.reverse();
//  for(var i in att_header_values) {
//    var member_name = att_header_values[i];
//    if (ChapterMemberObject["short_names"].indexOf(member_name) > -1){
//      continue;
//    }
//    var header_values = header_range.getValues()[0];
//    var member_name_short = shorten(member_name, 20, false);
//    var col = header_values.indexOf(member_name_short)+1;
//    if (col > 2){
//      sheet.deleteColumn(col);
//    }
//  }
//  var AttendanceObject = main_range_object("Attendance", "None", ss);
//  for(var i in ChapterMemberObject["short_names"]) {
//    var member_name_short = ChapterMemberObject["short_names"][i];
//    if (AttendanceObject["header_values"].indexOf(member_name_short) > -1){
//      continue;
//    }
//    previous_member = ChapterMemberObject["short_names"][i-1];
//    align_attendance_members(previous_member, member_name_short, sheet);
//  }
//  progress_update("Finished Updating Attendance Sheet");
//  var format_range = ss.getRangeByName("FORMAT");
//  var max_column = sheet.getLastColumn();
//  var max_row = sheet.getLastRow();
//  format_range.copyFormatToRange(sheet, 3, max_column, 2, 100);
//  sheet.getRange(3, 2, max_row, max_column).clearDataValidations();
//  sheet.setRowHeight(1, 100);
//}


//function align_attendance_members(previous_member, new_member, sheet){
////  var previous_member = "REALLYLONGNAMEFORTESTINGTHINGSLIKETHIS";
////  var new_member = "REALL3YLONGNAMEFORTESTINGTHINGSLIKETHIS";
//  var max_row = sheet.getLastRow() - 1;
//  max_row = (max_row != 0) ? max_row:1;
//  var max_column = sheet.getLastColumn();
//  var header_range = sheet.getRange(1, 1, 1, max_column);
//  var header_values = header_range.getValues()[0];
//  var previous_index = 2;
//  if (previous_member !== undefined){
//    for (var i in header_values){
//      var header_name = header_values[i];
//      if (header_name == "Event Name" || header_name == "Date"){
//        continue;
//      }
//      var new_string = att_name(header_name);
//      if (new_string == previous_member){
//        previous_index = parseInt(i)+1;
//      }
//    }
//  }
//  sheet.insertColumnAfter(previous_index);
//  var new_range = sheet.getRange(1, +previous_index+1);
//  new_range.setValue(new_member);
//  sheet.setColumnWidth(+previous_index+1, 50);
//}

function RESET() {
  unlock();
  var this_password = SCRIPT_PROP.getProperty("password");
  if (this_password != password){
    var ui = SpreadsheetApp.getUi();
    ui.alert('Incorrect Password!');
    return;
  }
  var target_doc = get_active_spreadsheet();
  var folder_id = SCRIPT_PROP.getProperty("folder");
  if (folder_id){
    var folder_chapter = DriveApp.getFolderById(folder_id);
    var file = DriveApp.getFileById(SCRIPT_PROP.getProperty("key"));
    folder_chapter.removeFile(file);
    var form_folder = DriveApp.getFolderById(get_form_id());
    var submit_folder = DriveApp.getFolderById(get_submit_id());
    folder_chapter.removeFolder(form_folder);
    folder_chapter.removeFolder(submit_folder);
    folder_chapter.setTrashed(true);
  }
    
  var sheets = target_doc.getSheets();
  var new_sheet = target_doc.insertSheet();
  
  for (var i in sheets){
    var sheet = sheets[i];
    target_doc.deleteSheet(sheet)
  }
  new_sheet.setName("Sheet1");
  var named_ranges = target_doc.getNamedRanges();
  for (var j in named_ranges){
    var named_range = named_ranges[j];
    named_range.remove();
  }
  target_doc.rename("BLANK");
  var survey_id = SCRIPT_PROP.getProperty("survey");
  var file = DriveApp.getFileById(survey_id);
  file.setTrashed(true);
  SCRIPT_PROP.deleteAllProperties();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
 }
}