function run_install(e){
  unlock();
  var this_password = SCRIPT_PROP.getProperty("password");
  if (this_password != password){
    var ui = SpreadsheetApp.getUi();
    ui.alert('Incorrect Password!');
    return;
  }
  onInstall(e);
}

function onInstall(e) {
  onOpen(e);
  setup();
}

var message = ""

function progress_update(this_message){
  message += "<br>" + this_message;
  var htmlOutput = HtmlService
     .createHtmlOutput(message)
     .setWidth(400)
     .setHeight(300)
  SpreadsheetApp.getUi()
    .showModalDialog(htmlOutput, 'Progress');
}

function createTriggers() {
  try {
    unlock();
    var this_password = SCRIPT_PROP.getProperty("password");
    if (this_password != password){
      var ui = SpreadsheetApp.getUi();
      ui.alert('Incorrect Password!');
      return;
    }
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
    Logger.log("(" + arguments.callee.name + ") " +error);
  } 
  try {
    progress_update("Creating Sync Trigger");
    ScriptApp.newTrigger("sync")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .create();
  } catch (error) {
    Logger.log("(" + arguments.callee.name + ") " +error);
  }
}

function get_chapter_name(){
  return SCRIPT_PROP.getProperty("chapter")
}

function get_chapter_fee(){
  return SCRIPT_PROP.getProperty("chapter_fee")
}

function chapter_name_process(form) {
  Logger.log("(" + arguments.callee.name + ") " +form);
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
  doc_name = doc_name.replace("- Chapter", "- "+chapter_name);
  Logger.log("(" + arguments.callee.name + ") " +doc_name);
  ss.rename(doc_name);
  var sheet = ss.getSheetByName("Chapter");
  var range = sheet.getRange(2, 2);
  range.setValue(chapter_name);
  var test = range.getValue();
  var chapter_object = main_range_object("MAIN", "Organization Name", ss_prop);
  var chapter_info = chapter_object[chapter_name];
  progress_update("Chapter Information: " + chapter_info.toString());
  Logger.log("(" + arguments.callee.name + ") " +chapter_info);
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
  range.setValue(chapter_name + " CHAPTER ANNUAL REPORT");
  range.getValue();
  create_submit_folder(chapter_name, region);
  get_chapter_members();
  setup_dataval();
  createTriggers();
  progress_update("Started Sync Main Info");
  sync_main()
  var ui = SpreadsheetApp.getUi();
  ui.alert('SETUP COMPLETE!\n'+
           'Next steps:\n'+
           '- Fill out Chapter Sheet\n'+
           '- Verify Membership\n'+
           '- Add Events & Attendance\n\n'+
           'Do not edit gray or black cells\n'+
           'Submit forms in menu "Add-ons-->ThetaTauReports"');
}

function protect_ranges(){
  // Can not have current user removed from protection.
  var ss = get_active_spreadsheet();
  var emailAddress = "venturafranklin@gmail.com";
  var sheet = ss.getSheetByName("Events");
  var range = sheet.getRange('C:E');
  var protection = range.protect().setDescription('EventScoreMemberPledge');
  ss.addEditor(emailAddress);
  protection.addEditor(emailAddress);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
   protection.setDomainEdit(false);
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
  for (var row in main_values){
    if (row.indexOf(chapter_name) > -1){
      found_chapter = true;
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
    folder_chapter = folder_chapter.next()
    progress_update("Found Chapter Folder: " + chapter_name);
  } else {
    folder_chapter = folder_region.createFolder(chapter_name);
    progress_update("Created Chapter Folder: " + chapter_name);
  }
  var folder_id = folder_chapter.getId();
  SCRIPT_PROP.setProperty("folder", folder_id);
  var file = DriveApp.getFileById(SCRIPT_PROP.getProperty("key"));
  DriveApp.addFolder(folder_chapter); // Adds the chapter folder to user drive
  folder_chapter.addFile(file); // Adds the ss to chapter folder
  DriveApp.removeFile(file); // Removes ss from user drive
}

function get_folder_id() {
  return SCRIPT_PROP.getProperty("folder");
  progress_update("Finished Submit Folder Creation");
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
  setup_sheets();
  chapter_name();
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
  var events = get_type_list("Events");
  var range = ss.getRangeByName("EventsType");
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(events)
    .setHelpText('Must be a valid event type.')
    .setAllowInvalid(false).build();
  range.setDataValidation(rule);

  var yes_no = ["EventsSTEM", "EventsPledge", "EventsHost",
                "MemberPro", "MemberHonor", "MemberOther"];
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'])
    .setAllowInvalid(false).build();
  for (var i in yes_no){
    var range_name = yes_no[i];
    var range = ss.getRangeByName(range_name);
    range.setDataValidation(rule);
  }

  var range = ss.getRangeByName("EventsDate");
  var rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setHelpText('Enter a valid date MM/DD/YYYY')
    .setAllowInvalid(false).build();
  range.setDataValidation(rule);
  
  var range = ss.getRange("Attendance!1:149");
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['P', 'E', 'U'], false)
    .setHelpText('P-Present; E-Excused; U-Unexcused')
    .setAllowInvalid(false).build();
  range.setDataValidation(rule);
  
  remove = ["Membership!1:1", "Events!1:1",
            "Attendance!1:1", "Attendance!A:B"];
  for (var i in remove) {
    ss.getRange(remove[i]).clearDataValidations();
  }
  ss.getRange("Membership!M2:V100").setBackground("white");
  progress_update("Finished Data Val Setup");
//requireNumberGreaterThan(number)
//requireTextIsEmail()
//requireTextIsUrl()
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

function get_chapter_members(){
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
  var csvData = CSVToArray(csvFile, ",");
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
    "Chapter Role": header.indexOf("Last Name"), //Organization Relation Relationship
    "Current Major": header.indexOf("Primary Education Major"),
    "School Status": header.indexOf("Primary Education Class of"),
    "Phone Number": header.indexOf("Mobile Phone Number"),
    "Email Address": header.indexOf("Email Address Number")
  };
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
          member_status = member_status=="Colony" ? "Student":member_status;
          member_object[col_name] = member_status;
          continue;
        }
        member_object[col_name] = row[indx[col_name]];
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
  Logger.log("(" + arguments.callee.name + ") " +ChapterMemberObject["object_header"]);
  var new_members = [];
  var verify_members = [];
  for (var k in CentralMemberObject['badge_numbers']){
    var badge_number = CentralMemberObject['badge_numbers'][k];
    if (ChapterMemberObject["object_header"].indexOf(badge_number) < 0){
      // Member is on Central list, not on chapter list
      Logger.log("(" + arguments.callee.name + ") " +"NEW MEMBER!");
      Logger.log("(" + arguments.callee.name + ") " +CentralMemberObject[badge_number]['Last Name']);
      new_members.push(badge_number);
    } else {
      Logger.log("(" + arguments.callee.name + ") " +"VERIFY MEMBER!");
      Logger.log("(" + arguments.callee.name + ") " +CentralMemberObject[badge_number]['Last Name']);
      verify_members.push(badge_number);
      // Member is already on chapter list, need to check for update
    }
  }
  progress_update("Found "+ new_members.length +" NEW Chapter Members");
  progress_update("Found "+ verify_members.length +" Previous Chapter Members");
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
    for (var col_name in indx){
      var col_val = ChapterMemberObject[badge][col_name][0];
      var member_val = CentralMemberObject[badge][col_name];
      if (col_val == "Away" || col_val == "Alumn"){continue;}
      if (col_val != member_val){
        var col = ChapterMemberObject[badge][col_name][1];
        sheet.getRange(this_row, col).setValue(member_val);
      }
    }
  }
  old_members.sort();
  old_members.reverse();
  var delete_att = [];
  for (var p in old_members){
    var badge = old_members[p];
    var this_row = ChapterMemberObject[badge]['object_row'];
    if (ChapterMemberObject[badge]['Chapter Status'][0] == "Alumn"){continue;}
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
    var range = sheet.getRange(this_row+1, 1, 1, 10);
    var range_note = sheet.getRange(this_row+1, 1);
    range_note.setNote("Member Info Updated: "+new_date);
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
    Logger.log("(" + arguments.callee.name + ") " +new_values);
    range.setValues([new_values]);
  }
  setup_attendance();
  progress_update("Finished Get Chapter Members");
}

function shorten_membership_list(object_header) {
  var short_list = [];
  for (var i in object_header){
    var short = shorten(object_header[i], 12, false);
    short_list.push(short);
  }
  return short_list;
}

function setup_attendance(){
//  var delete_att = ["Jeremy Faber", "Eugene Balaguer", "Jacob Landsiedel"];
  progress_update("Started Updating Attendance Sheet");
  Logger.log("(" + arguments.callee.name + ") ");
  var previous_member = undefined;
  var ChapterMemberObject = main_range_object("Membership");
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Attendance");
  var max_column = sheet.getLastColumn();
  var header_range = sheet.getRange(1, 1, 1, max_column);
  var header_values = header_range.getValues()[0];
  ChapterMemberObject["short_names"] = shorten_membership_list(ChapterMemberObject["object_header"]);
  progress_update("Removing From Att, not on Membership");
  var AttendanceObject = main_range_object("Attendance", "None", ss);
  var att_header_values = AttendanceObject["header_values"];
  att_header_values.reverse();
  for(var i in att_header_values) {
    var member_name = att_header_values[i];
    if (ChapterMemberObject["short_names"].indexOf(member_name) > -1){
      continue;
    }
    var header_values = header_range.getValues()[0];
    var member_name_short = shorten(member_name, 12, false);
    var col = header_values.indexOf(member_name_short)+1;
    if (col > 2){
      sheet.deleteColumn(col);
    }
  }
  var AttendanceObject = main_range_object("Attendance", "None", ss);
  for(var i in ChapterMemberObject["short_names"]) {
    var member_name_short = ChapterMemberObject["short_names"][i];
    if (AttendanceObject["header_values"].indexOf(member_name_short) > -1){
      continue;
    }
    previous_member = ChapterMemberObject["short_names"][i-1];
    align_attendance_members(previous_member, member_name_short, sheet);
  }
  progress_update("Finished Updating Attendance Sheet");
  var format_range = ss.getRangeByName("FORMAT");
  var max_column = sheet.getLastColumn();
  var max_row = sheet.getLastRow();
  format_range.copyFormatToRange(sheet, 3, max_column, 2, 100);
  sheet.getRange(3, 2, max_row, max_column).clearDataValidations();
  sheet.setRowHeight(1, 100);
}


function align_attendance_members(previous_member, new_member, sheet){
//  var previous_member = "REALLYLONGNAMEFORTESTINGTHINGSLIKETHIS";
//  var new_member = "REALL3YLONGNAMEFORTESTINGTHINGSLIKETHIS";
  var max_row = sheet.getLastRow() - 1;
  max_row = (max_row != 0) ? max_row:1;
  var max_column = sheet.getLastColumn();
  var header_range = sheet.getRange(1, 1, 1, max_column);
  var header_values = header_range.getValues()[0];
  var previous_index = 2;
  if (previous_member !== undefined){
    for (var i in header_values){
      var header_name = header_values[i];
      if (header_name == "Event Name" || header_name == "Date"){
        continue;
      }
      var new_string = att_name(header_name);
      if (new_string == previous_member){
        previous_index = parseInt(i)+1;
      }
    }
  }
  sheet.insertColumnAfter(previous_index);
  var new_range = sheet.getRange(1, +previous_index+1);
  new_range.setValue(new_member);
  sheet.setColumnWidth(+previous_index+1, 50);
}

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
  SCRIPT_PROP.deleteAllProperties();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
 }
}