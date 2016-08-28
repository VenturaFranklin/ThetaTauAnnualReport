function onInstall(e) {
  onOpen(e);
  setup();
  setup_sheets();
  chapter_name();
}

function createTriggers() {
  try {
    var ss = get_active_spreadsheet();
    ScriptApp.newTrigger('_onEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();
  } catch (error) {
    Logger.log(error);
  }    
}

function get_chapter_name(){
  return SCRIPT_PROP.getProperty("chapter")
}

function chapter_name_process(form) {
//  Logger.log(form);
//  var form = {'chapterslist': 'Chi Gamma'}
  var chapter_name = form.chapterslist;
  SCRIPT_PROP.setProperty("chapter", chapter_name);
  var properties_id = "1vCVKh8MExPxg8eHTEGYx7k-KTu9QUypGwbtfliLm58A";
  var ss_prop = SpreadsheetApp.openById(properties_id);
  var ss = get_active_spreadsheet();
  var default_id = "19aWLtjJJ-Uh6XOqOuseLpQcNJYslQHe9Y9Gaj2vSjEw";
  var default_doc = SpreadsheetApp.openById(default_id);
  var doc_name = default_doc.getName();
  doc_name = doc_name.replace("DEFAULT ", "");
  doc_name = doc_name.replace("- Chapter", "- "+chapter_name);
  Logger.log(doc_name);
  ss.rename(doc_name);
  var sheet = ss.getSheetByName("Chapter");
  var range = sheet.getRange(2, 2);
  range.setValue(chapter_name);
  var test = range.getValue();
  var chapter_object = main_range_object("MAIN", "Organization Name", ss_prop);
  var chapter_info = chapter_object[chapter_name];
  Logger.log(chapter_info);
  var region = chapter_info["Region Description"][0];
  range = sheet.getRange(2, 3);
  range.setValue(region);
  SCRIPT_PROP.setProperty("region", region);
  var balance = chapter_info["Balance Due"][0];
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
  range = sheet_dash.getRange(1, 1);
  range.setValue(chapter_name + " CHAPTER ANNUAL REPORT");
  range.getValue();
  create_submit_folder(chapter_name, region);
  get_chapter_members();
  createTriggers();
  var ui = SpreadsheetApp.getUi();
  ui.alert('SETUP COMPLETE!');
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
//  var chapter_name = "chapter_test";
//  var region = "region_test";
  folder_id = "0BwvK5gYQ6D4nTDRtY1prZG12UU0";
  var folder_submit = DriveApp.getFolderById(folder_id);
  var folder_region = folder_submit.getFoldersByName(region);
  if (folder_region.hasNext()) {
    folder_region = folder_region.next()
  } else {
    folder_region = folder_submit.createFolder(region);
  }
  var folder_chapter = folder_region.getFoldersByName(chapter_name);
  if (folder_chapter.hasNext()) {
    folder_chapter = folder_chapter.next()
  } else {
    folder_chapter = folder_region.createFolder(chapter_name);
  }
  var folder_id = folder_chapter.getId();
  SCRIPT_PROP.setProperty("folder", folder_id);
  var file = DriveApp.getFileById(SCRIPT_PROP.getProperty("key"));
  folder_chapter.addFile(file);
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
      .setWidth(125)
      .setHeight(75);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(htmlOutput, "Chapter Name");
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
  if (button == ui.Button.OK) {
    ui.alert('Verify you sheet id is: ' + text);
  } else {
    // User clicked "Cancel".
    ui.alert('The scripts will not work without the Sheet ID.');
  }
  SCRIPT_PROP.setProperty("key", text);
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
  var sheet = target_doc.getSheetByName("Sheet1");
  target_doc.deleteSheet(sheet);
  var sheet = target_doc.getSheetByName("Dashboard"); //A1
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
  var chapter_name = get_chapter_name();
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
    var new_badge = new_members[m];
    for (var k in ChapterMemberObject["object_header"]){
//      Logger.log(ChapterMemberObject["object_header"]);
      var badge_number = ChapterMemberObject["object_header"][k];
      var badge_next = ChapterMemberObject["object_header"][+k+1];
      badge_next = badge_next ? badge_next:new_badge+1;
//      Logger.log([badge_number, new_badge, badge_next]);
      if (new_badge > badge_number && new_badge < badge_next){
        this_row = ChapterMemberObject[badge_number]['object_row'];
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
  }
  var previous_member = undefined;
  ChapterMemberObject = main_range_object("Membership");
  for(var i = 0; i< ChapterMemberObject.object_count; i++) {
    var member_name = ChapterMemberObject.object_header[i];
    var member_badge = ChapterMemberObject[member_name]["Badge Number"][0];
    if (new_members.indexOf(member_badge) > -1){
      if (i>0){
        previous_member = ChapterMemberObject.object_header[i-1];;
      }
      align_attendance_members(previous_member, member_name);
    }
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

function RESET() {
  var target_doc = get_active_spreadsheet();
  var folder_id = SCRIPT_PROP.getProperty("folder");
  if (folder_id){
    var folder_chapter = DriveApp.getFolderById(folder_id);
    var file = DriveApp.getFileById(SCRIPT_PROP.getProperty("key"));
    folder_chapter.removeFile(file);
  }
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
  PropertiesService.getScriptProperties().deleteAllProperties();
}