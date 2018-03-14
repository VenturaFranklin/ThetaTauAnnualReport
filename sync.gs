function sync(local){
  try{
    if (!local){
      SILENT = true;
    }
    progress_update("STARTING TO SYNC");
    sync_region();
    sync_main();
    progress_update("FINISHED SYNC");
    } catch (e) {
    var message = Utilities.formatString('This error has automatically been sent to the developers. %s: %s (line %s, file "%s"). Stack: "%s" . While processing %s.',
                                         e.name||'', e.message||'', e.lineNumber||'', e.fileName||'',
                                         e.stack||'', arguments.callee.name||'');
    Logger = startBetterLog();
    Logger.severe(message);
//    var ui = SpreadsheetApp.getUi();
//    var result = ui.alert(
//     'ERROR',
//      message,
//      ui.ButtonSet.OK);
    return "";
  }
}

function submit_report(){
  var local = true;
  sync(local);
  progress_update("Emailing submission to RDs!");
  RD_email();
  progress_update("Submission complete!");
}

function RD_email() {
//  var form = get_survey();
//  var email = "venturafranklin@gmail.com";
  var email = SCRIPT_PROP.getProperty("director");
  var ss = get_active_spreadsheet();
  Logger.log("(" + arguments.callee.name + ") ");
  Logger.log("(" + arguments.callee.name + ") " + email);
  var url = ss.getUrl();
  Logger.log("(" + arguments.callee.name + ") " + url);
  // Fetch form's HTML
//  var response = UrlFetchApp.fetch(url);
//  var htmlBody = HtmlService.createHtmlOutput(response).getContent();
  var chapter_name = get_chapter_name();
  var subject = chapter_name + " Chapter Annual Report Submission";
  var email_chapter = SCRIPT_PROP.getProperty("email");
  var emailBody = "Your chapter has submitted their annual report. Please see the live report here:"+
    "\nSurvey("+url+")";

  var htmlBody = "Your chapter has submitted their annual report. Please see the live report here:"+
    '<br/><a href="'+url+'"> Survey</a> ('+url+')';
  var optAdvancedArgs = {name: chapter_name +" Chapter Scribe", htmlBody: htmlBody,
                         replyTo: email_chapter};
  if (!WORKING){
    try{
      MailApp.sendEmail(email, subject,
                        emailBody,
                        optAdvancedArgs);
    } catch (e){
      var ui = SpreadsheetApp.getUi();
      var result = ui.alert(
        'EMAIL ERROR',
        "There was an error with the email for: "+ email,
        ui.ButtonSet.OK);
    }
  }
}

function sync_region() {
//  var dash_id = "10ebwK7tTKgveVCEOpRle2S17d4UjwmsoXXCPFvC9A-A";
  progress_update("STARTED TO SYNC REGION INFO");
  var dash_id = SCRIPT_PROP.getProperty("dash");
  var dash_file = SpreadsheetApp.openById(dash_id);
  var chapter = SCRIPT_PROP.getProperty("chapter");
  var main_sheet = dash_file.getSheetByName("MAIN");
  var main_values = main_sheet.getDataRange().getValues();
  var main_chapter = filter_chapter(main_values, chapter);
  var submit_sheet = dash_file.getSheetByName("SUBMISSIONS");
  var submit_values = submit_sheet.getDataRange().getValues();
  var submit_chapter = filter_chapter(submit_values, chapter, "SUBMISSIONS");
  var event_sheet = dash_file.getSheetByName("EVENTS");
  var event_values = event_sheet.getDataRange().getValues();
  var event_chapter = filter_chapter(event_values, chapter, "EVENTS");
  var event_object = main_range_object("Events");
  var submit_object = main_range_object("Submissions");
  var event_extend = [];
  progress_update("Syncing Events");
  for (var event_name in event_object){
    var event = event_object[event_name];
    if (typeof event != typeof {}){continue;};
    if ('Event Name' in event){
      var test_name  = event;
      if (!(test_name in event_chapter)){
        event_extend.push([chapter,	event['Event Name'][0], event['Date'][0], event['Type'][0],
                           event['Description'][0]]);
      } else {
        var col = event_chapter.first_row.indexOf("Type")+1;
        var event_row = event_chapter[test_name][1];
        event_sheet.getRange(event_row, col)
        .setValue(event['Type'][0]);
        var col = event_chapter.first_row.indexOf("Description")+1;
        event_sheet.getRange(event_row, col)
        .setValue(event['Description'][0]);
      }
    }
  }
  progress_update("Syncing Submissions");
  var submit_extend = [];
  for (var submit in submit_object){
    submit = submit_object[submit];
    if (typeof submit != typeof {}){continue;};
    if ('Type' in submit){
      var test_name  = submit['Type'][0] + submit['Date'][0];
      if (!(test_name in submit_chapter)){
        submit_extend.push([chapter, submit['Date'][0], submit['File Name'][0],
                            submit['Type'][0], submit['Location of Upload'][0]]);
      }
    }
  }
  progress_update("Syncing Chapter info");
  var member_ranges = get_membership_ranges();
  var chapter_row = main_values.length + 1;
  if (chapter+chapter in main_chapter){
    chapter_row = main_chapter[chapter+chapter][1];
  } else {
    // Generally this should not happen, needed more for testing.
    main_sheet.getRange(chapter_row, 1).setValue(chapter);
    main_sheet.getRange(chapter_row, 2).setValue(SCRIPT_PROP.getProperty("email"));
    main_sheet.getRange(chapter_row, 3).setValue(SCRIPT_PROP.getProperty("tax"));
  }
  var update_cnt = 0;
  for (var member_range_year in member_ranges){
    for (var member_range_type in member_ranges[member_range_year]){
      var score_type = member_range_year + " " + member_range_type;
      var col = main_chapter.first_row.indexOf(score_type)+1;
      if (col == 0){
        main_sheet.insertColumns(11 + update_cnt);
        col = 11 + update_cnt;
        update_cnt += 1;
        main_sheet.getRange(1, col).setValue(score_type);
      }
      var val = member_ranges[member_range_year][member_range_type].value[0];
      main_sheet.getRange(chapter_row, col).setValue(val).setWrap(true);
    }
  }
  var ss = get_active_spreadsheet();
  var scores = ['Brotherhood', 'Service', 'Operate', 'ProDev'];
  progress_update("Syncing Scores");
  var score_total = 0;
  for (var score_num in scores){
    var score_type_raw  = scores[score_num];
    var score_type = "SCORE_" + score_type_raw.toUpperCase();
    var score = ss.getRangeByName(score_type).getValue();
    score_total += score;
    var score_col = main_chapter.first_row.indexOf(score_type_raw)+1;
    main_sheet.getRange(chapter_row, score_col).setValue(score);
  }
  var score_col = main_chapter.first_row.indexOf("Total")+1;
  main_sheet.getRange(chapter_row, score_col).setValue(score_total);
  var update_col = main_chapter.first_row.indexOf("Last Updated")+1;
  main_sheet.getRange(chapter_row, update_col).setValue(new Date());
  var address_col = main_chapter.first_row.indexOf("Address")+1;
  // var chapter_address_range = ss.getRangeByName("Address")
  //var chapter_address = chapter_address_range.getValue();
  //if (chapter_address != ""){
    //main_sheet.getRange(chapter_row, address_col).setValue(chapter_address);
    //chapter_address_range.clearNote()
    //  .setBackground('white');
  //} else {
    //chapter_address_range.setNote("Scribe should set this value")
     // .setBackground('red');
  //}
  var event_row_max = event_sheet.getLastRow();
  var event_col_max = event_sheet.getLastColumn();
  for (var row_ind in event_extend){
    var row = event_extend[row_ind];
    event_row_max++;
    event_sheet.getRange(event_row_max, 1, 1, event_col_max)
    .setValues([row]);
  }
  var submit_row_max = submit_sheet.getLastRow();
  var submit_col_max = submit_sheet.getLastColumn();
  for (var row_ind in submit_extend){
    var row = submit_extend[row_ind];
    submit_row_max++;
    submit_sheet.getRange(submit_row_max, 1, 1, submit_col_max)
    .setValues([row]);
  }
  progress_update("FINISHED SYNCING REGION");
}

function delete_rds(){
  SCRIPT_PROP.deleteProperty("Director A");
  SCRIPT_PROP.deleteProperty("Director B");
  SCRIPT_PROP.deleteProperty("Director C");
}

function sync_jewelry(ss_prop){
//  var properties_id = "1vCVKh8MExPxg8eHTEGYx7k-KTu9QUypGwbtfliLm58A";
//  var ss_prop = SpreadsheetApp.openById(properties_id);
  progress_update("Syncing Jewelry Costs");
  Logger.log("(" + arguments.callee.name + ") ");
  var chapter = SCRIPT_PROP.getProperty("chapter");
//  chapter = "Delta Gamma";
  var num = chapter.split(" ");
  var letter = (num.length > 1) ? 2:1;
  var main_object = main_range_object("Jewelry", "Item", ss_prop);
  var badges = {};
  var guards = {};
  for (var row in main_object.object_header){
    var item = main_object[main_object.object_header[row]];
    if (item.hasOwnProperty("Badge Style")){
      badges[item["Badge Style"][0]] = item["Badge Cost"][0];
      guards[item["Guard Type"][0]] = item[letter][0];
    }
  }
  Logger.log(badges);
  Logger.log(guards);
  delete badges[""];
  guards["None"] = 0;
  SCRIPT_PROP.setProperty("badges", JSON.stringify(badges));
  SCRIPT_PROP.setProperty("guards", JSON.stringify(guards));
}

function get_guard_badge(){
  var badges = SCRIPT_PROP.getProperty("badges");
  var guards = SCRIPT_PROP.getProperty("guards");
  var badges = JSON.parse(badges);
  var guards = JSON.parse(guards);
  Logger.log(badges);
  Logger.log(guards);
  return {badges: badges,
          guards: guards}
}
  

function sync_rds(ss_prop){
//  var properties_id = "1vCVKh8MExPxg8eHTEGYx7k-KTu9QUypGwbtfliLm58A";
//  var ss_prop = SpreadsheetApp.openById(properties_id);
  progress_update("Syncing RDs");
  var region = SCRIPT_PROP.getProperty("region");
  var main_object = main_range_object("REGIONS", "Region Description", ss_prop);
  var region_prop = main_object[region];
  for (var j in region_prop.object_header){
    var header = region_prop.object_header[j];
    if (header.indexOf("Director") > -1){
      var val = region_prop[header][0];
      var old_val = SCRIPT_PROP.getProperty(header);
      if (val != old_val && val!=""){
        SCRIPT_PROP.setProperty(header, val);
        share_calendar(val);
      }
    }
  }
}

function sync_main(){
  progress_update("STARTED TO SYNC NATIONAL INFO");
  var properties_id = "1vCVKh8MExPxg8eHTEGYx7k-KTu9QUypGwbtfliLm58A";
  var ss_prop = SpreadsheetApp.openById(properties_id);
  var ss = get_active_spreadsheet();
  var main_object = main_range_object("MAIN", "Organization Name", ss_prop);
  var top_avg = calc_top_average(main_object);
  progress_update("Updating Top/Avg Scores");
  for (var attr in top_avg.nat_avgs){
    var top_name = "TOP_"+attr.toUpperCase();
    var nat_name = "NAT_"+attr.toUpperCase();
    ss.getRangeByName(top_name).setValue(top_avg.top_avgs[attr].toFixed(1));
    ss.getRangeByName(nat_name).setValue(top_avg.nat_avgs[attr].toFixed(1));
  }
  var chapter_names = top_avg.top_chapter_names.toString()
  ss.getRangeByName("top_chapter_names").setNotes([
    [chapter_names],[chapter_names],[chapter_names],
    [chapter_names],[chapter_names],[chapter_names]]);
  var chapter = SCRIPT_PROP.getProperty("chapter");
  var chapter_row = main_object[chapter].object_row;
  var main_sheet = main_object.sheet;
  var scores = ['Brotherhood', 'Service', 'Operate', 'ProDev'];
  for (var score_num in scores){
    var score_type_raw  = scores[score_num];
    var score_type = "SCORE_" + score_type_raw.toUpperCase();
    var score = ss.getRangeByName(score_type).getValue();
    var score_col = main_object.header_values.indexOf(score_type_raw)+1;
    main_sheet.getRange(chapter_row, score_col).setValue(score);
  }
  var update_col = main_object.header_values.indexOf("Last Updated")+1;
  main_sheet.getRange(chapter_row, update_col).setValue(new Date());
  var school_col = main_object.header_values.indexOf("School Name")+1;
  var school = main_sheet.getRange(chapter_row, school_col).getValue();
  SCRIPT_PROP.setProperty("school", school);
  sync_rds(ss_prop);
  sync_jewelry(ss_prop);
  progress_update("FINISHED SYNCING NATIONALS");
}

function calc_top_average(main_object){
  var nat_avgs = {Brotherhood: 0,
                  Service: 0,
                  Operate: 0,
                  ProDev: 0};
  var top_avgs = {Brotherhood: 0,
                  Service: 0,
                  Operate: 0,
                  ProDev: 0};
  var totals = Object.keys(main_object).map(
    function (key) {
      if (typeof main_object[key] != typeof {}){return 0;};
      if ("Total" in main_object[key]){
        return main_object[key]["Total"][0];}
      else{return 0;}});
  var maxs = [];
  var maxs_length = 1;
  var next_max = 0;
  var this_max = 1;
  while (maxs_length < 5 && next_max <= this_max){
    this_max = Math.max.apply(Math, totals);
    maxs.push(this_max);
    totals.splice(totals.indexOf(this_max), 1);
    next_max = Math.max.apply(Math, totals);
    if (next_max != this_max){maxs_length++;}
  }
  Logger.log("(" + arguments.callee.name + ") " +maxs);
  maxs_length = maxs.length;
  var top_chapter_names = [];
  for (var chapter_count in main_object.object_header){
    var chapter = main_object.object_header[chapter_count];
    var Total = main_object[chapter]["Total"][0];
//    Logger.log("(" + arguments.callee.name + ") " +Total);
    for (var attr in nat_avgs){
      var attr_val = main_object[chapter][attr][0];
      if (attr_val ==""){continue;};
      nat_avgs[attr] += attr_val;
      if (maxs.indexOf(Total) > -1){
        if (top_chapter_names.indexOf(chapter) < 0){
          top_chapter_names.push(chapter);
        }
        top_avgs[attr] += attr_val;
      }
    }
  }
  for (var attr in nat_avgs){
    nat_avgs[attr] = nat_avgs[attr]/(parseInt(chapter_count)+1);
    top_avgs[attr] = top_avgs[attr]/maxs_length;
  }
  return {top_chapter_names: top_chapter_names,
          top_avgs:top_avgs,
          nat_avgs: nat_avgs,
         }
}

function sync_officers(oer){
//  var oer = {"Scribe":["Chi", "Scribe", "AlexanderNEW Gerwe", "(520) 405-3145",
//                       "agerwe75@gmail.com", "01/01/2017", "01/30/2017",
//                       "Fri Feb 03 12:25:30 GMT-07:00 2017"],
//             "Vice Regent":["Chi", "Vice Regent", "Alec Sonderman", "(520) 977-2560",
//                            "salmanderalec@gmail.com", "01/01/2017", "01/30/2017",
//                            "Fri Feb 03 12:25:30 GMT-07:00 2017"],
//             "Regent":["Chi", "Regent", "Adam Schilperoort", "(928) 420-4052",
//                       "adamschilperoort@gmail.com", "01/01/2017", "01/30/2017",
//                       "Fri Feb 03 12:25:30 GMT-07:00 2017"]};
  Logger.log("(" + arguments.callee.name + ") ");
  Logger.log(oer);
  var dash_id = SCRIPT_PROP.getProperty("dash");
//  var dash_id = "10ebwK7tTKgveVCEOpRle2S17d4UjwmsoXXCPFvC9A-A";
  var dash_file = SpreadsheetApp.openById(dash_id);
  var chapter = SCRIPT_PROP.getProperty("chapter");
  var officer_sheet = dash_file.getSheetByName("OFFICERS");
  var officer_values = officer_sheet.getDataRange().getValues();
  var officer_cols = officer_sheet.getLastColumn();
  var officer_rows = officer_sheet.getLastRow()+1;
  var officer_chapter = filter_chapter(officer_values, chapter, "OFFICERS");
  var officer_update = {};
  for (var officer in oer){
    var prev_name = '';
    var prev_phone = '';
    var prev_email = '';
    if (officer+officer in officer_chapter){
      var officer_row = officer_chapter[officer+officer][1];
      prev_name = officer_chapter[officer+officer][0][2];
      prev_phone = officer_chapter[officer+officer][0][3];
      prev_email = officer_chapter[officer+officer][0][4];
    } else {
      var officer_row = officer_rows++;
    }
    var officer_range = officer_sheet.getRange(officer_row, 1, 1, officer_cols);
    var new_values = oer[officer];
    new_values.push(prev_name, prev_phone, prev_email);
    officer_range.setValues([new_values]);
  }
}

function filter_chapter(array, chapter, type){
  var filtered = {};
  var first_row = array[0];
  filtered.first_row = first_row;
  var col1 = first_row.indexOf("Chapter");
  var col2 = first_row.indexOf("Chapter");
  switch (type){
    case "EVENTS":
      col1 = first_row.indexOf("Event Name");
      col2 = first_row.indexOf("Date");
      break;
    case "SUBMISSIONS":
      col1 = first_row.indexOf("Type");
      col2 = first_row.indexOf("Date");
      break;
    case "OFFICERS":
      col1 = first_row.indexOf("Office");
      col2 = first_row.indexOf("Office");
      break;
  }
  for (var row_ind in array){
    var row = array[row_ind];
    if (row.indexOf(chapter) > -1){
      var row_name = row[col1] + row[col2];
      filtered[row_name] = [row, parseInt(row_ind)+1];
    }
  }
  return filtered;
}