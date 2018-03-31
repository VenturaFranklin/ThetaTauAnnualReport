function update_scores_event(sheet_name, object){
//   var object = 4;
//   var sheet_name = "Events";
//  var object = range_object("Attendance", 3);
  Logger.log("(" + arguments.callee.name + ")");
  var att_obj = true;
  if (typeof(object)==typeof(2)){
    var user_row = object;
    var myObject = range_object(sheet_name, user_row);
//    var att_info = att_event_exists("Attendance", myObject)
    // This might mean that the attendance event has been deleted
//    if (typeof att_info.event_row == 'undefined'){
//      att_obj = false};
  } else {
    var event_info = att_event_exists(sheet_name, object);
    if (typeof event_info.event_row == 'undefined'){
      return;};
    var user_row = event_info.event_row;
    var myObject = range_object(sheet_name, user_row);
  }
  if (myObject.Type[0] == "" || myObject.Date[0] == "" ||
      myObject["Event Name"][0] == ""){
    return;
  }
  var date_range = myObject.sheet.getRange(myObject.object_row, myObject.Date[1]);
  if (!check_date_year_semester(myObject.Date[0])){
    var year_semesters = get_year_semesters();
    year_semesters = Object.keys(year_semesters);
    date_range.setBackground('red')
    .setNote("Date should be within year/semesters of Annual report.\n" + year_semesters.join(", "));
    return;
  } else {
    date_range.setBackground("white")
      .clearNote();
  }
  if (typeof myObject["# Members"][0] != typeof 2 || !att_obj){
//    attendance_add_event(myObject["Event Name"][0], myObject.Date[0]);
    event_add_calendar(myObject["Event Name"][0], myObject.Date[0],
                       myObject["Type"][0], myObject["Description"][0]);
  }
  if (!event_fields_set(myObject)){
    return;
  }
  var score_data = get_score_event(myObject);
  var other_type_rows = update_score(user_row, sheet_name, score_data, myObject);
  Logger.log("(" + arguments.callee.name + ") " +"OTHER ROWS" + other_type_rows);
//  if (refresh != true){
    for (i in other_type_rows){
      var other_row = other_type_rows[i][1];
      var sheet = other_type_rows[i][0];
      if (parseInt(other_row)!=parseInt(user_row)){
        var myObject = range_object(sheet, other_row);
        var score_data = get_score_event(myObject);
        update_score(other_row, sheet, score_data, myObject);
      }
    }
//  }
}

function get_semester(event_date){
  var month = event_date.getMonth();
  var semester = "FALL";
  if (month<5){
    var semester = "SPRING";
  }
  return semester
}

//function update_service_hours(){
//  var ss = get_active_spreadsheet();
//  var sheet = ss.getSheetByName("Membership");
//  var EventObject = main_range_object("Events");
//  var MemberObject = main_range_object("Membership");
//  var AttendanceObject = main_range_object("Attendance");
//  var score_obj = {};
//  for (var i = 0; i < EventObject.object_count; i++){
//    var event_name = EventObject.object_header[i];
//    var event_type = EventObject[event_name]["Type"][0];
//    if (event_type == "Service Hours"){
//      var event_hours = EventObject[event_name]["Event Hours"][0];
//      var event_date = EventObject[event_name]["Date"][0];
//      var semester = get_semester(event_date)
//      var att_obj = AttendanceObject[event_name];
//      if (typeof att_obj == 'undefined') {
//      // Event may no longer exists on Attendance sheet
//        Logger.log("(" + arguments.callee.name + ") " + "Missing Event: " + event_name);
//        ss.toast('Missing Event!' + event_name + 'Is missing from the attendance sheet',
//                 'ERROR', 5);
//        continue;
//      }
//      for (var j = 2; j < att_obj.object_count; j++){
//        var member_name_raw = AttendanceObject.header_values[j];
//        var member_name_short = att_name(member_name_raw);
//        var member_object = find_member_shortname(MemberObject, member_name_short);
//        if (typeof member_object == 'undefined') {
//          Logger.log("(" + arguments.callee.name + ") " + "Missing Member: " + member_name_short);
//          ss.toast('Missing Member!' + member_name_short + 'Is missing from the membership sheet',
//                   'ERROR', 5);
//          continue;
//        }
//        var member_name = member_object["Member Name"][0];
//        var att = att_obj[member_name_raw][0];
//        if (att == "P"){
//          score_obj[member_name] = score_obj[member_name] ? score_obj[member_name]:{};
//          score_obj[member_name][semester] = score_obj[member_name][semester] ?
//            score_obj[member_name][semester]+event_hours:event_hours;
//        }
////        Logger.log("(" + arguments.callee.name + ") " +score_obj);
//      }
//    }
//  }
//  for (var member_name in score_obj){
////    Logger.log("(" + arguments.callee.name + ") " +member_name);
//    var member_obj = MemberObject[member_name];
//    var member_row = member_obj.object_row;
//    var fall_col = member_obj["Service Hours Fall"][1];
//    var spring_col = member_obj["Service Hours Spring"][1];
//    var member_fall_range = sheet.getRange(member_row, fall_col);
//    var member_spring_range = sheet.getRange(member_row, spring_col);
//    var fall_score = score_obj[member_name]["FALL"] ? score_obj[member_name]["FALL"]:0;
//    member_fall_range.setValue(fall_score);
//    var spring_score = score_obj[member_name]["SPRING"] ? score_obj[member_name]["SPRING"]:0;
//    member_spring_range.setValue(spring_score);
//    Logger.log("(" + arguments.callee.name + ") " +"FALL: "+fall_score+" SPRING: "+spring_score+" ROW: "+member_row);
//  }
//  update_scores_org_gpa_serv();
//}

function update_score_att(){
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Scoring");
  var EventObject = main_range_object("Events");
  var ScoringObject = main_range_object("Scoring");
  var total_members = get_membership_ranges();
  var date_types = [];
  var counts = [];
  var year_semesters = get_year_semesters();
  for (var i = 0; i < EventObject.object_count; i++){
    var event_name = EventObject.object_header[i];
    var event_type = EventObject[event_name]["Type"][0];
    if (event_type == "Meetings"){
      var object_date = EventObject[event_name]["Date"][0];
      var meeting_att = EventObject[event_name]["# Members"][0];
      var semester = get_semester(object_date);
      var year = object_date.getFullYear();
      var year_semester = year + " " + semester
      if (!(year_semester in year_semesters)){
        continue;
      }
      var actives = total_members[year_semester]["Active Members"].value[0];
      actives = typeof actives === 'string' ? 1000:actives;
      actives = actives==0 ? 1000:actives;
      meeting_att = parseFloat(meeting_att / actives);
      date_types[year_semester] = date_types[year_semester] ? 
        date_types[year_semester] + meeting_att:meeting_att;
      counts[year_semester] = counts[year_semester] ? 
        counts[year_semester] + 1:1;
    }
  }
  var score_method_raw = ScoringObject["Meetings"]["Special"][0];
  var score_max = ScoringObject["Meetings"]["Max/ Semester"][0];
  var score_row = ScoringObject["Meetings"].object_row;
  var total_col = ScoringObject["Meetings"]["CHAPTER TOTAL"][1];
  var score_range_tot = sheet.getRange(score_row, total_col);
  var score_tot = 0;
  for (year_semester in date_types){
    var avg = date_types[year_semester]/counts[year_semester];
    Logger.log("(" + arguments.callee.name + ") " + year_semester + ": " + avg);
    var score_method = score_method_raw.replace("MEETINGS", avg);
    var score_range = sheet.getRange(score_row, ScoringObject["Meetings"][year_semester][1]);
    var score = eval_score(score_method, score_max);
    score = score >= 0 ? score:0;
    score_range.setValue(score);
    score_tot += score;
  }
  score_range_tot.setValue(score_tot);
  update_dash_score("Operate", total_col);
}

function update_score_member_pledge(){
  var ss = get_active_spreadsheet();
  var ScoringObject = main_range_object("Scoring");
  var sheet = ScoringObject.sheet;
  var score_method_pledge_raw = ScoringObject["Pledge Ratio"]["Special"][0];
  var score_pledge_max = ScoringObject["Pledge Ratio"]["Max/ Semester"][0];
  var score_method_raw = ScoringObject["Membership"]["Special"][0];
  var score_max = ScoringObject["Membership"]["Max/ Semester"][0];
  var score_row = ScoringObject["Membership"].object_row;
  var score_pledge_row = ScoringObject["Pledge Ratio"].object_row;
  var member_ranges = get_membership_ranges();
  var score_pledge_tot = 0;
  var score_tot = 0;
  for (var member_range_year in member_ranges){
    var score_method_pledge = score_method_pledge_raw;
    var score_method_member = score_method_raw;
    for (var member_range_type in member_ranges[member_range_year]){
      var value = member_ranges[member_range_year][member_range_type].value[0];
      value = typeof(value) == typeof(0) ? value:0;
      switch (member_range_type){
        case "Initiated Pledges":
          score_method_pledge = score_method_pledge.replace("INIT", value);
          score_method_member = score_method_member.replace("IN", value);
          break
        case "Total Pledges":
          score_method_pledge = score_method_pledge.replace("PLEDGE", value);
          break
        case "Graduated Members":
          score_method_member = score_method_member.replace("OUT", value);
          break
        case "Active Members":
          score_method_member = score_method_member.replace("MEMBERS", value);
          break
      }
    }
    var score_pledge = eval_score(score_method_pledge, score_pledge_max);
    score_pledge = score_pledge >= 0 ? score_pledge:0;
    score_pledge = !(isNaN(score_pledge)) ? score_pledge:0;
    var score = eval_score(score_method_member, score_max);
    score = score >= 0 ? score:0;
    score = !(isNaN(score)) >= 0 ? score:0;
    var score_range = sheet.getRange(score_row,
      ScoringObject["Membership"][member_range_year][1]);
    score_range.setValue(score);
    score_tot += score;
    var score_pledge_range = sheet.getRange(score_pledge_row,
      ScoringObject["Pledge Ratio"][member_range_year][1]);
    score_pledge_range.setValue(score_pledge);
    score_pledge_tot += score_pledge;
    }
  var total_col = ScoringObject["Membership"]["CHAPTER TOTAL"][1];
  var score_tot_range = sheet.getRange(score_row, total_col);
  var score_pledge_tot_range = sheet.getRange(score_pledge_row, total_col);
  score_tot_range.setValue(score_tot);
  score_pledge_tot_range.setValue(score_pledge_tot);
  update_dash_score("Operate", total_col);
  update_dash_score("Brotherhood", total_col);
}

function update_scores_org_gpa_serv(){
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Scoring");
  var score_data = get_scores_org_gpa_serv();
  var ScoringObject = main_range_object("Scoring");
  var total_col = ScoringObject["Societies"]["CHAPTER TOTAL"][1];
  var societies_range = sheet.getRange(ScoringObject["Societies"].object_row, total_col);
  var societies_method = ScoringObject["Societies"]["Special"][0];
  societies_method = societies_method.replace("ORG", score_data.percent_org);
  societies_method = societies_method.replace("OFFICER", score_data.officer_count);
  var societies_max = ScoringObject["Societies"]["Max/ Semester"][0];
  var socieities_score = eval_score(societies_method, societies_max);
  var year_semesters = get_year_semesters();
  var gpa_method_raw = ScoringObject["GPA"]["Special"][0];
  var gpa_max = ScoringObject["GPA"]["Max/ Semester"][0];
  var service_method_raw = ScoringObject["Service Hours"]["Special"][0];
  var service_max = ScoringObject["Service Hours"]["Max/ Semester"][0];
  var gpa_tot = 0;
  var service_tot = 0;
  for (var year_semester in year_semesters){
    var gpa_type = year_semester + " GPA";
    var service_type = year_semester + " Service";
    var gpa_col = ScoringObject["GPA"][year_semester][1];
    var service_col = ScoringObject["Societies"][year_semester][1];
    var gpa_range = sheet.getRange(ScoringObject["GPA"].object_row, gpa_col);
    var gpa_method = gpa_method_raw.replace("GPA", score_data[gpa_type]);
    var gpa_score = eval_score(gpa_method, gpa_max);
    var service_range = sheet.getRange(ScoringObject["Service Hours"].object_row, service_col);
    var service_method = service_method_raw.replace("HOURS", score_data[service_type]);
    var service_score = eval_score(service_method, service_max);
    gpa_range.setValue(gpa_score);
    gpa_tot += gpa_score;
    service_range.setValue(service_score);
    service_tot += service_score;
    Logger.log("(" + arguments.callee.name + ") " +"GPA: " + gpa_method + ", SCORE: " + gpa_score);
    Logger.log("(" + arguments.callee.name + ") " +"SERV: " + service_method + ", SCORE: " + service_score);
  }
  var gpa_tot_range = sheet.getRange(ScoringObject["GPA"].object_row, total_col);
  gpa_tot_range.setValue(gpa_tot);
  var service_tot_range = sheet.getRange(ScoringObject["Service Hours"].object_row, total_col);
  service_tot_range.setValue(service_tot);
  Logger.log("(" + arguments.callee.name + ") " +"SOC: " + societies_method + ", SCORE: " + socieities_score);
  societies_range.setValue(socieities_score);
  update_dash_score("ProDev", total_col);
  update_dash_score("Service", total_col);
}

function eval_score(score_method, score_max){
  var score = eval(score_method);
  score = parseFloat(score.toFixed(1));
  score = score > parseFloat(score_max) ? score_max: score;
  return score;
}

function get_scores_org_gpa_serv(){
  var officer_counts = {};
  var org_counts = {};
  var gpa_counts = {};
  var service_counts = {};
  var officer_count = 0;
  var org_count = 0;
  var officers = ["Officer (Pro/Tech)", "Officer (Honor)", "Officer (Other)"];
  var orgs = ["Professional/ Technical Orgs", "Honor Orgs", "Other Orgs"];
  var year_semesters = get_membership_ranges();
  var MemberObject = main_range_object("Membership");
  var gpa = 0;
  for (var i = 0; i < MemberObject.object_count; i++){
    var member_name = MemberObject.object_header[i];
    var org_true = false;
    var officer_true = false;
//     var status = MemberObject[member_name]["Chapter Status"][0];
//     var start = MemberObject[member_name]["Status Start"][0];
//     if (typeof(start) != typeof(new Date())){
//       if (start != ""){
//         if (start.indexOf("undefined") >= 0){
//           continue;
//         }
//       }
//     }
//     switch (status){
//       case "Pledge":
//         continue;
//         break;
//       case "Shiny":
//         var month = start.getMonth() + 1;
//         if (month<=6){fall_mult = 0;
//         } else {spring_mult = 0;
//         }
//         break;
//       case "Away":
//       case "Alumn":
//         var month = start.getMonth() + 1;
//         if (month<=6){spring_mult = 0;
//         } else {fall_mult = 0;
//         }
//         break;
//     }
//     active_total_fa += fall_mult;
//     active_total_sp += spring_mult;
//     active_total += 1;
    for (var j = 0; j <= 2; j++){
      var org_type = orgs[j];
      var this_org = MemberObject[member_name][org_type][0].toString();
      org_counts[org_type] = org_counts[org_type] ? org_counts[org_type]:1;
      org_counts[org_type] = ((this_org.toUpperCase()!="NONE") && (this_org!="")) ?
        org_counts[org_type]+1:org_counts[org_type];
      org_true = (this_org.toUpperCase()!="NONE" && this_org!="") ? true:org_true;
      var officer = MemberObject[member_name][officers[j]][0];
      officer_counts[officers[j]] = officer_counts[officers[j]] ? officer_counts[officers[j]]:0;
      officer_counts[officers[j]] = officer.toUpperCase()=="YES" ?
        officer_counts[officers[j]]+1:officer_counts[officers[j]];
      officer_true = officer.toUpperCase()=="YES" ? true:officer_true;
    }
    officer_count = officer_true ? officer_count + 1:officer_count;
    org_count = org_true ? org_count + 1:org_count;
    for (var year_semester in year_semesters){
      var gpa_type = year_semester + " GPA";
      if (!(gpa_type in gpa_counts)){gpa_counts[gpa_type] = 0;};
      var service_type = year_semester + " Service";
      if (!(service_type in service_counts)){service_counts[service_type] = 0;};
      var gpa_raw = MemberObject[member_name][gpa_type][0];
      gpa_raw = gpa_raw == "" ? 0:gpa_raw;
      var gpa = parseFloat(gpa_raw);
      if (isNaN(gpa)){
        gpa = 0.0
      }
//       if (gpa_type.indexOf("FALL") > -1){
//         if(!fall_mult){continue;
//         }
//       } else {
//         if(!spring_mult){continue;
//         }
//       }
      gpa_counts[gpa_type] = gpa_counts[gpa_type]+gpa;
      var service_hours = MemberObject[member_name][service_type][0];
      service_counts[service_type] = (+service_hours) >= 8 ?
        service_counts[service_type] + 1:service_counts[service_type];
    }
  }
  var percents = {};
  for (var year_semester in year_semesters){
    var gpa_type = year_semester + " GPA";
    var service_type = year_semester + " Service";
    var active_total = year_semesters[year_semester]["Active Members"].value[0];
    active_total = typeof active_total === 'string' ? 1000:active_total;
    active_total = active_total == 0 ? 1000:active_total;
    var service_count = service_counts[service_type];
    var gpa_count = gpa_counts[gpa_type];
    var percent_service = service_count / active_total;
    var percent_gpa = gpa_count / active_total;
    percents[gpa_type] = percent_gpa;
    percents[service_type] = percent_service;
  }
  var percent_org = active_total==0 ? 0:org_count / active_total;
  percents["percent_org"] = percent_org;
  percents["officer_count"] = officer_count;
  return percents
}

function get_score_submit(myScore){
  var event_type = myScore["Type"][0]
  if (~event_type.indexOf("Pledge Program")){
    var info = event_type.split(" - ");
    event_type = info[0];
    var mod = info[1]=="modified";
    Logger.log("(" + arguments.callee.name + ") " + "mod: " + mod);
  }
  var score_data = get_score_method(event_type, mod);
  Logger.log("(" + arguments.callee.name + ") ");
  Logger.log(score_data);
  var score = eval(score_data.score_method);
  score = score.toFixed(1);
  score_data.score = score;
  Logger.log("(" + arguments.callee.name + ") " +arguments.callee.name + "SCORE RAW: " + score);
  return score_data
}

function update_scores_submit(user_row){
//  var user_row = 2;
  Logger.log("(" + arguments.callee.name + ") " +"ROW: " + user_row);
  var myObject = range_object("Submissions", parseInt(user_row));
  var score_data = get_score_submit(myObject);
  var other_type_rows = update_score(user_row, "Submissions", score_data, myObject);
  Logger.log("(" + arguments.callee.name + ") " +other_type_rows);
}

function update_score(row, sheetName, score_data, myObject){
//  var row = 4
//  var shetName = "Events";
  Logger.log("(" + arguments.callee.name + ") " +"SHEET: " + sheetName + " ROW: " + row)
  var ss = get_active_spreadsheet();
  if (typeof sheetName === "string"){
    var sheet = ss.getSheetByName(sheetName);
  } else {
    var sheet = sheetName;
    var sheetName = sheet.getName();
  }
  var score_ind = myObject["Score"][1];
  var object_date = myObject["Date"][0];
  var object_type = myObject["Type"][0];
  var score_range = sheet.getRange(row, score_ind);
  score_range.setValue(0); // To protect the current score from affecting max
  Logger.log("(" + arguments.callee.name + ") " +"Date: " + object_date + " Type:" + object_type)
  var total_scores = get_current_scores(sheetName);
  Logger.log("(" + arguments.callee.name + ") " +total_scores)
  var semester = get_semester(object_date);
  var year = object_date.getFullYear();
  var year_semester = year + " " + semester
  score_data.semester = year_semester;
  var type_score = total_scores[year_semester][object_type][0];
  var other_type_rows = total_scores[year_semester][object_type][1];
  Logger.log("(" + arguments.callee.name + ") " +"Type Score: " + type_score);
  score_range.setNote(score_data.score_method_note);
  var score = score_data.score;
  if (score === null){
    score_range.setBackground("black");
    return [];
  } else {
    score_range.setBackground("dark gray 1");
  }
  var total = parseFloat(type_score) + parseFloat(score);
  Logger.log("(" + arguments.callee.name + ") " +total)
  if (total > parseFloat(score_data.score_max_semester)){
    score = score_data.score_max_semester - type_score;
    score = score > 0 ? score:0;
  }
  Logger.log("(" + arguments.callee.name + ") " +"FINAL SCORE: " + score);
  score_data.final_score = score;
  score_data.type_score = type_score;
//  score_range.setValue(0); This does not work, is not fast enough
  update_main_score(score_data);
  score_range.setValue(score);
  return other_type_rows;
}

function update_main_score(score_data){
//  var score_data = {
//    "score_ids": {"chapter": 9.0, "2017 FALL": 7.0, "score_row": 6.0,
//    "2016 FALL": 5.0, "2017 SPRING": 6.0, "2018 SPRING": 8.0},
//    "score": 0.1, "final_score": 0.1, "score_method": "memberATT*5+memberADD*0",
//    "type_score": 0.0, "score_method_note": "5*(% Attendance)",
//    "semester": "2017 SPRING", "score_max_semester": 10.0, "score_type": "ProDev"
//  }
  Logger.log("(" + arguments.callee.name + ") ");
  Logger.log(score_data);
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Scoring");
  var score_row = score_data.score_ids.score_row
  var semester_range = sheet.getRange(score_row, score_data.score_ids[score_data.semester]);
  var semester_years = get_year_semesters();
  var total_score = 0;
  for (var semester_year in semester_years){
    if (semester_year == score_data.semester){continue;}; // This avoids adding new score until later
    var semester_val = sheet.getRange(score_row, score_data.score_ids[semester_year]).getValue();
    semester_val = typeof semester_val === 'string' ? 0:semester_val;
    total_score += parseInt(semester_val);
  }
  var total_range = sheet.getRange(score_row, score_data.score_ids.chapter);
  var total_sem_score = parseFloat(score_data.final_score) + score_data.type_score;
  semester_range.setValue(total_sem_score);
  total_score += total_sem_score; // This semester is added to total
  total_range.setValue(total_score);
  update_dash_score(score_data.score_type, score_data.score_ids.chapter);
}

function update_dash_score(score_type, score_column){
  Logger.log("(" + arguments.callee.name + ") " +score_type);
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Scoring");
  if (score_type != undefined){
    var type_inds = get_ind_list(score_type);
    var type_count = type_inds.length;
  } else {
    var type_count = sheet.getLastRow();
    var type_inds = [];
    for (var i = 0; i <= type_count; i++) {
        type_inds.push(i);
    }
  }
  var total = 0;
  for (var j = 0; j < type_count; j++){
    var row = type_inds[j];
    var row_total = sheet.getRange(row, score_column).getValue();
    total = +total + row_total;
  }
  Logger.log("(" + arguments.callee.name + ") " +type_inds);
  var sheet = ss.getSheetByName("Dashboard");
  var RangeName = "SCORE" + "_" + score_type.toUpperCase();
  var dash_score_range = ss.getRangeByName(RangeName);
  dash_score_range.setValue(total);
}

function get_current_scores_event(){
  var EventObject = main_range_object("Events");
  var date_types = new Array();
  for(var j in EventObject.object_header) {
    var event_name = EventObject.object_header[j];
    var event = EventObject[event_name];
    var date = event.Date[0];
    var row = event.object_row;
    try{
      if (typeof date == 'undefined' || date == ''){
        continue;
      } else {
        var month = date.getMonth();
      }
    } catch (e) {
      var message = Utilities.formatString('This error has automatically been sent to the developers. DATE ERROR; Date Obj: %s; i: %s; Date Values: %s; Date Ind: %s; Stack: "%s"; While processing: %s.',
                                           date||'', row||'', date_values||'', date_ind||'',
                                           e.stack||'', arguments.callee.name||'');
      Logger = startBetterLog();
      Logger.severe(message);
//      var ui = SpreadsheetApp.getUi();
//      var result = ui.alert(
//        'ERROR',
//        message,
//        ui.ButtonSet.OK);
      return date_types;
    }
    var type_name = event["Type"][0];
    var score = event["Score"][0];
    var sheet = event.sheet;
    var semester = get_semester(date);
    var year = date.getFullYear();
    var year_semester = year + " " + semester
    if (!(year_semester in date_types)){
      date_types[year_semester] = {};
    }
    var old_score = date_types[year_semester][type_name] ? 
        date_types[year_semester][type_name][0] : 0;
    var new_score = parseFloat(old_score) + parseFloat(score);
    var old_rows = date_types[year_semester][type_name] ? 
        date_types[year_semester][type_name][1] : [];
    old_rows.push([sheet, parseInt(row)]);
    date_types[year_semester][type_name] = [new_score, old_rows]
    }
  return date_types;
}

function get_current_scores(sheetName){
  if (sheetName.indexOf('Event') < 0){
    return get_current_scores_orig(sheetName);
  } 
  return get_current_scores_event();
}

function get_current_scores_orig(sheetName){
//  var sheetName = "Events";
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName(sheetName);
//   var EventObject = main_range_object("Events");
  var max_column = sheet.getLastColumn();
  var max_row = sheet.getLastRow();
  var full_data_range = sheet.getRange(1, 1, max_row, max_column);
  var full_data_values = full_data_range.getValues();
  var score_ind = get_ind_from_string("Score", full_data_values);
  var date_ind = get_ind_from_string("Date", full_data_values);
  var type_ind = get_ind_from_string("Type", full_data_values);
  var score_values = get_column_values(score_ind-1, full_data_values);
  var date_values = get_column_values(date_ind-1, full_data_values);
  var type_values = get_column_values(type_ind-1, full_data_values);
  var date_types = new Array();
  for(var i = 1; i< date_values.length; i++) {
                var date = date_values[i];
    try{
      if (typeof date == 'undefined' || date == ''){
        continue;
      } else {
        var month = date.getMonth();
      }
    } catch (e) {
      var message = Utilities.formatString('This error has automatically been sent to the developers. DATE ERROR; Date Obj: %s; i: %s; Date Values: %s; Date Ind: %s; Stack: "%s"; While processing: %s.',
                                           date||'', i||'', date_values||'', date_ind||'',
                                           e.stack||'', arguments.callee.name||'');
      Logger = startBetterLog();
      Logger.severe(message);
//      var ui = SpreadsheetApp.getUi();
//      var result = ui.alert(
//        'ERROR',
//        message,
//        ui.ButtonSet.OK);
      return date_types;
    }
    var type_name = type_values[i];
    var score = score_values[i];
    var semester = get_semester(date);
    var year = date.getFullYear();
    var year_semester = year + " " + semester
    if (!(year_semester in date_types)){
      date_types[year_semester] = {};
    }
    var old_score = date_types[year_semester][type_name] ? 
            date_types[year_semester][type_name][0] : 0;
    var new_score = parseFloat(old_score) + parseFloat(score);
    var old_rows = date_types[year_semester][type_name] ? 
      date_types[year_semester][type_name][1] : [];
    old_rows.push(parseInt(i) + 1);
    date_types[year_semester][type_name] = [new_score, old_rows]
          }
  return date_types;
}

function get_score_event(myEvent){
  var event_type = myEvent["Type"][0];
  var score_data = get_score_method(event_type);
  Logger.log("(" + arguments.callee.name + ") ");
  Logger.log(score_data);
  var score_method_edit = edit_score_method_event(myEvent, score_data.score_method);
  var score = null
  if (score_method_edit !== null){
    score = eval(score_method_edit);
    score = score.toFixed(1);
  }
  score_data.score = score;
  Logger.log("(" + arguments.callee.name + ") " +"SCORE RAW: " + score);
  return score_data
}


function edit_score_method_event(myEvent, score_method, totals){
  var attend = myEvent["# Members"][0];
  var attend = (attend != "") ? attend:0;
  if (~score_method.indexOf("memberATT")){
    if (!totals){
      var total_members = get_membership_ranges();
    } else {
      var total_members = totals;
    }
      var event_date = myEvent["Date"][0];
      var semester = get_semester(event_date);
      var year = event_date.getFullYear();
      var year_semester = year + " " + semester;
      var actives = total_members[year_semester]["Active Members"].value[0];
      actives = typeof actives === 'string' ? 1000:actives;
      actives = actives==0 ? 1000:actives;
      var percent_attend = attend / actives;
      percent_attend = percent_attend>1 ? 0:percent_attend;
      score_method = score_method.replace("memberATT", percent_attend);
          }
  if (~score_method.indexOf("memberADD")){
      score_method = score_method.replace("memberADD", attend);
          }
  if (~score_method.indexOf("MILES")){
    var miles = myEvent["MILES"][0];
    miles = (miles != "") ? miles:0;
    score_method = score_method.replace("MILES", miles);
          }
  if (~score_method.indexOf("HOST")){
    var host = myEvent["HOST"][0];
    host = (host.toUpperCase() == "YES") ? 1:0;
    score_method = score_method.replace("HOST", host);
    score_method = score_method.replace("HOST", host);
    score_method = score_method.replace("HOST", host);
          }
  if (~score_method.indexOf("NON-MEMBER")){
      var non_members = myEvent["# Non- Members"][0];
      var non_members = (non_members != "") ? non_members:0;
      score_method = score_method.replace("NON-MEMBER", non_members);
          }
  if (~score_method.indexOf("ALUMNI")){
      var alumni_members = myEvent["# Alumni"][0];
      var alumni_members = (alumni_members != "") ? alumni_members:0;
      score_method = score_method.replace("ALUMNI", alumni_members);
          }
  if (~score_method.indexOf("STEM")){
      var stem = myEvent["STEM?"][0];
      var stem = (stem.toUpperCase() == "YES") ? 1:0;
      score_method = score_method.replace("STEM", stem);
          }
  if (~score_method.indexOf("MEETINGS")){
    update_score_att();
    return null;
          }
  if (~score_method.indexOf("HOURS")){
//    update_service_hours();
    return null;
          }
  if (~score_method.indexOf("MISC")){
    return null;
  }
  Logger.log("(" + arguments.callee.name + ") " +"Score Method Raw: " + score_method)
  return score_method
}

function get_score_method(event_type, mod, ScoringObject){
//  var event_type = "Alumni-Active";
  if (!ScoringObject){
    var ScoringObject = main_range_object("Scoring");
  }
  var score_object = ScoringObject[event_type];
  var score_type = score_object["Score Type"][0];
  var score_method_note = score_object["How points are calculated"][0];
  var att =  score_object["Attendance Multiplier"][0];
  var att = (att != "") ? att:0;
  var add = score_object["Member Add"][0];
  var add = (add != "") ? add:0;
  var base =  score_object["Base Points"][0];
  var special = score_object["Special"][0];
  if (score_type == "Events"){
   var score_method = "memberATT*" + att + "+memberADD*" + add;
  }
  if (score_type == "Submit"){
   var score_method = base;
  }
  if (event_type == "Pledge Program"){
    special = special.replace("UNMODIFIED", +!mod);
    special = special.replace("MODIFIED", +mod);
    var score_method = special;
  }

  if (score_type == "Events/Special" || score_type == "Special"){
   var score_method =  special;
  }
  var score_ids = {
     score_row: score_object.object_row,
     chapter: score_object["CHAPTER TOTAL"][1]
  }
  var year_semesters = get_year_semesters();
  for (var year_semester in year_semesters){
    score_ids[year_semester] = score_object[year_semester][1]
  }
  return {score_method: score_method,
          score_method_note: score_method_note,
          score_max_semester: score_object["Max/ Semester"][0],
          score_ids: score_ids,
          score_type: score_object["Type"][0]
         }
}

function refresh_main_scores(type_semester, ss, ScoringObject){
  // type_semester = semester.event_type
  if (!ss){
    var ss = get_active_spreadsheet();
    var ScoringObject = main_range_object("Scoring", undefined, ss);
  }
  var score_sheet = ScoringObject.sheet;
  var year_semesters = get_year_semesters();
  var cols = [];
  for (var year_semester in year_semesters){
    cols.push(ScoringObject.header_values.indexOf(year_semester));
    }
  var start_col = Math.min.apply(null, cols);
  var year_semester_cols = {};
  for (var year_semester in year_semesters){
    year_semester_cols[year_semester] = ScoringObject.header_values.indexOf(year_semester) - start_col;
  }
  cols.push(99);
  var all_scores = [];
  for (var i in ScoringObject.object_header){
    // This is so we do not have empty array/rows
    //var this_type = ScoringObject.object_header[i];
    var score_row = [];
    for (var year_semester in year_semesters){
      //score_row.push(ScoringObject[this_type][year_semester][0]);
      score_row.push(0)
    }
    //score_row.push(ScoringObject[this_type]["CHAPTER TOTAL"][0]);
    score_row.push(0)
    all_scores.push(score_row);
  }
  for (var year_semester in type_semester){
    for (var event_type in type_semester[year_semester]){
      var semester_col = year_semester_cols[year_semester];
      var event_score = type_semester[year_semester][event_type];
      var score_row = ScoringObject[event_type].object_row - 2 // 2 of header and row in sheet starts at 1 not 0;
      all_scores[score_row][semester_col]=event_score;
    }
  }
  for (var ind in all_scores){
    // This is the total score update
    all_scores[ind][cols.length-1] = all_scores[ind].reduce(function(pv, cv) { return pv + cv; }, 0);
  }
  var semester_range = score_sheet.getRange(2, start_col+1, ScoringObject.object_count, 5);
  semester_range.setValues(all_scores);
}

function refresh_scores_silent() {
  SILENT = true;
  refresh_scores();
  SILENT = false;
}

function refresh_scores() {
  try{
    progress_update("REFRESH EVENTS");
    update();
    var ss = get_active_spreadsheet();
//    var attendance_object = main_range_object("Attendance", undefined, ss);
    var EventObject = main_range_object("Events", undefined, ss);
    var SubmitObject = main_range_object("Submissions", undefined, ss);
//    events_to_att(ss, attendance_object, EventObject);
//    refresh_attendance(ss, attendance_object, EventObject);
//     EventObject = main_range_object("Events", undefined, ss);
    var ScoringObject = main_range_object("Scoring", undefined, ss);
    var total_members = get_membership_ranges();
    var all_scores = {};
    var all_backgrounds = {};
    var all_notes = {};
    var type_semester = {};
    var sheet_names = {};
    var exclude = ["Meetings", "Service Hours", "Misc"];
    var year_semesters = get_year_semesters();
    year_semesters = Object.keys(year_semesters);
    for (var j in EventObject.object_header){
      var event_name = EventObject.object_header[j];
      var event = EventObject[event_name];
      var event_type = event["Type"][0];
      var event_date = event["Date"][0];
      var event_sheet_name = event.sheet_name;
      sheet_names[event_sheet_name] = event.sheet;
      all_notes[event_sheet_name] = all_notes[event_sheet_name] ? all_notes[event_sheet_name]:[];
      all_backgrounds[event_sheet_name] = all_backgrounds[event_sheet_name] ? all_backgrounds[event_sheet_name]:[];
      all_scores[event_sheet_name] = all_scores[event_sheet_name] ? all_scores[event_sheet_name]:[];
      var not_set = false;
      if (event_type == ''){
        event_type = 'Misc';
        not_set = true;
      }
      if (!check_date_year_semester(event_date)){
        all_scores[event_sheet_name].push([0]);
        all_backgrounds[event_sheet_name].push(['red']);
        all_notes[event_sheet_name].push(["Date should be within year/semesters of Annual report.\n" + year_semesters.join(", ")]);
        continue;
      }
      var year = event_date.getFullYear();
      var semester = get_semester(event_date);
      var year_semester = year + " " + semester;
      if (!(year_semester in type_semester)){
        type_semester[year_semester] = {};
      }
      var score_data = get_score_method(event_type, undefined, ScoringObject);
      var score_method_edit = null;
      if (exclude.indexOf(event_type) < 0){
        score_method_edit = edit_score_method_event(event, score_data.score_method, total_members);
      }
      var background = "black";
      var score = 0
      if (score_method_edit !== null){
        score = eval(score_method_edit);
        score = score.toFixed(1);
        background = "dark gray 1";
      }
      var note = score_data.score_method_note;
      if (not_set){
        score = 0;
        background = "red";
        note = "The event type must be set.";
      }
      // Need to find the score, date, type to determine semester scoring
      var type_semester_score = type_semester[year_semester][event_type] ? type_semester[year_semester][event_type]:0;
      var combined_score = parseFloat(type_semester_score) + parseFloat(score);
      if (combined_score > parseFloat(score_data.score_max_semester)){
        combined_score = score_data.score_max_semester - type_semester_score;
        combined_score = combined_score > 0 ? combined_score:0;
        }
      type_semester[year_semester][event_type] = combined_score;
      all_scores[event_sheet_name].push([combined_score]);
      all_backgrounds[event_sheet_name].push([background]);
      all_notes[event_sheet_name].push([note]);
    }
  var score_col = EventObject.header_values.indexOf("Score") + 1;
  for (event_sheet_name in sheet_names){
    var event_sheet = sheet_names[event_sheet_name];
    var rows = all_scores[event_sheet_name].length;
    var score_range = event_sheet.getRange(2, score_col, rows, 1);
    score_range.setValues(all_scores[event_sheet_name]);
    score_range.setBackgrounds(all_backgrounds[event_sheet_name]);
    score_range.setNotes(all_notes[event_sheet_name]);
  }
    for (var j in SubmitObject.object_header){
      var submit_name = SubmitObject.object_header[j];
      var submit = SubmitObject[submit_name];
      var submit_type = submit["Type"][0];
      var ignore = ['OER', 'MSCR', 'INIT', 'DEPL', 'COOP'];
      if (ignore.indexOf(submit_type) >= 0){
        continue;
      }
      if (submit_type.indexOf("Pledge Program") >= 0){
        var mod = submit_type.split(" - ")[1];
        var mod = mod=="modified";
        submit_type = "Pledge Program";
      }
      var submit_date = submit["Date"][0];
      var submit_score = submit["Score"][0];
      var year = submit_date.getFullYear();
      var semester = get_semester(submit_date);
      var year_semester = year + " " + semester;
      if (!(year_semester in type_semester)){
        type_semester[year_semester] = {};
      }
      var score_data = get_score_method(submit_type, undefined, ScoringObject);
      var type_semester_score = type_semester[year_semester][submit_type] ? type_semester[year_semester][submit_type]:0;
      var combined_score = parseFloat(type_semester_score) + parseFloat(submit_score);
      if (combined_score > parseFloat(score_data.score_max_semester)){
        combined_score = score_data.score_max_semester - type_semester_score;
        combined_score = combined_score > 0 ? combined_score:0;
        }
      type_semester[year_semester][submit_type] = combined_score;
    }
    var ScoringObject = main_range_object("Scoring", undefined, ss);
    refresh_main_scores(type_semester, ss, ScoringObject);
    update_score_att();
    update_scores_org_gpa_serv();
    update_score_member_pledge();
    var total_col = ScoringObject["Meetings"]["CHAPTER TOTAL"][1];
    update_dash_score("ProDev", total_col);
    update_dash_score("Service", total_col);
    update_dash_score("Operate", total_col);
    update_dash_score("Brotherhood", total_col);
    progress_update("REFRESH EVENTS FINISHED");
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
