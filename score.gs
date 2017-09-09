function update_scores_event(object){
//  var object = 6;
//  var object = range_object("Attendance", 3);
  Logger.log("(" + arguments.callee.name + ")");
  var att_obj = true;
  if (typeof(object)==typeof(2)){
    var user_row = object;
    var myObject = range_object("Events", user_row);
//    var att_info = att_event_exists("Attendance", myObject)
    // This might mean that the attendance event has been deleted
//    if (typeof att_info.event_row == 'undefined'){
//      att_obj = false};
//  } else {
    var event_info = att_event_exists("Events", object);
    if (typeof event_info.event_row == 'undefined'){
      return;};
    var user_row = event_info.event_row;
    var myObject = range_object("Events", user_row);
  }
  if (myObject.Type[0] == "" || myObject.Date[0] == "" ||
      myObject["Event Name"][0] == ""){
    return;
  } else if (typeof myObject["# Members"][0] != typeof 2 || !att_obj){
//    attendance_add_event(myObject["Event Name"][0], myObject.Date[0]);
    event_add_calendar(myObject["Event Name"][0], myObject.Date[0],
                       myObject["Type"][0], myObject["Description"][0]);
    myObject = range_object("Events", user_row);
  }
  if (!event_fields_set(myObject)){
    return;
  }
  var score_data = get_score_event(myObject);
  var other_type_rows = update_score(user_row, "Events", score_data, myObject);
  Logger.log("(" + arguments.callee.name + ") " +"OTHER ROWS" + other_type_rows);
//  if (refresh != true){
    for (i in other_type_rows){
      if (parseInt(other_type_rows[i])!=parseInt(user_row)){
        var myObject = range_object("Events", other_type_rows[i]);
        var score_data = get_score_event(myObject);
        update_score(other_type_rows[i], "Events", score_data, myObject);
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
  var total_members = get_total_members(true);
  var date_types = [];
  var counts = [];
  for (var i = 0; i < EventObject.object_count; i++){
    var event_name = EventObject.object_header[i];
    var event_type = EventObject[event_name]["Type"][0];
    if (event_type == "Meetings"){
      var object_date = EventObject[event_name]["Date"][0];
      var meeting_att = EventObject[event_name]["# Members"][0];
      var semester = get_semester(object_date);
      meeting_att = parseFloat(meeting_att / total_members[semester]["Student"]);
      date_types[semester] = date_types[semester] ? 
        date_types[semester] + meeting_att:meeting_att;
      counts[semester] = counts[semester] ? 
        counts[semester] + 1:1;
    }
  }
  var fall_avg = date_types["FALL"]/counts["FALL"];
  var spring_avg = date_types["SPRING"]/counts["SPRING"];
  Logger.log("(" + arguments.callee.name + ") " +"FALL ATT: " + fall_avg + " SPRING ATT: " + spring_avg);
  var score_method_raw = ScoringObject["Meetings"]["Special"][0];
  var score_max = ScoringObject["Meetings"]["Max/ Semester"][0];
  var score_method_fa = score_method_raw.replace("MEETINGS", fall_avg);
  var score_row = ScoringObject["Meetings"].object_row;
  var total_col = ScoringObject["Meetings"]["CHAPTER TOTAL"][1];
  var score_range_fa = sheet.getRange(score_row, ScoringObject["Meetings"]["FALL SCORE"][1]);
  var score_range_sp = sheet.getRange(score_row, ScoringObject["Meetings"]["SPRING SCORE"][1]);
  var score_range_tot = sheet.getRange(score_row, total_col);
  var score_method_sp = score_method_raw.replace("MEETINGS", spring_avg);
  var score_fa = eval_score(score_method_fa, score_max);
  var score_sp = eval_score(score_method_sp, score_max);
  score_sp = score_sp >= 0 ? score_sp:0;
  score_fa = score_fa >= 0 ? score_fa:0;
  score_range_fa.setValue(score_fa);
  score_range_sp.setValue(score_sp);
  score_range_tot.setValue(+score_fa + score_sp);
  update_dash_score("Operate", total_col);
}

function update_score_member_pledge(){
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Scoring");
  var member_value_obj = get_membership_ranges();
  var init_sp_value = member_value_obj.init_sp_range.getValue();
  init_sp_value = typeof(init_sp_value) == typeof(0) ? init_sp_value:0;
  var init_fa_value = member_value_obj.init_fa_range.getValue();
  init_fa_value = typeof(init_fa_value) == typeof(0) ? init_fa_value:0;
  var pledge_sp_value = member_value_obj.pledge_sp_range.getValue();
  pledge_sp_value = typeof(pledge_sp_value) == typeof(0) ? pledge_sp_value:0;
  var pledge_fa_value = member_value_obj.pledge_fa_range.getValue();
  pledge_fa_value = typeof(pledge_fa_value) == typeof(0) ? pledge_fa_value:0;
  var grad_sp_value = member_value_obj.grad_sp_range.getValue();
  grad_sp_value = typeof(grad_sp_value) == typeof(0) ? grad_sp_value:0;
  var grad_fa_value = member_value_obj.grad_fa_range.getValue();
  grad_fa_value = typeof(grad_fa_value) == typeof(0) ? grad_fa_value:0;
  
  var totals = get_total_members(true);
  var act_sp_value = totals["SPRING"]["Student"];
  act_sp_value = typeof(act_sp_value) == typeof(0) ? act_sp_value:0;
  var act_fa_value = totals["FALL"]["Student"];
  act_fa_value = typeof(act_fa_value) == typeof(0) ? act_fa_value:0;
  
  var all_vals = [init_sp_value, init_fa_value, pledge_sp_value, pledge_fa_value,
                  grad_sp_value, grad_fa_value, act_sp_value, act_fa_value];
  Logger.log("(" + arguments.callee.name + ") " + all_vals)
  if (all_vals.indexOf("") > -1){return;};
  var ScoringObject = main_range_object("Scoring");
  var score_method_pledge_raw = ScoringObject["Pledge Ratio"]["Special"][0];
  var score_pledge_max = ScoringObject["Pledge Ratio"]["Max/ Semester"][0];
  var score_method_pledge_fa = score_method_pledge_raw.replace("INIT", init_fa_value);
  score_method_pledge_fa = score_method_pledge_fa.replace("PLEDGE", pledge_fa_value);
  var score_pledge_fa = eval_score(score_method_pledge_fa, score_pledge_max);
  var score_method_pledge_sp = score_method_pledge_raw.replace("INIT", init_sp_value);
  score_method_pledge_sp = score_method_pledge_sp.replace("PLEDGE", pledge_sp_value);
  var score_pledge_sp = eval_score(score_method_pledge_sp, score_pledge_max);
  var score_method_raw = ScoringObject["Membership"]["Special"][0];
  var score_max = ScoringObject["Membership"]["Max/ Semester"][0];
  var score_method_fa = score_method_raw.replace("OUT", grad_fa_value);
  score_method_fa = score_method_fa.replace("IN", init_fa_value);
  score_method_fa = score_method_fa.replace("MEMBERS", act_fa_value);
  var score_fa = eval_score(score_method_fa, score_max);
  var score_method_sp = score_method_raw.replace("OUT", grad_sp_value);
  score_method_sp = score_method_sp.replace("IN", init_sp_value);
  score_method_sp = score_method_sp.replace("MEMBERS", act_sp_value);
  var score_sp = eval_score(score_method_sp, score_max);
  var score_row = ScoringObject["Membership"].object_row;
  var score_fa_range = sheet.getRange(score_row,
                                      ScoringObject["Membership"]["FALL SCORE"][1]);
  var score_sp_range = sheet.getRange(score_row,
                                      ScoringObject["Membership"]["SPRING SCORE"][1]);
  var total_col = ScoringObject["Membership"]["CHAPTER TOTAL"][1];
  var score_tot_range = sheet.getRange(score_row,total_col);
  var score_pledge_row = ScoringObject["Pledge Ratio"].object_row;
  var score_pledge_fa_range = sheet.getRange(score_pledge_row,
                                      ScoringObject["Pledge Ratio"]["FALL SCORE"][1]);
  var score_pledge_sp_range = sheet.getRange(score_pledge_row,
                                      ScoringObject["Pledge Ratio"]["SPRING SCORE"][1]);
  var score_pledge_tot_range = sheet.getRange(score_pledge_row,total_col);
  score_sp = score_sp >= 0 ? score_sp:0;
  score_fa = score_fa >= 0 ? score_fa:0;
  score_sp = !(isNaN(score_sp)) ? score_sp:0;
  score_fa = !(isNaN(score_fa)) >= 0 ? score_fa:0;
  score_fa_range.setValue(score_fa);
  score_sp_range.setValue(score_sp);
  score_tot_range.setValue(score_fa + score_sp);
  score_pledge_sp = score_pledge_sp >= 0 ? score_pledge_sp:0;
  score_pledge_fa = score_pledge_fa >= 0 ? score_pledge_fa:0;
  score_pledge_sp = !(isNaN(score_pledge_sp)) ? score_pledge_sp:0;
  score_pledge_fa = !(isNaN(score_pledge_fa)) ? score_pledge_fa:0;
  score_pledge_fa_range.setValue(score_pledge_fa);
  score_pledge_sp_range.setValue(score_pledge_sp);
  score_pledge_tot_range.setValue(score_pledge_fa + score_pledge_sp);
  update_dash_score("Operate", total_col);
  update_dash_score("Brotherhood", total_col);
}

function update_scores_org_gpa_serv(){
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Scoring");
  var score_data = get_scores_org_gpa_serv();
  var ScoringObject = main_range_object("Scoring");
  var total_col = ScoringObject["Societies"]["CHAPTER TOTAL"][1];
  var fall_col = ScoringObject["Societies"]["FALL SCORE"][1];
  var spring_col = ScoringObject["Societies"]["SPRING SCORE"][1];
  var societies_range = sheet.getRange(ScoringObject["Societies"].object_row, total_col);
  var societies_method = ScoringObject["Societies"]["Special"][0];
  societies_method = societies_method.replace("ORG", score_data.percent_org);
  societies_method = societies_method.replace("OFFICER", score_data.officer_count);
  var societies_max = ScoringObject["Societies"]["Max/ Semester"][0];
  var socieities_score = eval_score(societies_method, societies_max);
  var gpa_fall_range = sheet.getRange(ScoringObject["GPA"].object_row, fall_col);
  var gpa_spring_range = sheet.getRange(ScoringObject["GPA"].object_row, spring_col);
  var gpa_range = sheet.getRange(ScoringObject["GPA"].object_row, total_col);
  var gpa_method_raw = ScoringObject["GPA"]["Special"][0];
  var gpa_fall_method = gpa_method_raw.replace("GPA", score_data.gpa_avg_fall);
  var gpa_spring_method = gpa_method_raw.replace("GPA", score_data.gpa_avg_spring);
  var gpa_max = ScoringObject["GPA"]["Max/ Semester"][0];
  var gpa_fall_score = eval_score(gpa_fall_method, gpa_max);
  var gpa_spring_score = eval_score(gpa_spring_method, gpa_max);
  var service_fall_range = sheet.getRange(ScoringObject["Service Hours"].object_row, fall_col);
  var service_spring_range = sheet.getRange(ScoringObject["Service Hours"].object_row, spring_col);
  var service_range = sheet.getRange(ScoringObject["Service Hours"].object_row, total_col);
  var service_method_raw = ScoringObject["Service Hours"]["Special"][0];
  var service_fall_method = service_method_raw.replace("HOURS", score_data.percent_service_fa);
  var service_spring_method = service_method_raw.replace("HOURS", score_data.percent_service_sp);
  var service_max = ScoringObject["Service Hours"]["Max/ Semester"][0];
  var service_fall_score = eval_score(service_fall_method, service_max);
  var service_spring_score = eval_score(service_spring_method, service_max);
  Logger.log("(" + arguments.callee.name + ") " +"SOC: " + societies_method + ", SCORE: " + socieities_score);
  Logger.log("(" + arguments.callee.name + ") " +"GPA_FALL: " + gpa_fall_method + ", SCORE: " + gpa_fall_score);
  Logger.log("(" + arguments.callee.name + ") " +"GPA_SPRING: " + gpa_spring_method + ", SCORE: " + gpa_spring_score);
  Logger.log("(" + arguments.callee.name + ") " +"SERV_FALL: " + service_fall_method + ", SCORE: " + service_fall_score);
  Logger.log("(" + arguments.callee.name + ") " +"SERV_SPRING: " + service_spring_method + ", SCORE: " + service_spring_score);
  societies_range.setValue(socieities_score);
  gpa_fall_range.setValue(gpa_fall_score);
  gpa_spring_range.setValue(gpa_spring_score);
  gpa_range.setValue(gpa_fall_score + gpa_spring_score);
  service_fall_range.setValue(service_fall_score);
  service_spring_range.setValue(service_spring_score);
  service_range.setValue(service_fall_score + service_spring_score);
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
  var gpa_counts = {};
  var officer_counts = {};
  var org_counts = {};
  var service_count_fa = 0;
  var service_count_sp = 0;
  var active_total_fa = 0;
  var active_total_sp = 0;
  var active_total = 0;
  var officer_count = 0;
  var org_count = 0;
  var officers = ["Officer (Pro/Tech)", "Officer (Honor)", "Officer (Other)"];
  var orgs = ["Professional/ Technical Orgs", "Honor Orgs", "Other Orgs"];
  var gpas = ["Fall GPA", "Service Hours Fall", "Spring GPA"];
  var MemberObject = main_range_object("Membership");
  var gpa = 0;
  for (var i = 0; i < MemberObject.object_count; i++){
    var member_name = MemberObject.object_header[i];
    var org_true = false;
    var officer_true = false;
    var spring_mult = 1;
    var fall_mult = 1;
    var status = MemberObject[member_name]["Chapter Status"][0];
    var start = MemberObject[member_name]["Status Start"][0];
    if (typeof(start) != typeof(new Date())){
      if (start != ""){
        if (start.indexOf("undefined") >= 0){
          continue;
        }
      }
    }
    switch (status){
      case "Pledge":
        continue;
        break;
      case "Shiny":
        var month = start.getMonth() + 1;
        if (month<=6){fall_mult = 0;
        } else {spring_mult = 0;
        }
        break;
      case "Away":
      case "Alumn":
        var month = start.getMonth() + 1;
        if (month<=6){spring_mult = 0;
        } else {fall_mult = 0;
        }
        break;
    }
    active_total_fa += fall_mult;
    active_total_sp += spring_mult;
    active_total += 1;
    for (var j = 0; j <= 2; j++){
      var gpa_type = gpas[j];
      var gpa_raw = MemberObject[member_name][gpa_type][0];
      gpa_raw = gpa_raw == "" ? 0:gpa_raw;
      var gpa = parseFloat(gpa_raw);
      if (gpa_type.indexOf("Fall") > -1){
        if(!fall_mult){continue;
        }
      } else {
        if(!spring_mult){continue;
        }
      }
      gpa_counts[gpa_type] = gpa_counts[gpa_type] ? gpa_counts[gpa_type]+gpa:gpa;
      var org_type = orgs[j];
      var this_org = MemberObject[member_name][org_type][0].toString();
      org_counts[org_type] = org_counts[org_type] ? org_counts[org_type]:1;
      org_counts[org_type] = ((this_org.toUpperCase()!="NONE") && (this_org!="")) ? org_counts[org_type]+1:org_counts[org_type];
      org_true = (this_org.toUpperCase()!="NONE" && this_org!="") ? true:org_true;
      var officer = MemberObject[member_name][officers[j]][0];
      officer_counts[officers[j]] = officer_counts[officers[j]] ? officer_counts[officers[j]]:0;
      officer_counts[officers[j]] = officer.toUpperCase()=="YES" ? officer_counts[officers[j]]+1:officer_counts[officers[j]];
      officer_true = officer.toUpperCase()=="YES" ? true:officer_true;
//      Logger.log("(" + arguments.callee.name + ") " +"GPA: " + gpa + " ORG: " + org_true + " OFFICER: " + officer);
    }
//    var service_hours_fa = MemberObject[member_name]["Service Hours Fall"][0];
//    var service_hours_sp = MemberObject[member_name]["Service Hours Spring"][0];
    var service_hours_self_fa = MemberObject[member_name]["Service Hrs FA"][0];
    var service_hours_self_sp = MemberObject[member_name]["Service Hrs SP"][0];
    var service_hours_fa = (+service_hours_self_fa) * fall_mult;
    var service_hours_sp = (+service_hours_self_sp) * spring_mult;
    var service_count_fa = service_hours_fa >= 8 ? service_count_fa + 1:service_count_fa;
    var service_count_sp = service_hours_sp >= 8 ? service_count_sp + 1:service_count_sp;
    officer_count = officer_true ? officer_count + 1:officer_count;
    org_count = org_true ? org_count + 1:org_count;
  }
  var percent_service_fa = active_total_fa==0 ? 0:service_count_fa / active_total_fa;
  var percent_service_sp = active_total_sp==0 ? 0:service_count_sp / active_total_sp;
  var percent_org = active_total==0 ? 0:org_count / active_total;
  var gpa_avg_fall = active_total_fa==0 ? 0:gpa_counts["Fall GPA"] / active_total_fa;
  var gpa_avg_spring = active_total_sp==0 ? 0:gpa_counts["Spring GPA"] / active_total_sp;
  return {percent_service_fa: percent_service_fa,
          percent_service_sp: percent_service_sp,
          percent_org: percent_org,
          officer_count: officer_count,
          gpa_avg_fall: gpa_avg_fall,
          gpa_avg_spring: gpa_avg_spring
          }
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
  var sheet = ss.getSheetByName(sheetName);
  var score_ind = myObject["Score"][1];
  var object_date = myObject["Date"][0];
  var object_type = myObject["Type"][0];
  var score_range = sheet.getRange(row, score_ind);
  score_range.setValue(0); // To protect the current score from affecting max
  Logger.log("(" + arguments.callee.name + ") " +"Date: " + object_date + " Type:" + object_type)
  var total_scores = get_current_scores(sheetName);
  Logger.log("(" + arguments.callee.name + ") " +total_scores)
  var semester = get_semester(object_date);
  score_data.semester = semester;
  var type_score = total_scores[semester][object_type][0];
  var other_type_rows = total_scores[semester][object_type][1];
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
  update_main_score(score_data);
  score_range.setValue(score);
  return other_type_rows;
}

function update_main_score(score_data){
  Logger.log("(" + arguments.callee.name + ") ");
  Logger.log(score_data);
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName("Scoring");
  var score_row = score_data.score_ids.score_row
  var semester_range = sheet.getRange(score_row, score_data.score_ids[score_data.semester]);
  var other_semester = score_data.semester=="FALL" ? "SPRING":"FALL";
  var other_semester_range = sheet.getRange(score_row, score_data.score_ids[other_semester]);
  var other_semester_value = other_semester_range.getValue();
  var other_semester_value = (other_semester_value != "") ? other_semester_value:0;
  var total_range = sheet.getRange(score_row, score_data.score_ids.chapter);
  var total_sem_score = parseFloat(score_data.final_score) + score_data.type_score;
  var total_score = parseFloat(other_semester_value) + total_sem_score;
  semester_range.setValue(total_sem_score);
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

function get_current_scores(sheetName){
//  var sheetName = "Events";
  var ss = get_active_spreadsheet();
  var sheet = ss.getSheetByName(sheetName);
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
  date_types["SPRING"] = {};
  date_types["FALL"] = {};
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
      var ui = SpreadsheetApp.getUi();
      var result = ui.alert(
        'ERROR',
        message,
        ui.ButtonSet.OK);
      return date_types;
    }
		var type_name = type_values[i];
		var score = score_values[i];
		var semester = get_semester(date);
        var old_score = date_types[semester][type_name] ? 
				date_types[semester][type_name][0] : 0;
        var new_score = parseFloat(old_score) + parseFloat(score);
        var old_rows = date_types[semester][type_name] ? 
				date_types[semester][type_name][1] : [];
        old_rows.push(parseInt(i) + 1);
		date_types[semester][type_name] = [new_score, old_rows]
	  }
  return date_types;
}

function get_score_event(myEvent){
  var event_type = myEvent["Type"][0]
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
      var totals = get_total_members(true);
    }
      var event_date = myEvent["Date"][0];
      var semester = get_semester(event_date);
      var percent_attend = attend / totals[semester]["Student"];
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
    update_service_hours();
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
		  FALL: score_object["FALL SCORE"][1],
		  SPRING: score_object["SPRING SCORE"][1],
		  chapter: score_object["CHAPTER TOTAL"][1]
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
  var all_scores = [];
  for (var i in ScoringObject.object_header){
    var this_type = ScoringObject.object_header[i];
    all_scores.push([ScoringObject[this_type]["FALL SCORE"][0],
                     ScoringObject[this_type]["SPRING SCORE"][0],
                     ScoringObject[this_type]["CHAPTER TOTAL"][0]]);
  }
//  var all_scores =
//        Array.apply(null, Array(ScoringObject.object_count)).map(function() { return [0, 0, 0] });
  for (var semester in type_semester){
    var semester_col = semester == "FALL" ? 0:1;
    for (var event_type in type_semester[semester]){
      var event_score = type_semester[semester][event_type];
      var score_row = ScoringObject[event_type].object_row - 2 // 2 of header and row in sheet starts at 1 not 0;
      all_scores[score_row][semester_col]=event_score;
    }
  }
  for (var ind in all_scores){
    all_scores[ind][2] = all_scores[ind][0] + all_scores[ind][1];
  }
  var fall_col = ScoringObject.header_values.indexOf("FALL SCORE");
  var semester_range = score_sheet.getRange(2, fall_col, ScoringObject.object_count, 3);
//  semester_range.setValues(all_scores);
//  update_dash_score(score_data.score_type, score_data.score_ids.chapter);
}

function refresh_scores() {
  try{
    progress_update("REFRESH EVENTS");
    var ss = get_active_spreadsheet();
//    var attendance_object = main_range_object("Attendance", undefined, ss);
    var EventObject = main_range_object("Events", undefined, ss);
    var SubmitObject = main_range_object("Submissions", undefined, ss);
//    events_to_att(ss, attendance_object, EventObject);
//    refresh_attendance(ss, attendance_object, EventObject);
    EventObject = main_range_object("Events", undefined, ss);
    var ScoringObject = main_range_object("Scoring", undefined, ss);
    var totals = get_total_members(true);
    var all_scores = [];
    var all_backgrounds = [];
    var all_notes = [];
    var type_semester = {};
    type_semester["FALL"] = {};
    type_semester["SPRING"] = {};
    var exclude = ["Meetings", "Service Hours", "Misc"];
    for (var j in EventObject.object_header){
      var event_name = EventObject.object_header[j];
      var event = EventObject[event_name];
      var event_type = event["Type"][0];
      var event_date = event["Date"][0];
      var semester = get_semester(event_date);
      var score_data = get_score_method(event_type, undefined, ScoringObject);
      var score_method_edit = null;
      if (exclude.indexOf(event_type) < 0){
        score_method_edit = edit_score_method_event(event, score_data.score_method, totals);
      }
      var background = "black";
      var score = 0
      if (score_method_edit !== null){
        score = eval(score_method_edit);
        score = score.toFixed(1);
        background = "dark gray 1";
      }
      // Need to find the score, date, type to determine semester scoring
      var type_semester_score = type_semester[semester][event_type] ? type_semester[semester][event_type]:0;
      var combined_score = parseFloat(type_semester_score) + parseFloat(score);
      if (combined_score > parseFloat(score_data.score_max_semester)){
        combined_score = score_data.score_max_semester - type_semester_score;
        combined_score = combined_score > 0 ? combined_score:0;
        }
      type_semester[semester][event_type] = combined_score;
      all_scores.push([combined_score]);
      all_backgrounds.push([background]);
      all_notes.push([score_data.score_method_note]);
    }
    var event_sheet = EventObject.sheet;
    var score_col = EventObject.header_values.indexOf("Score") + 1;
    var score_range = event_sheet.getRange(2, score_col, EventObject.object_count , 1);
    score_range.setValues(all_scores);
    score_range.setBackgrounds(all_backgrounds);
    score_range.setNotes(all_notes);
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
      var semester = get_semester(submit_date);
      var score_data = get_score_method(submit_type, undefined, ScoringObject);
      var type_semester_score = type_semester[semester][submit_type] ? type_semester[semester][submit_type]:0;
      var combined_score = parseFloat(type_semester_score) + parseFloat(submit_score);
      if (combined_score > parseFloat(score_data.score_max_semester)){
        combined_score = score_data.score_max_semester - type_semester_score;
        combined_score = combined_score > 0 ? combined_score:0;
        }
      type_semester[semester][submit_type] = combined_score;
    }
    update_score_att();
    update_service_hours();
    update_score_member_pledge();
    refresh_main_scores(type_semester, ss, ScoringObject);
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
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
     'ERROR',
      message,
      ui.ButtonSet.OK);
    return "";
  }
}