function sync() {
  var dash_id = SCRIPT_PROP.getProperty("dash");
//  var dash_id = "10ebwK7tTKgveVCEOpRle2S17d4UjwmsoXXCPFvC9A-A";
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
  for (var event in event_object){
    event = event_object[event];
    if (typeof event != typeof {}){continue;};
    if ('Event Name' in event){
      var test_name  = event['Event Name'][0] + event['Date'][0];
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
  var member_value_obj = get_membership_ranges();
  var score_dict= {
    init_sp_range: 'Spring Init',
    init_fa_range: 'Fall Init',
    pledge_sp_range: 'Spring Pledge',
    pledge_fa_range: 'Fall Pledge',
    grad_sp_range: 'Spring Graduated',
    grad_fa_range: 'Fall Graduated',
    act_sp_range: 'Spring Active',
    act_fa_range: 'Fall Active'
  }
  var chapter_row = main_chapter[chapter+chapter][1];
  for (var score_type_raw in member_value_obj){
    var score_type = score_dict[score_type_raw];
    var col = main_chapter.first_row.indexOf(score_type)+1;
    var val = member_value_obj[score_type_raw].getValue();
    var row_range = main_sheet.getRange(chapter_row, col).setValue(val);
  }
  var ss = get_active_spreadsheet();
  var scores = ['Brotherhood', 'Service', 'Operate', 'ProDev'];
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
}

function sync_officers(oer){
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
    if (officer+officer in officer_chapter){
      var officer_row = officer_chapter[officer+officer][1];
    } else {
      var officer_row = officer_rows++;
    }
    var officer_range = officer_sheet.getRange(officer_row, 1, 1, officer_cols);
    officer_range.setValues([oer[officer]]);
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