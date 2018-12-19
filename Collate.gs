function collatePlanningMaterial() {
  
  var sourceSpreadsheet = SpreadsheetApp.openById(PROMOTION_PLANNING_CALENDAR_FOR_TEAMS_ID_);
  var targetSheet = SpreadsheetApp.openById(COLLATED_PROMOTION_PLANNING_CALENDAR_ID_).getSheetByName(COLLATED_PROMOTION_PLANNING_CALENDAR_COLLATED_DATA_SHEET_NAME_);
  
  
  // clear rows on target sheet ignoring headers
  var numberHeaderRows = targetSheet.getFrozenRows();
  var firstRowNotIncludingHeaders = targetSheet.getFrozenRows() + 1;
  if (targetSheet.getMaxRows() > numberHeaderRows + 1) {
    targetSheet.deleteRows(firstRowNotIncludingHeaders, targetSheet.getMaxRows() - firstRowNotIncludingHeaders);
    targetSheet.getRange(firstRowNotIncludingHeaders, 1, 1, targetSheet.getLastColumn())
    .clearContent()
  } else {
    targetSheet.getRange(firstRowNotIncludingHeaders, 1, 1, targetSheet.getLastColumn())
    .clearContent()
  }
  
  
   // add ordinal Sundays 
  var monthOccurance = 0;
  var startYear = new Date()
  .getFullYear() + 1;
  var sundays = [];
  for (var m = 1; m <= 12; m++) {
    var startDate = new Date(startYear, m, 0);
    var totalDays = new Date(startYear, m, 0)
    .getDate()
    monthOccurance = 0;
    for (var i = 1; i < totalDays; i++) {
      var theDay = new Date(startYear, m - 1, i);
      if (theDay.getDay() == 0) {
        monthOccurance++;
        Logger.log(monthOccurance + ' - ' + theDay);
        var sunday = {
          date: theDay,
          occurance: monthOccurance
        }
        sundays.push(sunday);
      }
    }
  }
  for (var i = 0; i < sundays.length; i++) {
    var sunday = sundays[i];
    var occuranceText = ['First', 'Second', 'Third', 'Fourth', 'Fifth'];
    var targetRow = targetSheet.getMaxRows();
    targetSheet.insertRows(targetRow);
    var targetSheet_SundayEventSuggestedLevelOfPromotion = targetSheet.getRange(targetRow, 3);
    var targetSheet_SundayEventSuggestedNextYearDate = targetSheet.getRange(targetRow, 4);
    var targetSheet_SundayEventName = targetSheet.getRange(targetRow, 5);
    var targetSheet_SundayEventSponsorship = targetSheet.getRange(targetRow, 6);
    var yy = (sunday.date.getFullYear())
    .toString()
    .substr(-2)
    .toString();
    var mm = ("0" + (sunday.date.getMonth() + 1))
    .slice(-2)
    .toString();
    var dd = ("0" + (sunday.date.getDate()))
    .slice(-2)
    .toString();
    targetSheet_SundayEventSuggestedLevelOfPromotion.setValue('N/A');
    targetSheet_SundayEventSuggestedNextYearDate.setValue(mm + '/' + dd + '/' + yy); 
    targetSheet_SundayEventName.setValue(' ' + occuranceText[sunday.occurance - 1] + ' Sunday Service');
    targetSheet_SundayEventSponsorship.setValue('N/A');
    targetSheet.sort(4);
  }
  
  
  // loop across source sheets
  var targetRow = targetSheet.getLastRow() + 1;
  var sourceSheet_allSheets = sourceSpreadsheet.getSheets();
  for (var s in sourceSheet_allSheets) {
    var sourceSheet_currentSheet = sourceSheet_allSheets[s];
    // skip one iteration to ignore Instructions sheet
    if (
      (sourceSheet_currentSheet.getName() == "Instructions")) continue;
    // remove protections from source sheets 
    var me = Session.getEffectiveUser();
    var protections = sourceSheet_currentSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = 0; i < protections.length; i++) {
      var protection = protections[i];
      if (protection.canEdit()) {
        protection.remove();
      }
    }
    // transpose data from source sheets ... 
    var rows = sourceSheet_currentSheet.getLastRow();
    for (var row = 4; row < rows+1; row++) {
      var sourceSheet_eventSponsorship = sourceSheet_currentSheet.getName();
      var sourceSheet_eventThisYearsPromotionLevel = sourceSheet_currentSheet.getRange(row, 1);
      var sourceSheet_eventThisYearsDate = sourceSheet_currentSheet.getRange(row, 2);
      var sourceSheet_RecurringNegativeConfirmation = sourceSheet_currentSheet.getRange(row, 3);  // Yes == unanswered; No == IS NOT recurring next year
      var sourceSheet_RecurringPositiveConfirmation = sourceSheet_currentSheet.getRange(row, 5);  // No == unanswered; Yes == IS recurring next year 
      var sourceSheet_eventName = sourceSheet_currentSheet.getRange(row, 7);
      var sourceSheet_eventSuggestedLevelOfPromotion = sourceSheet_currentSheet.getRange(row, 8);
      var sourceSheet_eventSuggestedNextYearDate = sourceSheet_currentSheet.getRange(row, 9);
      var sourceSheet_eventComments = sourceSheet_currentSheet.getRange(row, 10);
      // ... onto target sheets if the event is confirmed to recur or if is not NOT confirmed to recur
      if ( (sourceSheet_RecurringNegativeConfirmation.getValue() == 'Yes' && sourceSheet_RecurringPositiveConfirmation.getValue() == 'Yes') || (sourceSheet_RecurringNegativeConfirmation.getValue() == 'Yes' && sourceSheet_RecurringPositiveConfirmation.getValue() == 'No') ) {
      var targetSheet_eventSuggestedLevelOfPromotion = targetSheet.getRange(targetRow, 3);
      var targetSheet_eventSuggestedNextYearDate = targetSheet.getRange(targetRow, 4);
      var targetSheet_eventName = targetSheet.getRange(targetRow, 5);
      var targetSheet_eventSponsorship = targetSheet.getRange(targetRow, 6);
      var targetSheet_eventComments = targetSheet.getRange(targetRow, 7);
      var targetSheet_eventThisYearsPromotionLevel = targetSheet.getRange(targetRow, 8);
      var targetSheet_eventThisYearsDate = targetSheet.getRange(targetRow, 9);
      targetRow++;
      /*NOTE Range.getValue is heavily used by the script, below.
      The script uses a method which is considered expensive. Each invocation generates a time consuming call to a remote server. 
      That may have critical impact on the execution time of the script, especially on large data. If performance is an issue for 
      the script, you should consider using another method, e.g. Range.getValues().      
      */
      targetSheet_eventSponsorship.setValue(sourceSheet_eventSponsorship);
      targetSheet_eventSuggestedLevelOfPromotion.setValue(sourceSheet_eventSuggestedLevelOfPromotion.getValue());
      targetSheet_eventSuggestedNextYearDate.setValue(sourceSheet_eventSuggestedNextYearDate.getValue());
      targetSheet_eventThisYearsPromotionLevel.setValue(sourceSheet_eventThisYearsPromotionLevel.getValue());
      targetSheet_eventThisYearsDate.setValue(sourceSheet_eventThisYearsDate.getValue());
      targetSheet_eventName.setValue(sourceSheet_eventName.getValue());
      targetSheet_eventComments.setValue(sourceSheet_eventComments.getValue());
     } 
   }
 }
    
 
  // unmerge all months, weeks 
  var targetSheet_columnBArray = targetSheet.getRange("A3:B")
  .getValues();
  var targetSheet_lastRowWithValuesinColumnB = targetSheet_columnBArray.filter(String)
  .length + 1;
  var unmergeRangeList = targetSheet.getRangeList(["A3:B" + targetSheet_lastRowWithValuesinColumnB]);
  unmergeRangeList.breakApart();

  
  //reset array for enumerating weeks, months 
  targetSheet.getRange('A2')
  .setFormula('={"MONTH"; ArrayFormula( if(D3:D, TEXT(MONTH(D3:D)&"-1","MMM"), IFERROR(1/0)) ) }');
  targetSheet.getRange('B2')
  .setFormula('={"WEEK"; ArrayFormula( if(D3:D, WEEKNUM(D3:D), IFERROR(1/0)) ) }');

  
  // sort sheet by date 
  targetSheet.sort(4);

  
  // vertically merge each month
  var month_number_of_rows = targetSheet_lastRowWithValuesinColumnB;
  var month_first_row = 3;
  var month_column = 1;
  var month_c = {};
  var month_k = "";
  var month_offset = 0;
  var month_data = targetSheet.getRange(month_first_row, month_column, month_number_of_rows, 1)
  .getValues()
  .filter(String);
  month_data.forEach(function(month_e) {
    month_c[month_e[0]] = month_c[month_e[0]] ? month_c[month_e[0]] + 1 : 1;
  });
  month_data.forEach(function(month_e) {
    if (month_k != month_e[0]) {
      targetSheet.getRange(month_first_row + month_offset, month_column, month_c[month_e[0]], 1)
      .merge();
      month_offset += month_c[month_e[0]];
    }
    month_k = month_e[0];
  });
 
  
  // vertically merge each week 
  var week_number_of_rows = targetSheet_lastRowWithValuesinColumnB; 
  var week_first_row = 3; // Start row number for values.
  var week_column = 2;
  var week_c = {};
  var week_k = "";
  var week_offset = 0;
  var week_data = targetSheet.getRange(week_first_row, week_column, week_number_of_rows, 1) // 4, 2, 185, 1 //
  .getValues()
  .filter(String);
  week_data.forEach(function(week_e) {
    week_c[week_e[0]] = week_c[week_e[0]] ? week_c[week_e[0]] + 1 : 1;
  });
  week_data.forEach(function(week_e) {
    if (week_k != week_e[0]) {
      targetSheet.getRange(week_first_row + week_offset, 2, week_c[week_e[0]], 1)
      .merge();
      week_offset += week_c[week_e[0]];
    }
    week_k = week_e[0];
  });

  
  //set dates format to MM.DD
  targetSheet.getRange("D:D").setNumberFormat('MM.DD');

}
