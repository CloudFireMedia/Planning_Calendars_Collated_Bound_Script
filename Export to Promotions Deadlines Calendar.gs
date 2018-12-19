/* This script will append data from the 'Collated Data' sheet in the current Spreadsheet 
* to the 'Communications Director Master' sheet in the 'Promotion Deadlines Calendar' Spreadsheet.
* It will also archive the 'Collated Data' sheet by cloning it to a new sheet with next year's date.
* Finally,the data on the the 'Collated Data' sheet below header rows will be deleted.


Redevelopment note:  POPUP MESSAGE TO WARN THE USER WHY BULLETINS ARE REPOPULATED

*///just look for the second instance of 'Jan'
//function deleteNextYearRows(){
//  var promoDeadlinesSheet = SpreadsheetApp.openById(PROMOTION_DEADLINES_CALENDAR_ID_).getSheetByName(PROMOTION_DEADLINES_CALENDAR_COMMUNICATIONS_DIRECTOR_MASTER_SHEET_NAME_); 
//  var data = promoDeadlinesSheet.getRange("A4:A").getValues();
//  for(var i = 0; i<data.length;i++)
//    if(data[i] === Jan) {//[1] because column B
//      Logger.log((i+1));
//      return i+1;
//    } 
//  
//}



//1546300800000

function exportData() {

  var collatedSheet = SpreadsheetApp.openById(COLLATED_PROMOTION_PLANNING_CALENDAR_ID_).getSheetByName(COLLATED_PROMOTION_PLANNING_CALENDAR_COLLATED_DATA_SHEET_NAME_);
  var promoDeadlinesSheet = SpreadsheetApp.openById(PROMOTION_DEADLINES_CALENDAR_ID_).getSheetByName(PROMOTION_DEADLINES_CALENDAR_COMMUNICATIONS_DIRECTOR_MASTER_SHEET_NAME_); 
  var tierDateSheet = SpreadsheetApp.openById(TIER_DUEDATE_SHEET_ID_).getSheetByName(TIER_DUEDATE_SHEET_NAME_);
  
  var number_of_full_rows_before_merge = promoDeadlinesSheet.getLastRow() - 3;
  var last_row_full_row_before_merge = promoDeadlinesSheet.getLastRow();
  
 
  // make sure the user knows what the script does
  var title = 'Notice! Script Actions';
  var prompt = Utilities.formatString("\
This script will execute the following actions:\n\n\
[ 1. ] Export data from the 'Collated Data' sheet in the current Spreadsheet to the 'Communications \n\
Director Master' sheet in the 'Promotion Deadlines Calendar' Spreadsheet. It will NOT export events \n\
at the bottom of this sheet that do not have dates assigned in Col D!\n\n\
[ 2. ] Archive the current 'Collated Data' sheet by cloning it to a new sheet within the current \n\
Spreadsheet and renaming the clone with next year's date. \n\n\
[ 3. ] Finally, all data on the the current 'Collated Data' sheet (except for header rows) will be deleted. \n\n\
Do you wish to continue?\
")
  if (SpreadsheetApp.getUi()
      .alert(title, prompt, Browser.Buttons.YES_NO) != 'YES') return;
  
  //delete events at the bottom of the source sheet that do not have dates assigned in Col D
  var collatedSheet_columnDArray = collatedSheet.getRange("D1:D")
  .getValues();
  var collatedSheet_lastRowWithValuesinColumnD = collatedSheet_columnDArray.filter(String)
  .length + 1;
  if (collatedSheet_lastRowWithValuesinColumnD != collatedSheet.getMaxRows()) {
    collatedSheet.deleteRows(collatedSheet_lastRowWithValuesinColumnD + 1, collatedSheet.getMaxRows() - collatedSheet_lastRowWithValuesinColumnD);
  };
  
  
  //delete events at the bottom of the target sheet that do not have dates assigned in Col D 
  var promoDeadlinesSheet_columnDArray = promoDeadlinesSheet.getRange("D1:D")
  .getValues();
  var promoDeadlinesSheet_lastRowWithValuesinColumnD = promoDeadlinesSheet_columnDArray.filter(String)
  .length + 2;
  if (promoDeadlinesSheet_lastRowWithValuesinColumnD != promoDeadlinesSheet.getMaxRows()) {
    promoDeadlinesSheet.deleteRows(promoDeadlinesSheet_lastRowWithValuesinColumnD + 1, promoDeadlinesSheet.getMaxRows() - promoDeadlinesSheet_lastRowWithValuesinColumnD);
  };
  
  
  // transpose data from source sheet to target sheet -- why not do this much more efficiently by column ranges rather than row ranges??
  var targetRow = promoDeadlinesSheet.getLastRow() + 1;
  var sourceRows = collatedSheet.getMaxRows();
  for (var sourceRow = 3; sourceRow < sourceRows; sourceRow++) {
    
    var collatedSheet_eventSuggestedLevelOfPromotion = collatedSheet.getRange(sourceRow, 3);
    var collatedSheet_eventSuggestedNextYearDate = collatedSheet.getRange(sourceRow, 4);
    var collatedSheet_eventName = collatedSheet.getRange(sourceRow, 5);
    var collatedSheet_eventSponsorship = collatedSheet.getRange(sourceRow, 6);
//    var collatedSheet_eventComments = collatedSheet.getRange(sourceRow, 7);
    
    var promoDeadlinesSheet_eventSuggestedLevelOfPromotion = promoDeadlinesSheet.getRange(targetRow, 3);
    var promoDeadlinesSheet_eventSuggestedNextYearDate = promoDeadlinesSheet.getRange(targetRow, 4);
    var promoDeadlinesSheet_eventName = promoDeadlinesSheet.getRange(targetRow, 5);
    var promoDeadlinesSheet_eventSponsorship = promoDeadlinesSheet.getRange(targetRow, 8);
//    var promoDeadlinesSheet_eventComments = promoDeadlinesSheet.getRange(targetRow, 12);
    targetRow++;
    
    promoDeadlinesSheet_eventSuggestedLevelOfPromotion.setValue(collatedSheet_eventSuggestedLevelOfPromotion.getValue());
    promoDeadlinesSheet_eventSuggestedNextYearDate.setValue(collatedSheet_eventSuggestedNextYearDate.getValue());
    promoDeadlinesSheet_eventName.setValue(collatedSheet_eventName.getValue());
    promoDeadlinesSheet_eventSponsorship.setValue(collatedSheet_eventSponsorship.getValue());
//    promoDeadlinesSheet_eventComments.setValue(collatedSheet_eventComments.getValue());
  }
  
  
  //copy source sheet and archive and rename
  var currentDate = new Date();
  var currentYear = currentDate.getFullYear();
  var newSheetName = currentYear + 1 + " Archive";
  var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentSheet = sourceSpreadsheet.getSheetByName('Collated Data')
  .copyTo(sourceSpreadsheet);
  var existantArchive = sourceSpreadsheet.getSheetByName(newSheetName);
  if (existantArchive) existantArchive.setName(newSheetName + " old copy");
  SpreadsheetApp.flush();
  currentSheet.setName(newSheetName);
  
  
  // clear last remaining row on current source sheet ignoring headers
  var numberHeaderRows = collatedSheet.getFrozenRows();
  var firstRowNotIncludingHeaders = collatedSheet.getFrozenRows() + 1;
  if (collatedSheet.getMaxRows() > numberHeaderRows + 1) {
    collatedSheet.deleteRows(firstRowNotIncludingHeaders, collatedSheet.getMaxRows() - firstRowNotIncludingHeaders)
    collatedSheet.getRange(firstRowNotIncludingHeaders, 1, 1, collatedSheet.getLastColumn())
    .clearContent()
  }
  
  
  // unmerge all months, weeks, and bulletin, on target sheet for this year and next
  var promoDeadlinesSheet_columnDArray = promoDeadlinesSheet.getRange("D4:D")
  .getValues();
  var promoDeadlinesSheet_lastRowWithValuesinColumnD = promoDeadlinesSheet_columnDArray.filter(String)
  .length + 5;
  var unmergeRangeList1 = promoDeadlinesSheet.getRangeList(["A4:B" + promoDeadlinesSheet_lastRowWithValuesinColumnD]);
  var unmergeRangeList2 = promoDeadlinesSheet.getRangeList(["M4:M" + promoDeadlinesSheet_lastRowWithValuesinColumnD]);
  unmergeRangeList1.breakApart();
  unmergeRangeList2.breakApart();
  
  //just in case, re-enumerate all weeks, months, and promo statuses, for this year and next
  promoDeadlinesSheet.getRange('A3')
  .setFormula('={"MONTH"; ArrayFormula( if(D4:D, TEXT(MONTH(D4:D)&"-1","MMM"), IFERROR(1/0)) ) }');
  promoDeadlinesSheet.getRange('B3')
  .setFormula('={"WEEK"; ArrayFormula( if(D4:D, WEEKNUM(D4:D), IFERROR(1/0)) ) }');
  promoDeadlinesSheet.getRange('L3')
  .setFormula('={"PROMO STATUS"; ArrayFormula(if((F4:F="No")*(G4:G="No"),"Awaiting Promotion Request",if((F4:F="Yes")*(G4:G="Yes"),"Scheduled",if((F4:F="Yes")*(G4:G="No"),"Awaiting Promotion Request",if((F4:F="No")*(G4:G="Yes"),"In Process",if((F4:F="Yes")*(G4:G="N/A"),"ERROR: COLUMNS F & G MUST NOT CONTAIN ONLY ONE N/A",if((F4:F="N/A")*(G4:G="Yes"),"ERROR: COLUMNS F & G MUST NOT CONTAIN ONLY ONE N/A",if((F4:F="N/A")*(G4:G="No"),"ERROR: COLUMNS F & G MUST NOT CONTAIN ONLY ONE N/A",if((F4:F="No")*(G4:G="N/A"),"ERROR: COLUMNS F & G MUST NOT CONTAIN ONLY ONE N/A",if((F4:F="N/A")*(G4:G="N/A"),"N/A","N/A")))))))))) }');
  
  
  // sort sheet by date or merges coming up won't work
  promoDeadlinesSheet.sort(4);

  
  // vertically merge each month for the current year
  var current_year_month_number_of_rows = number_of_full_rows_before_merge;
  var current_year_month_first_row = 4;
  var current_year_month_column = 1;
  var current_year_month_c = {};
  var current_year_month_k = "";
  var current_year_month_offset = 0;
  // Retrieve values of column A.
  var current_year_month_data = promoDeadlinesSheet.getRange(current_year_month_first_row, current_year_month_column, current_year_month_number_of_rows, 1)
  .getValues()
  .filter(String);
  current_year_month_data.forEach(function(current_year_month_e) {
    current_year_month_c[current_year_month_e[0]] = current_year_month_c[current_year_month_e[0]] ? current_year_month_c[current_year_month_e[0]] + 1 : 1;
  });
  current_year_month_data.forEach(function(current_year_month_e) {
    if (current_year_month_k != current_year_month_e[0]) {
      promoDeadlinesSheet.getRange(current_year_month_first_row + current_year_month_offset, current_year_month_column, current_year_month_c[current_year_month_e[0]], 1)
      .merge();
      current_year_month_offset += current_year_month_c[current_year_month_e[0]];
    }
    current_year_month_k = current_year_month_e[0];
  });
  
  
  //vertically merge each month for the new year
  var new_year_month_first_row = last_row_full_row_before_merge + 1; 
  var new_year_month_column = 1;
  var new_year_month_number_of_rows = promoDeadlinesSheet.getLastRow() - last_row_full_row_before_merge; // -184
  var new_year_month_c = {};
  var new_year_month_k = "";
  var new_year_month_offset = 0;
  // Retrieve values of column A.
  var new_year_month_data = promoDeadlinesSheet.getRange(new_year_month_first_row, new_year_month_column, new_year_month_number_of_rows, 1)
  .getValues()
  .filter(String);
  new_year_month_data.forEach(function(new_year_month_e) {
    new_year_month_c[new_year_month_e[0]] = new_year_month_c[new_year_month_e[0]] ? new_year_month_c[new_year_month_e[0]] + 1 : 1;
  });
  new_year_month_data.forEach(function(new_year_month_e) {
    if (new_year_month_k != new_year_month_e[0]) {
      promoDeadlinesSheet.getRange(new_year_month_first_row + new_year_month_offset, 1, new_year_month_c[new_year_month_e[0]], 1)
      .merge();
      new_year_month_offset += new_year_month_c[new_year_month_e[0]];
    }
    new_year_month_k = new_year_month_e[0];
  });
  
  
  // vertically merge each week for the current year
  var current_year_week_first_row = 4; // Start row number for values.
  var current_year_week_column = 2;
  var current_year_week_number_of_rows = number_of_full_rows_before_merge; 
  var current_year_week_c = {};
  var current_year_week_k = "";
  var current_year_week_offset = 0;
  // Retrieve values of column B.
  var current_year_week_data = promoDeadlinesSheet.getRange(current_year_week_first_row, current_year_week_column, number_of_full_rows_before_merge-1, 1) // 4, 2, 185, 1 //
  .getValues()
  .filter(String);
  current_year_week_data.forEach(function(current_year_week_e) {
    current_year_week_c[current_year_week_e[0]] = current_year_week_c[current_year_week_e[0]] ? current_year_week_c[current_year_week_e[0]] + 1 : 1;
  });
  current_year_week_data.forEach(function(current_year_week_e) {
    if (current_year_week_k != current_year_week_e[0]) {
      promoDeadlinesSheet.getRange(current_year_week_first_row + current_year_week_offset, 2, current_year_week_c[current_year_week_e[0]], 1)
      .merge();
      current_year_week_offset += current_year_week_c[current_year_week_e[0]];
    }
    current_year_week_k = current_year_week_e[0];
  });
  
  
  // break apart vertical merge for this year's Week 53+, otherwise the next merge will not run
  var ct = getFirstEmptyRow_();
  var n = ct + 1;
  var m = promoDeadlinesSheet.getLastRow();
  var unmergeRangeList = promoDeadlinesSheet.getRangeList(["B" + n + ":B" + m]);
  unmergeRangeList.breakApart();
  
  
  // vertically merge each week for the new year
  var new_year_week_first_row = last_row_full_row_before_merge + 1; // Start row number for values.
  var new_year_week_first_column = 2;
  var new_year_week_number_of_rows = promoDeadlinesSheet.getLastRow() - last_row_full_row_before_merge;
  var new_year_week_c = {};
  var new_year_week_k = "";
  var new_year_week_offset = 0;
  // Retrieve values of column B.
  var new_year_week_data = promoDeadlinesSheet.getRange(new_year_week_first_row, new_year_week_first_column, new_year_week_number_of_rows, 1)
  .getValues()
  .filter(String);
  new_year_week_data.forEach(function(new_year_week_e) {
    new_year_week_c[new_year_week_e[0]] = new_year_week_c[new_year_week_e[0]] ? new_year_week_c[new_year_week_e[0]] + 1 : 1;
  });
  new_year_week_data.forEach(function(new_year_week_e) {
    if (new_year_week_k != new_year_week_e[0]) {
      promoDeadlinesSheet.getRange(new_year_week_first_row + new_year_week_offset, 2, new_year_week_c[new_year_week_e[0]], 1)
      .merge();
      new_year_week_offset += new_year_week_c[new_year_week_e[0]];
    }
    new_year_week_k = new_year_week_e[0];
  });
  
  
  // set date format to MM.DD in Col D of Promo Deadlines Calendar
  promoDeadlinesSheet.getRange("D:D")
  .setNumberFormat('MM.DD');


  //merge bulletins vertically for the current year and new year
  var bulletins_range = promoDeadlinesSheet.getRange("B4:B" + promoDeadlinesSheet.getMaxRows());
  bulletins_range.copyFormatToRange(promoDeadlinesSheet, 13, 13, 4, promoDeadlinesSheet.getMaxRows());
  
  
  //populate web cal and pr req? columns for the new year
  var dataRange = promoDeadlinesSheet.getDataRange();
  var values = dataRange.getValues();
  var startRow = promoDeadlinesSheet.getFrozenRows();
  var numRows = values.length;
  var numColumns = values[0].length;
  for (var i = startRow; i < numRows; i++) {
    if (values[i][5] || values[i][6]) continue; //something already set, next please!
    var val = values[i][2] == "N/A" ? 'Yes' : 'No';
    promoDeadlinesSheet.getRange(i + 1, 6)
    .setValue(val); //LISTED ON WEB CAL
    promoDeadlinesSheet.getRange(i + 1, 7)
    .setValue(val); //PROMO REQUESTED
  }
  
  
  // populate due date deadlines 
  // getting tiers name from tiers due Date sheet
  var tier1Name = tierDateSheet.getRange('A2').getValue();
  var tier2Name = tierDateSheet.getRange('A3').getValue();
  var tier3Name = tierDateSheet.getRange('A4').getValue();
  // getting tiers due dates from tiers due Date sheet
  var tier1DueDate = tierDateSheet.getRange('B2').getValue();
  var tier2DueDate = tierDateSheet.getRange('B3').getValue();
  var tier3DueDate = tierDateSheet.getRange('B4').getValue();
  // writing teirs name to active sheet
  promoDeadlinesSheet.getRange('I3').setValue(tier1Name);
  promoDeadlinesSheet.getRange('J3').setValue(tier2Name);
  promoDeadlinesSheet.getRange('K3').setValue(tier3Name);
  // getting range
  var rangeColumnJ = promoDeadlinesSheet.getRange("J1");
  var rangeColumnK = promoDeadlinesSheet.getRange("K1");
  // hiding/ unhiding column J 
  if (tier2DueDate ==''){
    promoDeadlinesSheet.hideColumn(rangeColumnJ);
  }
  else{
    promoDeadlinesSheet.unhideColumn(rangeColumnJ)
  }
   // hiding/ unhiding column K 
  if (tier3DueDate ==''){
    promoDeadlinesSheet.hideColumn(rangeColumnK);
  }
  else{
    promoDeadlinesSheet.unhideColumn(rangeColumnK)
  }
  // save value of tier1, tier2 and tier3 for later use
  PropertiesService.getScriptProperties().setProperty('tier1DueDate', tier1DueDate);
  PropertiesService.getScriptProperties().setProperty('tier2DueDate', tier2DueDate);
  PropertiesService.getScriptProperties().setProperty('tier3DueDate', tier3DueDate);
  var lastRow = promoDeadlinesSheet.getLastRow();
  // loop through each cell in column D to update column I, J K
  for (var c = last_row_full_row_before_merge; c < lastRow+1; c++) {
    var startDate  = promoDeadlinesSheet.getRange('D' + c ).getValue();
    //finding value for column I
    var newTier1 = new Date(startDate.getTime()-tier1DueDate*3600000*24).toLocaleDateString(undefined, {day:'numeric', month: 'numeric', year: 'numeric'});
    var day = (newTier1+"").substring(0,3);
    // if day is Sunday than move it to Monday
    if(day=='Sun'){
      newTier1 = new Date(newTier1.getTime()+1*3600000*24).toLocaleDateString(undefined, {day:'numeric', month: 'numeric', year: 'numeric'});
    }
    // if day is Saturday than move it to Friday
    if(day=='Sat'){
      newTier1 = new Date(newTier1.getTime()-1*3600000*24).toLocaleDateString(undefined, {day:'numeric', month: 'numeric', year: 'numeric'});
    }
    promoDeadlinesSheet.getRange('I' + c ).setValue(newTier1); // setting value of active cell for Column I = tier1
    //Finding Value for Column J if tier2 is  not empty
    if(tier2DueDate != ''){
      var newTier2 = new Date(startDate.getTime()-tier2DueDate*3600000*24).toLocaleDateString(undefined, {day:'numeric', month: 'numeric', year: 'numeric'});
      day = (newTier2+"").substring(0,3);
      // if day is Sunday than move it to Monday
      if(day=='Sun'){
        newTier2 = new Date(newTier2.getTime()+1*3600000*24).toLocaleDateString(undefined, {day:'numeric', month: 'numeric', year: 'numeric'});
      }
      // if day is Saturday than move it to Friday
      if(day=='Sat'){
        newTier2 = new Date(newTier2.getTime()-1*3600000*24).toLocaleDateString(undefined, {day:'numeric', month: 'numeric', year: 'numeric'});
      }
      promoDeadlinesSheet.getRange('J' + c ).setValue(newTier2); // setting value of active cell for Column J = tier2
    }
    //finding value for column K if tier3 is not empty
    if(tier3DueDate != ''){
      var newTier3 = new Date(startDate.getTime()-tier3DueDate*3600000*24).toLocaleDateString(undefined, {day:'numeric', month: 'numeric', year: 'numeric'});
      day = (newTier3+"").substring(0,3);
      // if day is Sunday than move it to Monday
      if(day=='Sun'){
        newTier3 = new Date(newTier3.getTime()+1*3600000*24).toLocaleDateString(undefined, {day:'numeric', month: 'numeric', year: 'numeric'});
      }
      // if day is Saturday than move it to Friday
      if(day=='Sat'){
        var newTier3 = new Date(newTier3.getTime()-1*3600000*24).toLocaleDateString(undefined, {day:'numeric', month: 'numeric', year: 'numeric'});
      }
      promoDeadlinesSheet.getRange('K' + c ).setValue(newTier3); // setting value of active cell for Column K = tier2
    }
  }
  
  
  // copy current year's bulletin schedule onto next year's schedule
  var range = promoDeadlinesSheet.getRange(promoDeadlinesSheet.getFrozenRows() + 1, 13, promoDeadlinesSheet.getLastRow() - promoDeadlinesSheet.getFrozenRows() + 1);
  var bulletins = range.getValues()
  .reduce(function(bulletins, name) {
    if (name[0]) bulletins.push(name[0]);
    return bulletins;
  }, []);
  
  calendarMakeRepeat(bulletins, promoDeadlinesSheet.getFrozenRows() + 1);
}

function calendarMakeRepeat(arr, readFrom) {
  var targetSpreadsheet = SpreadsheetApp.openById(PROMOTION_DEADLINES_CALENDAR_ID_);
  var promoDeadlinesSheet = targetSpreadsheet.getSheetByName("Communications Director Master");
  var range = promoDeadlinesSheet.getDataRange();
  var values = range.getValues();
  var k = 0;
  var startFrom = 0;
  
  //gets the first row with data inb col 1
  for (var j = readFrom; j <= values.length; j++) {
    if (range.getCell(j, 1)
        .getValue()) {
      var startFrom = j;
      break;
    }
  }
  while (k < 52) { //yeah, only if there really /are/ 52... otherwise infinite loop
    if (startFrom > values.length) return; //happens during recursion
    if (range.getCell(startFrom, 2)
        .getValue()) {
      range.getCell(startFrom, 13)
      .setValue(arr[k]);
      k++
    }
    startFrom++;
  }
  calendarMakeRepeat(arr, startFrom)
}

//Private Functions
//================

function getFirstEmptyRow_() {
  var targetSpreadsheet = SpreadsheetApp.openById(PROMOTION_DEADLINES_CALENDAR_ID_);
  var promoDeadlinesSheet = targetSpreadsheet.getSheetByName("Communications Director Master");
  var column = promoDeadlinesSheet.getRange("B:B");
  var values = column.getValues();
  var ct = 0;
  while (values[ct][0] != "53") {
    ct++;
  }
  return ct;
}