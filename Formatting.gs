function sortCollatedData() {
  
  var sourceSheet = SpreadsheetApp.openById(COLLATED_PROMOTION_PLANNING_CALENDAR_ID_).getSheetByName(COLLATED_PROMOTION_PLANNING_CALENDAR_COLLATED_DATA_SHEET_NAME_);
  
  //break apart vertical merges and sort source sheet by Col D
  var sourceSheet_columnDArray = sourceSheet.getRange("D1:D").getValues();
  var sourceSheet_lastRowWithValuesinColumnD = sourceSheet_columnDArray.filter(String).length+1;
  var unmergeRangeList = sourceSheet.getRangeList(["A3:B" + sourceSheet_lastRowWithValuesinColumnD]);
  unmergeRangeList.breakApart();
  sourceSheet.sort(4);
  
  //re-enumerate next year's week numbers and re-merge
  var sourceSheet_dataRange = sourceSheet.getDataRange();
  var sourceSheet_allValues = sourceSheet_dataRange.getValues();
  var sundayRows = [];
  for (var i = 2; i < sourceSheet_allValues.length; i++) //skip two header rows
    if (sourceSheet_allValues[i][4].indexOf("Sunday Service") > -1) sundayRows.push(i);
  var maxValue = sundayRows[sundayRows.length - 1]; //the last entry
  for (var k = 0; k < sundayRows.length; k++) {
    var from = sundayRows[k] + 1; //+1 0-based array offset
    var to = sundayRows[k + 1]; //+1 next array index, gives us the row before the next Sunday
    var numRows = to - from + 1; //number of rows
    if (from > maxValue) break;
    //sheet.getRange(from, 1, 1, 2).mergeVertically();//uh, merging a single row doesn't do anything, just skip (or end in this case)
    //else
    sourceSheet.getRange(from, 2, numRows, 1)
    .mergeVertically();
    sourceSheet.getRange('B2')
    .setFormula('={"WEEK"; ArrayFormula( if(D3:D, WEEKNUM(D3:D), IFERROR(1/0)) ) }');
  }
  
  //re-enumerate next year's months and re-merge 
  sourceSheet.getRange('A2')
  .setFormula('={"MONTH"; ArrayFormula( if(D3:D, TEXT(MONTH(D3:D)&"-1","MMM"), IFERROR(1/0)) ) }');
  var start = 3; // Start row number for values
  var c = {};
  var k = "";
  var offset = 0;
  // retrieve values of column A
  var data = sourceSheet.getRange(start, 1, sourceSheet.getLastRow(), 1)
  .getValues()
  .filter(String);
  // retrieve the number of duplication values
  data.forEach(function(e) {
    c[e[0]] = c[e[0]] ? c[e[0]] + 1 : 1;
  });
  // merge cells
  data.forEach(function(e) {
    if (k != e[0]) {
      sourceSheet.getRange(start + offset, 1, c[e[0]], 1)
      .merge();
      offset += c[e[0]];
    }
    k = e[0];
  });
}


// resets month rows formatting in Col E by reseting the the month array in Col A, 
// this is necessary due to a quirk in the way that conditional formatting works on 
// merged rows
function onEdit(e) {
   var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();   
   var editRange = { // D4:G184
    top : 3,
    bottom : sourceSheet.getMaxRows(),
    left : 1,
    right : 7
  };
  var thisRow = e.range.getRow();
  if (thisRow < editRange.top || thisRow > editRange.bottom) return;
  var thisCol = e.range.getColumn();
  if (thisCol < editRange.left || thisCol > editRange.right) return;
  var sourceSheet = e.range.getSheet();
  commsDirectorMasterSheet.getRange("A4:A4")   
    .clearContent(); // Clear cell of values (not formatting)
}


