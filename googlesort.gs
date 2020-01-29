var START_COL = 'W';
var START_ROW = 2;

/** @OnlyCurrentDoc */
function SortMult() {
  var spreadsheet = SpreadsheetApp.getActive();
  var currentRow = START_ROW;
  var data = spreadsheet.getRange(START_COL + START_ROW.toString()).getValue();
  while(data != ""){
    data = spreadsheet.getRange(START_COL + currentRow.toString()).getValue();
    var lines = data.split('\n');

    if(lines.length > 0) {
      splitRow(currentRow);
      runReplaceInSheet();	
    }
    
    currentRow++;
  }
};

function splitRow(row) {
  var startRow = row;
  
  var spreadsheet = SpreadsheetApp.getActive();
  
  var leftValues = spreadsheet.getRange("A"+row.toString()+":V"+row.toString()).getValues();
  var rightValues = spreadsheet.getRange("X"+row.toString()+":BM"+row.toString()).getValues();
  
  var spreadsheet = SpreadsheetApp.getActive();
  var data = spreadsheet.getRange(START_COL + row.toString()).getValue();
  var lines = data.split('\n');

  for (var i = 0; i < lines.length; i++) {
    currentRow = i + row;
    lineValue = lines[i];
    
    // Paste in side by side data
    spreadsheet.getRange("A"+currentRow.toString()+":V"+currentRow.toString()).setValues(leftValues)
    spreadsheet.getRange("X"+currentRow.toString()+":BM"+currentRow.toString()).setValues(rightValues)
   
    // Set row with line data
    spreadsheet.getRange(START_COL+currentRow.toString()).setValue(lineValue);
   
    // Dont add row on last 
    if(i < lines.length-1){
      spreadsheet.insertRowAfter(currentRow);
    }
  }
};

function runReplaceInSheet(){
  var sheet = SpreadsheetApp.getActive();

  // Replace Subject Names
  replaceInSheet(sheet, "Date: ", "");

}


function replaceInSheet(sheet, to_replace, replace_with) {
  //get the current data range values as an array
  var values = sheet.getDataRange().getValues();
  var range = sheet.getRange("A2:BM")


  //loop over the rows in the array
  for (var row in values) {
    //use Array.map to execute a replace call on each of the cells in the row.
    var replaced_values = values[row].map(function(original_value) {
      return original_value.toString().replace(to_replace, replace_with);
    });

    //replace the original row values with the replaced values
    values[row] = replaced_values;
  }

  //write the updated values to the sheet
  sheet.getDataRange().setValues(values);
  range.sort({column: 23, ascending: true})
}
