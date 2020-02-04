function addNewFieldsForInput() {
  var spreadsheet = SpreadsheetApp.getActive();
  var numColumns = spreadsheet.getLastColumn()-5;
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Input'), true);
  spreadsheet.getActiveSheet().insertColumnsBefore(6, 1);
  spreadsheet.getRange('Format!F:F').copyTo(spreadsheet.getRange('Input!F:F'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  hideOldColumns();
  spreadsheet.getRange('F2').activate();
};


function hideEmptyRows() {
  var s = SpreadsheetApp.getActive().getSheetByName('Input');
  s.showRows(1, s.getMaxRows());
  s.getRange('F:F')
  .getValues()
  .forEach( function (r, i) {
    if (r[0] == '') 
      s.hideRows(i + 1);
  });
}

function showAllRows() {
  var s = SpreadsheetApp.getActive().getSheetByName('Input');
  var rRows = s.getRange("A:A");
  s.unhideRow(rRows);
}


function reformatSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var numColumns = sheet.getMaxColumns();  
  var numRows = sheet.getMaxRows();  
  sheet.getRange(1,1,numRows,numColumns).clearFormat();
  ss.getRange('Format!A:F').copyTo(sheet.getRange(1,1,1,6), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  ss.getRange('Format!G:G').copyTo(sheet.getRange(1,7,numRows,numColumns-6), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}



function hideOldColumns() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var numColumns = sheet.getLastColumn();
  var v = sheet.getRange(1,1,1,numColumns).getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  for (var i = sheet.getLastColumn(); i > 6; i--) {
    var t = v[0][i - 1];
    var u = new Date(t);
    if ((u < today) || (typeof t === 'string' || t instanceof String)) { 
      sheet.hideColumns(i);
    }
  }
}

function showAllColumns() {
  var spreadsheet = SpreadsheetApp.getActive();
  var numColumns = spreadsheet.getLastColumn();
  //  spreadsheet.getRange('F:F').activate();
  spreadsheet.getActiveSheet().showColumns(1, numColumns);
  spreadsheet.getRange('F2').activate();
};


function removeEmptyColumns() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Input');
  var numColumns = spreadsheet.getLastColumn();  
  for (var i = numColumns - 1; i>=6; i--) {
    if (sheet.getRange(1,i+1,sheet.getMaxRows(),1).isBlank()) {
      sheet.deleteColumn(i+1); 
      Logger.log(i+1)
    }  
  }
}



//****** Dead Code Below ******//


// This function does not work. Uses sheet.getRange().getValues() and transposes the matrix, 
// then looks for empty arrays. Fails in two ways: 1) the transpose function is another loop and therefore 
// inefficient (I could not figure out the transpose(a) 'map' method) and 2) the length of an empty transposed 
// array is not = 0 for some reason.
function removeEmptyColumns2() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Input');
  var numColumns = spreadsheet.getLastColumn(); 
  var numRows = spreadsheet.getLastRow(); 
  var newArray = transposeArray();
//  Logger.log(newArray);
//  Logger.log(newArray[0].length);
  for (var k = 0; k < newArray.length; k ++) {
    if(newArray[k].length == 0){
      // array is empty 
      Logger.log(k + 'is empty'); 
    } else {
      //array not empty
      Logger.log(k + 'is not empty');
    }  
  }
}


function transposeArray(array){
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName('Input');
  var numColumns = spreadsheet.getLastColumn(); 
  var numRows = spreadsheet.getLastRow(); 
  var array = sheet.getRange(2, 6, numRows, numColumns).getValues();
  var result = [];
  for (var col = 0; col < array[0].length; col++) { // Loop over array cols
    result[col] = [];
    for (var row = 0; row < array.length; row++) { // Loop over array rows
      result[col][row] = array[row][col]; // Rotate
    }
  }
  return result;
}




//function transpose(a) {
//  var spreadsheet = SpreadsheetApp.getActive();
//  var sheet = spreadsheet.getSheetByName('Input');
//  var numColumns = spreadsheet.getLastColumn(); 
//  var numRows = spreadsheet.getLastRow(); 
//  var a = sheet.getRange(2, 6, numRows, numColumns).getValues();
//  return Object.keys(a[0]).map(function (c) { return a.map(function (r) { return r[c]; }); });
//}
