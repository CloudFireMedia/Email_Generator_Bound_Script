//function Show() {
//  var spreadsheet = SpreadsheetApp.getActive();
//  spreadsheet.getRange('F:F').activate();
//  spreadsheet.getActiveSheet().showColumns(6, 42);
//};
//
//function new1() {
//  var spreadsheet = SpreadsheetApp.getActive();
//  spreadsheet.getRange('F:F').activate();
//  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Format'), true);
//  spreadsheet.getRange('F:F').activate();
//  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Input'), true);
//  spreadsheet.getActiveSheet().insertColumnsBefore(spreadsheet.getActiveRange().getColumn(), 1);
//  spreadsheet.getActiveRange().offset(0, 0, spreadsheet.getActiveRange().getNumRows(), 1).activate();
//  spreadsheet.getRange('Format!F:F').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
//  spreadsheet.getRange('G:G').activate();
//  spreadsheet.getActiveSheet().hideColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
//  spreadsheet.getRange('F:F').activate();
//};
//
////function FormatOldEmails() {
////  var spreadsheet = SpreadsheetApp.getActive();
////  spreadsheet.getRange('G:G').activate();
////  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Input'), true);
////  spreadsheet.getRange('Format!G:G').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
////};
//
//
//
//function FormatOld2() {
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sheet = ss.getSheets()[0];
//  var numColumns = sheet.getLastColumn();  
//  var numRows = sheet.getLastRow();  
//  sheet.getRange(1,7,numRows,numColumns);
//  ss.setActiveSheet(ss.getSheetByName('Format'), true);
//  ss.setActiveSheet(ss.getSheetByName('Input'), true);
//  ss.getActiveRangeList().clearFormat();
//  ss.getRange('Format!G:G').copyTo(ss.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
//};
//
////var sheet = ss.getSheets()[0];
////var range = sheet.getRange(1, 1, 3, 3);