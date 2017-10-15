function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Vacancy Listing Functions')
      .addItem('Import Yardi Vacancy Report', 'importYardiReport')
      .addItem('Set Listing Status', 'setListingStatusBtn')
//      .addSeparator()
//      .addSubMenu(ui.createMenu('Sub-menu')
//          .addItem('Second item', 'menuItem2'))
      .addToUi();
}


function menuItem1() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('You clicked the first menu item!');
}

function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
     .alert('You clicked the second menu item!');
}

/**
 * This code finds the folder in Google Drive called 'Vacancy Reports' and the file 
 downloaded from Yardi called 'DataGridExport.xls'.  It then converts this file into 
 a Google sheet called 'DataGridExport'.  The code then trashes the original DataGridExport.xls file.
 **/
 function ConvertYardiExcel2Sheets() {
  var folderIncoming = DriveApp.getFoldersByName('Vacancy Reports');
  var xlsFile = folderIncoming.next().getFilesByName('DataGridExport.xlsx');
//  var xlsId = '0B4jTPPJvvrqFMzc5aTZpUkN3enc'; // ID of Excel file to convert
//  var xlsFile = DriveApp.getFileById(xlsId); // File instance of Excel file
  var xlsBlob = xlsFile.next().getBlob(); // Blob source of Excel file for conversion
     var xlsFilename='DataGridExport' // File name to give to converted file; defaults to same as source file
  var destFolders = []; // array of IDs of Drive folders to put converted file in; empty array = root folder
  var ss = convertExcel2Sheets(xlsBlob, xlsFilename, destFolders);
  Logger.log(ss.getId());
 DriveApp.getFileById(ss.getId());

 // add the files to the correct folder
      var filesToMove = DriveApp.getFilesByName('DataGridExport');
      
        var fileToMove = filesToMove.next();
        var dest_folder = DriveApp.getFolderById('0B4jTPPJvvrqFMzc5aTZpUkN3enc')
        dest_folder.addFile(fileToMove);
        //remove the copy of the ticket from the drive
        fileToMove.getParents().next().removeFile(fileToMove);
        var FileToTrash = DriveApp.getFilesByName('DataGridExport.xlsx');
        FileToTrash.next().setTrashed(true);
      }

/**
 * Convert Excel file to Sheets
 * @param {Blob} excelFile The Excel file blob data; Required
 * @param {String} filename File name on uploading drive; Required
 * @param {Array} arrParents Array of folder ids to put converted file in; Optional, will default to Drive root folder
 * @return {Spreadsheet} Converted Google Spreadsheet instance
 **/
function convertExcel2Sheets(excelFile, filename, arrParents) {
  
  var parents  = arrParents || []; // check if optional arrParents argument was provided, default to empty array if not
  if ( !parents.isArray ) parents = []; // make sure parents is an array, reset to empty array if not
  
  // Parameters for Drive API Simple Upload request (see https://developers.google.com/drive/web/manage-uploads#simple)
  var uploadParams = {
    method:'post',
    contentType: 'application/vnd.ms-excel', // works for both .xls and .xlsx files
    contentLength: excelFile.getBytes().length,
    headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
    payload: excelFile.getBytes()
  };
  
  // Upload file to Drive root folder and convert to Sheets
  var uploadResponse = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v2/files/?uploadType=media&convert=true', uploadParams);
    
  // Parse upload&convert response data (need this to be able to get id of converted sheet)
  var fileDataResponse = JSON.parse(uploadResponse.getContentText());

  // Create payload (body) data for updating converted file's name and parent folder(s)
  var payloadData = {
    title: filename, 
    parents: []
  };
  if ( parents.length ) { // Add provided parent folder(s) id(s) to payloadData, if any
    for ( var i=0; i<parents.length; i++ ) {
      try {
        var folder = DriveApp.getFolderById(parents[i]); // check that this folder id exists in drive and user can write to it
        payloadData.parents.push({id: parents[i]});
      }
      catch(e){} // fail silently if no such folder id exists in Drive
    }
  }
  // Parameters for Drive API File Update request (see https://developers.google.com/drive/v2/reference/files/update)
  var updateParams = {
    method:'put',
    headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
    contentType: 'application/json',
    payload: JSON.stringify(payloadData)
  };
  
  // Update metadata (filename and parent folder(s)) of converted sheet
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v2/files/'+fileDataResponse.id, updateParams);
  
  return SpreadsheetApp.openById(fileDataResponse.id);
}

function checkFileDownload(filename,foldername){

  var folder = DriveApp.getFoldersByName("Vacancy Reports");

  Logger.log(folder.hasNext());

  //Folder does not exist
  if(!folder.hasNext()){

  Logger.log("No Folder Found");

  }
  //Folder does exist
  else{
    Logger.log("Folder Found")
    var file   = folder.next().getFilesByName("DataGridExport.xlsx");
    if(!file.hasNext()){
       Logger.log("No File Found");
      throw ("No File Found!");
    }
    else{
       Logger.log("File Found");
    }
  }

}

function importYardiReport() {
//  var ui = SpreadsheetApp.getUi();
//  var response = ui.alert('Confirm Download of Report From Yardi', 'Have you first downloaded the vacancy report from Yardi into the Vacancy Reports Google Drive Folder and is it named DataGridExport.xlsx?', ui.ButtonSet.YES_NO);
//   if (response == ui.Button.YES) {
//   Logger.log("Yes");
// } else if (response == ui.Button.NO) {
//   Logger.log("No");
//   throw ui.alert('Please ensure you have downloaded the file with the correct name to the correct Google Drive Folder before continuing.')
//   return;
// } else {
//   Logger.log('Clicked X box');
//   throw ui.alert('Please ensure you have downloaded the file with the correct name to the correct Google Drive Folder before continuing.')
// }
 //The code below lists the the functions in the order they need to run
  var folder = DriveApp.getFoldersByName("Vacancy Reports");

  Logger.log(folder.hasNext());

  //Folder does not exist
  if(!folder.hasNext()){

  Logger.log("No Folder Found");
    throw ("No Folder Found! Did Someone Delete It???")

  }
  //Folder does exist
  else{
    Logger.log("Folder Found")
    var file   = folder.next().getFilesByName("DataGridExport.xlsx");
    if(!file.hasNext()){
       Logger.log("No File Found");
      throw ("No File Found! Download the Available Units report from Yardi into the Vacancy Reports Google Drive Folder First ");
    }
    else{
       Logger.log("File Found");

  
  ConvertYardiExcel2Sheets();
//  Utilities.sleep(5000);
  makeCopyOfPFIVacancyReferral();
  makeCopyOfPFIVacancyReport();
  mergeSheets();
  removeDuplicateRows();
  mergeDuplicateUnits();
  mergeSheets2();
  amenitiesStartAdsStatusPreservation();
  deleteSS();
  deleteExtraRows();
  mergeDuplicateChangedData();
  makeCopyOfCopyOfChangedData();
//  Utilities.sleep(1000);
  setUpdatedChangedRentStatus();
  updatedChangedRentDelete();
  resetUpdatedRentSwitch();
  pushRentedUnits();
  mergeDuplicatePushedRented();
  setListingStatus();
  unlistedRentedUnitDelete();
  copyOrigNewNoticeColumn();
  pushNewNotice();
  mergeDuplicatePushedNewNotices();
  setListingStatus();
  listedNewNoticesDelete();
  pushTwoWeekNotice();
  setListingStatus();
  resetFormatting();
  setFormatting();
  setNumberFormatting();
}
    }
  }    
//This is the end of the code



function makeCopyOfPFIVacancyReferral() {
var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PFI Vacancy Referral List");
var numRowsSource = sourceSheet.getLastRow();
var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Copy of PFI Vacancy Referral List");
var numRowsDest = destSheet.getLastRow(); 
var sourceRange = sourceSheet.getRange(2, 1, numRowsSource, 22);
var destRange = destSheet.getRange(2, 1, numRowsDest, 22);
  destRange.clearContent();
  destRange = destSheet.getRange(2, 1, numRowsSource, 22)
  sourceRange.copyTo(destRange);  
}

/**
 * This code finds and opens the newly converted file named 'DataGridExport'.  
 Inside this file, there is a sheet called 'Report1'. The code then finds the file 
 'PFI Vacancy Report' and opens it.  It then copies 'Report1' from 'DataGridExport' 
 into 'PFI Vacancy Report'. This copiedsheet is named 'Copy of Report1'.  
 The code then merges the data from 'Copy of Report1' into the sheet called 
 'PFI Vacancy Report'. It then resets the array formula on the Vacancy Referral List
 **/
function mergeSheets() {
  
 var sourceFile = DriveApp.getFilesByName('DataGridExport');
 var sourceFile2 = DriveApp.getFileById(sourceFile.next().getId());
 var openedSourceFile = SpreadsheetApp.openById(sourceFile2.getId());
 var sourceSS = openedSourceFile.getSheetByName('Report1');
 var destinationFile = DriveApp.getFilesByName('PFI Vacancy Report'); 
 var openedDestinationSS = SpreadsheetApp.openById(destinationFile.next().getId());
 var destinationSS = openedDestinationSS.getSheetByName('PFI Vacancy Report');
 var vacancyRefSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Vacancy Referral List');
 var arrayFormulaCell = vacancyRefSS.getRange('D2'); 
  sourceSS.copyTo(openedDestinationSS);
 var report1 = openedDestinationSS.getSheetByName('Copy of Report1');
   var range = report1.getRange(2,1);
  if(range.getValue() == '') {while(range == '')break;}else{ 
//  report1.hideSheet();
 var report1NumRows = report1.getLastRow();
 var reportRangeToCopy = report1.getRange(2, 1, report1NumRows-1, 10);
    Logger.clear(); 
Logger.log(reportRangeToCopy);  
 var pfiReportLastRow = openedDestinationSS.getSheetByName('PFI Vacancy Report').getLastRow()+ 1;
    Logger.clear(); 
Logger.log(pfiReportLastRow);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('PFI Vacancy Report');
  var cell = sheet.getRange("A2:J2");
  cell.setBorder(false,false,false,false,false,false);
  sheet.insertRowsBefore(2, report1NumRows-1);
  var rangeToPasteTo = sheet.getRange(2,1)
  reportRangeToCopy.copyTo(rangeToPasteTo);
  }
var FileToTrash = DriveApp.getFilesByName('DataGridExport');
        FileToTrash.next().setTrashed(true);
  arrayFormulaCell.setValue("=ARRAYFORMULA('PFI Vacancy Report'!A2:I)");
  //  removeDuplicateRows();
//  removeDuplicateRows();
//  mergeDuplicateUnits();
//  deleteSS();
  //  deleteExtraRows();

  }

function setNumberFormatting() {
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PFI Vacancy Report");
var range = sheet.getRange('A2:A');  
range.setNumberFormat("0000");  
}

function removeDuplicateRows() {
  var destinationFile = DriveApp.getFilesByName('PFI Vacancy Report'); 
  var openedDestinationSS = SpreadsheetApp.openById(destinationFile.next().getId());
  var data = openedDestinationSS.getSheetByName('PFI Vacancy Report').getDataRange().getValues();
  var newData = new Array();
  var rented = new Array();
  var newNivUnit = new Array();
  
  var warning_count = 0;
  var warning_count1 = 0;
  var msg = "";  
  var msg1 = "";
  
  for(i in data){

var row = data[i];

var duplicate = false;

for(j in newData){
  
  if(row[0] == newData[j][0] && row[1] == newData[j][1] && row[2] == newData[j][2] && row[3] == newData[j][3] && row[4] == newData[j][4] && row[5] == newData[j][5] && row[6] == newData[j][6] && row[7] == newData[j][7] && row[8] == newData[j][8] && row[9] == newData[j][9] && row[10] == newData[j][10] && row[11] == newData[j][11] && row[12] == newData[j][12]){
   duplicate = true; 
  }

//if(row.join() == newData[j].join()){
//
//duplicate = true;
//
//}

}

  
    
if(!duplicate){

newData.push(row);

}

}

//if(warning_count || warning_count1) {
//      MailApp.sendEmail("pfigregg@pisf.com", 
//        "Rented Units Or No Longer Available", msg + msg1);
//  }
Logger.clear(); 
Logger.log(newData[i]);
var checkLogValue = Logger.getLog();
openedDestinationSS.getSheetByName('Variables').getRange("A4").setValue(checkLogValue);
var checkLog = openedDestinationSS.getSheetByName('Variables').getRange("B4");   
openedDestinationSS.getSheetByName('PFI Vacancy Report').clearContents();
//if(checkLog.getValue() == "undefined") {Logger.log("Empty")}else{  
openedDestinationSS.getSheetByName('PFI Vacancy Report').getRange(1, 1, newData.length, newData[0].length).setValues(newData);
//removeDuplicateRows2()
} 

function mergeDuplicateUnits() {
  var destinationFile = DriveApp.getFilesByName('PFI Vacancy Report'); 
   var changedDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ChangedData");
  var copyOfChangedDataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Copy of ChangedData");
  var targetCell = copyOfChangedDataSheet.getRange("D2")
  var openedDestinationSS = SpreadsheetApp.openById(destinationFile.next().getId());
  var data = openedDestinationSS.getSheetByName('PFI Vacancy Report').getDataRange().getValues();
  var newData = new Array();
  var changedData = new Array();
  
  for(i in data){

var row = data[i];

var duplicate = false;

for(j in newData){

if(row[0] == newData[j][0] && row[1] == newData[j][1]){
  duplicate = true;
changedData.push(row);

}

}

if(!duplicate){

newData.push(row);

}

}
Logger.clear();   
Logger.log(changedData);
var checkLogValue = Logger.getLog();
openedDestinationSS.getSheetByName('Variables').getRange("A2").setValue(checkLogValue);
var checkLog = openedDestinationSS.getSheetByName('Variables').getRange("B2");  
openedDestinationSS.getSheetByName('PFI Vacancy Report').clearContents();

openedDestinationSS.getSheetByName('PFI Vacancy Report').getRange(1, 1, newData.length, newData[0].length).setValues(newData);
  
  if(checkLog.getValue() == "undefined") {Logger.log("Empty");
//var numRows = SpreadsheetApp.getActive().getSheetByName('ChangedData').getDataRange().getNumRows();
//var numCols = SpreadsheetApp.getActive().getSheetByName('ChangedData').getDataRange().getNumColumns();
//SpreadsheetApp.getActive().getSheetByName('ChangedData').getRange(2, 1,numRows,numCols).clear();
                                         }else{ 
changedDataSheet.insertRowsBefore(2,changedData.length); 
SpreadsheetApp.getActive().getSheetByName('ChangedData').getRange(2, 1, changedData.length, changedData[0].length).clearContent(); 
SpreadsheetApp.getActive().getSheetByName('ChangedData').getRange(2, 1, changedData.length, changedData[0].length).setValues(changedData);     
targetCell.setValue("=ARRAYFORMULA(ChangedData!A2:I)")

  } 
}

function mergeSheets2() {
var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Copy of Report1");
var numRowsSource = sourceSheet.getLastRow();
var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PFI Vacancy Report");
var numRowsDest = destSheet.getLastRow(); 
//var varSheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Variables");
//var data = sourceSheet.getDataRange().getValues();
//var copiedData = new Array();
//var lastRow = destSheet.getLastRow();
//var lastCol = destSheet.getLastColumn();
//destSheet.getRange(2, 1, lastRow, lastCol).clearContent();   
var sourceRange = sourceSheet.getRange(2, 1, numRowsSource, 10);
var destRange = destSheet.getRange(2, 1, numRowsDest, 10);
  destRange.clearContent();
  destRange = destSheet.getRange(2, 1, numRowsSource, 10)
  sourceRange.copyTo(destRange);
  var column = destSheet.getRange("B2:B");
 
// Simple date format
column.setNumberFormat("0000");
}


function amenitiesStartAdsStatusPreservation() {
var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PFI Vacancy Referral List");
var numRowsSource = sourceSheet.getLastRow();
var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PFI Vacancy Referral List");
var numRowsDest = destSheet.getLastRow(); 
var sourceRange = sourceSheet.getRange(2, 23, numRowsSource, 2);
var destRange = destSheet.getRange(2, 21, numRowsDest, 2);
  destRange.clearContent();
  destRange = destSheet.getRange(2, 21, numRowsSource, 2)
  sourceRange.copyTo((destRange),{contentsOnly:true}) ;
}



function deleteSS() {
 var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  ss.getSheetByName('Copy of Report1').activate();
  ss.deleteActiveSheet();
}

function deleteExtraRows() {
        var sheet = SpreadsheetApp.getActive().getSheetByName('PFI Vacancy Report');
        var sheet2 = SpreadsheetApp.getActive().getSheetByName('ChangedData');
        var start=700;
        var start2=50;
        var end=sheet.getLastRow();
  var end2=sheet.getMaxRows() - 1;
  var end3=end2-start;
  var end4=sheet2.getMaxRows() - 1;
  var end5=end4-start2;
  Logger.clear(); 
  Logger.log(end2);
  
  
  if(end2 == 700) {while (end2 < 702)break;}else{
    sheet.deleteRows(start, end3);}
  end=sheet2.getLastRow();
  end2=sheet2.getMaxRows() - 1;
  end3=end2-start;
  if(end4 <= 50) {while (end4 < 52)break;}else{
    sheet2.deleteRows(start2, end5);}
}

function mergeDuplicateChangedData() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("ChangedData");
  var data = sheet.getDataRange().getValues();
  var newData = new Array();
  
  for(i in data){

var row = data[i];

var duplicate = false;

for(j in newData){

if(row[0] == newData[j][0] && row[1] == newData[j][1]){
  duplicate = true;
//changedData.push(row);

}

}

if(!duplicate){

newData.push(row);

}

}

sheet.clearContents();
sheet.getRange(1, 1, newData.length, newData[0].length) .setValues(newData); 
Utilities.sleep(1000);
//makeCopyOfCopyOfChangedData(); 
}

function setUpdatedChangedRentStatus() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("ListingManager");
  var ss = SpreadsheetApp.getActive().getSheetByName("Variables");
  var rows = sheet.getDataRange();
  var data = sheet.getDataRange().getValues();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var array = [];
  var lastRow = sheet.getLastRow();
//  var listingDataCol = sheet.getRange(26, column, numRows, numColumns)
//  data.forEach(function(row){
  for (var i=0; i < data.length; i++) {
//    var row = values[i];

    if (data [i][19]!="") {
      if (data [i][19] == data [i][7])
        array.push("Yes")
    }
    Logger.clear(); 
    Logger.log(array);
    if (array != "") {
      sheet.getRange(i+1,30).setValue(array);
    }
    array=[];

    }
}

function makeCopyOfCopyOfChangedData() {
var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Copy of ChangedData");
var numRowsSource = sourceSheet.getLastRow();
var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Copy2 of ChangedData");
var numRowsDest = destSheet.getLastRow(); 
var varSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Variables");  
var data = sourceSheet.getDataRange().getValues(); 
var copiedData = new Array();
var lastRow = destSheet.getLastRow();
var lastCol = destSheet.getLastColumn(); 
destSheet.getRange(2, 1, lastRow, lastCol).clearContent(); 

for(i in data){

var row = data[i];
    if(row[0] != '' && row[18] == "Yes" && row[17] != ""){
copiedData.push(row);

    }
}
Logger.clear(); 
Logger.log(copiedData);
var checkLogValue = Logger.getLog(); 
SpreadsheetApp.getActive().getSheetByName("Variables").getRange("A2").clear();  
SpreadsheetApp.getActive().getSheetByName("Variables").getRange("A2").setValue(checkLogValue);
var checkLog = SpreadsheetApp.getActive().getSheetByName("Variables").getRange("B2"); 
if(checkLog.getValue() == "undefined") {Logger.log("Empty")}else{  
// newBonuses.insertRowsBefore(2, newBonus.length-1); 
 
//SpreadsheetApp.getActive().getSheetByName('NewBonusesImport').getRange(2, 1, listedNewBonus.length, listedNewBonus[0].length).clearContent(); 
destSheet.getRange(2, 1, copiedData.length, copiedData[0].length).setValues(copiedData);

  }   
}

//This code is checking if the changed data (rents) have been updated on the Copy of ChangedData Sheet.  If they are not updated
//the code copies the row of the unit into an array and overwrites the data in the ChangedRent Page. 
function updatedChangedRentDelete() {
var sheet = SpreadsheetApp.getActive().getSheetByName("Copy2 of ChangedData");
var deleteFromSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Copy2 of ChangedData");
var varSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Variables");  
var data = sheet.getDataRange().getValues();
var deleteData = deleteFromSheet.getDataRange().getValues();
var changedRent = new Array();
var deleteRow = new Array();  
  
  for(i in data){

var row = data[i];
var deleteRow = deleteData[i];    
    if(row[0] != '' && row[16] != "Yes" && row[16] != "Rent Updated?" && row[16] != "Rented"){
changedRent.push(row);

    }
}
Logger.clear();  
Logger.log(changedRent);
var checkLogValue = Logger.getLog(); 
SpreadsheetApp.getActive().getSheetByName("Variables").getRange("A2").clear();  
SpreadsheetApp.getActive().getSheetByName("Variables").getRange("A2").setValue(checkLogValue);
var checkLog = SpreadsheetApp.getActive().getSheetByName("Variables").getRange("B2"); 
if(checkLog.getValue() == "undefined") {Logger.log("Empty")
var lastRow = deleteFromSheet.getLastRow();
var lastCol = deleteFromSheet.getLastColumn();
deleteFromSheet.getRange(2, 1, lastRow, lastCol).clearContent();
var numRows = SpreadsheetApp.getActive().getSheetByName('ChangedData').getDataRange().getNumRows();
var numCols = SpreadsheetApp.getActive().getSheetByName('ChangedData').getDataRange().getNumColumns();
SpreadsheetApp.getActive().getSheetByName('ChangedData').getRange(2, 1,numRows,numCols).clear();                                        
                                       }else{  
// newBonuses.insertRowsBefore(2, newBonus.length-1); 
var lastRow = deleteFromSheet.getLastRow();
var lastCol = deleteFromSheet.getLastColumn();  
//SpreadsheetApp.getActive().getSheetByName('NewBonusesImport').getRange(2, 1, listedNewBonus.length, listedNewBonus[0].length).clearContent(); 
deleteFromSheet.getRange(2, 1, lastRow, lastCol).clearContent(); 
deleteFromSheet.getRange(2, 1, changedRent.length, changedRent[0].length).setValues(changedRent);

  }   
}


function resetUpdatedRentSwitch() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("ListingManager");
  var ss = SpreadsheetApp.getActive().getSheetByName("Variables");
  var rows = sheet.getDataRange();
  var data = sheet.getDataRange().getValues();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var array = [];
  var lastRow = sheet.getLastRow();
//  var listingDataCol = sheet.getRange(26, column, numRows, numColumns)
//  data.forEach(function(row){
  for (var i=0; i < data.length; i++) {
//    var row = values[i];

    if (data [i][29]!="" && data [i][19]=="") {
         array.push("reset") 
//             Logger.log([i]);
    }
    Logger.clear(); 
    Logger.log(array);
    if (array != "") {
      sheet.getRange(i+1,30).setValue("");
    }
    array=[];

    }
}

function pushRentedUnits() {
var sheet = SpreadsheetApp.getActive().getSheetByName("Copy of PFI Vacancy Report");
var sheetRented = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RentedImport");
var range = sheetRented.getRange(2, 1);
var data = sheet.getDataRange().getValues();
var rented = new Array();
  
  for(i in data){

var row = data[i];
//Logger.log(row[15]);    

if(row[0] != "Unit ID" && row[15] == "Rented"){
//rented.push(row);
Logger.clear();   
Logger.log(row);
rented.push(row);
//Logger.log(rented);
}
}
Logger.clear();   
Logger.log(rented);
var checkLogValue = Logger.getLog();
SpreadsheetApp.getActive().getSheetByName("Variables").getRange("A2").clear();
SpreadsheetApp.getActive().getSheetByName("Variables").getRange("A2").setValue(checkLogValue);
var checkLog = SpreadsheetApp.getActive().getSheetByName("Variables").getRange("B2"); 
if(checkLog.getValue() == "undefined") {Logger.log("Empty")}else{
  if(range.getValue() == '') {while(range == '')break;}else{
    if(rented.length-1 != 0) {
      sheetRented.insertRowsBefore(2, rented.length-1);}else{
      sheetRented.insertRowsBefore(2, rented.length)
      }} 
SpreadsheetApp.getActive().getSheetByName('RentedImport').getRange(2, 1, rented.length, rented[0].length).clearContent(); 
  
SpreadsheetApp.getActive().getSheetByName('RentedImport').getRange(2, 1, rented.length, rented[0].length).setValues(rented);  
  }   
}

function mergeDuplicatePushedRented() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("RentedImport");
  var data = sheet.getDataRange().getValues();
  var newData = new Array();
  
  for(i in data){

var row = data[i];

var duplicate = false;

for(j in newData){

if(row[0] == newData[j][0] && row[1] == newData[j][1]){
  duplicate = true;
//changedData.push(row);

}

}

if(!duplicate){

newData.push(row);

}

}

sheet.clearContents();
sheet.getRange(1, 1, newData.length, newData[0].length) .setValues(newData); 
 
}

function makeCopyOfPFIVacancyReport() {
var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PFI Vacancy Report");
var numRowsSource = sourceSheet.getLastRow();
var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Copy of PFI Vacancy Report");
var numRowsDest = destSheet.getLastRow();
var vacancyRefSS = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Vacancy Referral List');
var arrayFormulaCell = vacancyRefSS.getRange('M2');   
var arrayFormulaCell2 = destSheet.getRange('K2');   
//var varSheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Variables");
//var data = sourceSheet.getDataRange().getValues();
//var copiedData = new Array();
//var lastRow = destSheet.getLastRow();
//var lastCol = destSheet.getLastColumn();
//destSheet.getRange(2, 1, lastRow, lastCol).clearContent();   
var sourceRange = sourceSheet.getRange(2, 1, numRowsSource, 9);
var destRange = destSheet.getRange(2, 2, numRowsDest, 9);
  destRange.clearContent();
  destRange = destSheet.getRange(2, 2, numRowsSource, 9)
  sourceRange.copyTo(destRange);
  arrayFormulaCell.setValue('=ArrayFormula(if(A2:A="","",iferror(vlookup(A2:A,UnitsAssetsManager!A:O,{15}*row(A2:A)^0,0))))');
  arrayFormulaCell2.setValue('=ArrayFormula(if(A2:A="","",iferror(vlookup(A2:A,UnitsAssetsManager!A:O,{15}*row(A2:A)^0,0))))');
}

//This code is checking the listing manager row by row to see if a unit is listed and where.  It will 
// then push the row to array if listed and set the value to the same row in listing status column
function setListingStatus() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("ListingManager");
  var ss = SpreadsheetApp.getActive().getSheetByName("Variables");
  var rows = sheet.getDataRange();
  var data = sheet.getDataRange().getValues();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var array = [];
  var lastRow = sheet.getLastRow();
//  var listingDataCol = sheet.getRange(26, column, numRows, numColumns)
//  data.forEach(function(row){
  for (var i=0; i < data.length; i++) {
//    var row = values[i];

    if (data [i][30]=="Yes") {
         array.push("Craigslist") 
//             Logger.log([i]);
    }
    if (data [i][31]=="Yes") {
          array.push(" Zillow ")
    }
     if (data [i][32]=="Yes") {
           array.push(" Apartments.com")
     }
      if (data [i][33]=="Yes") {
            array.push(" Adwords")
      }
     if (data [i][34]=="Yes") {
             array.push(" Facebook")
     }
      if (data [i][35]=="Yes") {
              array.push(" Other PPC")
      }
     if (data [i][36]=="Yes") {
                array.push(" Website(s)")
     }
    Logger.clear(); 
    Logger.log(array);
    if (array != "") {
      sheet.getRange(i+1,27).setValue([array].join());
    }else{sheet.getRange(i+1,27).setValue("");
         sheet.getRange(1,27).setValue("Listing Status");
         }
    array=[];

    }
}

//This code is checking the listing manager row by row to see if a unit is listed and where.  It will 
// then push the row to array if listed and set the value to the same row in listing status column
function setListingStatusBtn() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("ListingManager");
  var ss = SpreadsheetApp.getActive().getSheetByName("Variables");
  var rows = sheet.getDataRange();
  var data = sheet.getDataRange().getValues();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  var array = [];
  var lastRow = sheet.getLastRow();
//  var listingDataCol = sheet.getRange(26, column, numRows, numColumns)
//  data.forEach(function(row){
  for (var i=0; i < data.length; i++) {
//    var row = values[i];

    if (data [i][30]=="Yes") {
         array.push("Craigslist") 
//             Logger.log([i]);
    }
    if (data [i][31]=="Yes") {
          array.push(" Zillow ")
    }
     if (data [i][32]=="Yes") {
           array.push(" Apartments.com")
     }
      if (data [i][33]=="Yes") {
            array.push(" Adwords")
      }
     if (data [i][34]=="Yes") {
             array.push(" Facebook")
     }
      if (data [i][35]=="Yes") {
              array.push(" Other PPC")
      }
     if (data [i][36]=="Yes") {
                array.push(" Website(s)")
     }
    Logger.clear(); 
    Logger.log(array);
    if (array != "") {
      sheet.getRange(i+1,27).setValue([array].join());
    }else{sheet.getRange(i+1,27).setValue("");
         sheet.getRange(1,27).setValue("Listing Status");
         }
    array=[];

    }
  resetFormatting();
  setFormatting();
}

function resetFormatting() {
var sheet = SpreadsheetApp.getActive().getSheetByName("ListingManager");
var rows = sheet.getDataRange();
var data = sheet.getDataRange().getValues();
var numRows = rows.getNumRows();
var numCols = rows.getNumColumns();
sheet.getRange(2, 1, numRows, numCols).setBackground(null);  
}


function setFormatting() {
var sheet = SpreadsheetApp.getActive().getSheetByName("ListingManager");
var rows = sheet.getDataRange();
var data = sheet.getDataRange().getValues();
var numRows = rows.getNumRows();
var numCols = rows.getNumColumns();
var values = rows.getValues();
var startAdsArray = [];
var bonusArray = [];
var changedRentArray = [];
var changedBonusArray = [];
var rentedArray = [];
var removedBonusArray = [];
var listedNoStartArray = [];
var lastRow = sheet.getLastRow();  
var green = "#d9ead3"
var blue = "#c9daf8"
var gray = "#d9d9d9"
var yellow = "#fff2cc"
var orange =  "#fce5cd"
var red = "#e6b8af"
  for (var i=0; i < data.length; i++) {
    //start ads?
    if (data [i][20]=="Yes" && data [i][0] != "Unique Id") {
         startAdsArray.push("format") 
//         Logger.log([i]);
    }
    //bonus
    if (data [i][14]!="" && data [i][0] != "Unique Id") {
          bonusArray.push("format")
    }
    //changed rent
     if (data [i][19]!="" && data [i][0] != "Unique Id") {
           changedRentArray.push("format")
     }
    //changed bonus
      if (data [i][18]!="" && data [i][0] != "Unique Id") {
            changedBonusArray.push("format")
      }
    //rented
     if (data [i][13]!="" && data [i][0] != "Unique Id") {
             rentedArray.push("format")
     }
    //removed bonus
      if (data [i][16]!="" && data [i][0] != "Unique Id") {
              removedBonusArray.push("format")
      }
    //listed but no start ads
     if (data [i][20]!="Yes" && data [i][26] !="" && data [i][0] != "Unique Id") {
                listedNoStartArray.push("format")
     }
    
    if (startAdsArray != "") {
      sheet.getRange(i+1,1,1,numCols).setBackground(green);
    }
//    else{sheet.getRange(i+1,1,1,numCols).setBackground(null);
//         }
    startAdsArray=[];
    
    if (bonusArray != "") {
      sheet.getRange(i+1,1,1,numCols).setBackground(blue);
    }
//    else{sheet.getRange(i+1,1,1,numCols).setBackground(null);
//         }
    bonusArray=[];
    
    if (changedBonusArray != "") {
      sheet.getRange(i+1,1,1,numCols).setBackground(gray);
    }
//    else{sheet.getRange(i+1,1,1,numCols).setBackground(null);
//         }
    changedBonusArray=[];
    
    if (changedRentArray != "") {
      sheet.getRange(i+1,1,1,numCols).setBackground(red);
    }
//    else{sheet.getRange(i+1,1,1,numCols).setBackground(null);
//         }
    changedRentArray=[];
    
    if (listedNoStartArray != "") {
      sheet.getRange(i+1,1,1,numCols).setBackground(yellow);
    }
//    else{sheet.getRange(i+1,1,1,numCols).setBackground(null);
//         }
    listedNoStartArray=[];
    
    if (removedBonusArray != "") {
      sheet.getRange(i+1,1,1,numCols).setBackground(orange);
    }
//    else{sheet.getRange(i+1,1,1,numCols).setBackground(null);
//         }
    removedBonusArray=[];    
    
    if (rentedArray != "") {
      sheet.getRange(i+1,1,1,numCols).setBackground(yellow);
    }
//    else{sheet.getRange(i+1,1,1,numCols).setBackground(null);
//         }
    rentedArray=[];

    }
}

//This code is checking if rented units have been removed from listed on the Copy of RentedImport Sheet.  If they are still listed
//the code copies the row of the unit to the RentedImport Page (replaces content). 
function unlistedRentedUnitDelete() {
var sheet = SpreadsheetApp.getActive().getSheetByName("Copy of RentedImport");
var deleteFromSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("RentedImport");
var varSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Variables");  
var data = sheet.getDataRange().getValues();
var deleteData = deleteFromSheet.getDataRange().getValues();
var unlistedRentedUnit = new Array();
  
  for(i in data){

var row = data[i];
var deleteRow = deleteData[i];    
//Logger.log(row[15]);    

    if(row[0] != '' && row[16] != "" && row[16] != "Listing Status"){
unlistedRentedUnit.push(row);

    }
//    else{
//      if(row[16] != "Listing Status" && row[12] ==""){
//        unlistedRentedUnit.push(row);
//      }
//    }
}
Logger.clear();   
Logger.log(unlistedRentedUnit);
var checkLogValue = Logger.getLog(); 
SpreadsheetApp.getActive().getSheetByName("Variables").getRange("A2").clear();  
SpreadsheetApp.getActive().getSheetByName("Variables").getRange("A2").setValue(checkLogValue);
var checkLog = SpreadsheetApp.getActive().getSheetByName("Variables").getRange("B2"); 
if(checkLog.getValue() == "undefined") {Logger.log("Empty");
var lastRow = deleteFromSheet.getLastRow();
var lastCol = deleteFromSheet.getLastColumn();  
SpreadsheetApp.getActive().getSheetByName('RentedImport').getRange(2, 1, lastRow, lastCol).clearContent(); 
                                       }else{  
// newBonuses.insertRowsBefore(2, newBonus.length-1); 
var lastRow = deleteFromSheet.getLastRow();
var lastCol = deleteFromSheet.getLastColumn();  
//SpreadsheetApp.getActive().getSheetByName('RentedImport').getRange(2, 1, unlistedRentedUnit.length, unlistedRentedUnit[0].length).clearContent(); 
SpreadsheetApp.getActive().getSheetByName('RentedImport').getRange(2, 1, lastRow, lastCol).clearContent(); 
SpreadsheetApp.getActive().getSheetByName('RentedImport').getRange(2, 1, unlistedRentedUnit.length, unlistedRentedUnit[0].length).setValues(unlistedRentedUnit);
  }   
}

function copyOrigNewNoticeColumn() {
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Vacancy Referral List");
var numRows = sheet.getLastRow();
var origNewNoticeCol = sheet.getRange(2, 17, numRows, 1);
var copyNewNoticeCol = sheet.getRange(2, 18, numRows, 1);
copyNewNoticeCol.setValues(origNewNoticeCol.getValues());
}  

function pushNewNotice() {
var sheet = SpreadsheetApp.getActive().getSheetByName("Vacancy Referral List");
var newNotices = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NewNoticesImport");
var range = newNotices.getRange(2, 1);
var data = sheet.getDataRange().getValues();
var newNotice = new Array();
  
  for(i in data){

var row = data[i];
//Logger.log(row[15]);    

if(row[17] == "New Notice"){
//rented.push(row);
Logger.clear();   
Logger.log(row);
newNotice.push(row);
//Logger.log(rented);
}
}
Logger.clear();   
Logger.log(newNotice);
var checkLogValue = Logger.getLog(); 
SpreadsheetApp.getActive().getSheetByName("Variables").getRange("A2").setValue(checkLogValue);
var checkLog = SpreadsheetApp.getActive().getSheetByName("Variables").getRange("B2"); 
var checkLog2 = SpreadsheetApp.getActive().getSheetByName("Variables").getRange("C2");  
if(checkLog.getValue() == "undefined") {
  Logger.log("Empty")
  if(checkLog2.getValue() == "empty array") {
   Logger.log("Empty Array"); 
  }
}else{ 
  if(range.getValue() == '') {while(range == '')break;}else{
    if(newNotice.length-1 != 0) {
      newNotices.insertRowsBefore(2, newNotice.length-1)}else {
      newNotices.insertRowsBefore(2, newNotice.length)
      }}  
SpreadsheetApp.getActive().getSheetByName('NewNoticesImport').getRange(2, 1, newNotice.length, newNotice[0].length).clearContent(); 
  
SpreadsheetApp.getActive().getSheetByName('NewNoticesImport').getRange(2, 1, newNotice.length, newNotice[0].length).setValues(newNotice);
  }   
}

function mergeDuplicatePushedNewNotices() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("NewNoticesImport");
  var data = sheet.getDataRange().getValues();
  var newData = new Array();
  
  for(i in data){

var row = data[i];

var duplicate = false;

for(j in newData){

if(row[0] == newData[j][0] && row[1] == newData[j][1]){
  duplicate = true;
//changedData.push(row);

}

}

if(!duplicate){

newData.push(row);

}

}

sheet.clearContents();
sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData); 
 
}

//This code is checking if the new notices have been listed on the Copy of NewNoticesImport Sheet.  If they are not listed
//the code copies the row of the unit to the NewNoticesImport Page. 
function listedNewNoticesDelete() {
var sheet = SpreadsheetApp.getActive().getSheetByName("Copy of NewNoticesImport");
var deleteFromSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("NewNoticesImport");
var varSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Variables");  
var data = sheet.getDataRange().getValues();
var deleteData = deleteFromSheet.getDataRange().getValues();
var listedNotice = new Array();
  
  for(i in data){

var row = data[i];
var deleteRow = deleteData[i];    
//Logger.log(row[15]);    

    if(row[0] != '' && row[18] == "" && row[18] != "Listing Status"){
listedNotice.push(row);

}
}
Logger.clear();   
Logger.log(listedNotice);
var checkLogValue = Logger.getLog(); 
SpreadsheetApp.getActive().getSheetByName("Variables").getRange("A2").setValue(checkLogValue);
var checkLog = SpreadsheetApp.getActive().getSheetByName("Variables").getRange("B2"); 
var checkLog2 = SpreadsheetApp.getActive().getSheetByName("Variables").getRange("C2");  
if(checkLog.getValue() == "undefined") {
  Logger.log("Empty"); 
if(checkLog2.getValue() == "empty array") {
  Logger.log("true");
// newBonuses.insertRowsBefore(2, newBonus.length-1); 
var lastRow = deleteFromSheet.getLastRow();
var lastCol = deleteFromSheet.getLastColumn();  
//SpreadsheetApp.getActive().getSheetByName('NewBonusesImport').getRange(2, 1, listedNewBonus.length, listedNewBonus[0].length).clearContent(); 
SpreadsheetApp.getActive().getSheetByName('NewNoticesImport').getRange(2, 1, lastRow, lastCol).clearContent();   
                                        }
                                       }else{  
// newBonuses.insertRowsBefore(2, newBonus.length-1); 
var lastRow = deleteFromSheet.getLastRow();
var lastCol = deleteFromSheet.getLastColumn();  
//SpreadsheetApp.getActive().getSheetByName('NewBonusesImport').getRange(2, 1, listedNewBonus.length, listedNewBonus[0].length).clearContent(); 
SpreadsheetApp.getActive().getSheetByName('NewNoticesImport').getRange(2, 1, lastRow, lastCol).clearContent(); 
SpreadsheetApp.getActive().getSheetByName('NewNoticesImport').getRange(2, 1, listedNotice.length, listedNotice[0].length).setValues(listedNotice);
  }   
}

function pushTwoWeekNotice() {
var sheet = SpreadsheetApp.getActive().getSheetByName("Vacancy Referral List");
var newNotices = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("TwoWeekNoticeImport");
var data = sheet.getDataRange().getValues();
var today = sheet.getRange("T2")
var todayValue = today.getValue();
var twoWeekNotice = new Array();
  
  for(i in data){

var row = data[i];
//Logger.log(row[18]);    

//if date available is greater than today and less than today plus 14 days do something with the data
if(row[18] >= todayValue && row[18] < todayValue + 14){
//rented.push(row);
//Logger.log(row);
twoWeekNotice.push(row);
//Logger.log(rented);
}
}
Logger.clear();   
Logger.log(twoWeekNotice);
var checkLogValue = Logger.getLog(); 
SpreadsheetApp.getActive().getSheetByName("Variables").getRange("A2").setValue(checkLogValue);
var checkLog = SpreadsheetApp.getActive().getSheetByName("Variables").getRange("B2"); 
if(checkLog.getValue() == "undefined") {Logger.log("Empty")}else{  
  
SpreadsheetApp.getActive().getSheetByName('TwoWeekNoticeImport').getRange(2, 1, twoWeekNotice.length, twoWeekNotice[0].length).clearContent(); 
  
SpreadsheetApp.getActive().getSheetByName('TwoWeekNoticeImport').getRange(2, 1, twoWeekNotice.length, twoWeekNotice[0].length).setValues(twoWeekNotice);
  }   
}


function exportChangedRentNotifications() {
  // Set the Active Spreadsheet so we don't forget
  var originalSpreadsheet = SpreadsheetApp.getActive().getSheetByName("Copy of ChangedData");
  var sheetName = originalSpreadsheet.getSheetName();
  
  // Set the message to attach to the email.
  var message = SpreadsheetApp.getActive().getSheetByName("AutoEmails").getRange("F3");
  
//  // Construct the Subject Line
  var subject = SpreadsheetApp.getActive().getSheetByName("AutoEmails").getRange("E3");

      
  // Get contact details from "Contacts" sheet and construct To: Header
  // Would be nice to include "Name" as well, to make contacts look prettier, one day.
  var contacts = SpreadsheetApp.getActive().getSheetByName("AutoEmails");
  var numRows = contacts.getLastRow();
  var emailTo = contacts.getRange(3, 4, numRows, 1).getValues();

  // Google scripts can't export just one Sheet from a Spreadsheet
  // So we have this disgusting hack

  // Create a new Spreadsheet and copy the current sheet into it.
  var newSpreadsheet = SpreadsheetApp.create("Spreadsheet to export");
  var originalSpreadsheetCopy = SpreadsheetApp.getActive().getSheetByName("ChangedRent");
  var origSSLRow = originalSpreadsheet.getLastRow();
  var origSSLCol = originalSpreadsheet.getLastColumn();
  var origSSCopyLRow = originalSpreadsheetCopy.getLastRow();
  var origSSCopyLCol = originalSpreadsheetCopy.getLastColumn(); 
  var range = originalSpreadsheetCopy.getRange(2,1);
  range.clearContent();
  originalSpreadsheet.getRange(2,1,origSSLRow,origSSLCol).copyTo(range,{contentsOnly:true});
  originalSpreadsheetCopy.copyTo(newSpreadsheet);


  // Find and delete the default "Sheet 1", after the copy to avoid triggering an apocalypse
  newSpreadsheet.getSheetByName('Sheet1').activate();
  newSpreadsheet.deleteActiveSheet();
  
  // Make zee PDF, currently called "Weekly status.pdf"
  // When I'm smart, filename will include a date and project name
  var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();
  var attach = {fileName:sheetName+'.pdf',content:pdf, mimeType:'application/pdf'};

  // Send the freshly constructed email 
  MailApp.sendEmail(emailTo, subject.getValue(), message.getValue(), {attachments:[attach]});
  
  // Delete the wasted sheet we created, so our Drive stays tidy.
  DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);  
}

function exportRentedNotifications() {
  // Set the Active Spreadsheet so we don't forget
  var originalSpreadsheet = SpreadsheetApp.getActive().getSheetByName("Rented");
  var sheetName = originalSpreadsheet.getSheetName();
  
  // Set the message to attach to the email.
  var message = SpreadsheetApp.getActive().getSheetByName("AutoEmails").getRange("I3");
  
//  // Construct the Subject Line
  var subject = SpreadsheetApp.getActive().getSheetByName("AutoEmails").getRange("H3");

      
  // Get contact details from "Contacts" sheet and construct To: Header
  // Would be nice to include "Name" as well, to make contacts look prettier, one day.
  var contacts = SpreadsheetApp.getActive().getSheetByName("AutoEmails");
  var numRows = contacts.getLastRow();
  var emailTo = contacts.getRange(3, 7, numRows, 1).getValues();

  // Google scripts can't export just one Sheet from a Spreadsheet
  // So we have this disgusting hack

  // Create a new Spreadsheet and copy the current sheet into it.
  var newSpreadsheet = SpreadsheetApp.create("Spreadsheet to export");
  var originalSpreadsheetCopy = SpreadsheetApp.getActive().getSheetByName("Copy of Rented");
  var origSSLRow = originalSpreadsheet.getLastRow();
  var origSSLCol = originalSpreadsheet.getLastColumn();
  var origSSCopyLRow = originalSpreadsheetCopy.getLastRow();
  var origSSCopyLCol = originalSpreadsheetCopy.getLastColumn(); 
  var range = originalSpreadsheetCopy.getRange(2,1);
  range.clearContent();
  originalSpreadsheet.getRange(2,1,origSSLRow,origSSLCol).copyTo(range,{contentsOnly:true});
  originalSpreadsheetCopy.copyTo(newSpreadsheet);


  // Find and delete the default "Sheet 1", after the copy to avoid triggering an apocalypse
  newSpreadsheet.getSheetByName('Sheet1').activate();
  newSpreadsheet.deleteActiveSheet();
  
  // Make zee PDF, currently called "Weekly status.pdf"
  // When I'm smart, filename will include a date and project name
  var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();
  var attach = {fileName:sheetName+'.pdf',content:pdf, mimeType:'application/pdf'};

  // Send the freshly constructed email 
  MailApp.sendEmail(emailTo, subject.getValue(), message.getValue(), {attachments:[attach]});
  
  // Delete the wasted sheet we created, so our Drive stays tidy.
  DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);  
}


function exportNewNoticeNotifications() {
  // Set the Active Spreadsheet so we don't forget
  var originalSpreadsheet = SpreadsheetApp.getActive().getSheetByName("NewNoticesImport");
  var sheetName = originalSpreadsheet.getSheetName();
  
  // Set the message to attach to the email.
  var message = SpreadsheetApp.getActive().getSheetByName("AutoEmails").getRange("C3");
  
//  // Construct the Subject Line
  var subject = SpreadsheetApp.getActive().getSheetByName("AutoEmails").getRange("B3");

      
  // Get contact details from "Contacts" sheet and construct To: Header
  // Would be nice to include "Name" as well, to make contacts look prettier, one day.
  var contacts = SpreadsheetApp.getActive().getSheetByName("AutoEmails");
  var numRows = contacts.getLastRow();
  var emailTo = contacts.getRange(3, 1, numRows, 1).getValues();

  // Google scripts can't export just one Sheet from a Spreadsheet
  // So we have this disgusting hack

  // Create a new Spreadsheet and copy the current sheet into it.
  var newSpreadsheet = SpreadsheetApp.create("Spreadsheet to export");
  originalSpreadsheet.copyTo(newSpreadsheet);


  // Find and delete the default "Sheet 1", after the copy to avoid triggering an apocalypse
  newSpreadsheet.getSheetByName('Sheet1').activate();
  newSpreadsheet.deleteActiveSheet();
  
  // Make zee PDF, currently called "Weekly status.pdf"
  // When I'm smart, filename will include a date and project name
  var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();
  var attach = {fileName:sheetName+'.pdf',content:pdf, mimeType:'application/pdf'};

  // Send the freshly constructed email 
  MailApp.sendEmail(emailTo, subject.getValue(), message.getValue(), {attachments:[attach]});
  
  // Delete the wasted sheet we created, so our Drive stays tidy.
  DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);  
}



function exportTwoWeekNoticeNotifications() {
  // Set the Active Spreadsheet so we don't forget
  var originalSpreadsheet = SpreadsheetApp.getActive().getSheetByName("TwoWeekNoticeImport");
  var sheetName = originalSpreadsheet.getSheetName();
  
  // Set the message to attach to the email.
  var message = SpreadsheetApp.getActive().getSheetByName("AutoEmails").getRange("U3");
  
//  // Construct the Subject Line
  var subject = SpreadsheetApp.getActive().getSheetByName("AutoEmails").getRange("T3");

      
  // Get contact details from "Contacts" sheet and construct To: Header
  // Would be nice to include "Name" as well, to make contacts look prettier, one day.
  var contacts = SpreadsheetApp.getActive().getSheetByName("AutoEmails");
  var numRows = contacts.getLastRow();
  var emailTo = contacts.getRange(3, 19, numRows, 1).getValues();

  // Google scripts can't export just one Sheet from a Spreadsheet
  // So we have this disgusting hack

  // Create a new Spreadsheet and copy the current sheet into it.
  var newSpreadsheet = SpreadsheetApp.create("Spreadsheet to export");
  originalSpreadsheet.copyTo(newSpreadsheet);


  // Find and delete the default "Sheet 1", after the copy to avoid triggering an apocalypse
  newSpreadsheet.getSheetByName('Sheet1').activate();
  newSpreadsheet.deleteActiveSheet();
  
  // Make zee PDF, currently called "Weekly status.pdf"
  // When I'm smart, filename will include a date and project name
  var pdf = DriveApp.getFileById(newSpreadsheet.getId()).getAs('application/pdf').getBytes();
  var attach = {fileName:sheetName+'.pdf',content:pdf, mimeType:'application/pdf'};

  // Send the freshly constructed email 
  MailApp.sendEmail(emailTo, subject.getValue(), message.getValue(), {attachments:[attach]});
  
  // Delete the wasted sheet we created, so our Drive stays tidy.
  DriveApp.getFileById(newSpreadsheet.getId()).setTrashed(true);  
}

function saveAsCSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  // create a folder from the name of the spreadsheet
  var folder = DriveApp.createFolder(ss.getName().toLowerCase().replace(/ /g,'_') + '_csv_' + new Date().getTime());
  for (var i = 0 ; i < sheets.length ; i++) {
    var sheet = sheets[i];
    // append ".csv" extension to the sheet name
    fileName = sheet.getName() + ".csv";
    // convert all available sheet data to csv format
    var csvFile = convertRangeToCsvFile_(fileName, sheet);
    // create a file in the Docs List with the given name and the csv data
    folder.createFile(fileName, csvFile);
  }
  Browser.msgBox('Files are waiting in a folder named ' + folder.getName());
}

