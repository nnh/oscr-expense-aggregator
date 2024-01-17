function editSheetValues(){
  const inputSheetName = "List";
  const targetYears = [2022, 2023];
  targetYears.forEach(year => {
    const property_text = "ssId" + String(year);
    const inputValues = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty(property_text)).getSheetByName(inputSheetName).getDataRange().getValues();
    const sheet1 = getSheetAndSetValues_(String(year) + "_1", inputValues);
    sheet1.hideColumns(15, 5);
    const sheet2 = getSheetAndSetValues_(String(year) + "_2", inputValues);
    sheet2.hideColumns(2, 12);
  });
}
function getSheetAndSetValues_(sheetName, inputValues){
    let targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (targetSheet == null){
      targetSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
      SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().setName(sheetName);
    }  
    targetSheet.clearContents();
    targetSheet.getRange(1, 1, inputValues.length, inputValues[0].length).setValues(inputValues);
    return(targetSheet);
}

function exportSheetsToPDF(){
  const outputFolder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty("outputFolderId"));
  const todayText = getFormattedDate_();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();  
  const fileId = spreadsheet.getId();
  const baseUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?id=${fileId}`;
  const parameters = new Map([
    ['gridlines', 'true'],
    ['portrait', 'false'],
  ]);
  const options = generateOptionsString_(parameters);
  const url = baseUrl + options;
  const token = ScriptApp.getOAuthToken();
  const fetchOptions = {
    headers: {
      Authorization: `Bearer ${token}`,
    },
  };
  spreadsheet.getSheets().forEach(sheet => {
    const sheetName = sheet.getName(); 
    spreadsheet.getSheets().forEach(target => {
      if (target.getName() !== sheetName){
        target.hideSheet();
      }
    }); 
    const [year, seq] = sheetName.split("_");
    const newFileName = todayText + " OSCR理事会用" + seq + "(" + year + ")" + '.pdf';
    const blob = UrlFetchApp.fetch(url, fetchOptions)
      .getBlob()
      .setName(newFileName);
    outputFolder.createFile(blob);
    spreadsheet.getSheets().forEach(sheet => sheet.showSheet());
  });
}

function getFormattedDate_() {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0'); // Months are zero-based, so we add 1 and zero-pad to two digits
  const day = String(today.getDate()).padStart(2, '0'); // Zero-pad the day to two digits

  const formattedDate = `${year}${month}${day}`;
  return formattedDate;
}