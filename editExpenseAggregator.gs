const targetYears = [2022, 2023];

function editSheetValues(){
  const inputSheetName = "List";
  const colNumber = new Map([
    ["startMonth", 2],
    ["monthlyAverage", 15],
  ])
  targetYears.forEach(year => {
    const property_text = "ssId" + String(year);
    const inputValues = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty(property_text)).getSheetByName(inputSheetName).getDataRange().getValues();
    const sheet1 = getSheetAndSetValues_(String(year) + "_1", inputValues);
    sheet1.hideColumns(Number(colNumber.get("monthlyAverage")), sheet1.getLastColumn()-Number(colNumber.get("monthlyAverage"))+1);
    const sheet2 = getSheetAndSetValues_(String(year) + "_2", inputValues);
    const sheet2SumRowNumber = sheet2.getRange(`A2:A${sheet2.getLastRow()}`).getValues().map((x, idx) => x[0] === "計" ? idx : null).filter(x => x !== null)[0] + 2;
    for (let i = 2; i < sheet2SumRowNumber; i++){
      sheet2.getRange(i, Number(colNumber.get("monthlyAverage"))).setValue(Math.round(sheet2.getRange(i, Number(colNumber.get("monthlyAverage"))).getValue()));
      sheet2.getRange(i, Number(colNumber.get("monthlyAverage"))+1).setValue(Math.round(sheet2.getRange(i, Number(colNumber.get("monthlyAverage"))+1).getValue()));
      sheet2.getRange(i, Number(colNumber.get("monthlyAverage"))+2).setFormula(`=O${i}-P${i}`);
    }
    sheet2.getRange(sheet2SumRowNumber, Number(colNumber.get("monthlyAverage"))).setFormula(`=sum(O2:O${sheet2SumRowNumber-1})`);
    sheet2.getRange(sheet2SumRowNumber, Number(colNumber.get("monthlyAverage"))+1).setFormula(`=sum(P2:P${sheet2SumRowNumber-1})`);
    sheet2.getRange(sheet2SumRowNumber, Number(colNumber.get("monthlyAverage"))+2).setFormula(`=O${sheet2SumRowNumber}-P${sheet2SumRowNumber}`);
    if (sheet2.getName() === "2022_2"){
      sheet2.getRange("P29").setValue(115);
    }
    sheet2.hideColumns(Number(colNumber.get("startMonth")), Number(colNumber.get("monthlyAverage"))-Number(colNumber.get("startMonth"))-1);
  });
}

function getSheet_(spreadsheet, sheetName){
  let targetSheet = spreadsheet.getSheetByName(sheetName);
  if (targetSheet === null){
    targetSheet = spreadsheet.insertSheet();
  }  
  return(targetSheet);
}
function getSheetAndSetValues_(sheetName, inputValues){
    const targetSheet = getSheet_(SpreadsheetApp.getActiveSpreadsheet(), sheetName);
    targetSheet.setName(sheetName);
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