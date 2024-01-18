const sheetNamesForTest = new Map([
  ["csv", "csv"],
  ["items", "items"],  
  ["inputData", "filteredData"],
]);
const headerForTest = [["year_month"], ["date"], ["item"], ["filler1"], ["filler2"], ["filler3"], ["price"], ["filler4"]];
const headerIndex = new Map(headerForTest.map((x, idx) => [x[0], idx]));
const outputSpreadSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty("spreadsheetIdForTest"));
function getCsvValues_(attachment){
  const csvName = attachment.getName().toLowerCase();
  if (!/^\d{6}\.csv/.test(csvName)){
    return;
  }
  if (Number(csvName.substring(0, 4)) < 2022){
    return;
  }
  const csvText = attachment.getDataAsString('cp932');
  const csvTextEdit = csvText.replace(/, inc\./gi, " inc.");
  const splitLf = csvTextEdit.split(/\n/);
  let splitComma = splitLf.map(x => x.split(','));
  splitComma[0][0] = '';
  splitComma[0][1] = '';
  const maxIdx = splitComma.map(x => x.length).reduce((x, y) => Math.max(x, y));
  const setCsvValues = splitComma.map(x => {
    let res;
    if (x.length < maxIdx){
      const pushCount = maxIdx - x.length;
      const temp = new Array(pushCount).fill('');
      res = x.concat(temp);
    } else {
      res = x;
    }
    const ymText = "'" + csvName.substring(0, 4) + "年" + String(Number(csvName.substring(4, 6))) + "月";
    const csvdata = [ymText, ...res];
    return csvdata;
  });
  return(setCsvValues);
}
function getCreditCardInfo(){
  const targetTerm = 'subject:(クレジットカード明細)'
  let gmailThreads;
  gmailThreads = GmailApp.search(targetTerm);
  const threadsValues = gmailThreads.map(thread => 
  thread.getMessages().map(message => 
    message.getAttachments().map(attachment => getCsvValues_(attachment))
    )
  );
  const values = threadsValues.flat().flat().flat().filter(x => x !== undefined);
  const values2 = values.filter(x => x[6] !== "");
  const values2_1 = [headerForTest, ...values2];
  const outputSheet = getSheet_(outputSpreadSheet, sheetNamesForTest.get("csv"));
  outputSheet.setName(sheetNamesForTest.get("csv"));
  outputSheet.clearContents();
  outputSheet.getRange(1, 1, values2_1.length, values2_1[0].length).setValues(values2_1);
  outputSheet.hideColumns(4, 3);
  outputSheet.hideColumns(8, 1);
}
function test(){
  const latestYear = targetYears.slice(-1)[0];
  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(latestYear + "_1");
  const itemNames = targetSheet.getRange("A:A").getValues().filter(x => x[0] !== "" && x[0] !== "計");
  const itemNameSheet = getSheet_(outputSpreadSheet, sheetNamesForTest.get("items"));
  itemNameSheet.setName(sheetNamesForTest.get("items"));
  itemNameSheet.clearContents();
  itemNameSheet.getRange(2, 1, itemNames.length, 1).setValues(itemNames);
  const csvValues = outputSpreadSheet.getSheetByName(sheetNamesForTest.get("csv")).getDataRange().getValues();
  const inputDataSheetValues = targetYears.map(year => {
    const yearAndMonthText1 = [4, 5, 6, 7, 8, 9, 10, 11, 12].map(x => `${year}年${x}月`);  
    const yearAndMonthText2 = [1, 2, 3].map(x => `${year+1}年${x}月`); 
    const yearAndMonthText = [...yearAndMonthText1, ...yearAndMonthText2]; 
    const filteredData = yearAndMonthText.map(yearAndMonth => {
      const targetValues = csvValues.filter(x => x[headerIndex.get("year_month")] === yearAndMonth); 
      if (targetValues.length === 0){
        return(null);
      }
      const [itemAndPrice, errorValues] = getItemAndPrice_(itemNames, targetValues);
      if (errorValues !== null){
        console.log(`${yearAndMonth}:${[...errorValues]}`);
      }
      const yearMonthItemPrice = itemAndPrice.map(x => [yearAndMonth, ...x]);
      return(yearMonthItemPrice);
    }).flat();
    return(filteredData);
  }).flat().filter(x => x !== null);
  const inputDataSheet = getSheet_(outputSpreadSheet, sheetNamesForTest.get("inputData"));
  inputDataSheet.setName(sheetNamesForTest.get("inputData"));
  inputDataSheet.clearContents();
  const inputDataSheetHeader = [
    headerForTest[headerIndex.get("year_month")], 
    headerForTest[headerIndex.get("item")],
    headerForTest[headerIndex.get("price")],
    ["original_item"] 
  ].flat();
  const inputData = [inputDataSheetHeader, ...inputDataSheetValues];
  inputDataSheet.getRange(1, 1, inputData.length, inputData[0].length).setValues(inputData);
}
function getItemAndPrice_(itemNames, targetValues){
  const itemSheetItemIndex = 0; 
  const itemAndPricePriceIndex = 1; 
  const originalItemIndex = 2;
  const itemAndPrice = itemNames.map(itemName => {
    const values = targetValues.map(value => {
      if (itemName[itemSheetItemIndex] === "BOX"){
        if (value[headerIndex.get("item")] === "カブシキガイシヤボツクスジヤパ"){
          return([itemName[itemSheetItemIndex], value[headerIndex.get("price")], value[headerIndex.get("item")]])
        } else {
          return(null);
        }
      }
      if (RegExp(itemName[itemSheetItemIndex], "i").test(value[headerIndex.get("item")])){
        return([itemName[itemSheetItemIndex], value[headerIndex.get("price")], value[headerIndex.get("item")]])
      } else if (itemName[itemSheetItemIndex] === "PIVOTAL TRACKER" && value[headerIndex.get("item")] === "DRI*PVTLTRACKER (MY.VMWARE.COM)"){
        return([itemName[itemSheetItemIndex], value[headerIndex.get("price")], value[headerIndex.get("item")]])
      } else if (itemName[itemSheetItemIndex] === "書籍購入" && value[headerIndex.get("item")] === "SP PR GRAPHQL (MONTREAL )"){
        return([itemName[itemSheetItemIndex], value[headerIndex.get("price")], value[headerIndex.get("item")]])
      } else {
        return(null)
      }
    }).filter(x => x !== null);
    if (values.length > 0){
      const price = values.reduce((acc, row) => acc + row[itemAndPricePriceIndex], 0); 
      return([itemName[itemSheetItemIndex], price, values[0][originalItemIndex]]);
    } else {
      return(null)
    }
  }).filter(x => x !== null);
  const targetValuesItems = new Set(targetValues.map(value => value[headerIndex.get("item")]));
  const itemAndPriceItems = new Set(itemAndPrice.map(value => value[originalItemIndex])); 
  const originalSumPrice = targetValues.filter(value => value[headerIndex.get("item")] === "")[0][headerIndex.get("price")];
  const priceSum = itemAndPrice.reduce((acc, row) => acc + row[itemAndPricePriceIndex], 0);
  const errorValues = originalSumPrice === priceSum ? null : new Set([...targetValuesItems].filter(x => !itemAndPriceItems.has(x)).filter(x => x !== ""));
  return [itemAndPrice, errorValues];
}