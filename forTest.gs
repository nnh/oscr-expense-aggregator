function getCsvValues_(attachment){
  const csvname = attachment.getName().toLowerCase();
  if (!/^\d{6}\.csv/.test(csvname)){
    return;
  }
  if (Number(csvname.substring(0, 4)) < 2022){
    return;
  }
  const csvtext = attachment.getDataAsString('cp932');
  const splitLf = csvtext.split(/\n/);
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
    if (res[1] == 'カブシキガイシヤボツクスジヤパ'){
      res[1] = 'BOX';
    }
    if (/PVTLTRACKER/.test(res[1])){
      res[1] = res[1].replace('PVTLTRACKER', 'PIVOTAL TRACKER');
    }
    if (
      /^GITHUB$/.test(res[1]) && !/^[0-9]+$/.test(res[2]) ||
      /^DOCKER$/.test(res[1]) && !/^[0-9]+$/.test(res[2])
    ){
      res[2] = res[3];
      res[5] = res[6];
    }
    const ymText = "'" + csvname.substring(0, 4) + "年" + String(Number(csvname.substring(4, 6))) + "月";
    const csvdata = [ymText, ...res];
    return csvdata;
  });
  return(setCsvValues);
}
function getCreditCardInfo(){
  const outputSpreadSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty("spreadsheetIdForTest"));
  const targetTerm = 'subject:(クレジットカード明細)'
  let gmailThreads;
  gmailThreads = GmailApp.search(targetTerm);
  const threadsValues = gmailThreads.map(thread => 
  thread.getMessages().map(message => 
    message.getAttachments().map(attachment => getCsvValues_(attachment))
    )
  );
  const values = threadsValues.flat().flat().flat().filter(x => x !== undefined);
  const values2 = values.filter(x => x[1] !== "");
  const header = [["year_month"], ["date"], ["item"], ["filler1"], ["filler2"], ["filler3"], ["price"]];
  const values2_1 = [header, ...values2];
  const values3 = padArray_(values2_1);
  const outputSheet = outputSpreadSheet.getSheets()[0];
  outputSheet.clearContents();
  outputSheet.getRange(1, 1, values3.length, values3[0].length).setValues(values3);
}
function padArray_(arr) {
  // 一番長いサブ配列の長さを取得
  const maxLength = Math.max(...arr.map(subArray => subArray.length));
  
  // 各サブ配列を一番長いサブ配列に合わせる
  const paddedArray = arr.map(subArray => {
    const diff = maxLength - subArray.length;
    return subArray.concat(Array(diff).fill(null));
  });

  return paddedArray;
}
function setSumGroupByYearMonth(){
  const outputSpreadSheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty("spreadsheetIdForTest"));
  const inputSheet = outputSpreadSheet.getSheets()[0];
  const outputSheetName = "合計"
  if (inputSheet.getName === outputSheetName){
    return;
  }
  let outputSheet = outputSpreadSheet.getSheetByName(outputSheetName);
  if (outputSheet === null){
    outputSheet = outputSpreadSheet.insertSheet();
    outputSheet.setName(outputSheetName);
  }
  const targetYM = inputSheet.getRange(2, 1, inputSheet.getLastRow(), 1).getValues().flat().filter(x => x !== "");
  const setYm = Array.from([...new Set(targetYM)]);
  const sumList = setYm.map(ym => {
    const target = inputSheet.getDataRange().getValues().filter(x => x[0] === ym);
    const sum = target.reduce((acc, row) => acc + row[6], 0);
    return([ym, sum]);
  });
  outputSheet.clearContents();
  outputSheet.getRange(1, 1, sumList.length, 2).setValues(sumList);

}