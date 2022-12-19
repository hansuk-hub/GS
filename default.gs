var ui = SpreadsheetApp.getUi();
Array.prototype.findIndex = function(search)
{
  if(search == "") return -1;
  for(var t = 0; t < this.length; t++)
  {
    if(this[t].search(search) >=0)
    {break}
  }
  return t-1;
}
function doList() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var ogst = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("2022-12-15-발주");
  var tgst = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("박스-밀크런-적재리스트");
  var proNamest = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("상품명변경정리");
  var getLastRow = ogst.getLastRow();
  for ( var i = 2; i <= getLastRow; i++ ) {
      var getDT = ogst.getRange(i,1,1,10);
      var getFt = getDT.getValues().flat();
      // Logger.log(getFt[6]);
      // var setDataRow = [getFt[0],getFt[1],getFt[4],getFt[5],getFt[6],getFt[7]];
      var getRowPostion = proNamest.getRange("A:A").getValues().flat();
      var doCheck = getRowPostion.indexOf(getFt[6]);
      if (doCheck > 0 ) {
        var newProName = proNamest.getRange( getRowPostion.indexOf(getFt[6])+1 , 2).getValue();
        tgst.getRange(3+i,2,1,6).setValues([[getFt[0],getFt[1],getFt[4],getFt[5],newProName,getFt[7]]]);
      }
      else {
        tgst.getRange(3+i,2,1,6).setValues([[getFt[0],getFt[1],getFt[4],getFt[5],getFt[6],getFt[7]]]);
        tgst.getRange(3+i,6).setFontColor('#ff0000').setFontWeight('bold');
      }
      Logger.log ( i)
  }
  ui.alert('테스트 작업 완료!')
}
function doChgName(){
}
function doInsertBox(){
}
function setDefault(){
  var tgst = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("박스-밀크런-적재리스트");
    tgst.getRange(1,1,100,20).setFontColor('#000000').setFontWeight('normal');
    tgst.getRange(6,1,100,20).clear();
    // var getLastRow = tgst.getLastRow();
    ui.alert('데이터 삭제 완료!')
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('사무 자동 입력')
      .addItem('1-적재리스트만들기', 'doList')
      .addItem('2-수량 자동 입력','doSumBox')
      .addSeparator()
      .addItem('발주 데이터 삭제!','setDefault')
      .addToUi();  
}

