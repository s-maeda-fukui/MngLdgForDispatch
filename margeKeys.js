var book
  = SpreadsheetApp
    .openById(
      "1okkwLqDpr3rkYs6S_uxJ3iZghJT6udya4XX3avNiubM"
    );
var sheet1
  = book
    .getSheetByName("シート1");

var sheet2
  = book
    .getSheetByName("対照表");

//for Sheet1
var lastRowSheet1 = sheet1.getLastRow();
var lastColSheet1 = sheet1.getLastColumn();
var indexArraySheet1
  = sheet1
    .getRange( 2, 1, 1, lastColSheet1 )
    .getValues()[0];
var valuesSheet1
  = sheet1
    .getDataRange()
    .getValues();

//for Sheet2
var lastRowSheet2 = sheet2.getLastRow();
var lastColSheet2 = sheet2.getLastColumn();
var indexArraySheet2
  = sheet2
    .getRange( 2, 1, 1, lastColSheet2 )
    .getValues()[0];
var valuesSheet2
  = sheet2
    .getDataRange()
    .getValues();

function findLostKey(){
  var index
    = indexArraySheet1.indexOf("施設一覧");

  Logger.log("Sheet1");
  for(var i = 1;
      i < valuesSheet1.length;
      ++i){

    if( "" != valuesSheet1[i][index]){
      var check = 0 ;
      for(var j = 1;
          j < valuesSheet2.length;
          ++j){

        if( valuesSheet1[i][index]
            == valuesSheet2[j][index] ){
          ++check;
          break ;
        }//if_match

      }//for_j
      if(check == 0){
        Logger.log(valuesSheet1[i][index]);
      }
    }//if_void
  }//for_i

  Logger.log("Sheet2");
  for(var i = 1;
      i < valuesSheet2.length;
      ++i){

    if( "" != valuesSheet2[i][index]){
      var check = 0 ;
      for(var j = 1;
          j < valuesSheet1.length;
          ++j){

        if( valuesSheet2[i][index]
            == valuesSheet1[j][index] ){
          ++check;
          break ;
        }//if_match

      }//for_j
      if(check == 0){
        Logger.log(valuesSheet2[i][index]);
      }
    }//if_void
  }//for_i

}//func_findLostKey

//sheet2が対照表
function matchKeys(){
  //管理台帳との照合
  var bookMngLdg
    = SpreadsheetApp
      .openById(
        "1rUW-LiPis9cd900tzHED9R7U7xAO-inuV2nu3DJ1iKs"
      );

  var sheetName = "稼働管理";
  var sheetTgt
    = bookMngLdg
      .getSheetByName(sheetName);
  var index
    = indexArraySheet2.indexOf(sheetName);

  var arrayMngLdg
    = sheetTgt
      .getRange( 1, 1, 1, sheetTgt.getLastColumn() )
      .getValues()[0];

  Logger.log("管理台帳との照合");

  for(var i = 2;
      i < valuesSheet2.length;
      ++i){
    if( "" != valuesSheet2[i][index]
        &&
        -1 == arrayMngLdg.indexOf(valuesSheet2[i][1])
    ){
      //Logger.log(valuesSheet2[i][1]+"："+arrayMngLdg.indexOf(valuesSheet2[i][1]));
      Logger.log(valuesSheet2[i][index]);
    }//if_void
  }//for_i

}//func_maechKeys

//sheet2が対照表
function matchKeysOtheBook(){
  //他のブックとの照合
  var bookTgt
    = SpreadsheetApp
      .openById(
        "1Qr5vSvWOqOhJeT2IyyXmN3UQcNxmt7GKVdEPtnKpOtQ/"
      );
  var sheetTgt
    = bookTgt
      .getSheetByName("延長確認&勤怠修正");

  var sheetName = "営アシシート";
  var index
    = indexArraySheet2.indexOf(sheetName);

  var arrayTgt
    = sheetTgt
      .getRange( 1, 1, 1, sheetTgt.getLastColumn() )
      .getValues()[0];

  Logger.log(sheetName+"との照合");

  for(var i = 2;
      i < valuesSheet2.length;
      ++i){
    if( "" != valuesSheet2[i][index]
        &&
        -1 == arrayTgt.indexOf(valuesSheet2[i][1])
    ){
      //Logger.log(valuesSheet2[i][1]+"："+arrayTgt.indexOf(valuesSheet2[i][1]));
      Logger.log(valuesSheet2[i][index]);
    }//if_void
  }//for_i

}//func_maechKeys




