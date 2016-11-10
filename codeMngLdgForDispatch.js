function matchCopyAndSetValues(
  sheetTgt,indexArrayAppli,valuesArrayAppli
){
  var lastRowTgt
    = sheetTgt.getLastRow();
  var lastColTgt
    = sheetTgt.getLastColumn();
  var indexArrayTgt
    = sheetTgt
      .getRange(
        1, 1,
        1, lastColTgt
      )
      .getValues()[0];

  var insertArray = new Array();
  var tgtCol;

  for(var i = 0;
      i < indexArrayTgt.length;
      ++i){

    tgtCol
      = indexArrayAppli
        .indexOf(indexArrayTgt[i]);

    if( -1 != tgtCol ){
/*
      Logger.log(
        indexArrayAppli
           .indexOf(indexArrayTgt[i])
        +":"+
        indexArrayTgt[i]
        +"→"+
        valuesArrayAppli[tgtCol]
        );
*/
      insertArray[i]
        = valuesArrayAppli[tgtCol];
    }else{
      //Logger.log("-1:"+indexArrayTgt[i]);
      //どっかのキャッシュかなにかが入り込む
      //確実に空にしてあげる必要あり。
      insertArray[i]
        = "";
    }//if_indexFound
  }//for_i

  //データ移植
  sheetTgt
    .getRange(
      lastRowTgt+1,1,
      1,lastColTgt
    )
    .setValues([insertArray]);

}

function makeValueOnOrderReceive(
  mkValueIndexArray,indexArrayAppli,valuesArrayAppli
){
  var rtnValueArray = new Array();
  var hrs, mins;

  mkValueIndexArray
  .forEach(
  function(value,index){
    if("lcEmployeeUrl" === value){
      rtnValueArray[index]
        = lib.headerLcUrl
        + valuesArrayAppli
          [
            indexArrayAppli
            .indexOf("lcEmployeeId")
          ];
    }else if("dateAskExtensionWill" === value){
      //new Dateしないと、参照渡しになってしまうので。
      var dateDispatchEnd
        = new Date(
            valuesArrayAppli
            [
              indexArrayAppli
              .indexOf("dateDispatchEnd")
            ]
          );
      //単純に1月前を設定
      rtnValueArray[index]
        = new Date(
            dateDispatchEnd
            .setMonth(
              dateDispatchEnd.getMonth()-1
            )
          );
    }else if("dispatchContractPeriod" === value){
      var tmp
        = new Date(
              valuesArrayAppli
              [
                indexArrayAppli
                .indexOf("dateDispatchEnd")
              ]
            -
              valuesArrayAppli
              [
                indexArrayAppli
                .indexOf("dateDispatchStart")
              ]
          ) ;
      if( 1970 == tmp.getFullYear() ){
        //Logger.log("**" + (tmp.getUTCMonth()+1) );
        rtnValueArray[index]
          = tmp.getUTCMonth()+1;
      }else{
        rtnValueArray[index]
          = ( tmp.getFullYear()-1970 )*12
          + tmp.getUTCMonth() + 1;
      }//InSameYear?
    }else if("numEntension" === value){
      //受注申請から入るものは新規/復活orスライド
      rtnValueArray[index]=0;
    }else if("marginRatio" === value){
      rtnValueArray[index]
        = valuesArrayAppli
          [
            indexArrayAppli
            .indexOf("dispatchWage")
          ]
        /
          valuesArrayAppli
          [
            indexArrayAppli
            .indexOf("dispatchFee")
          ]
        * 100;
    }else if("workingHours" === value){
      hrs=0;
      mins=0;
      var count =0;
      lib.shiftTypeStart
      .forEach(
      function(val,index){
        var start
          = valuesArrayAppli
            [
              indexArrayAppli
              .indexOf(val)
            ];
        if( "" != start ){
          var end
            = valuesArrayAppli
              [
                indexArrayAppli
                .indexOf(
                  lib.shiftTypeEnd[index]
                )
              ];
          hrs  += (end.getHours() - start.getHours());
          mins += (end.getMinutes() - start.getMinutes());
          ++count;
        }else{
          //Do Nothing
        }
      }//function
      );//forEach
      //Logger.log(hrs+":"+mins);
      var tmp
        = new Date(
           ( hrs*60 + mins)
           * 60 * 1000
           / count
          );
      //1970年〜の初期値が9:00から始まってしまうので
      //9を引く必要あり
      rtnValueArray[index]
        = new Date(
            tmp
            .setHours(
              tmp.getHours()-9
            )
          );
/*
      Logger.log(
        new Date(
          tmp
          .setHours(
            tmp.getHours()-9
          )
        )
      );
*/
    }else if("workingDaysInMonth" === value){
      var workDays
      = valuesArrayAppli
        [
          indexArrayAppli
          .indexOf("workDaysInWeek")
        ];
      var avgWorkDaysInWeek
      if( typeof workDays === "string" ){
        avgWorkDaysInWeek
          = (
              Number(workDays[0])
              + Number(workDays[workDays.length-1] )
            ) / 2.0 ;
      }else{
        avgWorkDaysInWeek = workDays;
      }//calc_part
      rtnValueArray[index]
        = 4*avgWorkDaysInWeek;
    }//if_mkValueIndexArr
  }//function
  );//forEach_mkValueIndexArr

  //Logger.log(mkValueIndexArray);
  //Logger.log(rtnValueArray);

  //連結
  indexArrayAppli
    = indexArrayAppli
      .concat(mkValueIndexArray);
  valuesArrayAppli
    = valuesArrayAppli
      .concat(rtnValueArray);

  //配列で戻す
  return [indexArrayAppli,valuesArrayAppli] ;
}//func_makeValueOnOrderReceive

function calcEstimatedValues(
  estValuesIndexArray,indexArrayAppli,valuesArrayAppli
){
  var times
    = valuesArrayAppli
      [
        indexArrayAppli
        .indexOf("workingDaysInMonth")
      ];

  if( "時給" ===
      valuesArrayAppli
      [
        indexArrayAppli
        .indexOf("baseUnitPayment")
      ]
  ){
    var hrs
      = valuesArrayAppli
        [
          indexArrayAppli
          .indexOf("workingHours")
        ].getHours()
      +(valuesArrayAppli
        [
          indexArrayAppli
          .indexOf("workingHours")
        ].getMinutes() / 60);
    times *= hrs ;
  }//if_baseUnitPayment

  var rtnValueArray = new Array();
  estValuesIndexArray.forEach(
  function(value,index){
    if("estimatedSales" === value){
      rtnValueArray[index]
        = valuesArrayAppli
          [
            indexArrayAppli
            .indexOf("dispatchFee")
          ]
        * times ;
    }else if("estimatedCosts" === value){
      var tmp
        = valuesArrayAppli
          [
            indexArrayAppli
            .indexOf("dispatchWage")
          ]
        * times;
      if(
        "有" ==
        valuesArrayAppli
        [
          indexArrayAppli
          .indexOf("getSocialIns")
        ]
      ){
        tmp *= 1.13;
      }//socialIns
      rtnValueArray[index] = tmp;
    }else if("estimatedGrossProfitGrossProfit" === value){
      rtnValueArray[index]
        = rtnValueArray[0]
        - rtnValueArray[1];
    }//if_estValueIndexArray
  }//functio
  );//forEach_estValuesIndexArray

  //連結
  indexArrayAppli
    = indexArrayAppli
      .concat(estValuesIndexArray);
  valuesArrayAppli
    = valuesArrayAppli
      .concat(rtnValueArray);

  //配列で戻す
  return [indexArrayAppli,valuesArrayAppli] ;
}//func_calcEstimatedValues

function setWorkingId(){
  var lastRow
    = lib.sheetContractTable
      .getLastRow();
  var lastCol
    = lib.sheetContractTable
      .getLastColumn();
  var indexArray
    = lib.sheetContractTable
      .getRange(
        1, 1,
        1, lastCol
      )
      .getValues()[0];
  var values
    = lib.sheetContractTable
      .getDataRange()
      .getValues();
  var lastNum = 0;
  var tmp ;
  for(var i = lib.startRowContractTable;
      i < values.length;
      ++i){
    tmp
     = values[i][indexArray.indexOf("workingId")];
    if(lastNum < tmp){
      lastNum = tmp;
    }//if_lastNum
  }//for_i
  return (lastNum+1);
}//func_setWorkingId

function processOrderReceived(){
  //受注申請一覧関連
  //最終行
  var lastRowOrderReceiveList
    = lib.sheetOrderReceiveList
      .getLastRow();
  //最終列
  var lastColOrderReceiveList
    = lib.sheetOrderReceiveList
      .getLastColumn();
  //インデックス
  var indexArrayOrderReceiveList
    = lib.sheetOrderReceiveList
      .getRange(
        1, 1,
        1, lastColOrderReceiveList
      )
      .getValues()[0];

  //受注申請一覧に稼働IDをセット
  lib.sheetOrderReceiveList
  .getRange(
    lastRowOrderReceiveList,
    indexArrayOrderReceiveList
    .indexOf("workingId")
    +1
  )
  .setValue( setWorkingId() );

  //データ
  var valuesArrayOrderReceiveList
    = lib.sheetOrderReceiveList
      .getRange(
        lastRowOrderReceiveList, 1,
        1,lastColOrderReceiveList
      )
      .getValues()[0];

  //計算が必要な値の処理
  [
    indexArrayOrderReceiveList,
    valuesArrayOrderReceiveList
  ]
    = makeValueOnOrderReceive(
        lib.mkValueIndexArrOnOrderReceive,
        indexArrayOrderReceiveList,
        valuesArrayOrderReceiveList
      );
  //想定計算関係
  [
    indexArrayOrderReceiveList,
    valuesArrayOrderReceiveList
  ]
    = calcEstimatedValues(
        lib.estimatedValueIndexArr,
        indexArrayOrderReceiveList,
        valuesArrayOrderReceiveList
      );

  //スタッフ管理にデータを移植
  matchCopyAndSetValues(
    lib.sheetStaffTable,
    indexArrayOrderReceiveList,
    valuesArrayOrderReceiveList
  );

  //個別契約管理にデータを移植
  matchCopyAndSetValues(
    lib.sheetContractTable,
    indexArrayOrderReceiveList,
    valuesArrayOrderReceiveList
  );

}//func_processOrderReceived

function setTrigger(){
  var allTriggers = ScriptApp.getProjectTriggers();
  for( var i = 0; i < allTriggers.length; ++i ){
    ScriptApp.deleteTrigger(allTriggers[i]);
  }//delete_triggers

  //メニューバー生成のためのトリガ設定
  ScriptApp
    .newTrigger("addMenuBar")
    .forSpreadsheet(lib.bookMngLdg)
    .onOpen()
    .create();

  ScriptApp
    .newTrigger("processOrderReceived")
    .forForm(lib.formOrderReceive)
    .onFormSubmit()
    .create();
}//func_setTrigger

function hello(){
  Browser.msgBox("Hello");
}
function addMenuBar(){
  SpreadsheetApp
    .getUi()
    .createMenu("一覧変更関係")
    .addSubMenu(
      SpreadsheetApp
      .getUi()
      .createMenu("受注一覧")
        .addItem("e-naviスタッフID記入","hello")
        .addItem("承認","hello")
    )
    .addSubMenu(
      SpreadsheetApp
      .getUi()
      .createMenu("受注変更一覧")
        .addItem("承認","hello")
    )
    .addSubMenu(
      SpreadsheetApp
      .getUi()
      .createMenu("延長一覧")
        .addItem("承認","hello")
    )
    .addSubMenu(
      SpreadsheetApp
      .getUi()
      .createMenu("途中退場一覧")
        .addItem("LC ID再発行","hello")
        .addItem("承認","hello")
    )
    .addSubMenu(
      SpreadsheetApp
      .getUi()
      .createMenu("入場前キャン一覧")
        .addItem("LC ID再発行","hello")
        .addItem("承認","hello")
    )
    .addSubMenu(
      SpreadsheetApp
      .getUi()
      .createMenu("契約前キャン一覧")
        .addItem("LC ID再発行","hello")
        .addItem("承認","hello")
    )
    .addToUi();
}



