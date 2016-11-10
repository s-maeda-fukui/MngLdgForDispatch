var bookMngLdg
  = SpreadsheetApp
    .openById(
      "1rUW-LiPis9cd900tzHED9R7U7xAO-inuV2nu3DJ1iKs"
    );

//管理はTableで
var sheetContractTable
  = bookMngLdg
    .getSheetByName("個別契約管理");
var startRowContractTable
  = 10;

var sheetCommutingCostTable
  = bookMngLdg
    .getSheetByName("交通費管理");

var sheetStaffTable
  = bookMngLdg
    .getSheetByName("スタッフ管理");

//稼働管理は別book
//稼働=仕事中= on duty
var bookOnDutyTable
  = SpreadsheetApp
    .openById(
      "1ctoz2Hx5nmljYfs0EJLCdtoOyxYL8DjHjvujQfWQ_gI"
    );
var sheetAllDatasOnDutyTable
  = bookOnDutyTable
    .getSheetByName("全体");
//ここのシート管理をどうするかは問題ですね。。。
//特に担当者別のシートね・・・。
//いちいち、ここで記入するわけにもいかないし。
//シートを配列で取得してしまって、シート名検索とか・・・
//でも、シートの名前の付け方によっては動かないしなぁ。

//業サポ申請関係
var processIndexArray
  = [
      "stsWorkFlow",
      "stsEnavi",
      "stsLcOrderReceived",
      "stsEndQuitList"
    ];

//申請計算用
var shiftTypeStart
  = [
      "shiftType1Start",
      "shiftType2Start",
      "shiftType3Start",
      "shiftType3Start"
  ];

var shiftTypeEnd
  = [
      "shiftType1End",
      "shiftType2End",
      "shiftType3End",
      "shiftType4End"
  ];

//一覧はListで
var sheetOrderReceiveList
  = bookMngLdg
    .getSheetByName("受注一覧（フォーム実装）");
var formOrderReceive
  = FormApp
    .openById(
      "1veHNQA9PfiE_ycb3DPpcuJfsCbIKJLXhjdPWyiAqs_E"
    );
//受注申請時に計算が必要なカラム
//済→個別契,稼働管理,スタッフ管理

var mkValueIndexArrOnOrderReceive
  = [
      "lcEmployeeUrl",//LC URL
      "dateAskExtensionWill",//延長意志確認期限
      "dispatchContractPeriod",//契約期間
      "numEntension",//延長回数
      "marginRatio",//マージン率
      "workingHours",//就業時間@日
      "workingDaysInMonth",//就業日数@月
  ];
var estimatedValueIndexArr
  = [
      //想定関係の計算は、就業時間・日数より後にあること！
      "estimatedSales",//想定売上（派遣料金）
      "estimatedCosts",//想定原価
      "estimatedGrossProfitGrossProfit",//想定粗利
  ];

var sheetOrderChangeList    ;
var formOrderChange;

var sheetOrderExtensionList ;
var formOrderExtension;

var sheetEndList
  = bookMngLdg
    .getSheetByName("終了一覧（フォーム実装）");
var formEnd
  = FormApp
    .openById(
      "1PZgb4j57KPSg10Bp3L4X_EuNM7GDbaR2W5LxUtp1Vs8"
    );

var sheetQuitList
  = bookMngLdg
    .getSheetByName("途中退場一覧（フォーム実装）");
var formQuit
  = FormApp
    .openById(
      "1BWm-2k3yoEYyEINDl0h6PF-spbCoKDCyNOqDxgOjhyg"
    );

var sheetCanB4ContractList
  = bookMngLdg
    .getSheetByName("雇用契約前キャンセル（フォーム実装）");
var formCanB4Contract
  = FormApp
    .openById(
      "13TT1n_P0eiMsW2pmzv1W1D_b20idOGoOoe7dkcy8T8A"
    );

var sheetCanB4DispatchList
  = bookMngLdg
    .getSheetByName("入場前キャンセル一覧（フォーム実装）");
var formCanB4Dispatch
  = FormApp
    .openById(
      "1YpE2lTev_rhCVJlHRJfuNb8M2UUPR4V67fpjICbHcpg"
    );


//出口関連は後回しにします。
//出口関連を別のブックに移動させるかどうかも検討
//締め、売上などなど
var bookSa
  = SpreadsheetApp
      .openById(
        "1Qr5vSvWOqOhJeT2IyyXmN3UQcNxmt7GKVdEPtnKpOtQ"
      );

var sheetSaSupport
  = bookSa
    .getSheetByName("営アシシート");

var sheetEmployerDeptList
  = bookSa
    .getSheetByName("施設一覧");

//LC URL生成用
var headerLcUrl
  = "http://lc.leverages.jp/Employees/detail/";




