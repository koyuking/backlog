//テスト用メモ
// Browser.msgBox(a,a);

//共通関数
var backlogsheet = SpreadsheetApp.getActive().getSheetByName('backlogプロジェクト管理表');
var tempsheet = SpreadsheetApp.getActive().getSheetByName('メールテンプレート');
var logsheet = SpreadsheetApp.getActive().getSheetByName('ログ');
var logsheet10 = SpreadsheetApp.getActive().getSheetByName('10日前');
var logsheetlast = SpreadsheetApp.getActive().getSheetByName('最終通知');
var kanrisha = tempsheet.getRange("H2").getValue(); //管理者用メールアドレス設定箇所
var currentDate = new Date(); //今日の日付

// onOpenのインストール
function onInstall()
{
  onOpen();
}

//実行メニュー追加
function onOpen() {
  var aw = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries.push({name: "削除完了通知", functionName: "backlogsakujyo"});
  aw.addMenu("アクション", menuEntries);
}

//期限確認メール
function MailListCheck() {
  var today = Utilities.formatDate(currentDate,"JST","yyyy/M/d");
  var check30day; //30日前のデフォルト日付
  var check10day; //10日前のデフォルト日付
  var checktoday; //更新期限のデフォルト日
  var check30date; //30日前の変換後の日付
  var check10date; //10日前の変換後の日付
  var checktodate; //更新期限日

  //メール用情報格納
  var owner;
  var ownername;
  var bkproject;
  var maxdate;
  
  //メールテンプレートB列使用 
  var mailSubjectB = tempsheet.getRange(13,2).getValue();
  var mailbodyB = tempsheet.getRange(16,2).getValue();
  var optionB = {from: kanrisha};
  
  //メールテンプレートD列使用 
  var mailSubjectD = tempsheet.getRange(13,4).getValue();
  var mailbodyD = tempsheet.getRange(16,4).getValue();
  var optionD = {from: kanrisha,cc: kanrisha};
  
  //メールテンプレートF列使用
  var mailSubjectF = tempsheet.getRange(13,6).getValue();
  var mailbodyF = tempsheet.getRange(16,6).getValue();
  var optionF = {from: kanrisha,cc: kanrisha};
  
  // 最終行から1行ずつ上の行を参照
  for (var i = backlogsheet.getLastRow(); i > 6; i--) {
    check30day = backlogsheet.getRange("N" + i).getValue();
    check10day = backlogsheet.getRange("M" + i).getValue();
    checktoday = backlogsheet.getRange("L" + i).getValue();
    
    // 判定カラムの制御
    if(check30day >= 1 || check10day >= 1 || checktoday >= 1){
      check30date = Utilities.formatDate(check30day,"JST","yyyy/M/d");
      check10date = Utilities.formatDate(check10day,"JST","yyyy/M/d");
      checktodate = Utilities.formatDate(checktoday,"JST","yyyy/M/d");
    }
    
    //メール送信処理
    if(check30date == today){
      owner = backlogsheet.getRange("E" + i).getValue();
      ownername = backlogsheet.getRange("F" + i).getValue();
      bkproject = backlogsheet.getRange("D" + i).getValue();
      maxdate = backlogsheet.getRange("L" + i).getValue();
      maxdate = Utilities.formatDate(maxdate,"JST","yyyy/M/d"); 
      var mailbody = mailbodyB.replace('${"管理者"}',ownername).replace('${"プロジェクト"}',bkproject).replace('${"利用期限"}',maxdate);
      GmailApp.sendEmail(owner,mailSubjectB,mailbody,optionB);
      logsheet.appendRow([currentDate, bkproject, owner, ownername, "30日前メール"]);
    }
    if(check10date == today){
      owner = backlogsheet.getRange("E" + i).getValue();
      ownername = backlogsheet.getRange("F" + i).getValue();
      bkproject = backlogsheet.getRange("D" + i).getValue();
      maxdate = backlogsheet.getRange("L" + i).getValue();
      maxdate = Utilities.formatDate(maxdate,"JST","yyyy/M/d");
      var mailbody = mailbodyD.replace('${"管理者"}',ownername).replace('${"プロジェクト"}',bkproject).replace('${"利用期限"}',maxdate);
      GmailApp.sendEmail(owner,mailSubjectD,mailbody,optionD);
      logsheet10.appendRow([currentDate, bkproject, owner, ownername, "10日前メール"]);
    }
    if(checktodate == today){
      owner = backlogsheet.getRange("E" + i).getValue();
      ownername = backlogsheet.getRange("F" + i).getValue();
      bkproject = backlogsheet.getRange("D" + i).getValue();
      maxdate = backlogsheet.getRange("L" + i).getValue();
      maxdate = Utilities.formatDate(maxdate,"JST","yyyy/M/d"); 
      var mailbody = mailbodyF.replace('${"管理者"}',ownername).replace('${"プロジェクト"}',bkproject).replace('${"利用期限"}',maxdate);
      GmailApp.sendEmail(owner,mailSubjectF,mailbody,optionF);
      logsheetlast.appendRow([currentDate, bkproject, owner, ownername, "最終通知", "－"]);
    }
  }
}

//削除通知
function backlogsakujyo() {
  var lastrow = logsheetlast.getRange(1,10).getValue(); //メール配信最終行格納場所
  var sumirow = logsheetlast.getRange(2,10).getValue(); //済記載最終行格納場所
  
  //メールテンプレートH列使用
  var mailSubjectH = tempsheet.getRange(13,8).getValue();
  var mailbodyH = tempsheet.getRange(16,8).getValue();
  
  // 最終行から1行ずつ上の行を参照
  for (var i = lastrow; i > sumirow; i--) {
    var sumi = logsheetlast.getRange("G" + i).getValue(); //済記載チェック
    var kanriadress = logsheetlast.getRange("C" + i).getValue();
    var kanriname  = logsheetlast.getRange("D" + i).getValue();
    var bkprojectname = logsheetlast.getRange("B" + i).getValue();
     if(sumi != "済"){
       var mailbody = mailbodyH.replace('${"管理者名"}',kanriname).replace('${"プロジェクト"}',bkprojectname);
       GmailApp.sendEmail(kanriadress,mailSubjectH,mailbody,{from: kanrisha});
     }
    //メール送信したら「済」にする
    logsheetlast.getRange(i,7).setValue('済');
  }
}