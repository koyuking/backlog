//共通関数
var backlogsheet = SpreadsheetApp.getActive().getSheetByName('backlogプロジェクト管理表');
var tempsheet = SpreadsheetApp.getActive().getSheetByName('メールテンプレート');
var oversheet = SpreadsheetApp.getActive().getSheetByName('容量超過ログ');
var kanrisha = tempsheet.getRange("H2").getValue(); //管理者用メールアドレス設定箇所
var currentDate = new Date(); //今日の日付

function task() {
  // スプレットシート取得
  var mySS = SpreadsheetApp.openById("******");
  // スプレットシートの書き込む位置
  var range= mySS.getSheetByName("data").getRange(2, 1);
  // APIキーでBacklog認証&取得
  var diskusage = UrlFetchApp.fetch("https://*****.backlog.com/api/v2/space/diskUsage?apiKey=******"); 
  
  if (diskusage.getResponseCode() != 200) {
    return false;
  }
  var datelist = JSON.parse(diskusage.getContentText());
  
  // 一旦、シートをクリアにする 
  mySS.getSheetByName("data").getRange("A2:F"+datelist.details.length+1).clearContent();
  
  
  for(var i = 0; i<datelist.details.length; i++) {
    // スプレッドシートに書き込む
    range.offset(i, 0).setValue(datelist["details"][i]["projectId"]);
    range.offset(i, 1).setValue(datelist["details"][i]["issue"]);
    range.offset(i, 2).setValue(datelist["details"][i]["wiki"]);    
    range.offset(i, 3).setValue(datelist["details"][i]["file"]);
    range.offset(i, 4).setValue(datelist["details"][i]["subversion"]);
    range.offset(i, 5).setValue(datelist["details"][i]["git"]);
  }

}

function usageover() {
  var lastrow = 7 + backlogsheet.getRange(4,4).getValue();
  
  //メールテンプレートJ列使用
  var mailSubjectJ = tempsheet.getRange(13,10).getValue();
  var mailbodyJ = tempsheet.getRange(16,10).getValue();
  
  // 容量超過プロジェクトチェック
  for (var i = 7; i < lastrow ; i++) {
    // 1GB越えのプロジェクトにメール
    if( backlogsheet.getRange("o"+ i).getValue() >= 1000000000){
      var kanriadress = backlogsheet.getRange("E" + i).getValue();
      var kanriname  = backlogsheet.getRange("F" + i).getValue();
      var bkprojectname = backlogsheet.getRange("D" + i).getValue();
      var usage = backlogsheet.getRange("O" + i).getValue();
      var mailbody = mailbodyJ.replace('${"管理者名"}',kanriname).replace('${"プロジェクト"}',bkprojectname).replace('${"使用量"}',usage);
      GmailApp.sendEmail(kanriadress,mailSubjectJ,mailbody,{from: kanrisha});
      oversheet.appendRow([currentDate, bkprojectname, kanriadress, kanriname, "超過メール"]);
    }
  }
}
