//指定★行のみ共有設定を適用する(i行目（３，４・・・）)
//j = "適用"なら表のとおりに適用、それ以外なら入力の通り。（jのオプは★適用、編集、閲覧、制限の４種類。）
//新しい人が入った場合・退出した場合に使う
//ただしプロパティはこのファンクションでは操作しない
function share_setrow(i, j) {

  Logger.log(i + "行を、" + j + "にする。");
  if (i <= 2) {
    Logger.log("行数指定エラー");
    return;
  }

  //共有リストスプシ
  const lastcol = sheet.getLastColumn();//列数（１，２・・・）
  var values1 = sheet.getRange(1, 1, 1, lastcol).getValues();//１行目を読む
  var valuesi = sheet.getRange(i, 1, 1, lastcol).getValues();//i行目を読む

  var email = valuesi[0][2];//メール
  Logger.log(email);

  for (var col = 3; col <= valuesi[0].length - 1; col++) {//列（[3]以降）

    var file_url = values1[0][col];//ファイルのURL
    Logger.log(file_url);
    var setting = j;
    if (setting == "適用") {
      setting = valuesi[0][col];//編集、閲覧、制限
    }

    //共有設定を実行
    set_any(file_url, email, setting);

  }

  Logger.log(i + "行を、" + j + "にする。＞完了。");

}


//指定★列のみ共有設定を適用する。(i列目（４，５・・・）)
//j = "適用"なら表のとおりに適用、それ以外なら入力の通り。（jのオプは適用、編集、閲覧、制限の４種類。）
//新しいファイルができた場合に使う（新人教育ファイル？）
function share_setcol(i, j) {

  Logger.log(i + "列を、" + j + "にする。");
  if (i <= 3) {
    Logger.log("列数指定エラー");
    return;
  }

  const lastrow = sheet.getLastRow();//行数（１，２・・・）

  var values1 = sheet.getRange(1, 3, lastrow, 1).getValues();//３列目を読む（メールアドレス）
  var valuesi = sheet.getRange(1, i, lastrow, 1).getValues();//i列目を読む

  var file_url = valuesi[0][0];//ファイルのURL
  Logger.log(file_url);

  for (var row = 2; row <= valuesi.length - 1; row++) {

    var email = values1[row][0];
    var setting = j;
    if (setting == "適用") {
      setting = valuesi[row][0];//編集、閲覧、制限
    }

    //共有設定を実行
    set_any(file_url, email, setting);

  }

  Logger.log(i + "列を、" + j + "にする。＞完了。");

}


//ファイル共有リストスプシ編集トリガー
//function setTrigger2() {
//var file = SpreadsheetApp.openById("1TZ8pjp3Tc6M0BvoshszIGudfAIL4IBdmp4OUSMZGSHg");//ファイル共有リスト：済
//var functionName = "share_set"; //トリガーを設定したい関数名
//ScriptApp.newTrigger(functionName).forSpreadsheet(file).onEdit().create();//onEditにする
//}
//指定の１セルのみ共有設定を適用する。
//共有リストが★手動編集★された場合にトリガーされる。（灰色はスルー）
//なお行・列の削除ではoneditは実行されないらしい（ファイルや人を手動で削除する場合のため参考。）
//★注意：多数のセルにコピーして同時に入力するような場合はoneditは反応しない。
function share_set(e) {

  var sheetN = e.source.getSheetName();
  if (sheetN != "登録している人") { return; }//シート名が登録している人じゃなかったらスルー
  var sheetAct = e.source.getActiveSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var bgc = sheetAct.getRange(row, col).getBackground();
  if (bgc == "#b7b7b7") { return; }//編集されたセルが灰色だった場合スルー
  if (e.value == e.oldValue) { return; }//値の変更が無ければスルー

  var setting = e.value;
  var email = sheetAct.getRange(row, 3).getValue();
  var file_url = sheetAct.getRange(1, col).getValue();

  //共有設定を実行
  set_any(file_url, email, setting);

}


//任意の人の、任意のファイルへの権限を、任意に設定する。ファイルURL・メール・権限（編集・閲覧・制限）
function set_any(file_url, email, setting) {

  var file = getFileByUrl(file_url);//下のほうのパクリのやつを使ってる
  var file_name = file.getName();

  if (setting == "編集") {

    file.addEditor(email);
    Logger.log(file_name + " " + email + "＞編集");

  } else if (setting == "閲覧") {

    //editorリストにもしあれば削除する
    var editors = file.getEditors();
    for (let i in editors) {
      var e_email = editors[i].getEmail();
      if (e_email == email) {
        file.removeEditor(email);
        Logger.log("rem-e");
        break;
      }
    }

    file.addViewer(email);
    Logger.log(file_name + " " + email + "＞閲覧");

  } else if (setting == "制限") {

    //editorリストにもしあれば削除する（editorリストとviewerリストは重複しないらしい）
    var editors = file.getEditors();
    for (let i in editors) {
      var e_email = editors[i].getEmail();
      if (e_email == email) {
        file.removeEditor(email);
        Logger.log("rem-e");
        break;
      }
    }
    //viewerリストにもしあれば削除する
    //なお「一般アクセス」が閲覧のみだった場合はviewerリストには無いので以下は実行されず、ユーザーは閲覧できる。
    var viewers = file.getViewers();
    for (let i in viewers) {
      var v_email = viewers[i].getEmail();
      if (v_email == email) {
        file.removeViewer(email);
        Logger.log("rem-v");
        break;
      }
    }

    Logger.log(file_name + " " + email + "＞制限");

  } else {
    Logger.log(file_name + " " + email + "＞権限指定エラー：" + setting);
  }

}


//ファイルオブジェクトを入力→すべての個別アカウントの権限を削除（★一般アクセスはいじらない）
function set_remall(file) {

  //editorリストにもしあれば削除する（editorリストとviewerリストは重複しないらしい）
  var editors = file.getEditors();
  for (let i in editors) {
    var email = editors[i].getEmail();
    file.removeEditor(email);
    Logger.log("rem-e " + email);
  }

  //viewerリストにもしあれば削除する
  //なお「一般アクセス」が閲覧のみだった場合はviewerリストには無いので以下は実行されず、ユーザーは閲覧できる。
  var viewers = file.getViewers();
  for (let i in viewers) {
    var email = viewers[i].getEmail();
    file.removeViewer(email);
    Logger.log("rem-v " + email);
  }

}


//ファイル共有リストに氏名が存在しているかどうか
function simeiExist(simei) {

  var lastrow = sheet.getLastRow();
  var ary = sheet.getRange(3, 2, lastrow - 2, 1).getValues();//２列目を読む

  for (let row = 0; row <= ary.length - 1; row++) {
    if (ary[row][0] == simei) {
      return row + 3;//氏名のシート行を返す（３～）
    }
  }

  return -1;//存在しないない→-1を返す
}


//ファイル共有リストにメールアドレスが存在しているかどうか
function emailExist(email) {

  var lastrow = sheet.getLastRow();
  var ary = sheet.getRange(3, 3, lastrow - 2, 1).getValues();//３列目を読む

  for (let row = 0; row <= ary.length - 1; row++) {
    if (ary[row][0] == email) {
      return row + 3;//アドレスのシート行を返す（３～）
    }
  }

  return -1;//存在しないない→-1を返す
}