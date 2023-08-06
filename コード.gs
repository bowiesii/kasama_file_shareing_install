//フォーム投稿トリガー
//function setTrigger() {
//var file = FormApp.openById("1pXlrYwdyEQIR7pnT9ufmhBiAHJ0PzzF5pdf6ocUyPo4");//登録：済
//var functionName = "run_reg"; //トリガーを設定したい関数名
//var file = FormApp.openById("1cuFBbuZ61EIVUqFYQq9NCJR_TuxakoDSOVgjz1NCfpw");//登録解除；済
//var functionName = "run_unreg"; //トリガーを設定したい関数名
//ScriptApp.newTrigger(functionName).forForm(file).onFormSubmit().create();//onSubmitにする
//}

//ファイル共有リストスプシ編集トリガー
//function setTrigger2() {
//var file = SpreadsheetApp.openById("1TZ8pjp3Tc6M0BvoshszIGudfAIL4IBdmp4OUSMZGSHg");//ファイル共有リスト：済
//var functionName = "share_set"; //トリガーを設定したい関数名
//ScriptApp.newTrigger(functionName).forSpreadsheet(file).onEdit().create();//onEditにする
//}

//登録フォーム
//onformsubmitのはたらき
function run_reg(e) {

  //本日日付定義
  //var today = new Date("2023/7/2");
  //やっぱり本当の日付
  var today = new Date();
  var year = today.getFullYear();
  var month = today.getMonth() + 1;
  var date = today.getDate();
  //曜日付加
  var w_ary = ["日", "月", "火", "水", "木", "金", "土"];
  var w_jpn = w_ary[today.getDay()];
  var today_ymd = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd');
  var today_ymdd = today_ymd + " " + w_jpn;
  Logger.log("本日は " + today + " " + today_ymdd);

  var answer = e.response.getItemResponses();
  var simei = answer[0].getResponse();
  var email = e.response.getRespondentEmail();
  Logger.log(simei + " " + email);

  //もしブロックする場合はこのへんで弾く
  //if(email == "***"){
  //return;
  //}

  if (simei.length <= 1 || simei.length >= 9) {
    Logger.log("氏名文字数エラー");
    return;
  }

  //共有リストスプシ
  var sheet_file = SpreadsheetApp.openById('1TZ8pjp3Tc6M0BvoshszIGudfAIL4IBdmp4OUSMZGSHg');
  var sheet = sheet_file.getSheetByName('シート1');
  const lastrow = sheet.getLastRow();//行数
  const lastcol = sheet.getLastColumn();//列数

  //リストにすでにemailが存在したら氏名を変えるだけで終了
  var values3 = sheet.getRange(1, 3, lastrow, 1).getValues();//３列目を読む
  for (var row = 2; row <= lastrow - 1; row++) {
    if (values3[row][0] == email) {

      var simei_old = sheet.getRange(row + 1, 2).getValue();
      sheet.getRange(row + 1, 2).setValue(simei);
      Logger.log("氏名変更 " + simei_old + ">" + simei);

      sendmail(email, simei, 5, simei_old);
      sendmail(email, simei, 2, simei_old);
      return;
    }
  }

  //（リストにメールが存在せず、なおかつ）リストに同じ氏名が存在していたらスルー→メール
  var values_simei = sheet.getRange(1, 2, lastrow, 1).getValues();//２列目を読む
  for (var row = 2; row <= lastrow - 1; row++) {
    if (values_simei[row][0] == simei) {
      Logger.log("氏名存在 " + simei);

      sendmail(email, simei, 1);
      return;
    }
  }

  var values2 = sheet.getRange(2, 1, 1, lastcol).getValues();//２行目を読む
  //４列目以降、２行目と同じ値を書き込み
  var insertvalues = [today_ymd, simei, email];
  for (var col = 3; col <= lastcol - 1; col++) {
    insertvalues[col] = values2[0][col];
  }
  //共有リストの３行目に最新を挿入
  sheet.insertRowBefore(3);
  //２次元配列でないと書き込めないらしいので２次元化
  var insertvalues2 = [insertvalues];
  Logger.log(insertvalues2);
  Logger.log("長さ＞" + insertvalues2[0].length);
  sheet.getRange(3, 1, 1, insertvalues2[0].length).setValues(insertvalues2);

  //３行目を適用
  share_setrow(3, "適用");
  sendmail(email, simei, 3);

}

//登録解除フォーム
//onformsubmitのはたらき
function run_unreg(e) {

  var email = e.response.getRespondentEmail();
  Logger.log(email);

  //共有リストスプシ
  var sheet_file = SpreadsheetApp.openById('1TZ8pjp3Tc6M0BvoshszIGudfAIL4IBdmp4OUSMZGSHg');
  var sheet = sheet_file.getSheetByName('シート1');
  const lastrow = sheet.getLastRow();//行数
  const lastcol = sheet.getLastColumn();//列数

  //すべて共有設定を制限
  var values1 = sheet.getRange(1, 1, 1, lastcol).getValues();//１行目を読む
  for (var col = 3; col <= lastcol - 1; col++) {
    var setting = "制限";
    var file_url = values1[0][col];
    //共有設定を実行
    set_any(file_url, email, setting);
  }
  Logger.log("共有設定をすべて制限にしました");

  //リストからemailを探す
  var values3 = sheet.getRange(1, 1, lastrow, 3).getValues();//１～３列目を読む
  for (var row = 2; row <= lastrow - 1; row++) {
    if (values3[row][2] == email) {

      //解除日の情報
      var today = new Date();
      var today_ymd = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd');

      //共有リストから削除→別シートへ移動
      var hozon = [values3[row][0], values3[row][1], values3[row][2], today_ymd];
      var sheet2 = sheet_file.getSheetByName('登録解除した人');
      sheet2.appendRow(hozon);
      sheet.deleteRow(row + 1);

      //ユーザープロパティを削除（メール、氏名）
      var properties = PropertiesService.getUserProperties();
      properties.setProperty('email', 'REM');
      properties.setProperty('simei', 'REM');
      var p_email = properties.getProperty('email');
      var p_simei = properties.getProperty('simei');
      Logger.log("prop " + p_email + " " + p_simei);

      Logger.log("リストからemailを削除し、別シートに移動しました");
      sendmail(email, values3[row][1], 4);

      return;

    } else {
      if (row == lastrow - 1) {
        Logger.log("emailがリストにないよ");
      }
    }
  }

}


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
  var sheet_file = SpreadsheetApp.openById('1TZ8pjp3Tc6M0BvoshszIGudfAIL4IBdmp4OUSMZGSHg');
  var sheet = sheet_file.getSheetByName('シート1');
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

  var sheet_file = SpreadsheetApp.openById('1TZ8pjp3Tc6M0BvoshszIGudfAIL4IBdmp4OUSMZGSHg');
  var sheet = sheet_file.getSheetByName('シート1');
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

//指定の１セルのみ共有設定を適用する。
//共有リストが★手動編集★された場合にトリガーされる。（灰色はスルー）
//なお行・列の削除ではoneditは実行されないらしい（ファイルや人を手動で削除する場合のため参考。）
//★注意：多数のセルにコピーして同時に入力するような場合はoneditは反応しない。
function share_set(e) {

  var sheetN = e.source.getSheetName();
  //シート名がシート1じゃなかったらスルー
  if (sheetN != "シート1") { return; }
  var sheet = e.source.getActiveSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var bgc = sheet.getRange(row, col).getBackground();
  //編集されたセルが灰色だった場合スルー
  if (bgc == "#b7b7b7") { return; }
  //値の変更が無ければスルー
  if (e.value == e.oldValue) { return; }

  var setting = e.value;
  var email = sheet.getRange(row, 3).getValue();
  var file_url = sheet.getRange(1, col).getValue();

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

//opt=1 氏名存在
//opt=2 氏名変更(simei_old要)
//opt=3 登録→youseimale
//opt=4 登録解除→youseimale
//opt=5 氏名変更→youseimale(simei_old要)
function sendmail(address, simei, opt, simei_old) {

  var subject = "";
  var body = "";

  if (opt == 1) {

    subject = '同じ氏名で登録されているようです'; //件名
    body = `同じ氏名「` + simei + `」で登録している人が居るみたいなので、、ファイル共有登録を行いませんでした。

１，既にあなたが（ほかのGoogleアカウントで）共有登録済み、もしくは
２，他のスタッフとたまたま同じ氏名
の可能性があります。

１，２いずれの場合でも、他の氏名を指定することで共有登録することが可能です。

現在登録されている氏名のリスト
https://docs.google.com/spreadsheets/d/1topK0mauvhf4BptUYa7CsAx9RKiUTQmBcoDcJLH9X5g/edit#gid=0

ファイル共有登録フォーム
https://docs.google.com/forms/d/e/1FAIpQLSexh7ngMQJqgerMn4OK3QFNwTFKLCMilmEWj4dmp1MS7vwi5Q/viewform

※このメールは自動配信です。
`;

  } else if (opt == 2) {

    subject = '氏名を変更しました。'; //件名
    body = `お使いのGoogleアカウントに紐付けて登録されていた氏名「` + simei_old + `」を、「` + simei + `」に変更しました。
笠間店の共有ファイルは引き続き使用できます。

※このメールは自動配信です。
`;

  } else if (opt == 3) {//登録→youseimale

    subject = simei + "さんが笠間店ファイル共有に登録したようです"; //件名
    body = simei + "さん（" + address + "）が登録。";
    address = "youseimale@gmail.com";//送信先はyouseimaleに強制変更

  } else if (opt == 4) {//登録解除→youseimale

    subject = simei + "さんが笠間店ファイル共有登録を解除したようです"; //件名
    body = simei + "さん（" + address + "）が登録解除。";
    address = "youseimale@gmail.com";//送信先はyouseimaleに強制変更

  } else if (opt == 5) {//氏名変更→youseimale

    subject = simei_old + "さんが、" + simei + "さんに氏名変更登録したようです"; //件名
    body = simei_old + "さんが、" + simei + "さん（" + address + "）に氏名変更。";
    address = "youseimale@gmail.com";//送信先はyouseimaleに強制変更

  } else {
    Logger.log("opt指定エラー");
    return;
  }

  MailApp.sendEmail(address, subject, body);

}


//---------------------------------------------------------------------------------------------
//以下パクリ


// URLからファイルを取得する
function getFileByUrl(url) {
  const info = getIdAndResourcekeyByUrl(url, false)
  if (info['resourcekey']) {
    return DriveApp.getFileByIdAndResourceKey(info['id'], info['resourcekey'])
  } else {
    return DriveApp.getFileById(info['id'])
  }
}

// アイテム情報オブジェクトを取得する
function getIdAndResourcekeyByUrl(url, isFolder = true) {
  return {
    'id': getIdByUrl(url, isFolder),
    'resourcekey': getQueryParamsByUrl(url)['resourcekey']
  }
}

// URLからアイテムIDを取得する
function getIdByUrl(url, isFolder = true) {
  if (!url || url === '') {
    throw Error('無効なURL')
  }

  // スラッシュでURLを分割
  const splitedUrl = url.split('/')
  // idの前に来る特定の文字列
  let searchString = 'd'
  if (isFolder) {
    searchString = 'folders'
  }

  let id = ''
  for (let i = 0; i < splitedUrl.length; i++) {
    // 特定の文字列に一致する場合はidを取得
    if (splitedUrl[i] === searchString && splitedUrl[i + 1]) {
      id = splitedUrl[i + 1]
      break
    }
  }

  // クエリパラメータは除去
  return id.split('?')[0]
}

// URLからクエリパラメータを取得
function getQueryParamsByUrl(url) {
  const params = {}

  if (url.split('?').length < 0) {
    // クエリパラメータがない
    return params
  }
  // クエリパラメータの文字列を取得
  const queryUrl = url.split('?')[1]
  if (queryUrl) {
    // パラメータ毎にキーと値を抽出
    const queryRawParams = queryUrl.split('&')
    queryRawParams.forEach(function (value, index) {
      const kv = value.split('=')
      params[kv[0]] = kv[1]
    })
    return params
  }
  // 全てのキーと値を持ったオブジェクトを返却
  return params
}
