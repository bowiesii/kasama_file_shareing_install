//フォーム投稿トリガー
//function setTrigger() {
//var file = FormApp.openById("1pXlrYwdyEQIR7pnT9ufmhBiAHJ0PzzF5pdf6ocUyPo4");//登録：済
//var functionName = "run_reg"; //トリガーを設定したい関数名
//var file = FormApp.openById("1cuFBbuZ61EIVUqFYQq9NCJR_TuxakoDSOVgjz1NCfpw");//登録解除；済
//var functionName = "run_unreg"; //トリガーを設定したい関数名
//ScriptApp.newTrigger(functionName).forForm(file).onFormSubmit().create();//onSubmitにする
//}

//登録フォーム
//onformsubmitのはたらき
function run_reg(e) {

  var email = e.response.getRespondentEmail();
  var answer = e.response.getItemResponses();
  var simei = answer[0].getResponse();
  var pass = answer[1].getResponse();
  Logger.log(simei + " " + email + " " + pass);

  //★もしブロックする場合はこのへんで弾く★
  //if(email == "***"){
  //return;
  //}

  //半角に変換
  pass = bbsLib.convertCharacters(pass);
  Logger.log(pass);
  if (pass != form_password) {
    Logger.log("パスワードエラー");
    MailApp.sendEmail(email, "笠間店より", "パスワードエラー");
    return;
  }

  if (simei.length <= 1 || simei.length >= 9) {
    Logger.log("氏名文字数エラー");
    MailApp.sendEmail(email, "笠間店より", "氏名文字数エラー");
    return;
  }

  //共有リストスプシ
  const lastcol = sheet.getLastColumn();//列数

  var emailEx = emailExist(email);
  var simeiEx = simeiExist(simei);

  if (emailEx != -1) {//リストにすでにemailが存在したら
    if (simeiEx != -1) {//その氏名が存在していたら

      if (simeiEx == emailEx) {
        Logger.log("氏名もemailも変化なし" + simei);
        MailApp.sendEmail(email, "笠間店より", "既に登録されています。");
        return;

      } else {//変えようとした氏名が存在した
        Logger.log("氏名存在 " + simei);
        mail_share(email, simei, 1);//本人通知
        return;

      }

    } else {//氏名を変える
      var simei_old = sheet.getRange(emailEx, 2).getValue();
      sheet.getRange(emailEx, 2).setValue(simei);
      Logger.log("氏名変更 " + simei_old + ">" + simei);
      mail_share(email, simei, 2, simei_old);//本人通知
      mail_share(email, simei, 5, simei_old);//プッシュ通知
      return;

    }
  } else if (simeiEx != -1) {//新規登録しようとしたが氏名が存在した
    Logger.log("氏名存在 " + simei);
    mail_share(email, simei, 1);//本人通知
    return;

  }

  //以下普通の登録作業
  var values2 = sheet.getRange(2, 1, 1, lastcol).getValues();//２行目を読む
  var insertvalues = [today_ymddhm, simei, email];
  for (var col = 3; col <= lastcol - 1; col++) {
    insertvalues[col] = values2[0][col];//４列目以降、２行目と同じ値を書き込み
  }
  sheet.insertRowBefore(3);
  Logger.log(insertvalues);
  sheet.getRange(3, 1, 1, insertvalues.length).setValues([insertvalues]);//共有リストの３行目に最新を挿入

  share_setrow(3, "適用");//３行目を適用
  mail_share(email, simei, 3);//プッシュ通知
  mail_share(email, simei, 6);//本人通知

}


//登録解除フォーム
//onformsubmitのはたらき
function run_unreg(e) {
  var email = e.response.getRespondentEmail();
  Logger.log(email);

  //共有リストスプシ
  const lastcol = sheet.getLastColumn();//列数

  //すべて共有設定を制限
  var values1 = sheet.getRange(1, 1, 1, lastcol).getValues();//１行目を読む
  for (var col = 3; col <= lastcol - 1; col++) {
    var file_url = values1[0][col];
    set_any(file_url, email, "制限");//共有設定を実行
  }
  Logger.log("共有設定をすべて制限にしました");

  //プッシュ通知リストにあれば削除する
  var sheetPushList = bbsLib.getSheetByIdGid(id_bbLog, gid_pushList);//プッシュ登録者リスト
  var listPushRow = bbsLib.searchInCol(sheetPushList, 2, email);
  if (listPushRow != -1) {//登録済み
    sheetPushList.deleteRow(listPushRow);
  }

  //リストからemailを探す→別シートへ移動
  var emailEx = emailExist(email);
  if (emailEx != -1) {
    var hozon = sheet.getRange(emailEx, 1, 1, 3).getValues();
    hozon[0][3] = today_ymddhm;

    bbsLib.addLogFirst(sheet2, 2, hozon, 4, 10000);
    sheet.deleteRow(emailEx);

    Logger.log("リストからemailを削除し、別シートに移動しました");
    mail_share(email, hozon[0][1], 4);//プッシュ通知
    mail_share(email, hozon[0][1], 7);//本人通知

    return;
  } else {
    Logger.log("emailがリストにないよ");
  }

}
