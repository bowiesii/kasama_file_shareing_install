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
  const lastcol = sheet.getLastColumn();//列数

  var emailEx = emailExist(email);
  var simeiEx = simeiExist(simei);

  if (emailEx != -1) {//リストにすでにemailが存在したら
    if (simeiEx != -1) {//その氏名が存在していたら
      Logger.log("氏名存在 " + simei);
      sendmail(email, simei, 1);
      return;
    } else {//氏名を変えるだけで終了

      var simei_old = sheet.getRange(emailEx, 2).getValue();
      sheet.getRange(emailEx, 2).setValue(simei);
      Logger.log("氏名変更 " + simei_old + ">" + simei);
      sendmail(email, simei, 5, simei_old);
      sendmail(email, simei, 2, simei_old);
      return;

    }
  } else if (simeiEx != -1) {//リストにメールが存在せず、なおかつリストに同じ氏名が存在していたらスルー→メール
    Logger.log("氏名存在 " + simei);
    sendmail(email, simei, 1);
    return;
  }

  var values2 = sheet.getRange(2, 1, 1, lastcol).getValues();//２行目を読む
  var insertvalues = [today_ymddhm, simei, email];
  for (var col = 3; col <= lastcol - 1; col++) {
    insertvalues[col] = values2[0][col];//４列目以降、２行目と同じ値を書き込み
  }
  sheet.insertRowBefore(3);
  Logger.log(insertvalues);
  sheet.getRange(3, 1, 1, insertvalues.length).setValues([insertvalues]);//共有リストの３行目に最新を挿入

  share_setrow(3, "適用");//３行目を適用
  sendmail(email, simei, 3);

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
    //共有設定を実行
    set_any(file_url, email, "制限");
  }
  Logger.log("共有設定をすべて制限にしました");

  //リストからemailを探す→別シートへ移動
  var emailEx = emailExist(email);
  if (emailEx != -1) {

    var hozon = sheet.getRange(emailEx, 1, 1, 3).getValues();
    hozon[0][3] = today_ymddhm;
    sheet2.appendRow(hozon[0]);//appendrowは１次元配列
    sheet.deleteRow(emailEx);

    Logger.log("リストからemailを削除し、別シートに移動しました");
    sendmail(email, hozon[0][1], 4);

    return;

  } else {
    Logger.log("emailがリストにないよ");
  }

}
