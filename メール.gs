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
