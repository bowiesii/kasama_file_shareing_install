//opt=1 氏名存在
//opt=2 氏名変更(simei_old要)
//opt=3 登録→★bot
//opt=4 登録解除→★bot
//opt=5 氏名変更→★bot(simei_old要)
//opt=6 登録
//opt=7 登録解除
function mail_share(address, simei, opt, simei_old) {

  var subject = "";
  var body = "";

  if (opt == 1) {
    subject = '【笠間店】同じ氏名が登録されているようです'; //件名
    body = `同じ氏名「` + simei + `」で既に共有登録している人が居るため、、ファイル共有登録を行えませんでした。

１，既にあなたが（ほかのGoogleアカウントで）共有登録済み、もしくは
２，他のスタッフとたまたま同じ氏名
の可能性があります。

１，２いずれの場合でも、他の氏名を指定することで共有登録することが可能です。

現在登録されている氏名のリスト
https://docs.google.com/spreadsheets/d/1topK0mauvhf4BptUYa7CsAx9RKiUTQmBcoDcJLH9X5g/edit#gid=0

ファイル共有登録フォーム
https://docs.google.com/forms/d/e/1FAIpQLSexh7ngMQJqgerMn4OK3QFNwTFKLCMilmEWj4dmp1MS7vwi5Q/viewform

問題が発生しましたら、浦野（youseimale@gmail.com）まで連絡をお願いします。
※このメールは自動配信です。
`;

  } else if (opt == 2) {
    subject = '【笠間店】氏名を変更しました。'; //件名
    body = `お使いのGoogleアカウントに紐付けて登録されていた氏名「` + simei_old + `」を、「` + simei + `」に変更しました。
こちらで確認することが出来ます。
https://docs.google.com/spreadsheets/d/1topK0mauvhf4BptUYa7CsAx9RKiUTQmBcoDcJLH9X5g/edit#gid=0

笠間店の共有ファイルは引き続き使用できます。

なお、この氏名は「基本バインダー類」ファイル内シート（週タスク、鮮度、清掃など）の編集時にログ記録される氏名とは異なります。
ログ記録される氏名を変更したい場合は、「基本バインダー類」ファイル内シート内の、実行者氏名登録欄にて変更をお願いします。

問題が発生しましたら、浦野（youseimale@gmail.com）まで連絡をお願いします。
※このメールは自動配信です。
`;

  } else if (opt == 3) {//登録→★bot
    subject = "笠間店ファイル共有登録通知"; //件名
    body = simei + "さんが登録。";
    body = body + "\n※この氏名はファイル編集時にログ記録される「実行者氏名」とは異なる事があります。";
    address = "bot";

  } else if (opt == 4) {//登録解除→★bot
    subject = "笠間店ファイル共有解除通知"; //件名
    body = simei + "さんが登録解除。";
    body = body + "\n※この氏名はファイル編集時にログ記録される「実行者氏名」とは異なる事があります。";
    address = "bot";

  } else if (opt == 5) {//氏名変更→★bot
    subject = "笠間店共有登録氏名変更通知"; //件名
    body = simei_old + "さんが、" + simei + "さんに氏名変更。";
    body = body + "\n※この氏名はファイル編集時にログ記録される「実行者氏名」とは異なる事があります。";
    address = "bot";

  } else if (opt == 6) {//登録
    subject = '【笠間店】ファイル共有登録ありがとうございます。'; //件名
    body = simei + `さんのGoogleアカウントで笠間店のファイル共有を行えるように設定します。

なおこの笠間店スタッフ用LINEbotを友だち登録して頂くと情報アクセスに関して便利です。
https://liff.line.me/1645278921-kWRPP32q/?accountId=586bfamf
botの使用方法等についてはあいさつ文をご覧下さい。

【時間差について】
パソコンだと直ぐにファイルを編集できますが、スマホアプリ（Googleスプレッドシート）は、権限が反映されるまで１時間程かかる場合があるようです。

【アプリの取得について】
共有ファイルの編集にはGoogleスプレッドシートアプリが必要です。
Google Play
https://play.google.com/store/apps/details?id=com.google.android.apps.docs.editors.sheets&hl=ja&gl=US
App Store
https://apps.apple.com/jp/app/google-%E3%82%B9%E3%83%97%E3%83%AC%E3%83%83%E3%83%89%E3%82%B7%E3%83%BC%E3%83%88/id842849113

【実行者氏名登録について】
編集権限の反映後、「基本バインダー類」ファイル内の「発注」以外のシートの使用においては、再度シート内にて実行者氏名の登録が必要になります。
例えば、この「鮮度」シートですと１行H列が実行者氏名入力欄になります。
https://docs.google.com/spreadsheets/d/1sEKCFs6oNzbEkRgt2Z2aq_4mOGQXMU7dcFTXPNYf-wg/edit#gid=1405667253
先ほどの入力フォームで入力して頂いた氏名と完全に同一である必要はありません。
この入力作業は１度行えばOKです。

詳しくは、こちらの文書の【実行者氏名登録】章で把握をお願いします。
https://docs.google.com/document/d/19PtDzWvegrorxDipjWGl1XyUdC7phHL2bIYAeKhNBIg/edit#heading=h.yqw7j64q400e

※前述のLINEbotやスプレッドシートの使用方法等も上の文書に書かれています。

問題が発生しましたら、浦野（youseimale@gmail.com）まで連絡をお願いします。
※このメールは自動配信です。
`;

  } else if (opt == 7) {//登録解除
subject = '【笠間店】ファイル共有登録を解除しました。'; //件名
body = simei + `さんのGoogleアカウントで笠間店のファイル共有を解除します。

問題が発生しましたら、浦野（youseimale@gmail.com）まで連絡をお願いします。
※このメールは自動配信です。
`;

  } else {
    Logger.log("opt指定エラー");
    return;
  }

  //メールor通知
  if (address == "bot") {
    botLib.pushSB(subject, body);
  } else {
    MailApp.sendEmail(address, subject, body);
  }

}
