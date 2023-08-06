//本日日付定義
//const today = new Date("2023/7/2");
//やっぱり本当の日付
const today = new Date();
const year = today.getFullYear();
const month = today.getMonth() + 1;
const date = today.getDate();
//曜日付加
const w_ary = ["日", "月", "火", "水", "木", "金", "土"];
const w_jpn = w_ary[today.getDay()];
const today_ymd = Utilities.formatDate(today, 'JST', 'yyyy/MM/dd');
const today_ymdd = today_ymd + " " + w_jpn;
const today_hm = Utilities.formatDate(today, 'JST', 'HH:mm');
const today_ymddhm = today_ymdd + " " + today_hm;

//共有リスト
const sheet_file = SpreadsheetApp.openById("1TZ8pjp3Tc6M0BvoshszIGudfAIL4IBdmp4OUSMZGSHg");
const sheet = sheet_file.getSheetByName("登録している人");
const sheet2 = sheet_file.getSheetByName("登録解除した人");
