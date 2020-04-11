/*
 * *********************************************************
 * Copyright (c) 2020 Ken Hashi
 * Released under the MIT license
 * https://opensource.org/licenses/mit-license.php
 * *********************************************************
 */
//===========================================================
var debug = true;
var spApp = openSpreadsheet();
var usApp = Underscore.load();

/*
 * Debug用
 */
function debugTest() {
  //postSlackMessage('slack.g5hnsfwv','おは');
}
/*
 * ログ出力
 */
function writeLog(log_type, text) {
  // デバッグ出力しないときおわり
  if (!debug && log_type == 'debug') return;

  if (spApp) {
    // Logシート
    var sheet = spApp.getSheetByName('Log');
    if (sheet) {
      var lastRow = sheet.getLastRow();

      sheet.getRange(lastRow+1,1).setValue(new Date());
      sheet.getRange(lastRow+1,2).setValue(log_type);
      sheet.getRange(lastRow+1,3).setValue(text);
    }
  }
}
//===========================================================
var url = '';
var xlsName = '';
var xlsBlob;

/*
 * Postリクエスト処理
 */
function doGet(e){ // e にPOSTされたデータが入っている
  try {
    //writeLog('debug', e);
    //execQue(e.parameter.row);
    var user  = e.parameter.user;
    var year  = e.parameter.year;
    var month = e.parameter.month;
    viewKintai(user, Number(year), Number(month));
  } catch(err) {
    writeLog('error', err.message);
    return ContentService.createTextOutput(err.message);
  }
  var template = 'index';
  return HtmlService.createTemplateFromFile(template).evaluate();
}
/*
 * 外部ファイル取り込み
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
/*
 * HTMLに渡す
 */
function getUrl() {
  return url;
}
function getXlsName() {
  return xlsName;
}
function getXlsBlob() {
  writeLog('debug', xlsBlob);
  return xlsBlob;
}

/*
 * スプレッドシート開く
 */
function openSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREAD_SHEET_ID');
  return SpreadsheetApp.openById(id);
}

/*
 * que実行
 */
function execQue(row) {
  // queシート
  var sheet = spApp.getSheetByName('que');
  if (sheet) {
    var func_name = sheet.getRange(row,2).getValue();
    var func_para = sheet.getRange(row,3).getValue();
    func_para = func_para.split(',');
    if (func_name == 'viewKintai') {
      return viewKintai(func_para[0], Number(func_para[1]), Number(func_para[2]));
    }
  }
  return '処理済です。';
}

/*
 * １ヶ月分の勤務表
 */
function viewKintai(user, year, month) {
  //writeLog('debug', user + ':' + year + '/' + month);
  // スプレッドシートを開く
  //spApp = openSpreadsheet();
  // ユーザー情報取得
  var userInfo = readUserinfo(user);
  if (!userInfo) {
    throw new Error(user + ':ユーザー情報がありません。');
  }
  // ユーザーの勤怠シート
  var sheet = spApp.getSheetByName(userInfo[1]);
  if (!sheet) {
    throw new Error(userInfo[1] + ':勤怠情報がありません。');
  }
  // 作業用シート(前回の月報削除)
  var wkSheetList = spApp.getSheets();
  var wkSheet;
  var regex = new RegExp('^' + userInfo[1] + '_\\d+年\\d+月');
  for (var i in wkSheetList) {
    wkSheet = wkSheetList[i];
    var sheetName = wkSheet.getName();
    if (sheetName.match(regex)) {
      spApp.deleteSheet(wkSheet);
    }
  }
  // 雛型をコピーして新規作成
  var temp = spApp.getSheetByName('template');
  wkSheet = temp.copyTo(spApp);
  wkSheet.setName(userInfo[1] + '_' + year + '年' + month + '月');

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var aryKintai = sheet.getRange(2,1,lastRow,lastCol).getValues();
  var aryTrns = usApp.zip.apply(usApp, aryKintai);
  var aryDate = new Array();
  for (var i in aryTrns[0]) {
    if (aryTrns[0][i]) {
      aryDate.push(formatDate(aryTrns[0][i]));
    }
  }

  var wkDate     = new Date(year, month-1, 1);
  var wkDateStr  = formatDate(wkDate);
  var endDateStr = formatDate(new Date(year, month, 0));
  var wkDay      = wkDate.getDay();
  var aryWeek    = [ "日", "月", "火", "水", "木", "金", "土" ];
  //writeLog('debug', 'viewKintai:' + wkDateStr + '-' + endDateStr);
  var wkIdx = 2;
  var section = '';
  var fields  = '';
  while (wkDateStr <= endDateStr) {
    //var idxRow = findDateRow(sheet, wkDateStr);
    var idxRow = aryDate.indexOf(wkDateStr);
    var wkTime;
    var startTime = '';
    var endTime   = '';
    var breakTime = '';
    var holiday   = '';
    var comment   = '';
    if (idxRow > 0) {
      wkTime = aryTrns[4][idxRow];
      if (wkTime) startTime = formatTime(wkTime);
      wkTime = aryTrns[5][idxRow];
      if (wkTime) endTime   = formatTime(wkTime);
      wkTime = aryTrns[6][idxRow];
      if (wkTime) breakTime = formatTime(wkTime);
      holiday   = aryTrns[8][idxRow];
      comment   = aryTrns[9][idxRow];
    }
    // 西暦
    wkSheet.getRange(wkIdx,1).setValue(wkDateStr);
    // 年
    wkSheet.getRange(wkIdx,2).setValue('=text(A' + wkIdx + ',"yyyy年")');
    // 月
    wkSheet.getRange(wkIdx,3).setValue('=text(A' + wkIdx + ',"mm月")');
    // 日
    wkSheet.getRange(wkIdx,4).setValue('=text(A' + wkIdx + ',"dd(ddd)")');
    // 出勤時刻
    wkSheet.getRange(wkIdx,5).setValue(startTime);
    // 退勤時刻
    wkSheet.getRange(wkIdx,6).setValue(endTime);
    // 休憩時間
    wkSheet.getRange(wkIdx,7).setValue(breakTime);
    // 勤務時間
    wkSheet.getRange(wkIdx,8).setValue('=IF(F' +wkIdx+ '="","",F'+wkIdx+'-E'+wkIdx+'-G'+wkIdx+')');
    // 休暇
    wkSheet.getRange(wkIdx,9).setValue(holiday);
    // 補足
    wkSheet.getRange(wkIdx,10).setValue(comment);

    wkDate.setDate(wkDate.getDate()+1);
    wkDateStr  = formatDate(wkDate);
    wkDay      = wkDate.getDay();
    wkIdx++;
  }
  // 合計
  wkSheet.getRange(wkIdx,7).setValue('合計');
  wkSheet.getRange(wkIdx,8).setValue('=SUM(H2:H' + (wkIdx-1) + ')');
  wkSheet.getRange(wkIdx,8).setNumberFormat('[h]:mm:ss');

  // ダウンロード
  var sid = PropertiesService.getScriptProperties().getProperty('SPREAD_SHEET_ID');
  var gid = wkSheet.getSheetId();
  url = 'https://docs.google.com/spreadsheets/d/' + sid + '/export?gid=' + gid + '&exportFormat=xlsx'
  //url = 'https://docs.google.com/spreadsheets/d/' + sid + '/export?gid=' + gid + '&exportFormat=pdf'
  url += '&access_token=' + ScriptApp.getOAuthToken();
  xlsName = userInfo[1] + '_' + year + '年' + month + '月.xlsx';
  //xlsBlob = UrlFetchApp.fetch(url, {headers: {'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()}}).getBlob().setName(xlsName);
  //xlsBlob = UrlFetchApp.fetch(url, {headers: {'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()}}).getContent();
  //xlsBlob = UrlFetchApp.fetch(url, {headers: {'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()}}).getBlob();
  //DriveApp.createFile(xlsBlob);

  // 作業用シート削除
  //spApp.deleteSheet(wkSheet);
  //return '<a href="' + url + '">こちらからダウンロード</a>';
}

/*
 * ユーザー情報読み込み
 */
function readUserinfo(in_user) {
  // ユーザーシート
  var sheet = spApp.getSheetByName('user');
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  var aryUser = sheet.getRange(2,1,lastRow,lastCol).getValues();
  // slackユーザーの行番号を取得
  //var us = Underscore.load();
  var aryTrns = usApp.zip.apply(usApp, aryUser);
  var rowNum = aryTrns[2].indexOf(in_user);
  if (rowNum < 0) return null;

  return aryUser[rowNum];
}

/*
 * 年月日 フォーマット
 */
function formatDate(date) {
  date = valToDate(date);
  return date.getFullYear() + '/' + ('0' + (date.getMonth() + 1)).slice(-2) + '/' + ('0' + date.getDate()).slice(-2);
}

/*
 * 時分秒フォーマット
 */
function formatTime(date) {
  date = valToDate(date);
  return ('0' + date.getHours()).slice(-2) + ':' + ('0' + date.getMinutes()).slice(-2);
}

/*
 * date変換
 */
function valToDate(val) {
  if (val instanceof Date) {
    return val;
  } else if (!val) {
    return new Date(0);
  } else if (val == '') {
    return new Date(0);
  } else {
    return new Date(val);
  }
}

/*
 * slackにメッセージ送信
 */
function postSlackMessage(user, message) {
  var token = PropertiesService.getScriptProperties().getProperty('OAUTH_ACCESS_TOKEN');
  var slackApp = SlackApp.create(token); //SlackApp インスタンスの取得

  var options = {
    channelId: "#勤怠管理", //チャンネル名
    userName: user,
    message: message
  };

  slackApp.postMessage(options.channelId, options.message, {username: options.userName});
}
