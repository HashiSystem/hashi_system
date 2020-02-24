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
/*
 * Postリクエスト処理
 */
function doPost(e){ // e にPOSTされたデータが入っている
  try {
    //writeLog('debug', e.parameter);

    return postSlackMessage(e.parameter);
  } catch(e) {
    writeLog('error', e.message);
    var json_data = {
      text: e.message
    };
    return ContentService.createTextOutput(JSON.stringify(json_data)).setMimeType(ContentService.MimeType.JSON);
  }
}

/*
 * Slackメッセージ応答
 */
function postSlackMessage(param) {
  var in_userName = param.user_name;
  var in_text     = param.text;
  var json_data;
  
  // ユーザー情報取得
  var userInfo = readUserinfo(in_userName);
  if (!userInfo) {
    writeLog('info', in_userName + ':ユーザー情報がありません');
    var out_text = in_userName + "さんのユーザー情報がありません";
    json_data = {
      text: out_text
    };
  } else {
    // 勤怠入力表示
    var initial_date = formatDate2(new Date(), '-');
    json_data =
    {
        "blocks": [
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": " "
                },
                "accessory": {
                    "type": "button",
                    "action_id": "timecard",
                    "text": {
                        "type": "plain_text",
                        "text": "出退勤の登録",
                        "emoji": true
                    }
                }
            },
            {
                "type": "section",
                "text": {
                    "type": "mrkdwn",
                    "text": " "
                },
                "accessory": {
                    "type": "button",
                    "action_id": "geppou",
                    "text": {
                        "type": "plain_text",
                        "text": "月報の出力",
                        "emoji": true
                    }
                }
            }
        ]
    };  
  }
  return ContentService.createTextOutput(JSON.stringify(json_data)).setMimeType(ContentService.MimeType.JSON);
}

/*
 * スプレッドシート開く
 */
function openSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREAD_SHEET_ID');
  return SpreadsheetApp.openById(id);
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
function formatDate2(date, sep) {
  date = valToDate(date);
  return date.getFullYear() + sep + ('0' + (date.getMonth() + 1)).slice(-2) + sep + ('0' + date.getDate()).slice(-2);
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

