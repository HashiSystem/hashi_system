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
/*
 * Debug用
 */
function debugTest() {
  // デバッグ用
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

/*
 * Postリクエスト処理
 */
function doPost(e){ // e にPOSTされたデータが入っている
  try {
    var payload = JSON.parse(e["parameter"]["payload"]);
    //writeLog('debug', payload);
    
    var user = payload["user"]["name"];
    var type = payload["type"];
    if (type == 'block_actions') {
      // modal表示
      var action_id = payload["actions"][0]["action_id"];
      var trigger = payload["trigger_id"];
      var token = PropertiesService.getScriptProperties().getProperty('OAUTH_ACCESS_TOKEN');
      var json_data;
      var slackUrl;
      var title = '';
      var initial_date = '';
      if (action_id == 'timecard') {
        title = '勤怠入力';
        initial_date = formatDate(new Date());
        var view_json = makeKintaiView(title, user, initial_date);
        json_data = {
          "method": "post",
          "payload": {
            "token": token, // OAuth_token
            "trigger_id":  trigger,
            "view": JSON.stringify(view_json)
          }
        };
        slackUrl = "https://slack.com/api/views.open";
        var response = UrlFetchApp.fetch(slackUrl, json_data);
      } else if (action_id == 'geppou') {
        title = '月報';
        initial_date = formatDate(new Date());
        var view_json = makeGeppouView(title, user, initial_date, trigger);
        json_data = {
          "method": "post",
          "payload": {
            "token": token, // OAuth_token
            "trigger_id":  trigger,
            "view": JSON.stringify(view_json)
          }
        };
        slackUrl = "https://slack.com/api/views.open";
        var response = UrlFetchApp.fetch(slackUrl, json_data);
      } else if (action_id == 'modify_date') {
        title = '勤怠修正';
        initial_date = payload["actions"][0]["selected_date"];
        initial_date = formatDate3(initial_date);
        var view_json = makeKintaiView(title, user, initial_date);
        var view_id = payload["view"]["root_view_id"];
        var slackUrl = "https://slack.com/api/views.update";
        // いったん空にする
        var view_empty = {
          "type": "modal",
          "title": {
              "type": "plain_text",
              "text": title,
              "emoji": true
          },
          "close": {
              "type": "plain_text",
              "text": "閉じる",
              "emoji": true
          },
          "blocks": []
        };
        var json_data_update = {
          "method": "post",
          "payload": {
            "token": token, // OAuth_token
            "view_id":  view_id,
            "view": JSON.stringify(view_empty)
          }
        };
        var response = UrlFetchApp.fetch(slackUrl, json_data_update);
        var json_data_update = {
          "method": "post",
          "payload": {
            "token": token, // OAuth_token
            "view_id":  view_id,
            "view": JSON.stringify(view_json)
          }
        };
        var response = UrlFetchApp.fetch(slackUrl, json_data_update);
      }
    } else if (type == 'view_submission') {
      //writeLog('debug', JSON.stringify(payload["view"]));
      var callback_id = payload["view"]["callback_id"];
      if (callback_id == 'view_kintai') {
        var values = payload["view"]["state"]["values"];
        // 登録
        var timecard = new Array();
        var val = '';
        //timecard["date"]    = formatDate3(values["block_date"]["date"]["selected_date"]); // 日付
        timecard["date"]    = payload["view"]["private_metadata"]; // 日付
        val = values["block_start"]["start"];
        timecard["start"]   = empValue(values["block_start"]["start"]["value"]);   // 出勤時刻
        timecard["finish"]  = empValue(values["block_finish"]["finish"]["value"]); // 退勤時刻
        timecard["break"]   = empValue(values["block_break"]["break"]["value"]);   // 休憩
        var selOpt = values["block_holiday"]["holiday"]["selected_option"];
        timecard["holiday"] = selOpt ? selOpt["value"] : ''; // 休暇
        timecard["comment"] = empValue(values["block_comment"]["comment"]["value"]); // コメント
      
        //writeLog('debug', 'user:' + user + ',timecard:' + timecard["date"]);
        writeKintai(user, timecard);
      } else if (callback_id == 'view_geppou') {
        // 月報ダウンロード
        var token = PropertiesService.getScriptProperties().getProperty('OAUTH_ACCESS_TOKEN');
        var view_id = payload["view"]["root_view_id"];
        var trigger_id = payload["trigger_id"];
        
        // 出力対象
        var out_text = "";
        var values = payload["view"]["state"]["values"];
        //writeLog("debug", JSON.stringify(values));
        var viewYYYYMM = values["block_yyyymm"]["action_yyyymm"]["selected_option"]["value"];
        var viewUser   = values["block_user"]["action_user"]["selected_option"]["value"];

        // ユーザー情報
        var userInfo = readUserinfo(viewUser);
        if (!userInfo) {
          throw new Error(viewUser + "さんのユーザー情報がありません");
        }
        var matches = viewYYYYMM.match(/([0-9]{4})([0-9]{2})/);
        if (matches && matches.length >= 3) {
          var gas = PropertiesService.getScriptProperties().getProperty('SLACK_GET');
          var param = 'user=' + viewUser + '&year=' + matches[1] + '&month=' + matches[2];
          var url = 'https://script.google.com/macros/s/' + gas + '/exec?' + param;
          out_text = '<' + url + '|勤怠 - ' + userInfo[1] + '_' + matches[1] + '年' + matches[2] + '月>';
        }
        var view_json = {
          "type": "modal",
          "title": {
            "type": "plain_text",
            "text": "月報ダウンロード",
            "emoji": true
          },
          "close": {
            "type": "plain_text",
            "text": "閉じる",
            "emoji": true
          },
          "blocks": [
            {
              "type": "section",
              "text": {
                "type": "mrkdwn",
                "text": "新しくブラウザを開いて月報をダウンロードします。"
              }
            },
            {
              "type": "section",
              "text": {
                "type": "mrkdwn",
                "text": out_text
              }
            }
          ]
        };
        json_data = {
          "response_action": "update",
          "view": JSON.stringify(view_json)
        };
        return ContentService.createTextOutput(JSON.stringify(json_data)).setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput();
  } catch(e) {
    writeLog('error', e.message);
    var json_data = {
      text: e.message
    };
    return ContentService.createTextOutput(JSON.stringify(json_data)).setMimeType(ContentService.MimeType.JSON);
  }
}

/*
 * 勤怠データ取得
 */
function getTimecard(user, dateStr) {
  // 空データ
  var timecard = new Array();
  timecard["date"]    = dateStr; // 日付
  timecard["start"]   = ''; // 出勤時刻
  timecard["finish"]  = ''; // 退勤時刻
  timecard["break"]   = ''; // 休憩
  timecard["holiday"] = ''; // 休暇
  timecard["comment"] = ''; // コメント
  
  // スプレッドシートを開く
  spApp = openSpreadsheet();
  // ユーザー情報取得
  var userInfo = readUserinfo(user);
  if (!userInfo) {
    throw new Error(user + "さんのユーザー情報がありません");
  }
  // ユーザーの勤怠シート
  var sheet = spApp.getSheetByName(userInfo[1]);
  // データ検索
  var idxRow = -1;
  if (sheet) {
    idxRow = findDateRow(sheet, dateStr);
  }
  //writeLog('debug', 'idxRow:' + idxRow);
  if (idxRow > 0) {
    var startTime  = sheet.getRange(idxRow,5).getValue();
    var finishTime = sheet.getRange(idxRow,6).getValue();
    var breakTime  = sheet.getRange(idxRow,7).getValue();
    var holiday    = sheet.getRange(idxRow,9).getValue();
    var comment    = sheet.getRange(idxRow,10).getValue();
    if (startTime) {
      timecard["start"]   = formatTime(startTime);
    } else {
      timecard["start"]   = formatTime(new Date());
    };
    if (finishTime) {
      timecard["finish"]  = formatTime(finishTime);
    } else {
      if (startTime) timecard["finish"]  = formatTime(new Date());
    };
    if (breakTime) {
      timecard["break"]  = formatTime(breakTime);
    } else {
      timecard["break"]  = formatTime(userInfo[3]);
    };
    timecard["holiday"] = holiday2id(holiday);
    timecard["comment"] = comment;
  } else {
    timecard["start"]   = formatTime(new Date());
    timecard["break"]   = formatTime(userInfo[3]);
  }
  return timecard;
}

/*
 * 勤怠ダイアログ表示
 */
function makeKintaiView(title, user, initial_date) {
  //writeLog('debug', 'action_id:' + action_id + ',initial_date:' + initial_date);
  // 勤怠データの取得
  var timecard = getTimecard(user, initial_date);
  var optHoliday = '[';
  optHoliday += '{"text":{"type": "plain_text","text": "祝日"},"value": "public"},';
  optHoliday += '{"text":{"type": "plain_text","text": "午前半休"},"value": "before"},';
  optHoliday += '{"text":{"type": "plain_text","text": "午後半休"},"value": "after"},';
  optHoliday += '{"text":{"type": "plain_text","text": "有給休暇"},"value": "paid"},';
  optHoliday += '{"text":{"type": "plain_text","text": "特別休暇"},"value": "special"}]';

  var viewKintai = 
  {
    "type": "modal",
    "callback_id": "view_kintai",
    "private_metadata": initial_date,
    "title": {
      "type": "plain_text",
      "text": title,
      "emoji": true
    },
    "submit": {
      "type": "plain_text",
      "text": "登録",
      "emoji": true
    },
    "close": {
      "type": "plain_text",
      "text": "Cancel",
      "emoji": true
    },
    "blocks": [
      {
        "type": "section",
        "block_id": "block_date",
        "text": {
          "type": "mrkdwn",
          "text": "年月日"
        },
        "accessory": {
          "type": "datepicker",
          "action_id": "modify_date",
          "initial_date": formatDate2(initial_date,'-'),
          "placeholder": {
            "type": "plain_text",
            "text": "年月日選択",
            "emoji": true
          }
        }
      },
      {
        "type": "input",
        "optional": true,
        "block_id": "block_start",
        "element": {
          "type": "plain_text_input",
          "initial_value": timecard['start'],
          "action_id": "start"
        },
        "label": {
          "type": "plain_text",
          "text": "出勤時刻",
          "emoji": true
        }
      },
      {
        "type": "input",
        "optional": true,
        "block_id": "block_finish",
        "element": {
          "type": "plain_text_input",
          "initial_value": timecard['finish'],
          "action_id": "finish"
        },
        "label": {
          "type": "plain_text",
          "text": "退勤時刻",
          "emoji": true
        }
      },
      {
        "type": "input",
        "optional": true,
        "block_id": "block_break",
        "element": {
          "type": "plain_text_input",
          "initial_value": timecard['break'],
          "action_id": "break"
        },
        "label": {
          "type": "plain_text",
          "text": "休憩時間",
          "emoji": true
        }
      },
      {
        "type": "input",
        "optional": true,
        "block_id": "block_holiday",
        "element": {
          "type": "static_select",
          "action_id": "holiday",
          "placeholder": {
            "type": "plain_text",
            "text": "対象選択",
            "emoji": true
          },
          "options": JSON.parse(optHoliday)
        },
        "label": {
          "type": "plain_text",
          "text": "休暇選択",
          "emoji": true
        }
      },
      {
        "type": "input",
        "optional": true,
        "block_id": "block_comment",
        "element": {
          "type": "plain_text_input",
          "action_id": "comment",
          "initial_value": timecard['comment'],
          "multiline": true
        },
        "label": {
          "type": "plain_text",
          "text": "コメント",
          "emoji": true
        }
      }
    ]
  };
  // 休暇デフォルト
  if (timecard['holiday']) {
    var iniOptHoliday = '{"text":{"type": "plain_text","text": "' + id2holiday(timecard['holiday']) + '"},"value": "' + timecard['holiday'] + '"}';
    viewKintai['blocks'][4]['element']['initial_option'] = JSON.parse(iniOptHoliday);
  };
  //writeLog('debug', JSON.stringify(viewKintai));
  return viewKintai;
}

/*
 * 月報View表示
 */
function makeGeppouView(title, user, initial_date, trigger) {
  // 月報年月
  var wkDate = new Date(initial_date);
  var year = wkDate.getFullYear();
  var month = ('0' + (wkDate.getMonth() + 1)).slice(-2);
  var optYYYYMM = '';
  var iniOptYYYYMM = '{"text":{"type": "plain_text","text": "' + year + '年' + month + '月"},"value": "' + year + month + '"}';
  // 過去１年分出力できるよ
  for (var i=0; i<12; i++) {
    year = wkDate.getFullYear();
    month = ('0' + (wkDate.getMonth() + 1)).slice(-2);
    if (optYYYYMM) optYYYYMM += ',';
    //options += '{"label": "' + year + '年' + month + '月","value": "' + year + month + '"}';
    optYYYYMM += '{"text":{"type": "plain_text","text": "' + year + '年' + month + '月"},"value": "' + year + month + '"}';
    wkDate.setMonth(wkDate.getMonth()-1);
  }
  optYYYYMM = '[' + optYYYYMM + ']';
  var blockYYYYMM =         
      {
            "type": "input",
            "block_id": "block_yyyymm",
            "element": {
                "type": "static_select",
                "action_id": "action_yyyymm",
                "placeholder": {
                    "type": "plain_text",
                    "text": "年月選択",
                    "emoji": true
                },
                "initial_option": JSON.parse(iniOptYYYYMM),
                "options": JSON.parse(optYYYYMM)
            },
            "label": {
                "type": "plain_text",
                "text": "年月を選択してください（過去１年分）",
                "emoji": true
            }
        };
  
  // ユーザー選択
  var aryUser = getAryUser();
  var userInfo;
  var aryViewUser = new Array();
  for (var i in aryUser) {
    if (aryUser[i][2] == user) {
      userInfo = aryUser[i];
    }
  }
  if (!userInfo) {
    throw new Error(user + "さんのユーザー情報がありません");
  }
  // 参照権限
  if (userInfo[5] == 'all') {
    // 全員
    aryViewUser = aryUser;
  } else if (userInfo[5] == 'self') {
    // 自分だけ
    aryViewUser.push(userInfo);
  } else {
    // 自分のグループ参照
    for (var i in aryUser) {
      if (userInfo[4] == aryUser[i][4]) {
        aryViewUser.push(aryUser[i]);
      }
    }
  }
  var optUser = '';
  var iniOptUser = '{"text":{"type": "plain_text","text": "' + userInfo[1] + '"},"value": "' + userInfo[2] + '"}';
  for (var i in aryViewUser) {
    if (!aryViewUser[i][1]) continue;
    if (optUser) optUser += ',';
    optUser += '{"text":{"type": "plain_text","text": "' + aryViewUser[i][1] + '"},"value": "' + aryViewUser[i][2] + '"}';
  }
  optUser = '[' + optUser + ']';
  //writeLog('debug', optUser);
  var blockUser =         
      {
            "type": "input",
            "block_id": "block_user",
            "element": {
                "type": "static_select",
                "action_id": "action_user",
                "placeholder": {
                    "type": "plain_text",
                    "text": "対象選択",
                    "emoji": true
                },
                "initial_option": JSON.parse(iniOptUser),
                "options": JSON.parse(optUser)
            },
            "label": {
                "type": "plain_text",
                "text": "月報出力対象を選択してください",
                "emoji": true
            }
        };
      
  var view_json = {
	"type": "modal",
    "callback_id": "view_geppou",
	"title": {
		"type": "plain_text",
		"text": title,
		"emoji": true
	},
	"submit": {
		"type": "plain_text",
		"text": "選択",
		"emoji": true
	},
	"close": {
		"type": "plain_text",
		"text": "閉じる",
		"emoji": true
	},
	"blocks": [blockYYYYMM, blockUser]
  };
  return view_json;
}

/*
 * 勤怠書き込み
 */
function writeKintai(user, timecard) {
  //writeLog("debug", user + ':' + JSON.stringify(timecard));
  // 日付(yyyy/mm/dd)
  var dateStr = formatDate3(timecard["date"]);
  // スプレッドシートを開く
  spApp = openSpreadsheet();
  // ユーザー情報取得
  var userInfo = readUserinfo(user);
  if (!userInfo) {
    throw new Error(user + "さんのユーザー情報がありません");
  }
  // ユーザーの勤怠シート
  var sheet = spApp.getSheetByName(userInfo[1]);
  //writeLog('debug', 'userInfo:' + userInfo);
  var idxRow = -1;
  if (!sheet) {
    // ユーザーのシートがないとき、雛型をコピーして新規作成
    var temp = spApp.getSheetByName('template');
    sheet = temp.copyTo(spApp);
    sheet.setName(userInfo[1]);
  } else {
    idxRow = findDateRow(sheet, dateStr);
  }
  //writeLog('debug', 'idxRow:' + idxRow);
  if (idxRow < 0) {
    // 行追加
    idxRow = fillDateRow(sheet, dateStr, userInfo);
  }
  
  // 出勤
  sheet.getRange(idxRow,5).setValue(validHour(timecard["start"]));
  // 退勤
  sheet.getRange(idxRow,6).setValue(validHour(timecard["finish"]));
  // 休憩
  sheet.getRange(idxRow,7).setValue(validHour(timecard["break"]));
  // 休暇
  sheet.getRange(idxRow,9).setValue(id2holiday(timecard["holiday"]));
  // コメント
  sheet.getRange(idxRow,10).setValue(timecard["comment"]);
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
  var aryUser = getAryUser();
  // slackユーザーの行番号を取得
  var us = Underscore.load();
  var aryTrns = us.zip.apply(us, aryUser);
  var rowNum = aryTrns[2].indexOf(in_user);
  if (rowNum < 0) return null;
 
  return aryUser[rowNum];
}

function getAryUser() {
  // ユーザーシート
  var sheet = spApp.getSheetByName('user');
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  return sheet.getRange(2,1,lastRow,lastCol).getValues();
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
function formatDate3(dateStr) {
  return dateStr.replace(/-/g, '/');
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
 * 休日変換
 */
function holiday2id(holiday) {
  if (holiday == '祝日') return 'public';
  if (holiday == '前半') return 'before';
  if (holiday == '後半') return 'after';
  if (holiday == '有給') return 'paid';
  if (holiday == '特別') return 'special';
  return '';
}
function id2holiday(holiday) {
  if (holiday == 'public')  return '祝日';
  if (holiday == 'before')  return '前半';
  if (holiday == 'after')   return '後半';
  if (holiday == 'paid')    return '有給';
  if (holiday == 'special') return '特別';
  return '';
}

/*
 * 空データセット
 */
function empValue(val) {
  if (val) return val;
  else     return '';
}
/*
 * 有効時刻セット
 */
function validHour(str) {
  // 全角→半角
  str = str.replace(/[０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  str = str.replace(/[：]/g, ':');
  return str;
}

/*
 * 年月日行検索
 */
function findDateRow(sheet, dateStr) {
  // 最後の行の日付より大きかったら行追加
  var lastRow = sheet.getLastRow();
  var lastDate = formatDate(sheet.getRange(lastRow,1).getValue());
  if (dateStr > lastDate) {
    // 該当なし
    return -1;
  }
  // 最後の行の日付
  if (dateStr == lastDate) {
    return lastRow;
  }
  // 範囲から探す
  var aryDate = sheet.getRange(2,1,lastRow,1).getValues();
  for (var i=aryDate.length-1; i>0; i--) {
    var wkDate = formatDate(aryDate[i]);
    if (dateStr == wkDate) {
      // データみつけた
      return i+2;
    }
  }
  return -1;
}

/*
 * 行追加
 */
function fillDateRow(sheet, dateStr, userInfo){
  var lastRow = sheet.getLastRow();
  
  var wkDate      = valToDate(dateStr);
  var wkDateStr   = dateStr;
  var lastDate    = '';
  var lastDateStr = '';
  
  // 途中うめる
  var aryFill = new Array();
  if (lastRow <= 1) {
    // 新規作成(見出しだけ)
    idxRow = lastRow+1;
    lastDate = valToDate(dateStr);
    lastDate.setDate(1); // 1日からのデータ
    lastDateStr = formatDate(lastDate);
  } else {
    // 間のデータ埋める
    lastDate    = sheet.getRange(lastRow,1).getValue();
    lastDateStr = formatDate(lastDate);
  }
  //writeLog('debug', 'wkDate:' + wkDateStr + ',lastDate:' + lastDateStr);
  // 間のデータ埋める
  do {
    aryFill.push(wkDateStr);
    // 一日前
    wkDate.setDate(wkDate.getDate()-1);
    wkDateStr = formatDate(wkDate);
  } while(lastDateStr < wkDateStr);
  if (lastRow <= 1) {
    // 新規作成(見出しだけ)
    aryFill.push(lastDateStr);
  }

  // 書き込み
  var wkIdx = lastRow + aryFill.length;
  for (var i in aryFill) {
    // 西暦
    sheet.getRange(wkIdx,1).setValue(aryFill[i]);
    // 年
    sheet.getRange(wkIdx,2).setValue('=text(A' + wkIdx + ',"yyyy年")');
    // 月
    sheet.getRange(wkIdx,3).setValue('=text(A' + wkIdx + ',"mm月")');
    // 日
    sheet.getRange(wkIdx,4).setValue('=text(A' + wkIdx + ',"dd(ddd)")');
    // 休憩時間
    sheet.getRange(wkIdx,7).setValue(userInfo[3]);
    // 勤務時間
    sheet.getRange(wkIdx,8).setValue('=IF(F' +wkIdx+ '="","",F'+wkIdx+'-E'+wkIdx+'-G'+wkIdx+')');
    
    wkIdx--;
  }
  return lastRow + aryFill.length;
}
