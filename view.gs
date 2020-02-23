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
 * Debug�p
 */
function debugTest() {
  // �f�o�b�O�p
}
/*
 * ���O�o��
 */
function writeLog(log_type, text) {
  // �f�o�b�O�o�͂��Ȃ��Ƃ������
  if (!debug && log_type == 'debug') return;
  
  if (spApp) {
    // Log�V�[�g
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
 * Post���N�G�X�g����
 */
function doPost(e){ // e ��POST���ꂽ�f�[�^�������Ă���
  try {
    var payload = JSON.parse(e["parameter"]["payload"]);
    //writeLog('debug', payload);
    
    var user = payload["user"]["name"];
    var type = payload["type"];
    if (type == 'block_actions') {
      // modal�\��
      var action_id = payload["actions"][0]["action_id"];
      var trigger = payload["trigger_id"];
      var token = PropertiesService.getScriptProperties().getProperty('OAUTH_ACCESS_TOKEN');
      var json_data;
      var slackUrl;
      var title = '';
      var initial_date = '';
      if (action_id == 'timecard') {
        title = '�Αӓ���';
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
        title = '����';
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
        title = '�ΑӏC��';
        initial_date = payload["actions"][0]["selected_date"];
        initial_date = formatDate3(initial_date);
        var view_json = makeKintaiView(title, user, initial_date);
        var view_id = payload["view"]["root_view_id"];
        var slackUrl = "https://slack.com/api/views.update";
        // ���������ɂ���
        var view_empty = {
          "type": "modal",
          "title": {
              "type": "plain_text",
              "text": title,
              "emoji": true
          },
          "close": {
              "type": "plain_text",
              "text": "����",
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
        // �o�^
        var timecard = new Array();
        var val = '';
        //timecard["date"]    = formatDate3(values["block_date"]["date"]["selected_date"]); // ���t
        timecard["date"]    = payload["view"]["private_metadata"]; // ���t
        val = values["block_start"]["start"];
        timecard["start"]   = empValue(values["block_start"]["start"]["value"]);   // �o�Ύ���
        timecard["finish"]  = empValue(values["block_finish"]["finish"]["value"]); // �ދΎ���
        timecard["break"]   = empValue(values["block_break"]["break"]["value"]);   // �x�e
        var selOpt = values["block_holiday"]["holiday"]["selected_option"];
        timecard["holiday"] = selOpt ? selOpt["value"] : ''; // �x��
        timecard["comment"] = empValue(values["block_comment"]["comment"]["value"]); // �R�����g
      
        //writeLog('debug', 'user:' + user + ',timecard:' + timecard["date"]);
        writeKintai(user, timecard);
      } else if (callback_id == 'view_geppou') {
        // ����_�E�����[�h
        var token = PropertiesService.getScriptProperties().getProperty('OAUTH_ACCESS_TOKEN');
        var view_id = payload["view"]["root_view_id"];
        var trigger_id = payload["trigger_id"];
        
        // �o�͑Ώ�
        var out_text = "";
        var values = payload["view"]["state"]["values"];
        //writeLog("debug", JSON.stringify(values));
        var viewYYYYMM = values["block_yyyymm"]["action_yyyymm"]["selected_option"]["value"];
        var viewUser   = values["block_user"]["action_user"]["selected_option"]["value"];

        // ���[�U�[���
        var userInfo = readUserinfo(viewUser);
        if (!userInfo) {
          throw new Error(viewUser + "����̃��[�U�[��񂪂���܂���");
        }
        var matches = viewYYYYMM.match(/([0-9]{4})([0-9]{2})/);
        if (matches && matches.length >= 3) {
          var gas = PropertiesService.getScriptProperties().getProperty('SLACK_GET');
          var param = 'user=' + viewUser + '&year=' + matches[1] + '&month=' + matches[2];
          var url = 'https://script.google.com/macros/s/' + gas + '/exec?' + param;
          out_text = '<' + url + '|�Α� - ' + userInfo[1] + '_' + matches[1] + '�N' + matches[2] + '��>';
        }
        var view_json = {
          "type": "modal",
          "title": {
            "type": "plain_text",
            "text": "����_�E�����[�h",
            "emoji": true
          },
          "close": {
            "type": "plain_text",
            "text": "����",
            "emoji": true
          },
          "blocks": [
            {
              "type": "section",
              "text": {
                "type": "mrkdwn",
                "text": "�V�����u���E�U���J���Č�����_�E�����[�h���܂��B"
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
 * �ΑӃf�[�^�擾
 */
function getTimecard(user, dateStr) {
  // ��f�[�^
  var timecard = new Array();
  timecard["date"]    = dateStr; // ���t
  timecard["start"]   = ''; // �o�Ύ���
  timecard["finish"]  = ''; // �ދΎ���
  timecard["break"]   = ''; // �x�e
  timecard["holiday"] = ''; // �x��
  timecard["comment"] = ''; // �R�����g
  
  // �X�v���b�h�V�[�g���J��
  spApp = openSpreadsheet();
  // ���[�U�[���擾
  var userInfo = readUserinfo(user);
  if (!userInfo) {
    throw new Error(user + "����̃��[�U�[��񂪂���܂���");
  }
  // ���[�U�[�̋ΑӃV�[�g
  var sheet = spApp.getSheetByName(userInfo[1]);
  // �f�[�^����
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
 * �ΑӃ_�C�A���O�\��
 */
function makeKintaiView(title, user, initial_date) {
  //writeLog('debug', 'action_id:' + action_id + ',initial_date:' + initial_date);
  // �ΑӃf�[�^�̎擾
  var timecard = getTimecard(user, initial_date);
  var optHoliday = '[';
  optHoliday += '{"text":{"type": "plain_text","text": "�j��"},"value": "public"},';
  optHoliday += '{"text":{"type": "plain_text","text": "�ߑO���x"},"value": "before"},';
  optHoliday += '{"text":{"type": "plain_text","text": "�ߌ㔼�x"},"value": "after"},';
  optHoliday += '{"text":{"type": "plain_text","text": "�L���x��"},"value": "paid"},';
  optHoliday += '{"text":{"type": "plain_text","text": "���ʋx��"},"value": "special"}]';

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
      "text": "�o�^",
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
          "text": "�N����"
        },
        "accessory": {
          "type": "datepicker",
          "action_id": "modify_date",
          "initial_date": formatDate2(initial_date,'-'),
          "placeholder": {
            "type": "plain_text",
            "text": "�N�����I��",
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
          "text": "�o�Ύ���",
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
          "text": "�ދΎ���",
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
          "text": "�x�e����",
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
            "text": "�ΏۑI��",
            "emoji": true
          },
          "options": JSON.parse(optHoliday)
        },
        "label": {
          "type": "plain_text",
          "text": "�x�ɑI��",
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
          "text": "�R�����g",
          "emoji": true
        }
      }
    ]
  };
  // �x�Ƀf�t�H���g
  if (timecard['holiday']) {
    var iniOptHoliday = '{"text":{"type": "plain_text","text": "' + id2holiday(timecard['holiday']) + '"},"value": "' + timecard['holiday'] + '"}';
    viewKintai['blocks'][4]['element']['initial_option'] = JSON.parse(iniOptHoliday);
  };
  //writeLog('debug', JSON.stringify(viewKintai));
  return viewKintai;
}

/*
 * ����View�\��
 */
function makeGeppouView(title, user, initial_date, trigger) {
  // ����N��
  var wkDate = new Date(initial_date);
  var year = wkDate.getFullYear();
  var month = ('0' + (wkDate.getMonth() + 1)).slice(-2);
  var optYYYYMM = '';
  var iniOptYYYYMM = '{"text":{"type": "plain_text","text": "' + year + '�N' + month + '��"},"value": "' + year + month + '"}';
  // �ߋ��P�N���o�͂ł����
  for (var i=0; i<12; i++) {
    year = wkDate.getFullYear();
    month = ('0' + (wkDate.getMonth() + 1)).slice(-2);
    if (optYYYYMM) optYYYYMM += ',';
    //options += '{"label": "' + year + '�N' + month + '��","value": "' + year + month + '"}';
    optYYYYMM += '{"text":{"type": "plain_text","text": "' + year + '�N' + month + '��"},"value": "' + year + month + '"}';
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
                    "text": "�N���I��",
                    "emoji": true
                },
                "initial_option": JSON.parse(iniOptYYYYMM),
                "options": JSON.parse(optYYYYMM)
            },
            "label": {
                "type": "plain_text",
                "text": "�N����I�����Ă��������i�ߋ��P�N���j",
                "emoji": true
            }
        };
  
  // ���[�U�[�I��
  var aryUser = getAryUser();
  var userInfo;
  var aryViewUser = new Array();
  for (var i in aryUser) {
    if (aryUser[i][2] == user) {
      userInfo = aryUser[i];
    }
  }
  if (!userInfo) {
    throw new Error(user + "����̃��[�U�[��񂪂���܂���");
  }
  // �Q�ƌ���
  if (userInfo[5] == 'all') {
    // �S��
    aryViewUser = aryUser;
  } else if (userInfo[5] == 'self') {
    // ��������
    aryViewUser.push(userInfo);
  } else {
    // �����̃O���[�v�Q��
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
                    "text": "�ΏۑI��",
                    "emoji": true
                },
                "initial_option": JSON.parse(iniOptUser),
                "options": JSON.parse(optUser)
            },
            "label": {
                "type": "plain_text",
                "text": "����o�͑Ώۂ�I�����Ă�������",
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
		"text": "�I��",
		"emoji": true
	},
	"close": {
		"type": "plain_text",
		"text": "����",
		"emoji": true
	},
	"blocks": [blockYYYYMM, blockUser]
  };
  return view_json;
}

/*
 * �Αӏ�������
 */
function writeKintai(user, timecard) {
  //writeLog("debug", user + ':' + JSON.stringify(timecard));
  // ���t(yyyy/mm/dd)
  var dateStr = formatDate3(timecard["date"]);
  // �X�v���b�h�V�[�g���J��
  spApp = openSpreadsheet();
  // ���[�U�[���擾
  var userInfo = readUserinfo(user);
  if (!userInfo) {
    throw new Error(user + "����̃��[�U�[��񂪂���܂���");
  }
  // ���[�U�[�̋ΑӃV�[�g
  var sheet = spApp.getSheetByName(userInfo[1]);
  //writeLog('debug', 'userInfo:' + userInfo);
  var idxRow = -1;
  if (!sheet) {
    // ���[�U�[�̃V�[�g���Ȃ��Ƃ��A���^���R�s�[���ĐV�K�쐬
    var temp = spApp.getSheetByName('template');
    sheet = temp.copyTo(spApp);
    sheet.setName(userInfo[1]);
  } else {
    idxRow = findDateRow(sheet, dateStr);
  }
  //writeLog('debug', 'idxRow:' + idxRow);
  if (idxRow < 0) {
    // �s�ǉ�
    idxRow = fillDateRow(sheet, dateStr, userInfo);
  }
  
  // �o��
  sheet.getRange(idxRow,5).setValue(validHour(timecard["start"]));
  // �ދ�
  sheet.getRange(idxRow,6).setValue(validHour(timecard["finish"]));
  // �x�e
  sheet.getRange(idxRow,7).setValue(validHour(timecard["break"]));
  // �x��
  sheet.getRange(idxRow,9).setValue(id2holiday(timecard["holiday"]));
  // �R�����g
  sheet.getRange(idxRow,10).setValue(timecard["comment"]);
}

/*
 * �X�v���b�h�V�[�g�J��
 */
function openSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREAD_SHEET_ID');
  return SpreadsheetApp.openById(id);
}

/*
 * ���[�U�[���ǂݍ���
 */
function readUserinfo(in_user) {
  var aryUser = getAryUser();
  // slack���[�U�[�̍s�ԍ����擾
  var us = Underscore.load();
  var aryTrns = us.zip.apply(us, aryUser);
  var rowNum = aryTrns[2].indexOf(in_user);
  if (rowNum < 0) return null;
 
  return aryUser[rowNum];
}

function getAryUser() {
  // ���[�U�[�V�[�g
  var sheet = spApp.getSheetByName('user');
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  return sheet.getRange(2,1,lastRow,lastCol).getValues();
}

/*
 * �N���� �t�H�[�}�b�g
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
 * �����b�t�H�[�}�b�g
 */
function formatTime(date) {
  date = valToDate(date);
  return ('0' + date.getHours()).slice(-2) + ':' + ('0' + date.getMinutes()).slice(-2);
}

/*
 * date�ϊ�
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
 * �x���ϊ�
 */
function holiday2id(holiday) {
  if (holiday == '�j��') return 'public';
  if (holiday == '�O��') return 'before';
  if (holiday == '�㔼') return 'after';
  if (holiday == '�L��') return 'paid';
  if (holiday == '����') return 'special';
  return '';
}
function id2holiday(holiday) {
  if (holiday == 'public')  return '�j��';
  if (holiday == 'before')  return '�O��';
  if (holiday == 'after')   return '�㔼';
  if (holiday == 'paid')    return '�L��';
  if (holiday == 'special') return '����';
  return '';
}

/*
 * ��f�[�^�Z�b�g
 */
function empValue(val) {
  if (val) return val;
  else     return '';
}
/*
 * �L�������Z�b�g
 */
function validHour(str) {
  // �S�p�����p
  str = str.replace(/[�O-�X]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
  });
  str = str.replace(/[�F]/g, ':');
  return str;
}

/*
 * �N�����s����
 */
function findDateRow(sheet, dateStr) {
  // �Ō�̍s�̓��t���傫��������s�ǉ�
  var lastRow = sheet.getLastRow();
  var lastDate = formatDate(sheet.getRange(lastRow,1).getValue());
  if (dateStr > lastDate) {
    // �Y���Ȃ�
    return -1;
  }
  // �Ō�̍s�̓��t
  if (dateStr == lastDate) {
    return lastRow;
  }
  // �͈͂���T��
  var aryDate = sheet.getRange(2,1,lastRow,1).getValues();
  for (var i=aryDate.length-1; i>0; i--) {
    var wkDate = formatDate(aryDate[i]);
    if (dateStr == wkDate) {
      // �f�[�^�݂���
      return i+2;
    }
  }
  return -1;
}

/*
 * �s�ǉ�
 */
function fillDateRow(sheet, dateStr, userInfo){
  var lastRow = sheet.getLastRow();
  
  var wkDate      = valToDate(dateStr);
  var wkDateStr   = dateStr;
  var lastDate    = '';
  var lastDateStr = '';
  
  // �r�����߂�
  var aryFill = new Array();
  if (lastRow <= 1) {
    // �V�K�쐬(���o������)
    idxRow = lastRow+1;
    lastDate = valToDate(dateStr);
    lastDate.setDate(1); // 1������̃f�[�^
    lastDateStr = formatDate(lastDate);
  } else {
    // �Ԃ̃f�[�^���߂�
    lastDate    = sheet.getRange(lastRow,1).getValue();
    lastDateStr = formatDate(lastDate);
  }
  //writeLog('debug', 'wkDate:' + wkDateStr + ',lastDate:' + lastDateStr);
  // �Ԃ̃f�[�^���߂�
  do {
    aryFill.push(wkDateStr);
    // ����O
    wkDate.setDate(wkDate.getDate()-1);
    wkDateStr = formatDate(wkDate);
  } while(lastDateStr < wkDateStr);
  if (lastRow <= 1) {
    // �V�K�쐬(���o������)
    aryFill.push(lastDateStr);
  }

  // ��������
  var wkIdx = lastRow + aryFill.length;
  for (var i in aryFill) {
    // ����
    sheet.getRange(wkIdx,1).setValue(aryFill[i]);
    // �N
    sheet.getRange(wkIdx,2).setValue('=text(A' + wkIdx + ',"yyyy�N")');
    // ��
    sheet.getRange(wkIdx,3).setValue('=text(A' + wkIdx + ',"mm��")');
    // ��
    sheet.getRange(wkIdx,4).setValue('=text(A' + wkIdx + ',"dd(ddd)")');
    // �x�e����
    sheet.getRange(wkIdx,7).setValue(userInfo[3]);
    // �Ζ�����
    sheet.getRange(wkIdx,8).setValue('=IF(F' +wkIdx+ '="","",F'+wkIdx+'-E'+wkIdx+'-G'+wkIdx+')');
    
    wkIdx--;
  }
  return lastRow + aryFill.length;
}
