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
 * Slack���b�Z�[�W����
 */
function postSlackMessage(param) {
  var in_userName = param.user_name;
  var in_text     = param.text;
  var json_data;
  
  // ���[�U�[���擾
  var userInfo = readUserinfo(in_userName);
  if (!userInfo) {
    writeLog('info', in_userName + ':���[�U�[��񂪂���܂���');
    var out_text = in_userName + "����̃��[�U�[��񂪂���܂���";
    json_data = {
      text: out_text
    };
  } else {
    // �Αӓ��͕\��
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
                        "text": "�o�ދ΂̓o�^",
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
                        "text": "����̏o��",
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
  // ���[�U�[�V�[�g
  var sheet = spApp.getSheetByName('user');
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  var aryUser = sheet.getRange(2,1,lastRow,lastCol).getValues();
  // slack���[�U�[�̍s�ԍ����擾
  //var us = Underscore.load();
  var aryTrns = usApp.zip.apply(usApp, aryUser);
  var rowNum = aryTrns[2].indexOf(in_user);
  if (rowNum < 0) return null;
 
  return aryUser[rowNum];
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

