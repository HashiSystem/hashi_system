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
var url = '';
var xlsName = '';

/*
 * Post���N�G�X�g����
 */
function doGet(e){ // e ��POST���ꂽ�f�[�^�������Ă���
  try {
    //writeLog('debug', e);
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
 * �O���t�@�C����荞��
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
/*
 * HTML�ɓn��
 */
function getUrl() {
  return url;
}
function getXlsName() {
  return xlsName;
}

/*
 * �X�v���b�h�V�[�g�J��
 */
function openSpreadsheet() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREAD_SHEET_ID');
  return SpreadsheetApp.openById(id);
}

/*
 * �P�������̋Ζ��\
 */
function viewKintai(user, year, month) {
  //writeLog('debug', user + ':' + year + '/' + month);
  // ���[�U�[���擾
  var userInfo = readUserinfo(user);
  if (!userInfo) {
    throw new Error(user + ':���[�U�[��񂪂���܂���B');
  }
  // ���[�U�[�̋ΑӃV�[�g
  var sheet = spApp.getSheetByName(userInfo[1]);
  if (!sheet) {
    throw new Error(userInfo[1] + ':�Αӏ�񂪂���܂���B');
  }
  // ��Ɨp�V�[�g(�O��̌���폜)
  var wkSheetList = spApp.getSheets();
  var wkSheet;
  var regex = new RegExp('^' + userInfo[1] + '_\\d+�N\\d+��');
  for (var i in wkSheetList) {
    wkSheet = wkSheetList[i];
    var sheetName = wkSheet.getName();
    if (sheetName.match(regex)) {
      spApp.deleteSheet(wkSheet);
    }
  }
  // ���^���R�s�[���ĐV�K�쐬
  var temp = spApp.getSheetByName('template');
  wkSheet = temp.copyTo(spApp);
  wkSheet.setName(userInfo[1] + '_' + year + '�N' + month + '��');
  
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
  var aryWeek    = [ "��", "��", "��", "��", "��", "��", "�y" ];
  //writeLog('debug', 'viewKintai:' + wkDateStr + '-' + endDateStr);
  var wkIdx = 2;
  var section = '';
  var fields  = '';
  while (wkDateStr < endDateStr) {
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
    // ����
    wkSheet.getRange(wkIdx,1).setValue(wkDateStr);
    // �N
    wkSheet.getRange(wkIdx,2).setValue('=text(A' + wkIdx + ',"yyyy�N")');
    // ��
    wkSheet.getRange(wkIdx,3).setValue('=text(A' + wkIdx + ',"mm��")');
    // ��
    wkSheet.getRange(wkIdx,4).setValue('=text(A' + wkIdx + ',"dd(ddd)")');
    // �o�Ύ���
    wkSheet.getRange(wkIdx,5).setValue(startTime);
    // �ދΎ���
    wkSheet.getRange(wkIdx,6).setValue(endTime);
    // �x�e����
    wkSheet.getRange(wkIdx,7).setValue(breakTime);
    // �Ζ�����
    wkSheet.getRange(wkIdx,8).setValue('=IF(F' +wkIdx+ '="","",F'+wkIdx+'-E'+wkIdx+'-G'+wkIdx+')');
    // �x��
    wkSheet.getRange(wkIdx,9).setValue(holiday);
    // �⑫
    wkSheet.getRange(wkIdx,10).setValue(comment);
    
    wkDate.setDate(wkDate.getDate()+1);
    wkDateStr  = formatDate(wkDate);
    wkDay      = wkDate.getDay();
    wkIdx++;
  }
  // ���v
  wkSheet.getRange(wkIdx,7).setValue('���v');
  wkSheet.getRange(wkIdx,8).setValue('=SUM(H2:H' + (wkIdx-1) + ')');
  wkSheet.getRange(wkIdx,8).setNumberFormat('[h]:mm:ss');
  
  // �_�E�����[�h
  var sid = PropertiesService.getScriptProperties().getProperty('SPREAD_SHEET_ID');
  var gid = wkSheet.getSheetId();
  url = 'https://docs.google.com/spreadsheets/d/' + sid + '/export?gid=' + gid + '&exportFormat=xlsx'
  url += '&access_token=' + ScriptApp.getOAuthToken();
  xlsName = userInfo[1] + '_' + year + '�N' + month + '��.xlsx';
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
