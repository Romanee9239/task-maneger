function createTaskDocs(){
  var sht = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var task_sht = sht.getSheetByName("受注");
  var set_sht = sht.getSheetByName("設定・使い方");
  var last_row = task_sht.getLastRow();

  const TASK_NAME_COL = 3;
  const NEW_FLG_COL = 5;
  
  const CHK_STA_ROW = 6;
  
  var taskName = [];
  var newTaskRow = [];
  var task_tmp;
  var task_flg;
  var task_dir;
  var k = 0;
  for (var i = CHK_STA_ROW; i <= last_row; i++){
    task_tmp = task_sht.getRange(i, TASK_NAME_COL).getValue();
    task_flg = task_sht.getRange(i, NEW_FLG_COL).getValue();
    if (task_tmp == '' || task_flg !== ''){
      continue;
    }
    taskName[k] = task_tmp;
    newTaskRow[k] = i;
    k++;
  }
  
   const SAVE_DIR_ROW = 6;
   const TEMPLATE_ROW = 7;
   
   var mngDirId = set_sht.getRange(SAVE_DIR_ROW, 2).getValue();
   Logger.log(mngDirId);
   mngDirId = mngDirId.replace('https://drive.google.com/drive/folders/','');
   if (mngDirId == ''){
     mngDirId = '1fkTN2yFa1zqAdhOoApYQdpcR7N_nTu1Y';
   }
   var taskMngDir = DriveApp.getFolderById(mngDirId);
   var scenarioFmtId = set_sht.getRange(TEMPLATE_ROW, 2).getValue();
   scenarioFmtId = scenarioFmtId.replace('https://docs.google.com/document/d/','');
   scenarioFmtId = scenarioFmtId.replace('/edit','');
   Logger.log(scenarioFmtId);
   if (scenarioFmtId == ''){
     scenarioFmtId = '1jSDUJQsT_fOLqi0BBiqDOMpVR7EtWpcMZcjMNv2SVBY';
   }
   var scenarioFmt = DriveApp.getFileById(scenarioFmtId);
   var newDir;
   var newDirId;
   var newDocs;
   var newDocsId;
   var newDocsName;
   var newLinks = [];
   const DOCLINK = 'https://docs.google.com/document/d/@share@/edit?usp=sharing';
   const TITLE_LINK = '=HYPERLINK("@link@", "@title@")';
   var hyperlink = TITLE_LINK;
   
   var resource;
   
   for (var j = 0; j < taskName.length; j++){
     newDocsName = task_sht.getRange(newTaskRow[j], TASK_NAME_COL - 1).getValue() + '様_' + taskName[j];
     newDocs = scenarioFmt.makeCopy(newDocsName, taskMngDir);
     newDocsId = newDocs.getId();
     
     newDocs.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
     
     hyperlink = TITLE_LINK.replace('@title@',taskName[j]);
     hyperlink = hyperlink.replace('@link@',DOCLINK.replace('@share@', newDocsId));
     task_sht.getRange(newTaskRow[j], TASK_NAME_COL).setValue(hyperlink);
     task_sht.getRange(newTaskRow[j], TASK_NAME_COL).setFontColor('#1155cc');
     task_sht.getRange(newTaskRow[j], TASK_NAME_COL).setFontLine('underline');
     
     resource = {"requests": [
       {
         "repeatCell": {
           "cell": {"dataValidation": {"condition": {"type": "BOOLEAN"}}},
           "range": {"sheetId": task_sht.getSheetId(), "startRowIndex": newTaskRow[j] - 1, "endRowIndex": newTaskRow[j], "startColumnIndex": NEW_FLG_COL - 1, "endColumnIndex": NEW_FLG_COL},
           "fields": "dataValidation",
         },
       },
     ]};
     Sheets.Spreadsheets.batchUpdate(resource, sht.getId());
     task_sht.getRange(newTaskRow[j], NEW_FLG_COL).setValue('FALSE');
   }
   
}
