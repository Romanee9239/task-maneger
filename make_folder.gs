function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [
    {
      name : "このシートを共有する",
      functionName : "generateShareSheet"
    },
    {
      name : "新規フォルダ一括作成",
      functionName : "createTaskDir"
    },
    {
      name : "ガイドライン一括更新",
      functionName : "updateGaideToAll"
    }
  ];
  sheet.addMenu("スクリプト実行", entries);
  //メインメニュー部分に[スクリプト実行]メニューを作成して、
  //下位項目のメニューを設定している
};

function generateShareSheet(){
  const def_bookname = "タスク管理シート_";
  const def_manualSht = "ガイドライン";
  const def_taskSht = "タスク管理";
  
  const def_shareUrl = "https://docs.google.com/spreadsheets/d/@shareid@/edit?usp=sharing";
  
  var sht = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  //シート名→クリエイター名
  var sht_name = sht.getSheetName();
  
  //共有用原本のid
  var masterFileId = "19-3-rBpO73QjexnxJiLXuV21FnVmF4pvtx9wi96mEi0";
  
  //作成ファイル名→"タスク管理シート_" + "クリエイター名"
  var fileName = def_bookname + sht_name;
  
  //ブック作成
  var taskFolder = DriveApp.getFolderById("1ZaCardQi5aY8Q_ony7EHXCAbh6S8x4XR");
  var allFiles = taskFolder.getFiles();
  var file;
  
  //すでにファイルがないか判定
  while(allFiles.hasNext()) {
    file = allFiles.next();
    //同じファイル名があるか
    if (file.getName() == fileName){
      ui.alert(sht_name+"のファイルはもうあるよ");
      return;
    }
  }
  
  var newFile = DriveApp.getFileById(masterFileId).makeCopy(fileName, taskFolder);
  var newId = newFile.getId();
  var newBook = SpreadsheetApp.openById(newId);
  
  //最新のガイドライン取得
  updateSheet(newBook, def_manualSht);
  
  //タスク管理シートに共有するデータを入れる
  var taskSht = newBook.getSheetByName(def_taskSht);
  //クリエイター名入力
  taskSht.getRange("B1").setValue(sht_name);
  
  //共有リンク取得
  var file = DriveApp.getFileById(newId);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  var shareUrl = def_shareUrl.replace("@shareid@",newId);
  
  //リストに追加
  var list = sht.getSheetByName("一覧");
  var row = 2;
  var cel = list.getRange(row++, 1);
  while(cel.getValue() != ""){
    cel = list.getRange(row++, 1);
  }
  cel.setValue(sht_name);
  cel.offset(0,1).setValue(shareUrl);
  
}

/*******************************
 * 新規タスクのフォルダを一括作成
 *
 ******************************/
function createTaskDir(){
  var sht = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var task_sht = sht.getSheetByName("発注管理表");
  var last_row = task_sht.getLastRow();
  /*
  * 新規タスクを取得
  * フォルダリンクの欄が空白の場合、未作成とみなす
  * 未作成タスクを取得
  */
  var taskName = [];
  var newTaskRow = [];
  var task_num;
  var task_cel;
  var task_dir;
  var k = 0;
  for (var i = 2; i <= last_row; i++){
    task_num = task_sht.getRange(i, 1).getValue();
    task_cel = task_sht.getRange(i, 2).getValue();
    task_dir = task_sht.getRange(i, 3).getValue();
    if ((task_cel != '') && (task_dir == '')){
      taskName[k] = task_num + '_' + task_cel;
      newTaskRow[k] = i;
      k++;
    }
  }
  
  /*
   * フォルダ一括作成
   */
   //タスク管理親フォルダー取得
   var mngDirId = '11agkPJp7oZ5bBSIS2wpHpyPQYk0WRqTH';
   var taskMngDir = DriveApp.getFolderById(mngDirId);
   //台本フォーマット取得
   var scenarioFmtId = '1jSDUJQsT_fOLqi0BBiqDOMpVR7EtWpcMZcjMNv2SVBY';
   var scenarioFmt = DriveApp.getFileById(scenarioFmtId);
   var newDir;
   var newDirId;
   var newDocs;
   var newDocsId;
   var newLinks = [];
   const DIRLINK = 'https://drive.google.com/drive/folders/@share@?usp=sharing';
   const DOCLINK = 'https://docs.google.com/document/d/@share@/edit?usp=sharing';
   for (var j = 0; j < taskName.length; j++){
     //フォルダ作成
     newDir = taskMngDir.createFolder(taskName[j]);
     newDirId = newDir.getId();
     //権限設定
     newDir.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
     //台本フォーマットをコピー
     newDocs = scenarioFmt.makeCopy(taskName[j], newDir);
     newDocsId = newDocs.getId();
     
     //権限設定
     newDocs.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
     
     //newLinks[0] = DIRLINK.replace('@share@', newDirId);
     //newLinks[1] = DOCLINK.replace('@share@', newDocsId);
     //リンクをシートに張り付け
     task_sht.getRange(newTaskRow[j], 3).setValue(DIRLINK.replace('@share@', newDirId));
     task_sht.getRange(newTaskRow[j], 4).setValue(DOCLINK.replace('@share@', newDocsId));
     
   }
   
}

/*******************************
 * 同名シートを更新する
 *
 ******************************/
function updateSheet(book, shtName){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var del_sht = book.getSheetByName(shtName)
  if (del_sht != null){
    book.deleteSheet(del_sht);
  }
  var copy_sht = sheet.getSheetByName(shtName).copyTo(book);
  copy_sht.setName(shtName);
  
}
/*******************************
 * ガイドラインを一括更新
 *
 ******************************/
function updateGaideToAll(){
  var gaide = "ガイドライン";
  var taskFolder = DriveApp.getFolderById("1ZaCardQi5aY8Q_ony7EHXCAbh6S8x4XR");
  var allFiles = taskFolder.getFiles();
  var file;
  var sht;
  
  //すべてのファイルでガイドを更新
  while(allFiles.hasNext()) {
    file = allFiles.next();
    sht = SpreadsheetApp.openById(file.getId());
    updateSheet(sht, gaide);
  }
}
