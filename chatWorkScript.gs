function sendDairyTask(){
  const cw = ChatWorkClient.factory({token: '@token@'});
  
  var sht = SpreadsheetApp.openById('1o8s4VaFvgzAOUkEOALxNcEVVKqV9gLnQXXoDNLDjS4I');
  
  var task_sht = sht.getSheetByName("受注");
  var last_row = task_sht.getLastRow();
  
  const CLT_NAME_COL = 2;
  const TASK_NAME_COL = 3;
  const NEW_FLG_COL = 5;
  const DEAD_LINE_COL = 6;
  
  const CHK_STA_ROW = 6;
  
  var client = [];
  var taskName = [];
  var alartTaskRow = [];
  var alartType = [];
  var task_tmp;
  var task_flg;
  var task_date;
  var tmp_date;
  var ret = 0
  
  var dldate;
  var t = Moment.moment();
  var todate = Moment.moment([t.year(), t.month(), t.date()]);
  
  var k = 0;
  for (var i = CHK_STA_ROW; i <= last_row; i++){
    task_tmp = task_sht.getRange(i, TASK_NAME_COL).getValue();
    task_flg = task_sht.getRange(i, NEW_FLG_COL).getValue();
    task_date = task_sht.getRange(i, DEAD_LINE_COL).getValue();
    
    
    tmp_date = Object.prototype.toString.call(task_date);
    if (tmp_date === undefined || tmp_date !== "[object Date]"){
      continue;
    }

    if (task_tmp == '' || task_flg === true){
      continue;
    }
    
    client[k] = task_sht.getRange(i, CLT_NAME_COL).getValue();
    taskName[k] = task_tmp;
    alartTaskRow[k] = i; 
    dldate = Moment.moment(task_date);
    
    alartType[k] = dldate.diff(todate,'days');
    alartType[k] = (dldate - todate)/1000/60/60/24 + (2/3);
    
    k++;
  }
  
  const def_label = '【@label@】\r\n';
  
  var working = def_label.replace('@label@', '作業中のタスク');
  var tommorow = def_label.replace('@label@', '明日締め切りのタスク');
  var ondead = def_label.replace('@label@', '本日締め切りのタスク');
  var lineover = def_label.replace('@label@', '締め切りを過ぎたタスク');
  
  var set_sht = sht.getSheetByName("設定・使い方");
  
  const DEF_SET_COL = 3;
  const DEF_ALT_ROW = 1;
  
  var message = todate.format('YYYY/M/D') + '\r\n';
  
  for(var x = 0; x < k; x++){
    working += client[x] + '様_' + taskName[x] + '\r\n';
    switch (true){
      case alartType[x] == 1:
        tommorow += client[x] + '様_' + taskName[x] + '\r\n';
        break;
      case alartType[x] == 0:
        ondead += client[x] + '様_' + taskName[x] + '\r\n';
        break;
      case alartType[x] < 0:
        lineover += client[x] + '様_' + taskName[x] + '\r\n';
        break;
    }
  }
  
  var m = DEF_ALT_ROW;
  var nonmsg = 0;
  if (set_sht.getRange(m++, DEF_SET_COL)){
    message += working + '\r\n';
    nonmsg++;
  }
  if (set_sht.getRange(m++, DEF_SET_COL)){
    message += tommorow + '\r\n';
    nonmsg++;
  }
  if (set_sht.getRange(m++, DEF_SET_COL)){
    message += ondead + '\r\n';
    nonmsg++;
  }
  if (set_sht.getRange(m++, DEF_SET_COL)){
    message += lineover + '\r\n';
    nonmsg++;
  }
  
  if (nonmsg == 0){
    return;
  }
  
  const ROOM_ID_ROW = 5;
  var roomId = set_sht.getRange(ROOM_ID_ROW, 2).getValue();
  
  if (!isNaN(roomId) && roomId !== ''){
    cw.sendMessage({
      room_id: roomId,
      body: message
    });
  }else{
    cw.sendMessageToMyChat(message);
  }
  
}
