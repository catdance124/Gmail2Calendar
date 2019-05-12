var calendar_id = PropertiesService.getScriptProperties().getProperty("CALENDAR_ID");
var cal = CalendarApp.getCalendarById(calendar_id);

/*メールを取得*/
function getMessages(){
  // メール取得
  var terms = 'label:netz'
  var threads = GmailApp.search(terms, 0, 50);
  
  // まとめる
  var Messages = [];
  threads.forEach(function(thread){
    Messages = Messages.concat(thread.getMessages());
  });
  // スター付きか判定
  var StarMessages = [];
  Messages.forEach(function(message){
    if(!message.isStarred()){ StarMessages.push(message); }
  });
  // 時系列にソート
  StarMessages.sort(function(a,b){ return (a.getDate() > b.getDate() ? 1: -1) });
  
  return StarMessages;
}

/*開始時間と終了時間を取得*/
function getSEtime(date, time){
  var [start, end] = time.split('-');  // 開始と終了に分割
  var Sdt = new Date(date.getFullYear(),date.getMonth(),date.getDate(),start.split(':')[0],start.split(':')[1]);
  var Edt = new Date(date.getFullYear(),date.getMonth(),date.getDate(),end.split(':')[0],end.split(':')[1]);
  return [Sdt, Edt];
}

/*日付を年月日に拡張する*/
function rightDate(day){
  var today = new Date();
  var year = today.getFullYear();
  var date = new Date(year, Number(day.slice(0,2))-1, Number(day.slice(3,5)));
  today.setMonth(today.getMonth()-6);
  if (date.getTime() < today.getTime()){
    date.setFullYear(year+1);  // planが半年以上過去であるなら来年として扱う．
  }
  return date;
}

/*1コマごとに配列化して返す*/
function separate(text, day){
  var date = rightDate(day);  // 日付しかないので，1月とかを来年扱いに
  var str = text.split(/\n/);  // 改行で分割
  var array = []
  for (var row=0; row<str.length; row+=3){  // 3行単位でループ
    var time = str[row].trim().slice(0,-3);
    var student = str[row+1].trim();
    var detail = str[row+2];
    if(detail)  detail = detail.trim();  // 存在の可否
    var [start, end] = getSEtime(date,time);
    array.push([start, end, student, detail]);  // [時間，生徒名，詳細]をリストとして格納
  }
  return array;  // [[時間，生徒名，詳細], [時間，生徒名，詳細], ...] 2重配列
}

/*渡された辞書のkey(日付)を走査，eventを走査し登録*/
function addEvent(dic,end){
  var dicobj = Object.keys(dic);  // 辞書キー取得
  if(dicobj.length != 1){
    var first = dic[dicobj[0]][0][0];
    if(isToday(first)){delete dic[dicobj[0]];}  // 辞書中から今日を削除
    dicobj = Object.keys(dic)
  }
  deleteDayEvents(dic[dicobj[0]][0][0], dicobj.length, end);
  for (var key in dic){
    if(!(dic[key].length == 1 && dic[key][0][2] == '開校担当')){  // 開校担当だけの日は書き込まない
      for (var i in dic[key]){
        var start = dic[key][i][0];
        var end = dic[key][i][1];
        var student = dic[key][i][2];
        var detail = dic[key][i][3];
        if(['開校担当','閉校担当'].indexOf(student) != -1){ end.setMinutes(end.getMinutes()+30); }
        cal.createEvent(student, start, end, {description:detail});
      }
    }
  }
}

/*渡されたDateが今日かを判定*/
function isToday(date){
  var today = new Date();
  if(date.getFullYear() == today.getFullYear() &&
     date.getMonth() == today.getMonth() &&
     date.getDate() == today.getDate()){
       return true;
     }else{
       return false;
     }
}

/*日付ごとのイベントを削除*/
function deleteDayEvents(origin, len, end){
  var start = new Date(origin);
  if(len != 1){
    start.setHours(0);
    end.setHours(23);
    var events = cal.getEvents(start,end);
  }else{
    var events = cal.getEventsForDay(start);
  }
  if(events.length){
    for(var i in events){
      events[i].deleteEvent();
    }
  }
}

/*時間内重複イベントを削除*/
function deleteEvents(start, end){
  var events = cal.getEvents(start,end);
  if(events.length){
    for(var i in events){
      events[i].deleteEvent();
    }
  }  // 重複があれば削除
}

/*===========================↑メソッド↑=============================================================*/

/*○○/○○指導予定確認*/
function tommorow_plan(message){
  var day = message.getSubject().match(/\d{2}\/\d{2}/).toString();
  var text = message.getPlainBody().match(/\d{2}\:\d{2}-[\s\S]*?※/).toString().replace('※','').trim();
  var planDic = {};
  planDic[day] = separate(text, day);
  addEvent(planDic, null);
  message.star();
}

/*指導予定(○○/○○～○○/○○)*/
function future_plan(message){
  var maintext = message.getPlainBody().replace(/指導予定：\d*件/,'').trim();  // メイン箇所を切り抜き
  var planList = maintext.replace(/\d{2}\/\d{2}\(/mg,'※$&').split('※').slice(1);  // 一日ごとに切り分け
  var planDic = {};
  for(var i in planList){
    planList[i] = planList[i].trim();
    var day = planList[i].slice(0,5);  // ○○/○○　日付
    var text = planList[i].slice(10);  // 日付以外　時間\n名前\n科目\n *n
    planDic[day] = separate(text,day);  //  {key:day value:[[時間，生徒名，詳細], [時間，生徒名，詳細], ...]}
  }
  var end = message.getSubject().toString().slice(-6,-1);
  end = rightDate(end);
  addEvent(planDic, end);
  message.star();
}

/*振替受付*/
function delete_plan(message){
  var maintext = message.getPlainBody().match(/\d{2}\/\d{2}[\s\S]*?宇宿/).toString()
                                       .replace(/\(.\)/,'').replace(' 宇宿','');
  var [day,time] = maintext.split('\n');
  var date = rightDate(day);
  var [start, end] = getSEtime(date,time);
  deleteEvents(start,end);
  message.star();
}

/*振替決定*/
function add_plan(message){
  var student = message.getPlainBody().match(/生徒名：[\s\S]*?\n/).toString().replace('生徒名：','').replace('\n','');
  var maintext = message.getPlainBody().match(/\d{2}\/\d{2}[\s\S]*?※/).toString().replace('※','')
                                       .replace(/\(.\)/,'').replace(' 宇宿','').trim();
  var str = maintext.split(/\n/);  // 改行で分割
  var [start, end] = getSEtime(rightDate(str[0]), str[1]);
  start.setMinutes(start.getMinutes()+10);
  var detail = str[2].slice(3);
  cal.createEvent(student, start, end, {description:detail});
  message.star();
}

/*main script*/
function netz_plans(){
  // メール取得
  var messages = getMessages();
  // タイトルごとに処理
  messages.forEach(function(message){
    var sub = message.getSubject();
    if(sub.match(/指導予定確認/)){tommorow_plan(message);}
    else if(sub.match(/指導予定\(/)){future_plan(message);}
    else if(sub.match(/振替受付/)){delete_plan(message);}
    else if(sub.match(/振替決定/)){add_plan(message);}
    else{message.star();}
  });
}
