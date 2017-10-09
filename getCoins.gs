//トリガーを分単位で時間設定
function setTrigger() {
  var triggerDay = new Date();
  triggerDay.setHours(triggerDay.getHours()+1);
  triggerDay.setMinutes(00);
  ScriptApp.newTrigger("main").timeBased().at(triggerDay).create();
}

// その日のトリガーを削除する関数(消さないと残る)
function deleteTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for(var i=0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == "main") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

//シートからメール送信対象アドレスを取得
function getToAddress(id, ss) {
  var sheet   = ss.getSheetByName('YOUR SHEET NAME HERE');
  //シートの1列目にメールアドレスを記載
  var lastrow = sheet.getLastRow();
  var lastcol = sheet.getLastColumn();
  var sheetdata = sheet.getSheetValues(1, 1, lastrow, lastcol);
  var toAdr = [];
  for (i=0; i<sheetdata.length; i++) {
    toAdr.push(sheetdata[i][0]);
  }
  return toAdr;
}

function main() {
  //トリガーを削除
  deleteTrigger();
  
  //Parserを使ってnameを抽出
  var url = 'https://coinmarketcap.com/all/views/all/';
  var html = UrlFetchApp.fetch(url).getContentText('utf-8');    
  var doc = Parser.data(html)
            .from("<tbody>")
            .to("</tbody>")
            .build();
  var name = Parser.data(doc)
            .from('<td class="no-wrap currency-name">')
            .to("</td>")
            .iterate();
  var cap = Parser.data(doc)
            .from('<td class="no-wrap market-cap text-right"')
            .to("data-btc")
            .iterate();
  for(i=0; i<name.length; i++){
    name[i] = name[i].replace(/\s*<div class=\"[.\s\S]*\"*><\/div>\s*<a href=\"[\w-.\/?%&=]*\"*>([\S\s]+)<\/a>\s*/g, '$1');
    cap[i] = +cap[i].replace(/data-usd=\"([0-9\.]+)\"\s/g, '$1');
  }
  //Logger.log(name);
        
  //スプレッドシートを取得
  var id      = 'YOUR SHEET ID HERE', //シートID
      ss   = SpreadsheetApp.openById(id),
      sheet   = ss.getSheetByName('YOUR SHEET NAME HERE'),
      col = 1;

  var today = new Date();
  today = today.getFullYear() + "/" +  (today.getMonth() + 1) + "/"+ today.getDate();

  /*
  //初回のみスプレッドシートに書き込み
  var n = 1;
  for (i=0; i<name.length; i++) {
    //market capが$2000000以上のものを取得
    if (cap[i] >= 2000000) {
      sheet.getRange(n+1, col).setValue(name[i]);
      sheet.getRange(n+1, col+1).setValue(today);
      n++;
    }
  }
  */

  //スプレッドシートに記入済みのリストを取得  
  var startrow = 1;
  var startcol = 1;
  var lastrow = sheet.getLastRow();
  var lastcol = sheet.getLastColumn();
  var sheetdata = sheet.getSheetValues(startrow, startcol, lastrow, lastcol);
  
  //新しいものが増えていないかチェック
  var new_coin = [];
  for (i=0; i<name.length; i++) {
    if (cap[i] < 2000000) break;
    for (j=0; j<sheetdata.length; j++) {
      if (name[i] == sheetdata[j][0]) {
        break;
      } else {
        if (j==sheetdata.length-1) {
          new_coin.push([name[i], cap[i]]);
        }
      }
    }
  }

  //新しいものがあればスプレッドシートに書き込み
  for (i=0; i<new_coin.length; i++) {
    sheet.getRange(lastrow+i+1, col).setValue(new_coin[i]);
    sheet.getRange(lastrow+i+1, col+1).setValue(today);
  }
  
  //新しいものがあればメール送信
  var toAdr = getToAddress(id, ss),
      subject = '新しいコイン追加のお知らせ',
      body = '本日、以下のコインの追加を確認しました。\n';
  for (i=0; i<new_coin.length; i++) {
    body += new_coin[i][0] + '：$' + new_coin[i][1] + '\n';
  }
  if (new_coin.length>0) {
    MailApp.sendEmail(toAdr, subject, body);
  }
}
