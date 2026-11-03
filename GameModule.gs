// ==================== 💡 互動小遊戲與擴充模組 (GameModule.gs) ====================

// 💡 1. 取得設定 (預設抓第一款，但支援傳入「遊戲名稱」以利未來擴充)
function getGameSettings_(targetGameName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('遊戲設定');
  if (!sh) {
    sh = ss.insertSheet('遊戲設定');
    sh.appendRow(['遊戲名稱', '開場標題', '開場內容', '選項1', '選項2', '選項3', '選項4', '詢問生日(Y/N)', '結語標題', '結語內容', '歡迎回來標題', '歡迎回來內容']);
    sh.appendRow([
      '三禮',
      '敷籙功曹降臨！', '本神官已感應到貴客的到來💡\n\n請依照古禮選擇『香』『花』『果』三禮之一敬獻於吾，請教貴客要選哪一個？',
      '香：貴客將獲知「命主之色」', '花：貴客將獲知「招財之色」', '果：貴客將獲知「逢貴之色」', '', 'Y',
      '祈願已成', '本神官將請鬼磕頭儘速為貴客測算並予回覆🎁',
      '貴客再臨', '很開心再次遇見貴客，本神官已等候您多時。'
    ]);
    sh.getRange("A1:L1").setFontWeight("bold").setBackground("#d9ead3");
    sh.setFrozenRows(1);
  }
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return {};
  var headers = data[0];

  var nameColIdx = -1;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === '遊戲名稱') nameColIdx = c;
  }

  var targetRowIdx = 1; // 預設固定抓第2列(第一款遊戲)
  
  if (targetGameName && nameColIdx !== -1) {
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][nameColIdx]).trim() === targetGameName) {
        targetRowIdx = r;
        break;
      }
    }
  }

  var values = data[targetRowIdx] || data[1];
  var settings = {};
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim()) settings[String(headers[i]).trim()] = String(values[i] || '');
  }
  return settings;
}

// 💡 2. 輔助陣列發送
function replyLineArray_(token, msgArray) {
  var key = props_().getProperty('LINE_TOKEN');
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post', headers: { Authorization: 'Bearer ' + key }, contentType: 'application/json',
    payload: JSON.stringify({ replyToken: token, messages: msgArray }), muteHttpExceptions: true
  });
}

// 💡 3. 組裝遊戲卡片 (帶入遊戲名稱，解決「遊戲_預設遊戲」的問題)
function getGameCard_(targetGameName) {
  var settings = getGameSettings_(targetGameName);
  var gameName = settings['遊戲名稱'] || '預設遊戲';
  var gameTitle = settings['開場標題'] || '敷籙功曹降臨！';
  
  var opts = [];
  for (var i = 1; i <= 4; i++) {
    if (settings['選項' + i]) opts.push(settings['選項' + i]);
  }
  var btns = opts.map(function(o) { 
    // 👉 確保遊戲名稱被包進按鈕的隱藏指令中
    return pbBtn_(o, 'action=game_opt&game=' + encodeURIComponent(gameName) + '&choice=' + encodeURIComponent(o), o, 'secondary'); 
  });
  return btnCard_(gameTitle, '#8B7355', settings['開場內容'] || '請選擇：', btns);
}

// 💡 4. 帶有遊戲記憶的專屬日期選擇器
function getGameDatePicker_(overrideText, gameName) {
  var txt = overrideText || "請選擇日期";
  var pbData = "action=select_birthday&game=" + encodeURIComponent(gameName);
  var now = new Date();
  var initialDate = (now.getFullYear() - 10) + '-' + ('0' + (now.getMonth() + 1)).slice(-2) + '-' + ('0' + now.getDate()).slice(-2);

  return {
    type: "flex", altText: txt, contents: {
      type: "bubble", body: {
        type: "box", layout: "vertical", contents: [
          { type: "text", text: txt, weight: "bold", size: "md", wrap: true },
          { type: "button", action: { type: "datetimepicker", label: "填寫本人生辰", data: pbData, mode: "date", initial: initialDate, min: "1900-01-01", max: now.toISOString().split('T')[0] }, style: "primary", color: "#4A7C59", margin: "md" }
        ]
      }
    }
  };
}

// 💡 5. 新加入好友處理
function handleFollowEvent_(ev, isNewUser) {
  if (isNewUser) {
    replyLine_(ev.replyToken, getGameCard_()); // 預設抓第一款
  } else {
    var settings = getGameSettings_(); 
    replyLine_(ev.replyToken, textCard_(settings['歡迎回來標題'] || '故人歸來', '#8B7355', settings['歡迎回來內容'] || '歡迎回來'));
  }
}

// 💡 6. 防呆動態紀錄：幫每款遊戲建立專屬紀錄欄位
function checkAndSaveGameChoice_(uid, gameName, choice) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LINE_User');
    if (!sh) return { played: false };
    
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var colName = '遊戲_' + gameName;
    var colIdx = -1, uidColIdx = -1;

    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim();
      if (h === colName) colIdx = i;
      if (h === 'userId') uidColIdx = i;
    }

    if (colIdx === -1) {
      colIdx = sh.getLastColumn();
      sh.getRange(1, colIdx + 1).setValue(colName);
    }
    if (uidColIdx === -1) return { played: false };

    var userRowIdx = -1;
    for (var r = 1; r < data.length; r++) {
      if (String(data[r][uidColIdx]).trim() === uid) {
        userRowIdx = r + 1;
        if (colIdx < data[r].length && String(data[r][colIdx]).trim() !== '') {
           return { played: true, oldChoice: String(data[r][colIdx]).trim() };
        }
        break;
      }
    }

    if (userRowIdx !== -1) {
       sh.getRange(userRowIdx, colIdx + 1).setValue(choice);
    }
  } catch(e) {}
  return { played: false };
}

function handleGamePostback_(ev, pbData) {
  var uid = ((ev.source || {}).userId) || '';
  
  function formatYMD(dateVal) {
    if (!dateVal) return '';
    var d = new Date(dateVal);
    if (isNaN(d.getTime())) return String(dateVal); 
    return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
  }

  // 抓出這是哪一款遊戲
  var gameMatch = pbData.match(/game=([^&]+)/);
  var gameName = gameMatch ? decodeURIComponent(gameMatch[1]) : '預設遊戲';

  // === 處理生日 ===
  if (pbData.indexOf('action=select_birthday') === 0) {
    var bdayRaw = (ev.postback && ev.postback.params && ev.postback.params.date) || '';
    var bday = formatYMD(bdayRaw);
    
    var isUpdated = false;
    if (bday && uid) {
      isUpdated = saveUserBirthday_(uid, bday); // 💡 接收是否有成功更新的狀態
    }
    
    var settings = getGameSettings_(gameName); 
    var savedChoice = getUserGameChoice_(uid, gameName);
    var choiceText = savedChoice ? '同時，喜悅的鑑賞貴客奉獻的『' + savedChoice + '』\n\n' : '';
    
    if (isUpdated) {
      // 第一次填寫成功
      var brithdayText = '貴客的生辰已如實登冊：' + bday + '\n（若有錯漏，再呼喚本座）';  
      replyLineArray_(ev.replyToken, [
        toMsg_(textCard_('敷籙功曹慎重的寫下貴客的生辰', '#4A7C59',  choiceText + brithdayText)),
        toMsg_(textCard_(settings['結語標題'] || '祈願已成', '#8B7355', settings['結語內容'] || '已收到回覆'))
      ]);
    } else {
      // 💡 防呆：已經有資料了，拒絕修改
      var ur = userRole_(uid);
      var oldBday = (ur && ur.birthday) ? formatYMD(ur.birthday) : '';
      if (oldBday) {
        replyLine_(ev.replyToken, textCard_('一人怎可二生？！', '#B04A3A',  '貴客生辰已存於玉牒之中，為：' + oldBday + '\n\n命數既定，豈可易改。若有特殊緣由，請另外呼喚本座。')); 
      } else {
        replyLine_(ev.replyToken, textCard_('近日必有怪事', '#B04A3A', '玉牒缺貨中，請呼喚本神官。'));
      }
    }
    return true;
  }
  
  // === 處理選項 ===
  if (pbData.indexOf('action=game_opt') === 0) {
    var choiceMatch = pbData.match(/choice=([^&]+)/);
    var choice = choiceMatch ? decodeURIComponent(choiceMatch[1]) : '禮物';
    var shortChoice = choice.split('：')[0]; 
    
    var status = checkAndSaveGameChoice_(uid, gameName, shortChoice);
    if (status.played) {
       replyLine_(ev.replyToken, textCard_('敷籙功曹看著貴客送來的『' + shortChoice + "』", '#8B7355', '貴客的美意，本神官甚感欣悅。既然心意已足，往後便無須如此拘禮。'));
       return true;
    }

    var settings = getGameSettings_(gameName);
    var actualGameName = settings['遊戲名稱'] || gameName || '預設遊戲';
    var ur = userRole_(uid); 

    if (String(settings['詢問生日(Y/N)']).toUpperCase() === 'Y') {
      if (ur && ur.birthday) {
        var formattedBday = formatYMD(ur.birthday);
        replyLineArray_(ev.replyToken, [
            toMsg_(textCard_('敷籙功曹慎重的寫下貴客的生辰', '#4A7C59', '同時，喜悅的鑑賞貴客奉獻的『' + shortChoice + '』\n\n貴客的生辰已如實登冊：' + formattedBday + '\n（若有錯漏，再呼喚本座）')),
            toMsg_(textCard_(settings['結語標題'], '#8B7355', settings['結語內容']))
        ]);
      } else {
        replyLine_(ev.replyToken, getGameDatePicker_("敷籙功曹滿意的收下貴客的『" + shortChoice + "』\n\n請您告訴本座您的生辰吧！", actualGameName));
      }
    } else {
      replyLineArray_(ev.replyToken, [
         toMsg_(textCard_(settings['結語標題'], '#8B7355', settings['結語內容']))
      ]);
    }
    return true;
  }
  return false;
}

function getUserGameChoice_(uid, gameName) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LINE_User');
    if (!sh) return '';
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var colName = '遊戲_' + gameName;
    var colIdx = -1, uidColIdx = -1;

    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim();
      if (h === colName) colIdx = i;
      if (h === 'userId') uidColIdx = i;
    }
    if (colIdx === -1 || uidColIdx === -1) return '';

    for (var r = 1; r < data.length; r++) {
      if (String(data[r][uidColIdx]).trim() === uid) {
         return String(data[r][colIdx] || '').trim();
      }
    }
  } catch(e) {}
  return '';
}