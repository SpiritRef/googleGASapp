// ===== GAS-LINE：通知＋指令＋訓練以後也放這。屬性：LINE_TOKEN / ADMIN_EMAIL / RELAY_SECRET / GROUP_ID(備用) =====
var SHEET_MAIN = '諮詢表單內容';
var SHEET_GROUP = 'LINE_Group';
var SHEET_USER = 'LINE_User';
var STATUS_COL = 12;
var REMIND_DAYS = { '待處理': 3, '處理中': 7, '無法聯繫': 5 };
var VIP_QUOTA_RESERVE = 100;
var SHEET_SOP = 'SOP與庫存表';
var SHEET_TRAIN = '訓練紀錄';

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function props_() { return PropertiesService.getScriptProperties(); }

function doPost(e) {
  try {
    var data = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (data.relay) {
      if (data.relay !== props_().getProperty('RELAY_SECRET')) {
        return json_({ success: false, message: 'forbidden' });
      }
      return json_(onRelay_(data.data || {}));
    }
    if (data.events) return handleLine_(data);
    return json_({ success: false, message: '未知動作' });
  } catch (err) {
    return json_({ success: false, message: '系統錯誤: ' + err });
  }
}

function onRelay_(d) {
  var card = flexCard_('新諮詢・' + d.tier + ' ' + d.caseId, d.tier === 'VIP' ? '#8B7355' : '#6d6d6d', [
    ['姓名', d.name], ['項目', d.cat], ['聯絡', d.contact], ['內容', String(d.msg || '').slice(0, 300)]
  ], null);
  routeNotify_(d.tier, card);
  return { success: true };
}

function isAllowedGroup_(gid) {
  try {
    gid = String(gid).trim();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_GROUP);
    if (!sh || sh.getLastRow() < 2) return false;
    if (String(sh.getRange(1, 1).getValue() || '').trim() === 'groupId') {
      var v = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
      for (var i = 0; i < v.length; i++) {
        if (String(v[i][0]).trim() === gid) return String(v[i][2]).toUpperCase() === 'Y';
      }
      return false;
    }
    var ids = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues();
    for (var k = 0; k < ids.length; k++) {
      if (String(ids[k][0]).trim() === gid) return true;
    }
    return false;
  } catch (e) { return false; }
}

function groupRole_(gid) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_GROUP);
    if (!sh || sh.getLastRow() < 2) return 'guest';
    if (String(sh.getRange(1, 1).getValue() || '').trim() !== 'groupId') return 'guest';
    var n = sh.getLastRow() - 1;
    var v = sh.getRange(2, 1, n, 4).getValues();
    var roles = sh.getLastColumn() >= 6 ? sh.getRange(2, 6, n, 1).getValues() : null;
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][0]).trim() === String(gid).trim()) {
        return String((roles && roles[i][0]) || 'guest').trim() || 'guest';
      }
    }
  } catch (e) {}
  return 'guest';
}
function userSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_USER) || ss.insertSheet(SHEET_USER);
  if (sh.getLastRow() === 0) sh.appendRow(['userId', '名稱', '角色', '啟用', '綁定', '備註']);
  return sh;
}

function userRole_(uid) {
  try {
    var sh = userSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return null;
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var colIdx = {};
    for (var c = 0; c < headers.length; c++) colIdx[String(headers[c]).trim()] = c;
    var uidCol = colIdx['userId'] !== undefined ? colIdx['userId'] : 0;
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][uidCol]).trim() === String(uid).trim()) {
        var r = data[i];
        return { 
          name: colIdx['名稱'] !== undefined ? String(r[colIdx['名稱']] || '') : '', 
          role: colIdx['角色'] !== undefined ? String(r[colIdx['角色']] || '').trim() || 'guest' : 'guest', 
          on: colIdx['啟用'] !== undefined ? String(r[colIdx['啟用']]).toUpperCase() === 'Y' : true, 
          bind: colIdx['綁定'] !== undefined ? String(r[colIdx['綁定']] || '').trim() : '',
          birthday: colIdx['生日'] !== undefined ? String(r[colIdx['生日']] || '').trim() : ''
        };
      }
    }
  } catch (e) {}
  return null;
}

function guestEvent_(ev, ur, inGroup) {
  if (!ev.replyToken) return;
  if (ev.type === 'postback' && ev.postback) {
    var pbData = String(ev.postback.data || '');
    if (handleGamePostback_(ev, pbData)) return;
  }

  var t = (ev.type === 'message' && ev.message && ev.message.type === 'text') ? String(ev.message.text || '').trim() : '';

  // ==================== 💡 呼叫翻譯模組 ====================
  var transCard = checkTranslation_(t);
  if (transCard) {
    replyLine_(ev.replyToken, transCard);
    return;
  }
  // =========================================================

  if (inGroup) {
    if (/^(指令|說明)$/i.test(t) || ev.type === 'postback') {
      replyLine_(ev.replyToken, textCard_('敬請期待', '#8B7355', '本神官尚在煉製新法，請貴客稍候'));
    } else if (/^(綁定|我的案件|查我的案件|查詢我的案件)/.test(t)) {
      replyLine_(ev.replyToken, textCard_('敷籙功曹專司項', '#8B7355', '貴客所詢乃本神官與貴客間之冥契\n請貴客獨自向本神官查閱'));
    }
    return;
  }
  var bm = t.match(/^綁定\s+([A-Za-z0-9\-]+)\s+(.+)$/);
  if (bm) {
    var uid = ((ev.source || {}).userId) || '';
    var found = findCaseRow_(bm[1]);
    if (!found) { replyLine_(ev.replyToken, textCard_('玉牒搜檢無果', '#B04A3A', '尚無貴客所述案件編號😑')); return; }
    var r = found.sheet.getRange(found.row, 1, 1, STATUS_COL).getValues()[0];
    if (norm_(r[2]) !== norm_(bm[2])) { replyLine_(ev.replyToken, textCard_('玉牒匹配無果', '#B04A3A', '貴客尊諱與案件編號懸殊😓')); return; }
    var cur = (ur && ur.bind) ? ur : userRole_(uid);
    var lst = (cur && cur.bind) ? String(cur.bind).split(',').map(function(s) { return s.trim(); }).filter(Boolean) : [];
    var nid = String(r[1]).trim();
    if (lst.indexOf(nid) === -1) lst.push(nid);
    
    var ush = userSheet_();
    var uidx = userRowIdx_(uid);
    
    // 💡 動態抓取欄位索引
    var headers = ush.getRange(1, 1, 1, ush.getLastColumn()).getValues()[0];
    var colMap = {};
    for (var c = 0; c < headers.length; c++) colMap[String(headers[c]).trim()] = c;
    
    var keepRole = 'customer';
    if (uidx > 0 && colMap['角色'] !== undefined) {
      var curRole = String(ush.getRange(uidx, colMap['角色'] + 1).getValue() || '').trim();
      if (curRole === 'admin') keepRole = 'admin';
    }
    
    // 💡 組裝動態列資料（如果是舊用戶，先抓整列當底底避免覆寫到其他如生日的資料）
    var newRow = new Array(headers.length).fill('');
    if (uidx > 0) {
       newRow = ush.getRange(uidx, 1, 1, headers.length).getValues()[0];
    }
    
    if (colMap['userId'] !== undefined) newRow[colMap['userId']] = uid;
    if (colMap['名稱'] !== undefined) newRow[colMap['名稱']] = String(r[2]).trim();
    if (colMap['角色'] !== undefined) newRow[colMap['角色']] = keepRole;
    if (colMap['啟用'] !== undefined) newRow[colMap['啟用']] = 'Y';
    if (colMap['綁定'] !== undefined) newRow[colMap['綁定']] = lst.join(',');
    if (colMap['備註'] !== undefined) newRow[colMap['備註']] = '綁定 ' + new Date();
    
    if (uidx > 0) ush.getRange(uidx, 1, 1, headers.length).setValues([newRow]);
    else ush.appendRow(newRow);
    
    replyLine_(ev.replyToken, textCard_('玉牒成冊', '#4A7C59', r[2] + ' 貴客您好，共裝幀' + lst.length + '篇\n' + lst.join('\n').slice(0, 500)));
    return;
  }
  if (/^(我的案件|查我的案件|查詢我的案件)$/.test(t)) {
    var me2 = (ur && ur.on) ? ur : userRole_(((ev.source || {}).userId) || '');
    if (!me2 || !me2.bind) {
      replyLine_(ev.replyToken, textCard_('玉牒雅卷', '#8B7355', '查詢進度前，請先填寫『諮詢表單』，方可啟用綁定功能：\n綁定 案件編號 姓名\n若有其他疑惑，請誠心呼喚本神官名號三稱。\n吾身去也~'));
      return;
    }
    var ids = String(me2.bind).split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    if (!ids.length) { replyLine_(ev.replyToken, textCard_('玉牒字號訛誤', '#B04A3A', '[' + me2.bind + ']\n本神官恐有筆誤，請向鬼磕頭查證。')); return; }
    if (ids.length === 1) {
      var f = findCaseRow_(ids[0]);
      if (!f) { replyLine_(ev.replyToken, textCard_('玉牒字號訛誤', '#B04A3A', '[' + ids[0] + ']\n本神官恐有筆誤，請向鬼磕頭查證。')); return; }
      var rr = f.sheet.getRange(f.row, 1, 1, STATUS_COL).getValues()[0];
      replyLine_(ev.replyToken, flexCard_('我的案件 ' + rr[1], '#4A5A6A', [
        ['項目', rr[7]], ['狀態', String(rr[11] || '').trim() || '待處理'],
        ['內容', String(rr[8] || '').slice(0, 200)]
      ], null));
      return;
    }
    var lines = [];
    ids.reverse().slice(0, 10).forEach(function(id) {
      var ff = findCaseRow_(id);
      if (!ff) {
        lines.push('【' + id + '】' + '\n狀態：本神官恐有筆誤，請向鬼磕頭查證。');
        return;
      }
      var qq = ff.sheet.getRange(ff.row, 1, 1, STATUS_COL).getValues()[0];
      var type = qq[7] || '未分類';
      var status = String(qq[11] || '').trim() || '待處理';

      // 擷取內容前 30 個字，超過則加上刪節號，避免洗版
      var fullContent = String(qq[8] || '');
      var contentSummary = fullContent.slice(0, 30);
      if (fullContent.length > 30) {
        contentSummary += '...';
      }

      // 將每一筆的資訊格式化，用換行排版比較清楚
      lines.push('【' + id + '】' + type + '\n狀態：' + status + '\n內容：' + contentSummary);
    });
    replyLine_(ev.replyToken, textCard_('我的玉牒共' + lines.length + '篇', '#4A5A6A', lines.join('\n\n') || '玉牒尚未纂修'));
    return;
  }
  if (t === '' || ev.type === 'postback' || /^(指令|說明|綁定)$/i.test(t)) {
    replyLine_(ev.replyToken, textCard_('貴客玉牒', '#8B7355', '僅供已填寫『諮詢表單』的貴客方可使用\n填寫諮詢表單後即可綁定案件\n綁定案件後可查進度：\n綁定 案件編號 姓名\n例：綁定 20261013-K9X8A2 王小明\n查詢進度請打：『我的案件』\n本神官將為您詳細記載祈願與牽掛\n敷籙功曹的威德持續放送中，敬請期待！'));
  }
}

function addAdmins_(ev) {
  if (!ev.replyToken) return;
  var me = userRole_(((ev.source || {}).userId) || '');
  var by = ((me && me.name) || '管理員');
  var ids = [];
  var men = (ev.message && ev.message.mention && ev.message.mention.mentionees) || [];
  men.forEach(function(x) { if (x && x.userId && !x.isSelf && ids.indexOf(x.userId) === -1) ids.push(x.userId); });
  var t = (ev.message && ev.message.text) ? String(ev.message.text) : '';
  var m2 = t.match(/U[0-9a-f]{32}/i);
  if (m2 && ids.indexOf(m2[0]) === -1) ids.push(m2[0]);
  
  var doneN = 0;
  var sh = userSheet_();
  var gid0 = ((ev.source || {}).groupId) || '';
  
  // 💡 動態抓取欄位索引
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var colMap = {};
  for (var c = 0; c < headers.length; c++) colMap[String(headers[c]).trim()] = c;
  
  ids.forEach(function(id) {
    var idx = userRowIdx_(id);
    var nm = '';
    
    if (idx > 0) {
      // 舊用戶：只針對指定欄位更新，不動其他資料
      var rowData = sh.getRange(idx, 1, 1, headers.length).getValues()[0];
      var curRole = (colMap['角色'] !== undefined) ? String(rowData[colMap['角色']]).trim() : '';
      var curActive = (colMap['啟用'] !== undefined) ? String(rowData[colMap['啟用']]).toUpperCase() : '';
      var curName = (colMap['名稱'] !== undefined) ? String(rowData[colMap['名稱']]).trim() : '';
      
      if (curRole === 'admin' && curActive === 'Y') return; // 已經是管理員
      
      if (!curName) nm = gid0 ? lineGroupMemberName_(gid0, id) : lineProfile_(id);
      
      if (colMap['角色'] !== undefined) sh.getRange(idx, colMap['角色'] + 1).setValue('admin');
      if (colMap['啟用'] !== undefined) sh.getRange(idx, colMap['啟用'] + 1).setValue('Y');
      if (colMap['備註'] !== undefined) sh.getRange(idx, colMap['備註'] + 1).setValue('加 by ' + by + ' ' + new Date());
      if (nm && colMap['名稱'] !== undefined) sh.getRange(idx, colMap['名稱'] + 1).setValue(nm);
      doneN++;
    } else {
      // 新用戶：依照標題組合陣列新增
      nm = gid0 ? lineGroupMemberName_(gid0, id) : lineProfile_(id);
      var newRow = new Array(headers.length).fill('');
      if (colMap['userId'] !== undefined) newRow[colMap['userId']] = id;
      if (colMap['名稱'] !== undefined) newRow[colMap['名稱']] = nm;
      if (colMap['角色'] !== undefined) newRow[colMap['角色']] = 'admin';
      if (colMap['啟用'] !== undefined) newRow[colMap['啟用']] = 'Y';
      if (colMap['備註'] !== undefined) newRow[colMap['備註']] = '加 by ' + by + ' ' + new Date();
      sh.appendRow(newRow);
      doneN++;
    }
  });
  
  if (!ids.length) {
    replyLine_(ev.replyToken, textCard_('加管理員', '#8B7355', '用法（一句話）：@要加入的人 加管理員\n或：加管理員 U開頭ID'));
    return;
  }
  if (!doneN) {
    replyLine_(ev.replyToken, textCard_('加管理員', '#8B7355', '這些人都已經是管理員了'));
    return;
  }
  replyLine_(ev.replyToken, textCard_('已加入管理員', '#4A7C59', '共' + doneN + '位\n打「管理員名單」查看'));
}

function adminList_() {
  try {
    var sh = userSheet_();
    if (sh.getLastRow() >= 2) {
      // 💡 動態抓取欄位索引
      var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      var colMap = {};
      for (var c = 0; c < headers.length; c++) colMap[String(headers[c]).trim()] = c;
      
      var v = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
      var out = [];
      v.forEach(function(r) {
        var role = (colMap['角色'] !== undefined) ? String(r[colMap['角色']]).trim() : '';
        var active = (colMap['啟用'] !== undefined) ? String(r[colMap['啟用']]).toUpperCase() : '';
        
        if (role === 'admin' && active === 'Y') {
          var id = (colMap['userId'] !== undefined) ? String(r[colMap['userId']]) : '';
          var nm = (colMap['名稱'] !== undefined) ? String(r[colMap['名稱']]) : '未命名';
          out.push((nm || '未命名') + '（' + id.slice(0, 6) + '…' + id.slice(-4) + '）');
        }
      });
      if (out.length) return { card: textCard_('管理員共' + out.length + '位', '#8B7355', out.join('\n')), copy: null };
    }
  } catch (e) {}
  return { card: textCard_('管理員名單', '#8B7355', '尚無管理員，請手動新增第一位'), copy: null };
}
function handleAllowedEvent_(ev) {
  var t0 = (ev.type === 'message' && ev.message && ev.message.type === 'text') ? String(ev.message.text || '').trim() : '';
  if (/加管理員/.test(t0)) { addAdmins_(ev); return; }
  if (ev.type === 'postback' && ev.postback && ev.replyToken) {
    var pbData = String(ev.postback.data || '');

    if (handleGamePostback_(ev, pbData)) return;

    var ansPb = handleTrainPostback_(String(ev.postback.data || ''));
    if (ansPb === true) return; // ignore 偽停用按鈕，不回覆
    if (ansPb) {
      if (ansPb.card) replyLine_(ev.replyToken, ansPb.card, ansPb.copy);
      else replyLine_(ev.replyToken, ansPb);
    }
    return;
  }
  if (ev.type === 'message' && ev.message && ev.message.type === 'text' && ev.replyToken) {
    var ans = handleCommand_(String(ev.message.text || ''));
    if (ans) {
      if (ans && typeof ans === 'object' && ans.card) replyLine_(ev.replyToken, ans.card, ans.copy);
      else replyLine_(ev.replyToken, ans);
    }
  }
}

function handleLine_(data) {
  var events = data.events || [];
  events.forEach(function(ev) {
    var src = ev.source || {};
    var gid = src.groupId || '';
    var uid = src.userId || '';
    
    var existingUser = uid ? userRole_(uid) : null;
    var isNewUser = (existingUser === null);
    
    if (gid) ensureGroupRow_(gid);
    if (uid) ensureUserRow_(uid, gid || null);
    
    var ur = uid ? userRole_(uid) : null;
    var isAdmin = !!(ur && ur.on && ur.role === 'admin');

    // if (ev.type === 'follow') {
    //   handleFollowEvent_(ev, isNewUser);
    //   return;
    // }

    if (ev.type === 'message' && ev.message && ev.message.type === 'text') {
      var testText = String(ev.message.text || '').trim();
      if (testText === '測試開場白') {
        handleFollowEvent_(ev, true); // 傳入 true，強制模擬「新用戶」看到小遊戲
        return;
      }
      if (testText === '測試舊客') {
        handleFollowEvent_(ev, false); // 傳入 false，強制模擬「舊客」看到歡迎回來
        return;
      }
    }
    if (gid) {
      if (!isAllowedGroup_(gid)) return;
      if ((groupRole_(gid) || 'guest') !== 'admin') { guestEvent_(ev, ur, true); return; }
      handleAllowedEvent_(ev);
    } else if (uid) {
      if (!ev.replyToken) return;
      if (isAdmin) handleAllowedEvent_(ev);
      else guestEvent_(ev, ur, false);
    }
  });
  return json_({ ok: true });
}

function findCaseRow_(caseId) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN);
  if (!sh || sh.getLastRow() < 2) return null;
  var ids = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim().toUpperCase() === String(caseId).trim().toUpperCase()) {
      return { sheet: sh, row: i + 2 };
    }
  }
  return null;
}

// ==================== 💡 翻譯共用模組 ====================
function checkTranslation_(text) {
  if (!text) return null;
  var t = String(text).trim();
  
  // 1. 判斷是否有明確的翻譯指令 (支援: 翻譯 / 中翻英 / 英翻中)
  var trMatch = t.match(/^(翻譯|中翻英|英翻中)\s+(.+)$/);
  if (trMatch) {
    var mode = trMatch[1];
    var sourceText = trMatch[2].trim();
    var targetLang = (mode === '英翻中') ? 'zh-TW' : 'en';
    try {
      var translatedText = LanguageApp.translate(sourceText, '', targetLang);
      return textCard_('翻譯結果', '#4A5A6A', translatedText);
    } catch (e) {
      return textCard_('翻譯發生錯誤', '#B04A3A', String(e.message));
    }
  }

  // 2. 自動偵測：智慧比例判斷法
  if (!/^(指令|說明|綁定|我的案件|查詢)/i.test(t)) {
    try {
      // 先把網址(http/https)移除，避免網址被當成大量英文
      var textWithoutUrl = t.replace(/https?:\/\/[^\s]+/gi, '');
      
      // 計算純英文字母數量 (算字母總數)
      var engChars = textWithoutUrl.match(/[a-zA-Z]/g);
      var engLen = engChars ? engChars.length : 0;
      
      // 計算中文字數量
      var chiChars = textWithoutUrl.match(/[\u4e00-\u9fa5]/g);
      var chiLen = chiChars ? chiChars.length : 0;
      
      // 💡 判斷是否需要翻譯的條件：
      // (A) 純英文：完全沒中文，且至少有 2 個英文字母 (例如 hi, ok)
      var isPureEnglish = (chiLen === 0 && engLen >= 2);
      
      // (B) 英文為主：有中文，但「英文字母數量」大於「中文字數的 2 倍」
      // (例如 "hello hello 你好阿" -> 10個英文 > 3個中文*2)
      // (不會誤判 "官方Line帳號" -> 4個英文 沒有大於 5個中文*2)
      var isEnglishDominated = (chiLen > 0 && engLen > (chiLen * 2));

      if (isPureEnglish || isEnglishDominated) {
        var autoTranslated = LanguageApp.translate(t, 'en', 'zh-TW');
        
        // 確保翻譯結果跟原本不一樣才發送
        if (t.toLowerCase() !== autoTranslated.toLowerCase()) {
          return textCard_('自動翻譯', '#4A5A6A', autoTranslated);
        }
      }
    } catch (e) {
      // 靜默忽略錯誤
    }
  }

  return null;
}

function handleCommand_(text) {
  var t = text.trim();
  if (/^(指令|說明)$/i.test(t)) return menuCard_();
  if (/^(訓練總覽|全體進度)$/.test(t)) return trainOverview_();
  if (/^(庫存|庫存查詢)$/.test(t)) return stockText_();
  if (/^管理員名單$/.test(t)) return adminList_();

  var tm = t.match(/^(敕令|領取|完成)\s+(\S+)\s+(.+)$/);
  if (tm) return setTrain_(tm[2], tm[3].trim(), tm[1] === '完成' ? 'V' : tm[1] === '敕令' ? '已敕令' : '已領取');
  var am = t.match(/^補庫存\s+(.+)\s+(蓋章|列印)$/);
  if (am) {
    ensureTrainSheets_();
    var mm2 = sopMap_().map[norm_(am[1])];
    if (!mm2) return textCard_('找不到單元', '#B04A3A', am[1] + '（用「庫存」看正確名稱）');
    var sh2 = sheetByName_(SHEET_SOP);
    var isC = am[2] === '蓋章';
    var newB, newC;
    if (isC) {
      if (mm2.b > 0) { newB = mm2.b - 1; newC = mm2.c + 1; }
      else { newB = mm2.b; newC = mm2.c + 1; }
    } else {
      newB = mm2.b + 1; newC = mm2.c;
    }
    sh2.getRange(mm2.row, 2).setValue(newB);
    sh2.getRange(mm2.row, 3).setValue(newC);
    return textCard_('已補庫存', '#4A7C59', am[1] + ' ' + am[2] + '+1（現印' + newB + '章' + newC + '）');
  }
  var sm = t.match(/^設庫存\s+(.+)\s+(列印|蓋章)\s+(\d+)$/);
  if (sm) return setStock_(sm[1].trim(), sm[2], Number(sm[3]));
  if (/^(庫存盤點|盤點庫存)$/.test(t)) return stockCheckAll_();
  var vm = t.match(/^盤點\s+(.+)$/);
  if (vm) return stockCheckVol_(vm[1].trim());
  var lm = t.match(/^(列表|清單)\s*(待處理|處理中|無法聯繫|未結案|全部|結案|已結案)?$/);
  if (lm) {
    var lr = listCases_(lm[2] || '未結案');
    return (lr && lr.card) ? lr : { card: lr, copy: null };
  }

  var m = t.match(/^(結案|處理中|待處理|無法聯繫|查詢)\s+([A-Za-z0-9\-]+)$/);
  if (!m) {
    // ==================== 💡 呼叫翻譯模組 ====================
    var transCard = checkTranslation_(t);
    if (transCard) return transCard;
    // =========================================================
    return null;
  }
  var found = findCaseRow_(m[2]);
  if (!found) return textCard_('找不到案件', '#B04A3A', m[2].toUpperCase());
  if (m[1] === '查詢') {
    var r = found.sheet.getRange(found.row, 1, 1, STATUS_COL).getValues()[0];
    var qcard = flexCard_('案件查詢 ' + r[1], '#4A5A6A', [
      ['姓名', r[2]], ['項目', r[7]], ['狀態', (String(r[11] || '').trim() || '待處理')],
      ['聯絡', r[3]], ['內容', String(r[8] || '').slice(0, 300)]
    ], null);
    return { card: qcard, copy: r[1] };
  }
  found.sheet.getRange(found.row, STATUS_COL).setValue(m[1] === '結案' ? '已結案' : m[1]);
  return textCard_('已更新', '#4A7C59', m[2].toUpperCase() + ' 已改為「' + (m[1] === '結案' ? '已結案' : m[1]) + '」');

}

function listCases_(which) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN);
  if (!sh || sh.getLastRow() < 2) return { card: textCard_('案件清單', '#8B7355', '目前沒有案件'), copy: null };
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, STATUS_COL).getValues();
  var now = Date.now();
  var out = [];
  rows.forEach(function(r) {
    if (!r[1]) return;
    var st = String(r[STATUS_COL - 1] || '').trim();
    if (which === '未結案' && (st === '已結案' || st === '結案')) return;
    if ((which === '結案' || which === '已結案') && st !== '已結案' && st !== '結案') return;
    if (which !== '未結案' && which !== '全部' && which !== '結案' && which !== '已結案' && st !== which) return;
    var age = Math.floor((now - new Date(r[10]).getTime()) / 86400000);
    out.push(r[1] + '｜' + r[2] + '｜' + (st || '待處理') + '｜' + age + '天');
  });
  if (!out.length) return { card: textCard_('案件清單', '#8B7355', '沒有符合「' + which + '」的案件'), copy: null };
  var more = out.length > 15 ? '\n…等共 ' + out.length + ' 筆' : '';
  var body = out.slice(0, 15).join('\n') + more;
  return { card: textCard_('【' + which + '】共 ' + out.length + ' 筆', '#8B7355', body), copy: body };
}

function flexCard_(title, headColor, rows, foot) {
  var body = rows.map(function(r) {
    return { type: 'box', layout: 'baseline', spacing: 'sm', contents: [
      { type: 'text', text: String(r[0]), size: 'sm', color: '#8b7355', flex: 2 },
      { type: 'text', text: String(r[1]), size: 'sm', color: '#333333', flex: 5, wrap: true }
    ]};
  });
  var b = { type: 'bubble',
    header: { type: 'box', layout: 'vertical', backgroundColor: headColor,
      contents: [{ type: 'text', text: title, size: 'md', color: '#ffffff', weight: 'bold', wrap: true }] },
    body: { type: 'box', layout: 'vertical', spacing: 'md', contents: body } };
  if (foot) b.footer = { type: 'box', layout: 'vertical',
    contents: [{ type: 'text', text: foot, size: 'xs', color: '#999999', wrap: true }] };
  return { type: 'flex', altText: title, contents: b };
}

function textCard_(title, headColor, text) {
  return { type: 'flex', altText: title, contents: { type: 'bubble',
    header: { type: 'box', layout: 'vertical', backgroundColor: headColor,
      contents: [{ type: 'text', text: title, size: 'md', color: '#ffffff', weight: 'bold', wrap: true }] },
    body: { type: 'box', layout: 'vertical',
      contents: [{ type: 'text', text: String(text).slice(0, 1500), size: 'sm', color: '#333333', wrap: true }] } } };
}

function toMsg_(t) {
  return (t && typeof t === 'object') ? t : { type: 'text', text: String(t).slice(0, 1000) };
}

function replyLine_(token, text, copy) {
  var key = props_().getProperty('LINE_TOKEN');
  var msgs = [toMsg_(text)];
  if (copy) msgs.push({ type: 'text', text: String(copy).slice(0, 1000) });
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + key },
    contentType: 'application/json',
    payload: JSON.stringify({ replyToken: token, messages: msgs }),
    muteHttpExceptions: true
  });
}

function targetGroup_() {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_GROUP);
    if (sh && sh.getLastRow() >= 2) {
      if (String(sh.getRange(1, 1).getValue() || '').trim() === 'groupId') {
        var v = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
        for (var i = 0; i < v.length; i++) {
          if (String(v[i][0]).trim() && String(v[i][3]).toUpperCase() === 'Y') return String(v[i][0]).trim();
        }
      } else {
        var old = sh.getRange(2, 3).getValue();
        if (old) return String(old).trim();
      }
    }
  } catch (e) {}
  return props_().getProperty('GROUP_ID') || '';
}

function pushLine_(to, text, copy) {
  var token = props_().getProperty('LINE_TOKEN');
  var msgs = [toMsg_(text)];
  if (copy) msgs.push({ type: 'text', text: String(copy).slice(0, 1000) });
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + token },
    contentType: 'application/json',
    payload: JSON.stringify({ to: to, messages: msgs }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('LINE ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 120));
}

function lineUsage_() {
  try {
    var token = props_().getProperty('LINE_TOKEN');
    var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/quota/consumption', {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return Infinity;
    return Number(JSON.parse(res.getContentText()).totalUsage || Infinity);
  } catch (e) { return Infinity; }
}

function mailAdmin_(subject, body) {
  GmailApp.sendEmail(props_().getProperty('ADMIN_EMAIL'), subject, body);
}

function routeNotify_(tier, card) {
  var gid = targetGroup_();
  try {
    if (tier === 'VIP') {
      pushLine_(gid, card);
    } else {
      if (lineUsage_() < VIP_QUOTA_RESERVE) {
        pushLine_(gid, card);
      } else {
        throw new Error('quota-reserved');
      }
    }
  } catch (e) {
    try { mailAdmin_('【諮詢通知備援 ' + tier + '】' + e.message, '[Flex卡片] 請看試算表'); } catch (e2) {}
  }
}

function dailyUnclosed_() {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN);
    if (!sh || sh.getLastRow() < 2) return;
    if (sh.getRange(1, STATUS_COL + 1).getValue() !== '上次提醒')
      sh.getRange(1, STATUS_COL + 1).setValue('上次提醒');

    var n = sh.getLastRow() - 1;
    var rows = sh.getRange(2, 1, n, STATUS_COL + 1).getValues(); // A~M
    var now = Date.now();
    var unclosedIdx = [];
    var needPush = false;
    var lines = [];

    rows.forEach(function(r, i) {
      if (!r[1]) return;
      var st = String(r[STATUS_COL - 1] || '').trim() || '待處理';
      if (st === '已結案' || st === '結案') return;
      var limit = REMIND_DAYS[st];
      if (limit === undefined) return;
      var ct = new Date(r[10]).getTime();
      if (isNaN(ct)) return;
      var lr = r[STATUS_COL];
      var lt = 0;
      if (lr) {
        var d = lr instanceof Date ? lr : new Date(lr);
        lt = d.getTime();
        if (isNaN(lt)) lt = 0;
      }
      var base = Math.max(ct, lt);
      var totalAge = Math.floor((now - ct) / 86400000);
      unclosedIdx.push(i);
      lines.push(r[1] + '｜' + r[2] + '｜' + st + '｜' + totalAge + '天');
      if ((now - base) >= limit * 86400000) needPush = true;
    });

    if (!needPush || !lines.length) return;

    var more = lines.length > 15 ? '\n…等共 ' + lines.length + ' 筆' : '';
    var body = lines.slice(0, 15).join('\n') + more;
    pushLine_(targetGroup_(), flexCard_('未結案提醒', '#B04A3A', [['明細', body]], null));

    var nowDate = new Date(now);
    var col = sh.getRange(2, STATUS_COL + 1, n, 1).getValues();
    unclosedIdx.forEach(function(i) { col[i][0] = nowDate; });
    sh.getRange(2, STATUS_COL + 1, n, 1).setValues(col);
  } catch (err) {
    try { mailAdmin_('提醒任務失敗 ' + err, ''); } catch (e2) {}
  }
}

// ===== 徒弟訓練管理：SOP庫存＋訓練紀錄＋備課提醒 =====
function norm_(s) { return String(s == null ? '' : s).replace(/[\u200B\u200C\u200D\uFEFF]/g, '').replace(/[ \t\n\v\f\r   -   　﻿​‌‍]+/g, '').trim(); }
function sheetByName_(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (sh) return sh;
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    if (norm_(all[i].getName()) === norm_(name)) return all[i];
  }
  return null;
}
function pbBtn_(label, data, display, style, color) {
  var b = { type: 'button', action: { type: 'postback', label: label, data: data, displayText: display || label }, height: 'sm', margin: 'sm' };
  if (style) b.style = style;
  if (color) b.color = color;
  return b;
}

function saveUserBirthday_(uid, birthday) {
  try {
    var sh = userSheet_();
    var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    
    var bdayColIdx = -1, uidColIdx = -1, nameColIdx = -1;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim();
      if (h === '生日') bdayColIdx = i;
      if (h === 'userId') uidColIdx = i;
      if (h === '名稱') nameColIdx = i;
    }
    
    if (bdayColIdx === -1) {
      bdayColIdx = sh.getLastColumn(); 
      sh.getRange(1, bdayColIdx + 1).setValue('生日');
      headers.push('生日'); 
    }

    var idx = userRowIdx_(uid);
    var currentUserName = "新客人";
    
    // 💡 核心防呆：如果已經有填過生日，就直接中斷，不允許修改！
    if (idx > 0) {
      var existingBday = String(sh.getRange(idx, bdayColIdx + 1).getValue() || '').trim();
      if (existingBday !== '') {
        return false; // 回傳 false 代表「拒絕更新」
      }
    }

    var allData = sh.getDataRange().getValues();
    var sameBirthdayUsers = [];
    
    for (var r = 1; r < allData.length; r++) {
      var rowUid = (uidColIdx !== -1) ? String(allData[r][uidColIdx]).trim() : '';
      var rowName = (nameColIdx !== -1) ? String(allData[r][nameColIdx]).trim() : '未知用戶';
      
      var rawDate = (bdayColIdx !== -1) ? allData[r][bdayColIdx] : '';
      var rowBday = '';
      if (rawDate instanceof Date) {
        rowBday = rawDate.getFullYear() + '-' + ('0' + (rawDate.getMonth() + 1)).slice(-2) + '-' + ('0' + rawDate.getDate()).slice(-2);
      } else {
        rowBday = String(rawDate || '').trim();
      }
      
      if (rowBday === birthday && rowUid !== uid) {
        sameBirthdayUsers.push(rowName);
      }
    }
    
    if (idx > 0) {
      sh.getRange(idx, bdayColIdx + 1).setValue(birthday);
      if (nameColIdx !== -1) {
        currentUserName = String(sh.getRange(idx, nameColIdx + 1).getValue() || "未知");
      }
    } else {
      var newRow = new Array(headers.length).fill('');
      var colMap = {};
      for (var c = 0; c < headers.length; c++) colMap[String(headers[c]).trim()] = c;
      
      if (colMap['userId'] !== undefined) newRow[colMap['userId']] = uid;
      if (colMap['角色'] !== undefined) newRow[colMap['角色']] = 'customer';
      if (colMap['啟用'] !== undefined) newRow[colMap['啟用']] = 'Y';
      if (colMap['備註'] !== undefined) newRow[colMap['備註']] = '補填生日自動加入 ' + new Date();
      if (colMap['生日'] !== undefined) newRow[colMap['生日']] = birthday;
      
      sh.appendRow(newRow);
      currentUserName = "新加入之客人";
    }

    if (sameBirthdayUsers.length > 0) {
      var subject = '【系統提醒】有客人填寫了相同的生日！';
      var body = '剛剛填寫生日的客人：' + currentUserName + '\n' +
                 '填寫的日期為：' + birthday + '\n\n' +
                 '資料庫中目前已有以下客人也是同一天生日：\n' + 
                 sameBirthdayUsers.join('、') + '\n\n' +
                 '請留意是否為同一人或特殊巧合。';
      try { mailAdmin_(subject, body); } catch (mailErr) {}
    }
    
    return true; // 回傳 true 代表「成功更新」
    
  } catch (e) {
    return false;
  }
}

function menuCard_() {
  return { type: 'flex', altText: '指令選單', contents: { type: 'bubble',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#8B7355',
      contents: [{ type: 'text', text: '請選擇查詢類型', size: 'md', color: '#ffffff', weight: 'bold' }] },
    body: { type: 'box', layout: 'vertical', spacing: 'sm',
      contents: [
        { type: 'text', text: '第一步：請先選擇「諮詢」或「訓練」類別', size: 'sm', color: '#333333', wrap: true },
        { type: 'separator', margin: 'md' },
        // 💡 這裡可以列出所有未做成按鈕的隱藏指令
        { type: 'text', text: '【其他隱藏指令】\n可點擊下方按鈕，或手動打字輸入：\n• 翻譯 (例: 英翻中 apple)', size: 'xs', color: '#777777', wrap: true }
      ] },
    footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: [
      pbBtn_('諮詢查詢', 'action=menu&go=consult', '諮詢查詢', 'primary', '#8B7355'),
      pbBtn_('訓練管理', 'action=menu&go=train', '訓練管理', 'secondary'),
      // 💡 加入「管理員名單」按鈕，點擊後會自動在對話框送出文字
      { type: 'button', action: { type: 'message', label: '管理員名單', text: '管理員名單' }, style: 'secondary', height: 'sm', margin: 'sm' }
    ] } } };
}

function consultMenuCard_() {
  var filters = ['未結案', '待處理', '處理中', '無法聯繫', '已結案'];
  return { type: 'flex', altText: '諮詢查詢', contents: { type: 'bubble',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#4A5A6A',
      contents: [{ type: 'text', text: '諮詢查詢', size: 'md', color: '#ffffff', weight: 'bold' }] },
    body: { type: 'box', layout: 'vertical',
      contents: [{ type: 'text', text: '第二步：選狀態，會列出單號按鈕，點單號即可操作，不用打字', size: 'sm', color: '#333333', wrap: true }] },
    footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: filters.map(function(f) {
      return pbBtn_(f === '未結案' ? '未結案列表' : f, 'action=caselist&filter=' + encodeURIComponent(f), '諮詢' + f, f === '未結案' ? 'primary' : 'secondary', f === '未結案' ? '#4A5A6A' : undefined);
    }) } } };
}
function trainMenuCard_() {
  return { type: 'flex', altText: '訓練管理', contents: { type: 'bubble',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#8B7355',
      contents: [{ type: 'text', text: '訓練管理', size: 'md', color: '#ffffff', weight: 'bold' }] },
    body: { type: 'box', layout: 'vertical', spacing: 'sm', contents: [
      { type: 'text', text: '打字更新進度：敕令/領取/完成 姓名 單元', size: 'sm', color: '#333333', wrap: true },
      { type: 'text', text: '補庫存 單元 蓋章/列印', size: 'sm', color: '#333333', wrap: true },
      { type: 'text', text: '例：領取 Ren 第一冊1-32', size: 'xs', color: '#999999', wrap: true }
    ] },
    footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: [
      pbBtn_('訓練總覽', 'action=progress', '訓練總覽', 'primary', '#8B7355'),
      pbBtn_('庫存查詢', 'action=stock', '庫存查詢', 'secondary'),
      pbBtn_('庫存盤點', 'action=stockcheck', '庫存盤點', 'secondary'),
      pbBtn_('更新進度', 'action=tupd', '更新進度', 'secondary'),
      pbBtn_('補庫存', 'action=tadd', '補庫存', 'secondary')
    ] } } };
}
function caseListCard_(filter) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MAIN);
  if (!sh || sh.getLastRow() < 2) return { card: textCard_('案件清單', '#8B7355', '目前沒有案件'), copy: null };
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, STATUS_COL).getValues();
  var now = Date.now();
  var match = [];
  rows.forEach(function(r) {
    if (!r[1]) return;
    var st = String(r[STATUS_COL - 1] || '').trim() || '待處理';
    if (filter === '未結案' && (st === '已結案' || st === '結案')) return;
    if ((filter === '結案' || filter === '已結案') && st !== '已結案' && st !== '結案') return;
    if (filter !== '未結案' && filter !== '全部' && filter !== '結案' && filter !== '已結案' && st !== filter) return;
    var age = Math.floor((now - new Date(r[10]).getTime()) / 86400000);
    match.push({ id: String(r[1]).trim(), name: String(r[2]).trim(), st: st, age: age });
  });
  if (!match.length) return { card: textCard_('案件清單', '#8B7355', '沒有符合「' + filter + '」的案件'), copy: null };
  var btns = match.slice(0, 10).map(function(c) {
    return pbBtn_(c.id, 'action=case&caseId=' + encodeURIComponent(c.id), '查詢 ' + c.id, 'secondary');
  });
  var card = { type: 'flex', altText: '案件清單', contents: { type: 'bubble',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#4A5A6A',
      contents: [{ type: 'text', text: '諮詢・' + filter + '（' + match.length + '筆）', size: 'md', color: '#ffffff', weight: 'bold', wrap: true }] },
    body: { type: 'box', layout: 'vertical',
      contents: [{ type: 'text', text: '第三步：點單號進入操作' + (match.length > 10 ? '（僅顯示前10筆，全文在下一則）' : ''), size: 'sm', color: '#333333', wrap: true }] },
    footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: btns } } };
  var copy = match.map(function(c) { return c.id + '｜' + c.name + '｜' + c.st + '｜' + c.age + '天'; }).join('\n').slice(0, 1000);
  return { card: card, copy: copy };
}
function caseCard_(caseId) {
  var found = findCaseRow_(caseId);
  if (!found) return textCard_('找不到案件', '#B04A3A', String(caseId).toUpperCase());
  var r = found.sheet.getRange(found.row, 1, 1, STATUS_COL).getValues()[0];
  var st = String(r[11] || '').trim() || '待處理';
  var card = flexCard_('案件 ' + r[1], '#4A5A6A', [
    ['姓名', r[2]], ['項目', r[7]], ['狀態', st],
    ['聯絡', r[3]], ['內容', String(r[8] || '').slice(0, 100)]
  ], null);
  var ops = ['查詢', '結案', '處理中', '待處理', '無法聯繫'];
  card.contents.footer = { type: 'box', layout: 'vertical', spacing: 'sm', contents: ops.map(function(op) {
    return pbBtn_(op === '查詢' ? '重整摘要' : '改為' + (op === '結案' ? '已結案' : op),
      'action=casestatus&caseId=' + encodeURIComponent(String(r[1]).trim()) + '&to=' + encodeURIComponent(op),
      op + ' ' + String(r[1]).trim(), op === '結案' ? 'primary' : 'secondary', op === '結案' ? '#4A7C59' : undefined);
  }) };
  return card;
}
function doCaseStatus_(caseId, to) {
  if (to === '查詢') return caseCard_(caseId);
  var found = findCaseRow_(caseId);
  if (!found) return textCard_('找不到案件', '#B04A3A', String(caseId).toUpperCase());
  found.sheet.getRange(found.row, STATUS_COL).setValue(to === '結案' ? '已結案' : to);
  return caseCard_(found.sheet.getRange(found.row, 2).getValue());
}
function btnCard_(title, color, hint, btns) {
  return { type: 'flex', altText: title, contents: { type: 'bubble',
    header: { type: 'box', layout: 'vertical', backgroundColor: color,
      contents: [{ type: 'text', text: title, size: 'md', color: '#ffffff', weight: 'bold', wrap: true }] },
    body: { type: 'box', layout: 'vertical',
      contents: [{ type: 'text', text: hint, size: 'sm', color: '#333333', wrap: true }] },
    footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: btns } } };
}

function lineProfile_(uid) {
  try {
    var token = props_().getProperty('LINE_TOKEN');
    var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/profile/' + encodeURIComponent(uid), {
      headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return '';
    return String(JSON.parse(res.getContentText()).displayName || '');
  } catch (e) { return ''; }
}
function lineGroupName_(gid) {
  try {
    var p = props_();
    var cached = p.getProperty('gname_' + gid);
    if (cached) return cached;
    var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/group/' + encodeURIComponent(gid) + '/summary', {
      headers: { Authorization: 'Bearer ' + p.getProperty('LINE_TOKEN') }, muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return '';
    var nm = String(JSON.parse(res.getContentText()).groupName || '');
    if (nm) p.setProperty('gname_' + gid, nm);
    return nm;
  } catch (e) { return ''; }
}
function lineGroupMemberName_(gid, uid) {
  try {
    var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/group/' + encodeURIComponent(gid) + '/member/' + encodeURIComponent(uid), {
      headers: { Authorization: 'Bearer ' + props_().getProperty('LINE_TOKEN') }, muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return lineProfile_(uid);
    return String(JSON.parse(res.getContentText()).displayName || '');
  } catch (e) { return ''; }
}
function userRowIdx_(uid) {
  try {
    var sh = userSheet_();
    if (sh.getLastRow() < 2) return 0;
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][0]).trim() === String(uid).trim()) return i + 2;
    }
  } catch (e) {}
  return 0;
}
function ensureGroupRow_(gid) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_GROUP) || ss.insertSheet(SHEET_GROUP);
    if (sh.getLastRow() === 0) {
      sh.appendRow(['groupId', '名稱', '啟用', '預設推送', '備註', '角色']);
      sh.appendRow([gid, lineGroupName_(gid), 'Y', '', '自動加入 ' + new Date(), 'guest']);
      return;
    }
    if (String(sh.getRange(1, 1).getValue() || '').trim() !== 'groupId') return;
    var n = sh.getLastRow();
    if (n >= 2) {
      var v = sh.getRange(2, 1, n - 1, 2).getValues();
      for (var i = 0; i < v.length; i++) {
        if (String(v[i][0]).trim() === gid) {
          if (!String(v[i][1]).trim()) {
            var gn = lineGroupName_(gid);
            if (gn) sh.getRange(i + 2, 2).setValue(gn);
          }
          return;
        }
      }
    }
    sh.appendRow([gid, lineGroupName_(gid), 'Y', '', '自動加入 ' + new Date(), 'guest']);
  } catch (e) {}
}

function ensureUserRow_(uid, gid) {
  if (!uid) return;
  try {
    var cache = CacheService.getScriptCache();
    var sh = userSheet_();
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var colIdx = {};
    for (var c = 0; c < headers.length; c++) colIdx[String(headers[c]).trim()] = c;
    var uidCol = colIdx['userId'] !== undefined ? colIdx['userId'] : 0;
    var nameCol = colIdx['名稱'];

    if (data.length >= 2) {
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][uidCol]).trim() === uid) {
          if (nameCol !== undefined && !String(data[i][nameCol]).trim() && !cache.get('nm_' + uid)) {
            cache.put('nm_' + uid, '1', 21600);
            var nm0 = gid ? lineGroupMemberName_(gid, uid) : lineProfile_(uid);
            if (nm0) sh.getRange(i + 1, nameCol + 1).setValue(nm0);
          }
          return;
        }
      }
    }
    var newRow = new Array(headers.length).fill('');
    if (colIdx['userId'] !== undefined) newRow[colIdx['userId']] = uid;
    if (colIdx['名稱'] !== undefined) newRow[colIdx['名稱']] = gid ? lineGroupMemberName_(gid, uid) : lineProfile_(uid);
    if (colIdx['角色'] !== undefined) newRow[colIdx['角色']] = 'guest';
    if (colIdx['啟用'] !== undefined) newRow[colIdx['啟用']] = 'Y';
    if (colIdx['備註'] !== undefined) newRow[colIdx['備註']] = '自動加入 ' + new Date();
    sh.appendRow(newRow);
  } catch (e) {}
}
function setupTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    var fn = t.getHandlerFunction();
    if (fn === 'onStatusEdit_' || fn === 'dailyUnclosed_' || fn === 'dailyTrainAlert_') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyUnclosed_').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('dailyTrainAlert_').timeBased().everyDays(1).atHour(8).create();
}

function diagPush() {
  var gid = 'Ue63699996ba4fdaef1bd1f2964b37fd0';
  Logger.log('gid=' + gid);
  try {
    // pushLine_(gid, textCard_('敷籙功曹降臨！', '#8B7355', '本神官已感應到貴客的疑惑？！\n\n請依照古禮選擇『香』『花』『果』三禮其中之一，奉獻於吾，請教貴客要選哪一個'));
    pushLine_(gid, textCard_('敷籙功曹降臨！', '#8B7355', '本神官已感應到貴客的到來？！\n\n請依照古禮選擇『香』『花』『果』三禮其中之一，奉獻於吾，請教貴客要選哪一個'));
    Logger.log('push ok');
  } catch (e) { Logger.log('push FAIL: ' + e); }
}

function diagPush_game() {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LINE_User');
    if (!sh || sh.getLastRow() < 2) {
      Logger.log('沒有使用者資料可推播');
      return;
    }

    // 1. 動態抓取欄位位置
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var colIdx = {};
    for (var c = 0; c < headers.length; c++) colIdx[String(headers[c]).trim()] = c;

    var uidCol = colIdx['userId'];
    var activeCol = colIdx['啟用'];
    if (uidCol === undefined) {
      Logger.log('找不到 userId 欄位，推播中止');
      return;
    }

    // 2. 組裝帶有按鈕的卡片 (使用 getGameCard_，按鈕已含 game 參數)
    var card = getGameCard_();

    // 3. 逐一推播給所有有效用戶
    var successCount = 0;
    for (var i = 1; i < data.length; i++) {
      var uid = String(data[i][uidCol]).trim();
      
      // 確保 ID 是有效的，且狀態為啟用 (Y)
      var isActive = (activeCol !== undefined) ? String(data[i][activeCol]).toUpperCase() === 'Y' : true;
      if (!uid || uid.indexOf('U') !== 0 || !isActive) continue; 

      try {
        pushLine_(uid, card);
        successCount++;
        Utilities.sleep(100); // 暫停 100 毫秒，避免觸發 LINE API 發送頻率限制
      } catch (e) {
        Logger.log('Push FAIL 傳送失敗 (' + uid + '): ' + e);
      }
    }
    
    Logger.log('✅ 推播完成，共成功發送給 ' + successCount + ' 位顧客');
  } catch (err) {
    Logger.log('diagPush 發生錯誤：' + err);
  }
}