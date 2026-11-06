// ===== 耗材庫存與盤點模組 =====
var SHEET_CONSUMABLE = '耗材總表';
var SHEET_CONSUMABLE_LOG = '耗材異動紀錄';

function ensureConsumableSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s1 = ss.getSheetByName(SHEET_CONSUMABLE) || ss.insertSheet(SHEET_CONSUMABLE);
  if (s1.getLastRow() === 0) {
    s1.appendRow(['耗材名稱', '當前庫存', '安全庫存', '最後更新時間']);
  } else if (String(s1.getRange(1, 1).getValue() || '').trim() !== '耗材名稱') {
    s1.getRange(1, 1, 1, 4).setValues([['耗材名稱', '當前庫存', '安全庫存', '最後更新時間']]);
  }
  var s2 = ss.getSheetByName(SHEET_CONSUMABLE_LOG) || ss.insertSheet(SHEET_CONSUMABLE_LOG);
  if (s2.getLastRow() === 0) {
    s2.appendRow(['時間', '操作人', '耗材名稱', '動作類型', '變動數量', '結算後庫存', '備註']);
  }
}

function findConsumable_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONSUMABLE);
  if (!sh || sh.getLastRow() < 2) return null;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  for (var i = 0; i < v.length; i++) {
    if (norm_(v[i][0]) === norm_(name)) {
      return { sheet: sh, row: i + 2, name: String(v[i][0]).trim(), stock: Number(v[i][1]) || 0, safe: Number(v[i][2]) || 0 };
    }
  }
  return null;
}

function logConsumable_(operator, name, action, qty, after, note) {
  ensureConsumableSheets_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONSUMABLE_LOG);
  sh.appendRow([new Date(), operator, name, action, qty, after, note || '']);
}

function consumableAdd_(name, qty, operator) {
  ensureConsumableSheets_();
  var found = findConsumable_(name);
  if (found) return textCard_('耗材新增', '#B04A3A', '項目已存在，請使用進貨');
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONSUMABLE);
  sh.appendRow([name, qty, 0, new Date()]);
  logConsumable_(operator, name, '新增', qty, qty, '');
  return textCard_('耗材新增', '#4A7C59', name + ' 已新增，初始數量 ' + qty);
}

function consumableIn_(name, qty, operator) {
  ensureConsumableSheets_();
  var found = findConsumable_(name);
  if (!found) return textCard_('耗材進貨', '#B04A3A', '找不到項目：' + name);
  var newStock = found.stock + qty;
  found.sheet.getRange(found.row, 2).setValue(newStock);
  found.sheet.getRange(found.row, 4).setValue(new Date());
  logConsumable_(operator, name, '進貨', qty, newStock, '');
  return textCard_('耗材進貨', '#4A7C59', name + ' +' + qty + '（現有 ' + newStock + '）');
}

function consumableOut_(name, qty, operator) {
  ensureConsumableSheets_();
  var found = findConsumable_(name);
  if (!found) return textCard_('耗材消耗', '#B04A3A', '找不到項目：' + name);
  var newStock = found.stock - qty;
  if (newStock < 0) newStock = 0;
  found.sheet.getRange(found.row, 2).setValue(newStock);
  found.sheet.getRange(found.row, 4).setValue(new Date());
  logConsumable_(operator, name, '消耗', -qty, newStock, '');
  return textCard_('耗材消耗', '#4A7C59', name + ' -' + qty + '（現有 ' + newStock + '）');
}

function consumableCheck_(name, qty, operator) {
  ensureConsumableSheets_();
  var found = findConsumable_(name);
  if (!found) return textCard_('耗材盤點', '#B04A3A', '找不到項目：' + name);
  var diff = qty - found.stock;
  found.sheet.getRange(found.row, 2).setValue(qty);
  found.sheet.getRange(found.row, 4).setValue(new Date());
  logConsumable_(operator, name, '盤點', diff, qty, '盤點校正');
  return textCard_('耗材盤點', '#4A7C59', name + ' 盤點為 ' + qty + '（差異 ' + diff + '）');
}

function consumableQuery_() {
  ensureConsumableSheets_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONSUMABLE);
  if (!sh || sh.getLastRow() < 2) return textCard_('耗材查詢', '#8B7355', '耗材總表無資料');
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  var rows = [];
  v.forEach(function(r, idx) {
    var name = String(r[0]).trim(); if (!name) return;
    var stock = Number(r[1]) || 0;
    var safe = Number(r[2]) || 0;
    var warn = stock < safe ? ' ⚠️' : '';
    var upd = r[3];
    var updStr = '';
    try { updStr = (upd instanceof Date) ? Utilities.formatDate(upd, Session.getScriptTimeZone(), 'MM/dd HH:mm') : String(upd || '').slice(0, 16); } catch (e) { updStr = String(upd || '').slice(0, 16); }
    rows.push({ type: 'box', layout: 'vertical', margin: idx > 0 ? 'sm' : 'none', contents: [
      { type: 'box', layout: 'horizontal', contents: [
        { type: 'text', text: name + warn, size: 'sm', color: '#333333', flex: 3, wrap: true },
        { type: 'text', text: '庫存 ' + stock, size: 'sm', color: '#333333', flex: 2, wrap: true },
        { type: 'text', text: '安全 ' + safe, size: 'sm', color: '#333333', flex: 2, wrap: true }
      ]},
      { type: 'text', text: '更新 ' + updStr, size: 'xs', color: '#999999', wrap: true, margin: 'xs' }
    ]});
  });
  if (!rows.length) return textCard_('耗材查詢', '#8B7355', '耗材總表無資料');
  var card = { type: 'flex', altText: '耗材查詢', contents: { type: 'bubble',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#8B7355', contents: [{ type: 'text', text: '耗材庫存查詢', size: 'md', color: '#ffffff', weight: 'bold', wrap: true }] },
    body: { type: 'box', layout: 'vertical', spacing: 'xs', contents: rows }
  }};
  return card;
}

function consumableMenuCard_() {
  return { type: 'flex', altText: '耗材管理', contents: { type: 'bubble',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#8B7355', contents: [{ type: 'text', text: '耗材管理', size: 'md', color: '#ffffff', weight: 'bold', wrap: true }] },
    body: { type: 'box', layout: 'vertical', spacing: 'sm', contents: [
      { type: 'text', text: '選擇操作類型', size: 'sm', color: '#333333', wrap: true }
    ] },
    footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: [
      pbBtn_('庫存查詢', 'action=consumable_query', '耗材查詢', 'primary', '#8B7355'),
      pbBtn_('進貨', 'action=consumable_select&type=in', '耗材進貨', 'secondary'),
      pbBtn_('消耗', 'action=consumable_select&type=out', '耗材消耗', 'secondary'),
      pbBtn_('盤點', 'action=consumable_select&type=check', '耗材盤點', 'secondary'),
      { type: 'button', action: { type: 'postback', label: '新增耗材', data: 'action=consumable_prefill&hint=' + encodeURIComponent('請輸入：耗材新增 物品名稱 初始數量'), inputOption: 'openKeyboard', fillInText: '耗材新增 ' }, style: 'secondary', height: 'sm', margin: 'sm' }
    ] } } };
}

function consumableSelectCard_(type) {
  ensureConsumableSheets_();
  var prefix = type === 'in' ? '進貨' : type === 'out' ? '消耗' : '盤點';
  var cmdPrefix = '耗材' + prefix;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CONSUMABLE);
  var items = [];
  if (sh && sh.getLastRow() >= 2) {
    var v = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < v.length; i++) {
      var nm = String(v[i][0]).trim();
      if (nm && items.indexOf(nm) === -1) items.push(nm);
    }
  }
  if (!items.length) return textCard_('耗材', '#8B7355', '耗材總表無資料，請先新增耗材');
  var btns = items.slice(0, 20).map(function(name) {
    var hint = '請輸入數量（整數）';
    return { type: 'button', action: { type: 'postback', label: name, data: 'action=consumable_prefill&hint=' + encodeURIComponent(hint), inputOption: 'openKeyboard', fillInText: cmdPrefix + ' ' + name + ' ' }, style: 'secondary', height: 'sm', margin: 'sm' };
  });
  return btnCard_('耗材' + prefix + '・選物品', '#8B7355', '點選物品後輸入數量', btns);
}

function handleConsumablePostback_(data) {
  var q = {};
  String(data || '').split('&').forEach(function(kv) {
    var p = kv.split('=');
    if (p[0]) q[decodeURIComponent(p[0])] = decodeURIComponent(p[1] || '');
  });
  if (q.action === 'consumable_menu') return consumableMenuCard_();
  if (q.action === 'consumable_query') return consumableQuery_();
  if (q.action === 'consumable_select' && q.type) return consumableSelectCard_(q.type);
  if (q.action === 'consumable_prefill') return textCard_('耗材輸入', '#8B7355', q.hint || '請輸入指令');
  return null;
}

function handleConsumableCommand_(text, uid) {
  text = String(text || '').trim();
  var ur = uid ? userRole_(uid) : null;
  var isAdmin = !!(ur && ur.on && ur.role === 'admin');
  var operator = (ur && ur.name) ? ur.name : (uid || '未知');

  var m1 = text.match(/^耗材新增\s+(\S+)\s+(\d+)$/);
  if (m1) {
    if (!isAdmin) return textCard_('權限不足', '#B04A3A', '此功能僅限管理員使用');
    return consumableAdd_(m1[1].trim(), Number(m1[2]), operator);
  }
  var m2 = text.match(/^耗材進貨\s+(\S+)\s+(\d+)$/);
  if (m2) {
    if (!isAdmin) return textCard_('權限不足', '#B04A3A', '此功能僅限管理員使用');
    return consumableIn_(m2[1].trim(), Number(m2[2]), operator);
  }
  var m3 = text.match(/^耗材消耗\s+(\S+)\s+(\d+)$/);
  if (m3) {
    if (!isAdmin) return textCard_('權限不足', '#B04A3A', '此功能僅限管理員使用');
    return consumableOut_(m3[1].trim(), Number(m3[2]), operator);
  }
  var m4 = text.match(/^耗材盤點\s+(\S+)\s+(\d+)$/);
  if (m4) {
    if (!isAdmin) return textCard_('權限不足', '#B04A3A', '此功能僅限管理員使用');
    return consumableCheck_(m4[1].trim(), Number(m4[2]), operator);
  }
  if (/^耗材查詢$/.test(text)) {
    if (!isAdmin) return textCard_('權限不足', '#B04A3A', '此功能僅限管理員使用');
    return consumableQuery_();
  }
  if (/^耗材管理$/.test(text)) {
    if (!isAdmin) return textCard_('權限不足', '#B04A3A', '此功能僅限管理員使用');
    return consumableMenuCard_();
  }
  return null;
}