// ===== 客戶推播管理模組 =====
function parsePb_(s) {
  var o = {};
  String(s || '').split('&').forEach(function(part) {
    var kv = part.split('=');
    if (kv.length === 2) {
      try { o[decodeURIComponent(kv[0])] = decodeURIComponent(kv[1]); } catch (e) { o[kv[0]] = kv[1]; }
    }
  });
  return o;
}

function isAdminUid_(uid) {
  var ur = userRole_(uid);
  return !!(ur && ur.on && ur.role === 'admin');
}

function clientUsers_() {
  var sh = userSheet_();
  if (!sh || sh.getLastRow() < 2) return [];
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var colMap = {};
  for (var c = 0; c < headers.length; c++) colMap[String(headers[c]).trim()] = c;
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, headers.length).getValues();
  var out = [];
  rows.forEach(function(r) {
    var uid = (colMap['userId'] !== undefined) ? String(r[colMap['userId']]).trim() : '';
    var name = (colMap['名稱'] !== undefined) ? String(r[colMap['名稱']]).trim() : '';
    var role = (colMap['角色'] !== undefined) ? String(r[colMap['角色']]).trim() : '';
    var on = (colMap['啟用'] !== undefined) ? String(r[colMap['啟用']]).toUpperCase() === 'Y' : true;
    if (!uid || uid.indexOf('U') !== 0) return;
    if (!on) return;
    out.push({ uid: uid, name: name || uid, role: role });
  });
  return out;
}

function filterClients_(kw) {
  var all = clientUsers_();
  if (kw === 'all' || kw === 'ALL') return all;
  var k = String(kw).toLowerCase();
  return all.filter(function(c) {
    return String(c.name).toLowerCase().indexOf(k) !== -1 || String(c.uid).toLowerCase().indexOf(k) !== -1;
  });
}

function b64urlEncodeStr_(str) {
  return Utilities.base64EncodeWebSafe(str, Utilities.Charset.UTF_8).replace(/=+$/, '');
}
function b64urlEncodeBytes_(bytes) {
  return Utilities.base64EncodeWebSafe(bytes).replace(/=+$/, '');
}
function b64urlDecodeStr_(str) {
  var s = String(str || '').replace(/-/g, '+').replace(/_/g, '/');
  var pad = s.length % 4;
  if (pad) s += '='.repeat(4 - pad);
  var bytes = Utilities.base64Decode(s);
  return Utilities.newBlob(bytes).getDataAsString(Utilities.Charset.UTF_8);
}

function makePushToken_(uid, name) {
  var secret = props_().getProperty('PUSH_LINK_SECRET');
  if (!secret) throw new Error('缺少 PUSH_LINK_SECRET');
  var payload = { uid: uid, name: name, exp: Date.now() + 600000, n: Math.random().toString(36).slice(2) };
  var body = b64urlEncodeStr_(JSON.stringify(payload));
  var sig = b64urlEncodeBytes_(Utilities.computeHmacSha256Signature(body, secret));
  return body + '.' + sig;
}

function clientMenuCard_() {
  return { type: 'flex', altText: '客戶管理', contents: { type: 'bubble',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#8B7355', contents: [{ type: 'text', text: '客戶管理', size: 'md', color: '#ffffff', weight: 'bold' }] },
    body: { type: 'box', layout: 'vertical', contents: [{ type: 'text', text: '請選擇查詢方式', size: 'sm', color: '#333333', wrap: true }] },
    footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: [
      pbBtn_('查詢全部客戶', 'action=client_list&kw=all&page=1', '查詢全部客戶', 'primary', '#8B7355'),
      { type: 'button', action: { type: 'uri', label: '關鍵字查詢', uri: 'line://msg/text/查客戶%20' }, style: 'secondary', height: 'sm', margin: 'sm' }
    ] } } };
}

function clientListCarousel_(kw, page) {
  var list = filterClients_(kw);
  page = Math.max(1, Number(page) || 1);
  var pageSize = 60;
  var total = list.length;
  var start = (page - 1) * pageSize;
  var items = list.slice(start, start + pageSize);
  if (!items.length) return textCard_('客戶列表', '#8B7355', '沒有符合的客戶');
  var bubbles = [];
  var chunk = 5;
  for (var i = 0; i < items.length; i += chunk) {
    if (bubbles.length >= 12) break;
    var part = items.slice(i, i + chunk);
    var btns = part.map(function(c) {
      return { type: 'button', action: { type: 'uri', label: String(c.name || '').slice(0, 18), uri: 'https://spiritref.github.io/client-push/?t=' + makePushToken_(c.uid, c.name) }, style: 'secondary', height: 'sm', margin: 'sm' };
    });
    if (i + chunk >= items.length && total > page * pageSize) {
      btns.push({ type: 'button', action: { type: 'postback', label: '下一頁', data: 'action=client_list&kw=' + encodeURIComponent(kw) + '&page=' + (page + 1) }, style: 'primary', height: 'sm', margin: 'sm' });
    }
    bubbles.push({ type: 'bubble',
      header: { type: 'box', layout: 'vertical', backgroundColor: '#8B7355', contents: [{ type: 'text', text: '客戶列表', size: 'md', color: '#ffffff', weight: 'bold' }] },
      body: { type: 'box', layout: 'vertical', contents: [{ type: 'text', text: '第 ' + page + ' 頁（共 ' + total + ' 人）', size: 'sm', color: '#333333', wrap: true }] },
      footer: { type: 'box', layout: 'vertical', spacing: 'sm', contents: btns }
    });
  }
  return { type: 'flex', altText: '客戶列表', contents: { type: 'carousel', contents: bubbles } };
}

function handleClientPostback_(pbData, ev) {
  if (!pbData) return false;
  var p = parsePb_(pbData);
  if (p.action !== 'client_menu' && p.action !== 'client_list') return false;
  var uid = (ev.source && ev.source.userId) || '';
  if (!isAdminUid_(uid)) {
    if (ev.replyToken) replyLine_(ev.replyToken, textCard_('權限不足', '#B04A3A', '此功能僅限管理員使用'));
    return true;
  }
  if (p.action === 'client_menu') {
    if (ev.replyToken) replyLine_(ev.replyToken, clientMenuCard_());
    return true;
  }
  if (p.action === 'client_list') {
    if (ev.replyToken) replyLine_(ev.replyToken, clientListCarousel_(p.kw || 'all', p.page || 1));
    return true;
  }
  return false;
}