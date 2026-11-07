// ===== GAS-WEB：諮詢表單專用（含LINE通知）。屬性：RECAPTCHA_SECRET / LINE_TOKEN / ADMIN_EMAIL / GROUP_ID(備用) =====
var SHEET_MAIN = '諮詢表單內容';
var SHEET_PUSH_LOG = '推播紀錄';
var SHEET_GROUP = 'LINE_Group';
var OTP_TTL = 600;
var RECAPTCHA_MIN = 0.5;
var VIP_QUOTA_RESERVE = 100;

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function props_() { return PropertiesService.getScriptProperties(); }

function doPost(e) {
  try {
    var data = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (data.action === 'sendOtp') return json_(sendOtp_(data));
    if (data.action === 'submit') return json_(submit_(data));
  if (data.action === 'resolve_push_token') return json_(resolvePushToken_(data));
  if (data.action === 'push_message') return json_(pushMessage_(data));
  if (data.action === 'push_message_manual') return json_(pushMessageManual_(data));
    return json_({ success: false, message: '未知動作' });
  } catch (err) {
    return json_({ success: false, message: '系統錯誤: ' + err });
  }
}

function sendOtp_(d) {
  var email = String(d.email || '').trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return { success: false, message: '信箱格式錯誤' };
  if (!verifyRecaptcha_(d.recaptcha)) return { success: false, message: '機器人驗證失敗' };
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get('otp_n_' + email) || 0);
  if (n >= 3) return { success: false, message: '發送太頻繁，請稍後再試' };
  cache.put('otp_n_' + email, String(n + 1), 600);
  var code = String(Math.floor(100000 + Math.random() * 900000));
  cache.put('otp_' + email, code, OTP_TTL);
  GmailApp.sendEmail(email, '【鬼磕頭】信箱驗證碼', '您的驗證碼是：' + code + '\n10 分鐘內有效，請勿轉發。');
  return { success: true };
}

function submit_(d) {
  var name = String(d.name || '').trim();
  var email = String(d.email || '').trim();
  var phone = String(d.phone || '').trim();
  var lineId = String(d.line_id || '').trim();
  var cat = String(d.category || '').trim();
  var msg = String(d.message || '').trim();
  if (!name) return { success: false, message: '缺少姓名' };
  if (!email && !phone && !lineId) return { success: false, message: '聯絡方式至少填一項' };
  if (!cat) return { success: false, message: '缺少諮詢項目' };
  if (!msg) return { success: false, message: '缺少問題說明' };
  if (!verifyRecaptcha_(d.recaptcha)) return { success: false, message: '機器人驗證失敗' };

  var tier = 'General';
  if (email) {
    var cache = CacheService.getScriptCache();
    var tries = Number(cache.get('otp_t_' + email) || 0);
    if (tries >= 5) return { success: false, message: '驗證錯誤太多次，請重發驗證碼' };
    if (String(d.otp || '').trim() !== cache.get('otp_' + email) || !cache.get('otp_' + email)) {
      cache.put('otp_t_' + email, String(tries + 1), 600);
      return { success: false, message: '驗證碼錯誤' };
    }
    cache.remove('otp_' + email);
    tier = 'VIP';
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_MAIN);
  if (!sh) return { success: false, message: '找不到資料表' };

  var item = Math.max(1, sh.getLastRow());
  var day = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyyMMdd');
  var caseId = '';
  var existed = {};
  try {
    if (sh.getLastRow() >= 2) {
      var ids = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) existed[String(ids[i][0])] = true;
    }
  } catch (e) {}
  var CHARS = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  for (var t = 0; t < 5; t++) {
    var rand = '';
    for (var k = 0; k < 6; k++) rand += CHARS.charAt(Math.floor(Math.random() * CHARS.length));
    caseId = day + '-' + rand;
    if (!existed[caseId]) break;
  }
  var contact = [email && ('Email:' + email), phone && ('電話:' + phone), lineId && ('LINE:' + lineId)].filter(Boolean).join(' / ');
  var level = (tier === 'VIP') ? 'VIP' : '';
  var appTime = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  sh.appendRow([item, caseId, name, contact, lineId, email, phone, cat, msg, level, appTime, '待處理']);

  if (email && d.copy !== false) {
    try {
      GmailApp.sendEmail(email,
        '【鬼磕頭】已收到您的諮詢（' + caseId + '）',
        name + ' 您好：\n\n我們已收到您的諮詢，會盡快與您聯繫。\n\n案件編號：' + caseId + '\n諮詢項目：' + cat + '\n\n您填寫的內容：\n' + msg.slice(0, 1500));
    } catch (e) {}
  }
  // LINE 通知（失敗不影響已存的單）
  try {
    var card = flexCard_('新諮詢・' + tier + ' ' + caseId, tier === 'VIP' ? '#8B7355' : '#6d6d6d', [
      ['姓名', name], ['項目', cat], ['聯絡', contact], ['內容', String(msg).slice(0, 300)]
    ], null);
    routeNotify_(tier, card);
  } catch (e) {}
  return { success: true, caseId: caseId };
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

function targetGroup_() {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_GROUP);
    if (sh && sh.getLastRow() >= 2) {
      var v = sh.getRange(2, 1).getValue();
      var result = String(v).trim();
      if (v) return String(v).trim();
    }
  } catch (e) {}
  return props_().getProperty('GROUP_ID') || '';
}

function pushLine_(to, text) {
  var token = props_().getProperty('LINE_TOKEN');
  var msg = (text && typeof text === 'object') ? text : { type: 'text', text: String(text || '').slice(0, 5000) };
  var res = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + token },
    contentType: 'application/json',
    payload: JSON.stringify({ to: to, messages: [msg] }),
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
    try {
      // 將錯誤訊息與當時要發送的卡片內容（或你的表單資料變數）合併
      var errorDetails = "錯誤原因: " + e.message + "\n\n--- 完整填寫內容/卡片資料 ---\n";
      for (var key in card) {
        if (card.hasOwnProperty(key)) {
          errorDetails += key + ": " + card[key] + "\n";
        }
      }
      mailAdmin_('【諮詢通知備援 ' + tier + '】發生錯誤', errorDetails);
    } catch (e2) {}
  }
}

function verifyRecaptcha_(token) {
  try {
    var secret = props_().getProperty('RECAPTCHA_SECRET');
    if (!secret || !token) return false;
    var res = UrlFetchApp.fetch('https://www.google.com/recaptcha/api/siteverify', {
      method: 'post',
      payload: { secret: secret, response: token },
      muteHttpExceptions: true
    });
    var j = JSON.parse(res.getContentText());
    return j.success === true && (j.score === undefined || j.score >= RECAPTCHA_MIN);
  } catch (e) { return false; }
}

// ===== 客戶推播 API =====
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

function verifyPushToken_(token) {
  if (!token || typeof token !== 'string') return null;
  var parts = token.split('.');
  if (parts.length !== 2) return null;
  var body = parts[0], sig = parts[1];
  var secret = props_().getProperty('PUSH_LINK_SECRET');
  if (!secret) return null;
  var expected = b64urlEncodeBytes_(Utilities.computeHmacSha256Signature(body, secret));
  if (!constantTimeEqual_(sig, expected)) return null;
  var payload;
  try { payload = JSON.parse(b64urlDecodeStr_(body)); } catch (e) { return null; }
  if (!payload || !payload.uid || !payload.exp) return null;
  if (Date.now() > Number(payload.exp)) return null;
  return payload;
}

function constantTimeEqual_(a, b) {
  a = String(a || ''); b = String(b || '');
  if (a.length !== b.length) return false;
  var r = 0;
  for (var i = 0; i < a.length; i++) r |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return r === 0;
}

function resolvePushToken_(d) {
  var payload = verifyPushToken_(String((d && d.token) || '').trim());
  if (!payload) return { success: false, message: '連結已失效，請重新從 LINE 取得' };
  return { success: true, name: payload.name };
}

function rateLimit_(key, limit, windowSec) {
  var cache = CacheService.getScriptCache();
  var k = 'rl_' + key;
  var n = Number(cache.get(k) || 0);
  if (n >= limit) return false;
  cache.put(k, String(n + 1), windowSec);
  return true;
}

function verifyGithubToken_(token) {
  if (!token) return false;
  try {
    var res = UrlFetchApp.fetch('https://api.github.com/user', {
      method: 'get',
      headers: { Authorization: 'token ' + token, Accept: 'application/vnd.github.v3+json' },
      muteHttpExceptions: true
    });
    return res.getResponseCode() === 200;
  } catch (e) { return false; }
}

function auditPushSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_PUSH_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PUSH_LOG);
    sh.appendRow(['時間', '客戶名稱(遮罩)', '訊息長度', '訊息雜湊', '結果', '錯誤']);
  }
  return sh;
}

function auditPush_(payload, message, ok, err) {
  try {
    var sh = auditPushSheet_();
    var name = (payload && payload.name) ? String(payload.name) : '';
    var masked = name.length > 1 ? name.charAt(0) + '**' : name;
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(message || ''), Utilities.Charset.UTF_8);
    var hex = digest.map(function(b) { return (b < 0 ? b + 256 : b).toString(16).padStart(2, '0'); }).join('').slice(0, 12);
    sh.appendRow([new Date(), masked, String(message || '').length, hex, ok ? '成功' : '失敗', err || '']);
  } catch (e) {}
}

function pushMessage_(d) {
  var token = String((d && d.token) || '').trim();
  var password = String((d && d.password) || '');
  var message = String((d && d.message) || '').trim();
  var payload = verifyPushToken_(token);
  if (!payload) {
    auditPush_(null, message, false, 'token_invalid');
    return { success: false, message: '連結已失效，請重新從 LINE 取得' };
  }
  if (!rateLimit_('push_token_' + token.slice(0, 16), 5, 900)) {
    auditPush_(payload, message, false, 'rate_limit');
    return { success: false, message: '嘗試次數過多，請稍後再試' };
  }
  if (!message) return { success: false, message: '請輸入訊息內容' };
  if (message.length > 1000) return { success: false, message: '訊息過長（上限 1000 字）' };
  var githubToken = String((d && d.githubToken) || '').trim();
  if (!githubToken || !verifyGithubToken_(githubToken)) {
    auditPush_(payload, message, false, 'token_invalid');
    return { success: false, message: '驗證失敗，請重新登入後台' };
  }
  try {
    pushLine_(payload.uid, message);
    auditPush_(payload, message, true, '');
    return { success: true };
  } catch (e) {
    auditPush_(payload, message, false, String(e).slice(0, 120));
    return { success: false, message: '推播失敗：' + e };
  }
}

function pushMessageManual_(d) {
  var uid = String((d && d.uid) || '').trim();
  var name = String((d && d.name) || '').trim();
  var password = String((d && d.password) || '');
  var message = String((d && d.message) || '').trim();
  if (!/^U[0-9a-fA-F]{32}$/.test(uid)) return { success: false, message: 'LINE 使用者 ID 格式錯誤' };
  if (!message) return { success: false, message: '請輸入訊息內容' };
  if (message.length > 1000) return { success: false, message: '訊息過長（上限 1000 字）' };
  if (!rateLimit_('push_manual_' + uid, 5, 900)) return { success: false, message: '嘗試次數過多，請稍後再試' };
  var githubToken = String((d && d.githubToken) || '').trim();
  if (!githubToken || !verifyGithubToken_(githubToken)) return { success: false, message: '驗證失敗，請重新登入後台' };
  try {
    pushLine_(uid, message);
    auditPush_({ name: name || uid }, message, true, '');
    return { success: true };
  } catch (e) {
    auditPush_({ name: name || uid }, message, false, String(e).slice(0, 120));
    return { success: false, message: '推播失敗：' + e };
  }
}