function sopMap_() {
  var sh = sheetByName_(SHEET_SOP);
  if (!sh) return { sh: null, list: [], map: {} };
  var n = sh.getLastRow() - 1;
  if (n < 1) return { sh: sh, list: [], map: {} };
  var v = sh.getRange(2, 1, n, 3).getValues();
  var list = [], map = {};
  v.forEach(function(r, i) {
    if (!norm_(r[0])) return;
    var o = { unit: String(r[0]).trim(), key: norm_(r[0]), b: Number(r[1]) || 0, c: Number(r[2]) || 0, row: i + 2, idx: list.length };
    list.push(o); map[o.key] = o;
  });
  return { sh: sh, list: list, map: map };
}

function trainRows_() {
  var sh = sheetByName_(SHEET_TRAIN);
  if (!sh) return { sh: null, rows: [] };
  var n = sh.getLastRow() - 1;
  if (n < 1) return { sh: sh, rows: [] };
  return { sh: sh, rows: sh.getRange(2, 1, n, 5).getValues() };
}

function ensureTrainSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s1 = sheetByName_(SHEET_SOP) || ss.insertSheet(SHEET_SOP);
  if (norm_(s1.getRange(1, 1).getValue()) !== '訓練單元') s1.getRange(1, 1, 1, 3).setValues([['訓練單元', '已列印份數', '已蓋章份數']]);
  var s2 = sheetByName_(SHEET_TRAIN) || ss.insertSheet(SHEET_TRAIN);
  if (norm_(s2.getRange(1, 1).getValue()) !== 'ID') s2.getRange(1, 1, 1, 5).setValues([['ID', '姓名', '訓練單元', '狀態', '更新時間']]);
}

function deductStock_(unit) {
  var m = sopMap_();
  var o = m.map[norm_(unit)];
  if (!o) return false;
  if (o.c > 0) m.sh.getRange(o.row, 3).setValue(o.c - 1);
  else if (o.b > 0) m.sh.getRange(o.row, 2).setValue(o.b - 1);
  else return false;
  return true;
}

function trainAppend_(name, unit, status) {
  ensureTrainSheets_();
  var ss = sheetByName_(SHEET_TRAIN);
  // 產生 YYYYMMDD-序號 的 ID
  var ymd = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  var maxSeq = 0;
  // 防重複：相同 姓名+單元+狀態 已存在則不新增，同時計算當日最大序號
  try {
    var rows = ss.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (norm_(rows[i][1]) === norm_(name) && norm_(rows[i][2]) === norm_(unit) && norm_(rows[i][3]) === norm_(status)) {
        return false;
      }
      var m = String(rows[i][0]).match(new RegExp('^' + ymd + '-(\\d+)$'));
      if (m) {
        var n = parseInt(m[1], 10);
        if (n > maxSeq) maxSeq = n;
      }
    }
  } catch (e) {}
  var seq = maxSeq + 1;
  var seqStr = (seq < 10) ? '0' + seq : String(seq);
  var id = ymd + '-' + seqStr;
  ss.appendRow([id, name, unit, status, new Date()]);
  return true;
}

function hasSting_(name, unit) {
  var t = trainRows_().rows;
  var nk = norm_(name), uk = norm_(unit);
  for (var i = 0; i < t.length; i++) {
    var st = norm_(t[i][3]);
    if (norm_(t[i][1]) === nk && norm_(t[i][2]) === uk && st === '已敕令') return true;
  }
  return false;
}

function setTrain_(name, unit, status) {
  var m = sopMap_();
  var o = m.map[norm_(unit)];
  if (!o) return textCard_('找不到單元', '#B04A3A', unit + '（用「庫存」看正確名稱）');
  if (status === '已敕令' || status === '已領取') {
    if (!hasSting_(name, unit)) {
      var isDecree = (status === '已敕令');
      var appended = false;
      if (isDecree) appended = trainAppend_(name, o.unit, '已敕令');
      else { var a = trainAppend_(name, o.unit, '已敕令'); var b = trainAppend_(name, o.unit, '已領取'); appended = a || b; }
      var ok = true;
      if (appended) ok = deductStock_(o.unit);
      return textCard_('已更新', '#4A7C59', name + '・' + o.unit + '→' + status + (appended ? (ok ? '（已扣庫存）' : '（庫存不足0）') : '（重複，未重複扣庫存）'));
    }
  }
  var ap = trainAppend_(name, o.unit, status);
  return textCard_('已更新', '#4A7C59', name + '・' + o.unit + '→' + status + (ap ? '' : '（重複紀錄）'));
}

function apprenticeState_() {
  var m = sopMap_();
  var pos = {}; m.list.forEach(function(o) { pos[o.key] = o.idx; });
  var rows = trainRows_().rows;
  // 先以「姓名+單元」為 key，保留最後一筆（最新）狀態，避免流水帳舊紀錄干擾
  var latest = {};
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    var nm = String(r[1]).trim(); if (!nm) continue;
    var uk = norm_(r[2]); if (!uk) continue;
    latest[norm_(nm) + '|' + uk] = { name: nm, unit: String(r[2]).trim(), status: norm_(r[3]), time: r[4] };
  }
  var by = {};
  var unmatched = [];
  var lkeys = Object.keys(latest);
  for (var i = 0; i < lkeys.length; i++) {
    var rec = latest[lkeys[i]];
    var nm = rec.name;
    var uk = norm_(rec.unit);
    var st = rec.status;
    var g = by[nm] || (by[nm] = { name: nm, active: 0, done: 0, latestIdx: -1, earliest: Infinity, units: [], stung: {}, have: {}, completed: {} });
    if (st === '已領取') {
      g.active++;
      g.units.push(rec.unit);
      var tm = new Date(rec.time).getTime();
      if (!isNaN(tm) && tm < g.earliest) g.earliest = tm;
    }
    if (st === 'V' || st === 'v') g.done++;
    if (st === '已敕令') g.stung[uk] = true;
    if (st === 'V' || st === 'v' || st === '已領取' || st === '已敕令') g.have[uk] = true;
    if (st === 'V' || st === 'v') g.completed[uk] = true;
    if (st === '已領取' || st === 'V' || st === 'v') {
      if (uk && pos[uk] !== undefined) { if (pos[uk] > g.latestIdx) g.latestIdx = pos[uk]; }
      else if (uk && unmatched.length < 3 && unmatched.indexOf(rec.unit) === -1) unmatched.push(rec.unit);
    }
  }
  return { sop: m, by: by, unmatched: unmatched };
}

function trainOverview_() {
  var s = apprenticeState_();
  var names = Object.keys(s.by);
  if (!names.length) return { card: textCard_('全體進度', '#8B7355', '訓練紀錄無資料'), copy: null };
  var show = names.slice(0, 15);
  var contents = [];
  show.forEach(function(nm, idx) {
    var g = s.by[nm];
    var summary = '';
    var nextLine = '';
    if (!g.active) {
      if (g.done > 0 && g.latestIdx >= 0 && g.latestIdx >= s.sop.list.length - 1) {
        summary = nm + '｜完訓（共完' + g.done + '）';
      } else {
        summary = nm + '｜⚠LOSS（無進行中，完' + g.done + '）';
      }
    } else {
      var nu = s.sop.list[g.latestIdx + 1];
      var nxt = '完訓';
      if (nu) {
        var tags = [];
        if (g.stung[nu.key]) {
          tags.push('已敕令');
        } else {
          if (nu.c > 0) tags.push('蓋章' + nu.c);
          else if (nu.b > 0) tags.push('列印' + nu.b);
          else tags.push('未備料');
        }
        nxt = nu.unit + '(' + tags.join('/') + ')';
      }
      summary = nm + '｜完成' + g.done + '/進行中' + g.active + '｜' + g.units.join('、');
      nextLine = '    下:' + nxt;
    }
    var boxContents = [{ type: 'text', text: summary, size: 'sm', color: '#333333', wrap: true }];
    if (nextLine) boxContents.push({ type: 'text', text: nextLine, size: 'sm', color: '#777777', wrap: true, margin: 'xs' });
    contents.push({ type: 'box', layout: 'vertical', margin: idx > 0 ? 'lg' : 'none', contents: boxContents });
    if (idx < show.length - 1) contents.push({ type: 'separator', margin: 'md' });
  });
  var card = { type: 'flex', altText: '全體進度', contents: { type: 'bubble',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#8B7355', contents: [{ type: 'text', text: '全體進度共' + show.length + '人', size: 'md', color: '#ffffff', weight: 'bold', wrap: true }] },
    body: { type: 'box', layout: 'vertical', spacing: 'md', contents: contents }
  }};
  return { card: card, copy: null };
}

function stockText_() {
  var s = apprenticeState_();
  var m = s.sop;
  if (!m.list.length) return textCard_('庫存', '#8B7355', 'SOP表無資料');
  var dem = {}, i, g, k;
  var names = Object.keys(s.by);
  var contributors = [];
  for (i = 0; i < names.length; i++) {
    g = s.by[names[i]];
    if (!g.active || g.latestIdx < 0) continue;
    contributors.push(names[i]);
    for (k = 1; k <= 3; k++) {
      var o = m.list[g.latestIdx + k];
      if (!o) break;
      dem[o.key] = true;
    }
  }
  var keys = Object.keys(dem);
  if (!keys.length) {
    var hasActive = names.some(function(nm) { return s.by[nm].active > 0; });
    if (hasActive && s.unmatched.length) {
      return textCard_('單元名稱對不上', '#B04A3A', '紀錄：' + s.unmatched.join('、') + '\nSOP首筆：' + m.list[0].unit + '\n請兩邊用完全相同名稱');
    }
    var minLatest = Infinity;
    names.forEach(function(nm) { var li = s.by[nm].latestIdx; if (li >= 0 && li < minLatest) minLatest = li; });
    var zero = m.list.filter(function(o) { return (o.b + o.c) < 1 && (minLatest === Infinity || o.idx > minLatest); });
    if (!zero.length) return textCard_('庫存', '#4A7C59', '目前無進行中徒弟，且全關都有庫存');
    var maxZ = 0;
    zero.forEach(function(o) { var l = pageLabel_(o.unit).length; if (l > maxZ) maxZ = l; });
    var zlines = zero.slice(0, 15).map(function(o) {
      var label = pageLabel_(o.unit);
      while (label.length < maxZ) label += ' ';
      return label + '｜庫0（印' + o.b + '/章' + o.c + '）';
    });
    return textCard_('待補庫存' + zero.length + '關', '#B04A3A', zlines.join('\n'));
  }
  keys.sort(function(a, b) { return m.map[a].idx - m.map[b].idx; });
  var rows = [];
  var currentVol = '';
  keys.forEach(function(key) {
    var o = m.map[key];
    var need = 0, j;
    for (j = 0; j < contributors.length; j++) { if (!s.by[contributors[j]].have[key]) need++; }
    if (need <= 0) return;
    var vol = (o.unit.match(/^(.+冊)/) || ['', ''])[1] || '其他';
    if (vol !== currentVol) {
      currentVol = vol;
      rows.push({ type: 'text', text: '【' + vol + '】', size: 'sm', color: '#8B7355', weight: 'bold', wrap: true, margin: rows.length > 0 ? 'sm' : 'none' });
    }
    var stock = o.b + o.c;
    var lack = Math.max(0, need - stock);
    var label = pageLabel_(o.unit);
    rows.push({ type: 'box', layout: 'horizontal', margin: 'xs', contents: [
      { type: 'text', text: label, size: 'sm', color: '#333333', flex: 2, wrap: true },
      { type: 'text', text: '庫' + stock + '（印' + o.b + '/章' + o.c + '）', size: 'sm', color: '#333333', flex: 4, wrap: true },
      { type: 'text', text: '需' + need, size: 'sm', color: '#333333', flex: 1, wrap: true },
      { type: 'text', text: '缺' + lack, size: 'sm', color: lack > 0 ? '#B04A3A' : '#4A7C59', flex: 1, wrap: true }
    ]});
    rows.push({ type: 'separator', margin: 'xs' });
  });
  if (!rows.length) return textCard_('庫存', '#4A7C59', '3關內人人都有料，不用補');
  var card = { type: 'flex', altText: '3關內需求', contents: { type: 'bubble',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#8B7355', contents: [{ type: 'text', text: '3關內需求（' + contributors.length + '人）', size: 'md', color: '#ffffff', weight: 'bold', wrap: true }] },
    body: { type: 'box', layout: 'vertical', spacing: 'xs', contents: rows.slice(0, 30) }
  }};
  return card;
}

function sopVolumes_() {
  var m = sopMap_();
  var vols = [], byVol = {};
  m.list.forEach(function(o) {
    var mt = o.unit.match(/^(.+冊)\s*(.+)$/);
    var v = mt ? mt[1] : '其他';
    if (!byVol[v]) { byVol[v] = []; vols.push(v); }
    byVol[v].push(o);
  });
  return { vols: vols, byVol: byVol };
}

function opStatus_(op) { return op === '完成' ? 'V' : op === '敕令' ? '已敕令' : '已領取'; }

function pageLabel_(unit) { return unit.replace(/^.+冊\s*/, '') || unit; }

function tUpdCards_(q) {
  if (!q.op) {
    return btnCard_('更新進度・選動作', '#8B7355', '第一步：選動作', ['敕令', '領取', '完成'].map(function(op) {
      return pbBtn_(op, 'action=tupd&step=name&op=' + encodeURIComponent(op), op, 'secondary');
    }));
  }
  if (!q.name) {
    var names = Object.keys(apprenticeState_().by);
    if (!names.length) return textCard_('無徒弟資料', '#B04A3A', '訓練紀錄是空的，請直接打字：領取 姓名 單元');
    return btnCard_('更新' + q.op + '・選人', '#8B7355', '第二步：選徒弟（新人直接打字：' + q.op + ' 姓名 單元）', names.slice(0, 10).map(function(nm) {
      return pbBtn_(nm, 'action=tupd&step=vol&op=' + encodeURIComponent(q.op) + '&name=' + encodeURIComponent(nm), q.op + ' ' + nm, 'secondary');
    }));
  }
  if (!q.vol && q.step !== 'done') {
    var sv = sopVolumes_();
    if (!sv.vols.length) return textCard_('無SOP資料', '#B04A3A', 'SOP表是空的');
    return btnCard_('更新' + q.op + '・選冊 (' + q.name + ')', '#8B7355', '第三步：選冊', sv.vols.map(function(v) {
      return pbBtn_(v, 'action=tupd&step=unit&op=' + encodeURIComponent(q.op) + '&name=' + encodeURIComponent(q.name) + '&vol=' + encodeURIComponent(v), v, 'secondary');
    }));
  }
  if (!q.unit) {
    var sv2 = sopVolumes_();
    var list = sv2.byVol[q.vol] || [];
    if (!list.length) return textCard_('無此冊', '#B04A3A', q.vol);
    var st = apprenticeState_().by[q.name] || { have: {}, stung: {}, completed: {} };
    return btnCard_('更新' + q.op + '・選單元 (' + q.name + ' - ' + q.vol + ')', '#8B7355', '第四步：選單元（括號內為現庫存）', list.slice(0, 20).map(function(o) {
      var uk = norm_(o.unit);
      var hasTaken = !!st.have[uk];
      var isStung = !!st.stung[uk];
      var disabled = false;
      var disabledLabel = '';
      if (q.op === '敕令' && (hasTaken || isStung)) { disabled = true; disabledLabel = pageLabel_(o.unit) + ' (已敕令)'; }
      if (q.op === '領取' && (hasTaken && !isStung)) { disabled = true; disabledLabel = pageLabel_(o.unit) + ' (已領取)'; }
      if (q.op === '完成' && !!st.completed[uk]) { disabled = true; disabledLabel = pageLabel_(o.unit) + ' (已完成)'; }
      if (disabled) {
        // 反黑按鈕：不帶 displayText，點擊後不會在聊天室送出文字
        return { type: 'button', action: { type: 'postback', label: disabledLabel, data: 'action=ignore' }, height: 'sm', margin: 'sm', style: 'secondary', color: '#666666' };
      }
      return pbBtn_(pageLabel_(o.unit) + '(庫' + (o.b + o.c) + ')',
        'action=tupd&step=done&op=' + encodeURIComponent(q.op) + '&name=' + encodeURIComponent(q.name) + '&unit=' + encodeURIComponent(o.unit),
        q.op + ' ' + q.name + ' ' + o.unit, 'secondary');
    }));
  }
  return setTrain_(q.name, q.unit, opStatus_(q.op));
}

function tAddCards_(q) {
  var kindLabel = (q.kind === 'C') ? '蓋章' : '列印';
  if (!q.kind) {
    return btnCard_('補庫存・選種類', '#8B7355', '第一步：選種類', [
      pbBtn_('列印', 'action=tadd&step=vol&kind=B', '補庫存 列印', 'secondary'),
      pbBtn_('蓋章', 'action=tadd&step=vol&kind=C', '補庫存 蓋章', 'primary', '#8B7355')
    ]);
  }
  if (!q.vol && q.step !== 'done') {
    var sv = sopVolumes_();
    if (!sv.vols.length) return textCard_('無SOP資料', '#B04A3A', 'SOP表是空的');
    return btnCard_('補庫存・選冊 (' + kindLabel + ')', '#8B7355', '第二步：選冊', sv.vols.map(function(v) {
      return pbBtn_(v, 'action=tadd&step=unit&kind=' + q.kind + '&vol=' + encodeURIComponent(v), v, 'secondary');
    }));
  }
  if (!q.unit) {
    var sv2 = sopVolumes_();
    var list = sv2.byVol[q.vol] || [];
    if (!list.length) return textCard_('無此冊', '#B04A3A', q.vol);
    return btnCard_('補庫存・選單元 (' + kindLabel + ')', '#8B7355', '第三步：選單元', list.slice(0, 20).map(function(o) {
      var fillText = '補庫存 ' + o.unit + ' ' + kindLabel + ' ';
      return { type: 'button', action: { type: 'postback', label: pageLabel_(o.unit) + '(庫' + (o.b + o.c) + ')', data: 'action=tadd_prefill&hint=' + encodeURIComponent('請輸入數量（整數）'), inputOption: 'openKeyboard', fillInText: fillText }, style: 'secondary', height: 'sm', margin: 'sm' };
    }));
  }
  ensureTrainSheets_();
  var mm = sopMap_().map[norm_(q.unit)];
  if (!mm) return textCard_('找不到單元', '#B04A3A', q.unit);
  var isC = q.kind === 'C';
  var newB, newC;
  if (isC) {
    // 蓋章：優先將已列印的轉為蓋章，避免重複計數
    if (mm.b > 0) { newB = mm.b - 1; newC = mm.c + 1; }
    else { newB = mm.b; newC = mm.c + 1; }
  } else {
    newB = mm.b + 1; newC = mm.c;
  }
  var sh = sheetByName_(SHEET_SOP);
  sh.getRange(mm.row, 2).setValue(newB);
  sh.getRange(mm.row, 3).setValue(newC);
  return textCard_('已補庫存', '#4A7C59', mm.unit + ' ' + (isC ? '蓋章' : '列印') + '+1（現印' + newB + '章' + newC + '）');
}

function handleTrainPostback_(data) {
  var q = {};
  String(data || '').split('&').forEach(function(kv) {
    var p = kv.split('=');
    if (p[0]) q[decodeURIComponent(p[0])] = decodeURIComponent(p[1] || '');
  });
  if (q.action === 'ignore') return true;
  if (q.action === 'train_prefill') return textCard_('補庫存', '#8B7355', q.hint || '請輸入：補庫存 單元 列印/蓋章 數量');
  if (q.action === 'tadd_prefill') return textCard_('補庫存', '#8B7355', q.hint || '請輸入數量（整數）');
  if (q.action === 'progress') return trainOverview_();
  if (q.action === 'stock') return stockText_();
  if (q.action === 'stockcheck') return stockCheckAll_();
  if (q.action === 'stockcheckvol' && q.vol) return stockCheckVol_(q.vol);
  if (q.action === 'list') return listCases_('未結案');
  if (q.action === 'tupd') return tUpdCards_(q);
  if (q.action === 'tadd') return tAddCards_(q);
  if (q.action === 'menu') return q.go === 'train' ? trainMenuCard_() : consultMenuCard_();
  if (q.action === 'caselist') return caseListCard_(q.filter || '未結案');
  if (q.action === 'case') return caseCard_(q.caseId);
  if (q.action === 'casestatus') return doCaseStatus_(q.caseId, q.to);
  if ((q.action === 'claim' || q.action === 'done' || q.action === 'sting') && q.name && q.unit) {
    return setTrain_(q.name, q.unit, q.action === 'done' ? 'V' : q.action === 'sting' ? '已敕令' : '已領取');
  }
  if (q.action === 'addStock' && q.unit && q.kind) {
    ensureTrainSheets_();
    var mm = sopMap_().map[norm_(q.unit)];
    if (!mm) return textCard_('找不到單元', '#B04A3A', q.unit);
    var sh = sheetByName_(SHEET_SOP);
    var isC = q.kind === 'C';
    var newB, newC;
    if (isC) {
      if (mm.b > 0) { newB = mm.b - 1; newC = mm.c + 1; }
      else { newB = mm.b; newC = mm.c + 1; }
    } else {
      newB = mm.b + 1; newC = mm.c;
    }
    sh.getRange(mm.row, 2).setValue(newB);
    sh.getRange(mm.row, 3).setValue(newC);
    return textCard_('已補庫存', '#4A7C59', q.unit + ' ' + (isC ? '蓋章+1' : '列印+1') + '（現印' + newB + '章' + newC + '）');
  }
  return null;
}

function dailyTrainAlert_() {
  try {
    ensureTrainSheets_();
    var m = sopMap_();
    if (!m.list.length) return;
    var pos = {}; m.list.forEach(function(o) { pos[o.key] = o.idx; });
    var rows = trainRows_().rows;
    var by = {};
    rows.forEach(function(r) {
      var nm = String(r[1]).trim(); if (!nm) return;
      var st = norm_(r[3]);
      var un = norm_(r[2]);
      var tm = new Date(r[4]).getTime(); if (isNaN(tm)) return;
      var g = by[nm] || (by[nm] = { n: 0, earliest: Infinity, latestIdx: -1 });
      if (st === '已領取') { g.n++; if (tm < g.earliest) g.earliest = tm; }
      if ((st === '已領取' || st === 'V' || st === 'v') && pos[un] !== undefined && pos[un] > g.latestIdx) g.latestIdx = pos[un];
    });
    var now = Date.now();
    var out = [];
    Object.keys(by).forEach(function(nm) {
      var g = by[nm];
      if (!g.n || g.earliest === Infinity || g.latestIdx < 0) return;
      var passed = Math.floor((now - g.earliest) / 86400000);
      var th = (g.n - 1) * 31 + 25;
      if (passed < th) return;
      var warns = [];
      var t1 = m.list[g.latestIdx + 1], t2 = m.list[g.latestIdx + 2], t3 = m.list[g.latestIdx + 3];
      var t1hit = false, t2hit = false;
      if (t1 && (t1.b + t1.c) < 1) { warns.push('下一關 ' + t1.unit + ' 無庫存'); t1hit = true; }
      if (t2 && (t2.b + t2.c) < 1 && ((passed - th) % 5 === 0)) { warns.push('下下關 ' + t2.unit + ' 無庫存'); t2hit = true; }
      if ((t1hit || t2hit) && t3 && (t3.b + t3.c) < 1) warns.push('（另：' + t3.unit + '亦無庫存）');
      if (warns.length) out.push(nm + '｜N=' + g.n + '｜D' + passed + '：' + warns.join('；'));
    });
    if (!out.length) return;
    pushLine_(targetGroup_(), flexCard_('備課提醒', '#B04A3A', [['缺漏', out.slice(0, 10).join('\n').slice(0, 1400)]], null));
  } catch (err) { try { mailAdmin_('備課提醒失敗 ' + err, ''); } catch (e2) {} }
}

// ===== 庫存盤點 / 確認 =====
function stockCheckAll_() {
  var sv = sopVolumes_();
  if (!sv.vols.length) return textCard_('庫存盤點', '#8B7355', 'SOP表無資料');
  var btns = sv.vols.map(function(v) {
    return pbBtn_(v, 'action=stockcheckvol&vol=' + encodeURIComponent(v), '盤點 ' + v, 'secondary');
  });
  return btnCard_('庫存盤點', '#8B7355', '選擇冊別查看庫存盤點（共 ' + sv.vols.length + ' 冊）', btns);
}

function stockCheckVol_(vol) {
  var s = apprenticeState_();
  var sv = sopVolumes_();
  var list = sv.byVol[vol];
  if (!list || !list.length) return textCard_('無此冊', '#B04A3A', vol);
  var names = Object.keys(s.by);
  var rows = [];
  list.forEach(function(o, idx) {
    var need = 0;
    for (var i = 0; i < names.length; i++) {
      var g = s.by[names[i]];
      if (!g.active) continue;
      if (!g.have[o.key]) need++;
    }
    var stock = o.b + o.c;
    var lack = Math.max(0, need - stock);
    var label = pageLabel_(o.unit);
    rows.push({ type: 'box', layout: 'horizontal', margin: idx > 0 ? 'sm' : 'none', contents: [
      { type: 'text', text: label, size: 'sm', color: '#333333', flex: 2, wrap: true },
      { type: 'text', text: '庫' + stock + '（印' + o.b + '/章' + o.c + '）', size: 'sm', color: '#333333', flex: 4, wrap: true },
      { type: 'text', text: '需' + need, size: 'sm', color: '#333333', flex: 1, wrap: true },
      { type: 'text', text: '缺' + lack, size: 'sm', color: lack > 0 ? '#B04A3A' : '#4A7C59', flex: 1, wrap: true }
    ]});
    rows.push({ type: 'separator', margin: 'xs' });
  });
  var card = { type: 'flex', altText: '庫存盤點・' + vol, contents: { type: 'bubble',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#8B7355', contents: [{ type: 'text', text: '庫存盤點・' + vol, size: 'md', color: '#ffffff', weight: 'bold', wrap: true }] },
    body: { type: 'box', layout: 'vertical', spacing: 'xs', contents: rows }
  }};
  return card;
}

function setStock_(unit, kind, qty) {
  ensureTrainSheets_();
  var m = sopMap_();
  var o = m.map[norm_(unit)];
  if (!o) return textCard_('找不到單元', '#B04A3A', unit + '（用「庫存」看正確名稱）');
  var sh = sheetByName_(SHEET_SOP);
  var isC = (kind === '蓋章');
  if (isC) sh.getRange(o.row, 3).setValue(Number(qty));
  else sh.getRange(o.row, 2).setValue(Number(qty));
  return textCard_('已設定庫存', '#4A7C59', o.unit + ' ' + kind + '=' + qty);
}
