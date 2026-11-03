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
  // 防重複：相同 姓名+單元+狀態 已存在則不新增
  try {
    var rows = ss.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (norm_(rows[i][1]) === norm_(name) && norm_(rows[i][2]) === norm_(unit) && norm_(rows[i][3]) === norm_(status)) {
        return false;
      }
    }
  } catch (e) {}
  var id = new Date().getTime().toString(36) + Math.floor(Math.random() * 99);
  ss.appendRow([id, name, unit, status, new Date()]);
  return true;
}

function hasSting_(name, unit) {
  var t = trainRows_().rows;
  var nk = norm_(name), uk = norm_(unit);
  for (var i = 0; i < t.length; i++) {
    if (norm_(t[i][1]) === nk && norm_(t[i][2]) === uk && norm_(t[i][3]) === '已刺令') return true;
  }
  return false;
}

function setTrain_(name, unit, status) {
  var m = sopMap_();
  var o = m.map[norm_(unit)];
  if (!o) return textCard_('找不到單元', '#B04A3A', unit + '（用「庫存」看正確名稱）');
  if (status === '已刺令' || status === '已領取') {
    if (!hasSting_(name, unit)) {
      var appended = false;
      if (status === '已刺令') appended = trainAppend_(name, o.unit, '已刺令');
      else { var a = trainAppend_(name, o.unit, '已刺令'); var b = trainAppend_(name, o.unit, '已領取'); appended = a || b; }
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
  var by = {};
  var unmatched = [];
  rows.forEach(function(r) {
    var nm = String(r[1]).trim(); if (!nm) return;
    var st = norm_(r[3]);
    var uk = norm_(r[2]);
    var g = by[nm] || (by[nm] = { name: nm, active: 0, done: 0, latestIdx: -1, earliest: Infinity, units: [], stung: {}, have: {} });
    if (st === '已領取') {
      g.active++;
      g.units.push(String(r[2]).trim());
      var tm = new Date(r[4]).getTime();
      if (!isNaN(tm) && tm < g.earliest) g.earliest = tm;
    }
    if (st === 'V' || st === 'v') g.done++;
    if (st === '已刺令') g.stung[uk] = true;
    if (st === 'V' || st === 'v' || st === '已領取' || st === '已刺令') g.have[uk] = true;
    if ((st === '已領取' || st === 'V' || st === 'v')) {
      if (uk && pos[uk] !== undefined) { if (pos[uk] > g.latestIdx) g.latestIdx = pos[uk]; }
      else if (uk && unmatched.length < 3 && unmatched.indexOf(String(r[2]).trim()) === -1) unmatched.push(String(r[2]).trim());
    }
  });
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
        if (g.stung[nu.key]) tags.push('已刺令');
        if (nu.c > 0) tags.push('蓋章' + nu.c);
        else if (nu.b > 0) tags.push('列印' + nu.b);
        else tags.push('未備料');
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
  try {
  var m = sopMap_();
  if (!m.list.length) return textCard_('庫存', '#8B7355', 'SOP表無資料');

  // 建立「單元 -> 已領取人名」對照
  var claimers = {};
  var tr = trainRows_().rows;
  for (var i = 0; i < tr.length; i++) {
    var r = tr[i];
    if (norm_(r[3]) !== '已領取') continue;
    var uk = norm_(r[2]);
    var nm = String(r[1]).trim();
    if (!claimers[uk]) claimers[uk] = [];
    if (claimers[uk].indexOf(nm) === -1) claimers[uk].push(nm);
  }

  // 計算 3 關內需求（用於 需/缺）
  var s2 = apprenticeState_();
  var names = Object.keys(s2.by);
  var contributors = [];
  for (var i = 0; i < names.length; i++) {
    var g = s2.by[names[i]];
    if (!g.active || g.latestIdx < 0) continue;
    contributors.push(names[i]);
  }
  var dem = {};
  for (var i = 0; i < contributors.length; i++) {
    var g = s2.by[contributors[i]];
    for (var k = 1; k <= 3; k++) {
      var o = m.list[g.latestIdx + k];
      if (!o) break;
      dem[o.key] = true;
    }
  }

  var sv = sopVolumes_();
  var rows = [];
  sv.vols.forEach(function(v) {
    rows.push({ type: 'text', text: '【' + v + '】', size: 'sm', color: '#8B7355', weight: 'bold', wrap: true, margin: rows.length > 0 ? 'sm' : 'none' });
    (sv.byVol[v] || []).forEach(function(o) {
      var need = 0;
      if (dem[o.key]) {
        for (var j = 0; j < contributors.length; j++) {
          if (!s2.by[contributors[j]].have[o.key]) need++;
        }
      }
      var stock = o.b + o.c;
      var lack = Math.max(0, need - stock);
      var c = claimers[norm_(o.unit)] || [];
      var namesStr = c.length ? c.join('、') : '';
      var label = pageLabel_(o.unit);
      rows.push({ type: 'box', layout: 'horizontal', margin: 'xs', contents: [
        { type: 'text', text: label, size: 'xs', color: '#333333', flex: 2, wrap: true },
        { type: 'text', text: '庫' + stock + '（印' + o.b + '/章' + o.c + '）', size: 'xs', color: '#333333', flex: 3, wrap: true },
        { type: 'text', text: '需' + need, size: 'xs', color: '#333333', flex: 1, wrap: true },
        { type: 'text', text: '缺' + lack, size: 'xs', color: lack > 0 ? '#B04A3A' : '#4A7C59', flex: 1, wrap: true },
        { type: 'text', text: namesStr, size: 'xs', color: '#555555', flex: 3, wrap: true }
      ]});
    });
  });

  var bodyContents = rows;
  if (bodyContents.length > 80) {
    bodyContents = bodyContents.slice(0, 80);
    bodyContents.push({ type: 'text', text: '…（資料過多，僅顯示前 80 筆）', size: 'xs', color: '#999999', wrap: true, margin: 'sm' });
  }
  var card = { type: 'flex', altText: '庫存查詢', contents: { type: 'bubble',
    header: { type: 'box', layout: 'vertical', backgroundColor: '#8B7355', contents: [{ type: 'text', text: '庫存查詢', size: 'md', color: '#ffffff', weight: 'bold', wrap: true }] },
    body: { type: 'box', layout: 'vertical', spacing: 'xs', contents: bodyContents }
  }};
  return card;
  } catch (e) { return textCard_('庫存查詢錯誤', '#B04A3A', String(e)); }
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

function opStatus_(op) { return op === '完成' ? 'V' : op === '刺令' ? '已刺令' : '已領取'; }

function pageLabel_(unit) { return unit.replace(/^.+冊\s*/, '') || unit; }

function tUpdCards_(q) {
  if (!q.op) {
    return btnCard_('更新進度・選動作', '#8B7355', '第一步：選動作', ['刺令', '領取', '完成'].map(function(op) {
      return pbBtn_(op, 'action=tupd&step=name&op=' + encodeURIComponent(op), op, 'secondary');
    }));
  }
  if (!q.name) {
    var names = Object.keys(apprenticeState_().by);
    if (!names.length) return textCard_('無徒弟資料', '#B04A3A', '訓練紀錄是空的，請直接打字：領取 姓名 單元');
    return btnCard_('更新進度・選人（' + q.op + '）', '#8B7355', '第二步：選徒弟（新人直接打字：' + q.op + ' 姓名 單元）', names.slice(0, 10).map(function(nm) {
      return pbBtn_(nm, 'action=tupd&step=vol&op=' + encodeURIComponent(q.op) + '&name=' + encodeURIComponent(nm), q.op + ' ' + nm, 'secondary');
    }));
  }
  if (!q.vol && q.step !== 'done') {
    var sv = sopVolumes_();
    if (!sv.vols.length) return textCard_('無SOP資料', '#B04A3A', 'SOP表是空的');
    return btnCard_('更新進度・選冊', '#8B7355', '第三步：選冊', sv.vols.map(function(v) {
      return pbBtn_(v, 'action=tupd&step=unit&op=' + encodeURIComponent(q.op) + '&name=' + encodeURIComponent(q.name) + '&vol=' + encodeURIComponent(v), v, 'secondary');
    }));
  }
  if (!q.unit) {
    var sv2 = sopVolumes_();
    var list = sv2.byVol[q.vol] || [];
    if (!list.length) return textCard_('無此冊', '#B04A3A', q.vol);
    return btnCard_('更新進度・選單元', '#8B7355', '第四步：選單元（括號內為現庫存）', list.slice(0, 20).map(function(o) {
      return pbBtn_(pageLabel_(o.unit) + '(庫' + (o.b + o.c) + ')',
        'action=tupd&step=done&op=' + encodeURIComponent(q.op) + '&name=' + encodeURIComponent(q.name) + '&unit=' + encodeURIComponent(o.unit),
        q.op + ' ' + q.name + ' ' + o.unit, 'secondary');
    }));
  }
  return setTrain_(q.name, q.unit, opStatus_(q.op));
}

function tAddCards_(q) {
  if (!q.kind) {
    return btnCard_('補庫存・選種類', '#8B7355', '第一步：選種類', [
      pbBtn_('列印+1', 'action=tadd&step=vol&kind=B', '補庫存 列印', 'secondary'),
      pbBtn_('蓋章+1', 'action=tadd&step=vol&kind=C', '補庫存 蓋章', 'primary', '#8B7355')
    ]);
  }
  if (!q.vol && q.step !== 'done') {
    var sv = sopVolumes_();
    if (!sv.vols.length) return textCard_('無SOP資料', '#B04A3A', 'SOP表是空的');
    return btnCard_('補庫存・選冊', '#8B7355', '第二步：選冊', sv.vols.map(function(v) {
      return pbBtn_(v, 'action=tadd&step=unit&kind=' + q.kind + '&vol=' + encodeURIComponent(v), v, 'secondary');
    }));
  }
  if (!q.unit) {
    var sv2 = sopVolumes_();
    var list = sv2.byVol[q.vol] || [];
    if (!list.length) return textCard_('無此冊', '#B04A3A', q.vol);
    return btnCard_('補庫存・選單元', '#8B7355', '第三步：選單元', list.slice(0, 20).map(function(o) {
      return pbBtn_(pageLabel_(o.unit) + '(庫' + (o.b + o.c) + ')',
        'action=tadd&step=done&kind=' + q.kind + '&unit=' + encodeURIComponent(o.unit),
        '補庫存 ' + o.unit, 'secondary');
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
    return setTrain_(q.name, q.unit, q.action === 'done' ? 'V' : q.action === 'sting' ? '已刺令' : '已領取');
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
