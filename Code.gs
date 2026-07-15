// ===== Bybit KZ Affiliate Champions API =====
// Google Apps Script Web App backend.
// Deployed as: Execute as Me, Who has access Anyone.
// After any edit: Deploy > Manage deployments > Edit > Version: New version > Deploy.

function doGet(e) {
  var action = e.parameter.action;
  if (action === 'validateAID') { var resolved = resolveToAID(e.parameter.aid); return json({ valid: !!resolved, aid: resolved || '' }); }
  if (action === 'leaderboard') return json({ leaderboard: getLeaderboard(), last_updated: getLastUpdated() });
  if (action === 'stats')       return json({ stats: getStats(e.parameter.aid) });
  if (action === 'tasks')       return json({ tasks: getTasks() });
  if (action === 'pool')        return json({ pool: getPool() });
  if (action === 'saveMapping') { saveMapping(e.parameter.userId, e.parameter.aid, e.parameter.nick); return json({ ok: true }); }
  if (action === 'getByUserId')   return json({ data: getByUserId(e.parameter.userId) });
  if (action === 'submitLink')    { submitLink(e.parameter.aid, e.parameter.taskId, e.parameter.link, e.parameter.social, e.parameter.views); return json({ ok: true }); }
  if (action === 'getSubmissions') return json({ submissions: getSubmissions(e.parameter.aid) });
  if (action === 'ping') { pingUser(e.parameter.userId); return json({ ok: true }); }
  return json({ error: 'unknown action' });
}

function sheet(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function json(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

function rows(name) {
  var sh = sheet(name);
  var data = sh.getDataRange().getValues();
  var headers = data.shift();
  return data.map(function(r) {
    var o = {};
    headers.forEach(function(h, i) { o[String(h).trim()] = r[i]; });
    return o;
  });
}

// read config into key/value object
function getConfig() {
  var sh = sheet('config');
  var data = sh.getDataRange().getValues();
  var c = {};
  for (var i = 1; i < data.length; i++) {
    var key = String(data[i][0]).trim();
    if (key) c[key] = Number(data[i][1]) || 0;
  }
  return c;
}

// read raw_export, skip "Sum" row, GROUP BY AID and SUM metrics (dupes = different commission rows)
var EXCLUDED_AIDS = ['160869', '161151'];

function affiliateRows() {
  var raw = rows('raw_export').filter(function(r) {
    var aid = String(r['AID']).trim();
    return aid && aid.toLowerCase() !== 'sum' && EXCLUDED_AIDS.indexOf(aid) === -1;
  });
  var grouped = {};
  raw.forEach(function(r) {
    var aid = String(r['AID']).trim();
    if (!grouped[aid]) grouped[aid] = { AID: aid, AFF_Name: r['AFF_Name'] || '', eFTD: 0, ftd: 0, reg: 0, vol: 0, approvedDate: '', uid: '' };
    grouped[aid].eFTD += Number(r['eFTD']) || 0;
    grouped[aid].ftd  += Number(r['ftd']) || 0;
    grouped[aid].reg  += Number(r['Reg Users']) || 0;
    grouped[aid].vol  += Number(r['Vol_Portal_Client_Non_MT5']) || 0;
    if (!grouped[aid].approvedDate && r['Aff_Approved_date']) grouped[aid].approvedDate = String(r['Aff_Approved_date']);
    if (!grouped[aid].uid && r['UID']) grouped[aid].uid = String(r['UID']).trim();
  });
  return Object.keys(grouped).map(function(k) { return grouped[k]; });
}

function calcPoints(a, c) {
  var volUnit = c.vol_unit_size || 100000;
  return a.reg  * (c.points_per_reg  || 0)
       + a.ftd  * (c.points_per_ftd  || 0)
       + a.eFTD * (c.points_per_eftd || 0)
       + Math.floor(a.vol / volUnit) * (c.points_per_vol_unit || 0);
}

function levelFor(points, c) {
  if (points >= (c.level_elite  || 1500)) return 'Elite';
  if (points >= (c.level_gold   || 700))  return 'Gold';
  if (points >= (c.level_silver || 300))  return 'Silver';
  if (points >= (c.level_bronze || 100))  return 'Bronze';
  return 'Starter';
}

function isNewcomer(approvedDateStr) {
  if (!approvedDateStr) return false;
  var today = new Date();
  var d = new Date(approvedDateStr);
  if (isNaN(d)) return false;
  var diff = today - d;
  return diff >= 0 && diff <= 90 * 24 * 60 * 60 * 1000;
}

function resolveToAID(input) {
  if (!input) return null;
  var str = String(input).trim();
  var all = affiliateRows();
  var byAID = all.filter(function(a) { return a.AID === str; })[0];
  if (byAID) return byAID.AID;
  var byUID = all.filter(function(a) { return a.uid && a.uid === str; })[0];
  if (byUID) return byUID.AID;
  return null;
}

function aidExists(aid) {
  return !!resolveToAID(aid);
}

function nickMap() {
  var m = {};
  rows('mapping').forEach(function(x) { if (x['aid']) m[String(x['aid']).trim()] = x['nickname']; });
  return m;
}

function getAllTaskPoints() {
  var sh = sheet('submissions');
  if (!sh) return {};
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return {};
  var headers = data[0];
  var aidIdx = headers.indexOf('aid');
  var taskIdIdx = headers.indexOf('task_id');
  var statusIdx = headers.indexOf('status');
  var approvedByAid = {};
  data.slice(1).forEach(function(r) {
    if (String(r[statusIdx]).trim() === 'approved') {
      var aid = String(r[aidIdx]).trim();
      var taskId = String(r[taskIdIdx]).trim();
      if (!approvedByAid[aid]) approvedByAid[aid] = {};
      approvedByAid[aid][taskId] = true;
    }
  });
  var taskPtsMap = {};
  rows('tasks').filter(function(t){ return t['id']; }).forEach(function(t) {
    taskPtsMap[String(t['id']).trim()] = Number(t['pts']) || 0;
  });
  var result = {};
  Object.keys(approvedByAid).forEach(function(aid) {
    var total = 0;
    Object.keys(approvedByAid[aid]).forEach(function(tid) { total += taskPtsMap[tid] || 0; });
    result[aid] = total;
  });
  return result;
}

function getLeaderboard() {
  var c = getConfig();
  var nicks = nickMap();
  var taskPtsMap = getAllTaskPoints();
  return affiliateRows().map(function(a) {
    var pts = calcPoints(a, c) + (taskPtsMap[a.AID] || 0);
    return { aid: a.AID, nick: nicks[a.AID] || ('AID ' + a.AID), points: pts, eftd: a.eFTD, level: levelFor(pts, c), isNewcomer: isNewcomer(a.approvedDate) };
  }).sort(function(x, y) { return y.points - x.points; });
}

function getStats(aid) {
  var aidStr = String(aid).trim();
  var c = getConfig();
  var affRows = affiliateRows();
  var taskPtsMap = getAllTaskPoints();
  var a = affRows.filter(function(x){ return x.AID === aidStr; })[0];
  if (!a) return null;
  var volUnit = c.vol_unit_size || 100000;
  var taskPts = taskPtsMap[aidStr] || 0;
  var pts = calcPoints(a, c) + taskPts;
  var lb = affRows.map(function(x){
    return { aid: x.AID, points: calcPoints(x, c) + (taskPtsMap[x.AID] || 0) };
  }).sort(function(x, y){ return y.points - x.points; });
  var rank = 1;
  for (var i = 0; i < lb.length; i++) { if (lb[i].aid === aidStr) { rank = i + 1; break; } }
  return {
    aid: aidStr, points: pts, eftd: a.eFTD, ftd: a.ftd, regs: a.reg, vol: a.vol,
    earned: a.eFTD * (c.usd_per_eftd || 10),
    level: levelFor(pts, c), rank: rank, total: lb.length,
    isNewcomer: isNewcomer(a.approvedDate),
    approvedDate: String(a.approvedDate || ''),
    breakdown: {
      reg:  a.reg  * (c.points_per_reg  || 0),
      ftd:  a.ftd  * (c.points_per_ftd  || 0),
      eftd: a.eFTD * (c.points_per_eftd || 0),
      vol:  Math.floor(a.vol / volUnit) * (c.points_per_vol_unit || 0),
      tasks: taskPts
    },
    multipliers: {
      reg:  c.points_per_reg      || 0,
      ftd:  c.points_per_ftd      || 0,
      eftd: c.points_per_eftd     || 0,
      vol:  c.points_per_vol_unit || 0
    }
  };
}

function getPool() {
  var c = getConfig();
  var nicks = nickMap();
  var taskPtsMap = getAllTaskPoints();
  var base = c.pool_base || 500;
  var perEftd = c.pool_per_eftd || 10;
  var totalEftd = 0;
  var bestNewcomer = null;
  var bestNewcomerPts = -1;
  affiliateRows().forEach(function(a) {
    totalEftd += a.eFTD;
    if (isNewcomer(a.approvedDate)) {
      var pts = calcPoints(a, c) + (taskPtsMap[a.AID] || 0);
      if (pts > bestNewcomerPts) {
        bestNewcomerPts = pts;
        bestNewcomer = { aid: a.AID, nick: nicks[a.AID] || ('AID ' + a.AID), points: pts };
      }
    }
  });
  return { base: base, eftd: totalEftd, topup: totalEftd * perEftd, total: base + totalEftd * perEftd, bestNewcomer: bestNewcomer };
}

function getTasks() {
  var today = new Date(); today.setHours(0,0,0,0);
  return rows('tasks').filter(function(t){ return t['id']; }).map(function(t) {
    var dl = t['deadline'];
    var dlDate = dl ? new Date(dl) : null;
    var dlPassed = dlDate && !isNaN(dlDate) ? dlDate < today : false;
    var dlStr = dlDate && !isNaN(dlDate)
      ? dlDate.toLocaleDateString('ru-RU', {day:'numeric', month:'long'})
      : String(dl || '');
    return {
      id: t['id'], status: t['status'] || 'open',
      title: { ru: t['title_ru'], kz: t['title_kz'], en: t['title_en'] || t['title_ru'] },
      desc:  { ru: t['desc_ru'],  kz: t['desc_kz'],  en: t['desc_en']  || t['desc_ru']  },
      pts: Number(t['pts']) || 0, cash: Number(t['cash']) || 0,
      threshold: Number(t['threshold']) || 0,
      deadline: dlStr, deadlinePassed: dlPassed
    };
  });
}

function submitLink(aid, taskId, link, social, views) {
  var sh = sheet('submissions');
  if (!sh) return;
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  sh.appendRow([String(aid).trim(), String(taskId).trim(), link, social, Number(views) || 0, now, 'pending', '']);

  var taskTitle = String(taskId).trim();
  try {
    var taskRow = rows('tasks').filter(function(t){ return String(t['id']).trim() === taskTitle; })[0];
    if (taskRow && taskRow['title_ru']) taskTitle = taskRow['title_ru'];
  } catch(e) {}
  notifyManager(
    '📋 Новый сабмит таска\n\n' +
    'AID: ' + String(aid).trim() + '\n' +
    'Таск: ' + taskTitle + '\n' +
    'Соцсеть: ' + social + '\n' +
    'Просмотры: ' + (Number(views) || 0) + '\n' +
    'Ссылка: ' + link
  );
}

function getSubmissions(aid) {
  var sh = sheet('submissions');
  if (!sh) return [];
  var data = sh.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0];
  return data.slice(1).filter(function(r) {
    var i = headers.indexOf('aid');
    return String(r[i]).trim() === String(aid).trim();
  }).map(function(r) {
    function val(h) { var i = headers.indexOf(h); return i >= 0 ? r[i] : ''; }
    return {
      taskId: String(val('task_id')).trim(),
      link: String(val('link') || ''),
      social: String(val('social') || ''),
      views: Number(val('views')) || 0,
      submittedAt: String(val('submitted_at') || ''),
      status: String(val('status') || 'pending'),
      comment: String(val('comment') || '')
    };
  });
}

function getByUserId(userId) {
  var data = rows('mapping');
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (String(row['user_id']).trim() === String(userId).trim()) {
      return { aid: String(row['aid']).trim(), nickname: String(row['nickname']).trim() };
    }
  }
  return null;
}

function notifyManager(text) {
  var token = PropertiesService.getScriptProperties().getProperty('BOT_TOKEN');
  if (!token) return;
  try {
    UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: '8660141727', text: text })
    });
  } catch(e) {}
}

function getLastUpdated() {
  var sh = sheet('config');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'last_updated') return String(data[i][1] || '');
  }
  return '';
}

function onEdit(e) {
  if (!e || !e.source) return;
  if (e.source.getActiveSheet().getName() !== 'raw_export') return;
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
  var sh = sheet('config');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'last_updated') {
      sh.getRange(i + 1, 2).setValue(now);
      return;
    }
  }
  sh.appendRow(['last_updated', now]);
}

// ── ANNOUNCEMENT ──────────────────────────────────────────────────────
// Run from the editor after updating raw_export. Copy the result from Logs.
function generateAnnouncement() {
  var lb = getLeaderboard();
  var pool = getPool();
  var prizes = [0.40, 0.25, 0.15];
  var medals = ['🥇', '🥈', '🥉'];
  var top = lb.slice(0, 3);
  var lines = [];
  top.forEach(function(p, i) {
    var prize = Math.round(pool.total * prizes[i]);
    lines.push(medals[i] + ' ' + p.nick + ' — ' + p.points + ' ұпай (~$' + prize + ')');
  });
  var daysLeft = daysLeftInMonth();
  var text =
    '📊 Affiliate Champions — ағымдағы рейтинг\n\n' +
    lines.join('\n') +
    '\n\nЖүлде қоры: $' + pool.total +
    ' · Айдың соңына ' + daysLeft + ' күн қалды\n' +
    '👉 https://kuanyshadilkhan.github.io/affiliate-champion/';
  Logger.log(text);
}

function daysLeftInMonth() {
  var now = new Date();
  var last = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  return last.getDate() - now.getDate();
}

// ── DAILY REMINDER ────────────────────────────────────────────────────
// Sends a reminder to the manager at 10:00 to update raw_export.
function dailyReminder() {
  notifyManager('📋 Деректерді жаңартуды ұмытпаңыз — raw_export парағын жүктеп салыңыз.');
}

// Run ONCE from the editor to set up the daily trigger at 10:00.
function createDailyReminderTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyReminder') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyReminder')
    .timeBased()
    .atHour(10)
    .everyDays(1)
    .create();
  Logger.log('Trigger created: dailyReminder at 10:00');
}

function saveMapping(userId, aid, nick) {
  var sh = sheet('mapping');
  var data = sh.getDataRange().getValues();
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(userId).trim()) {
      sh.getRange(i + 1, 2).setValue(aid);
      sh.getRange(i + 1, 3).setValue(nick);
      sh.getRange(i + 1, 5).setValue(now); // last_seen_at
      return;
    }
  }
  sh.appendRow([userId, aid, nick, now, now]); // registered_at, last_seen_at
}

function pingUser(userId) {
  if (!userId) return;
  var sh = sheet('mapping');
  var data = sh.getDataRange().getValues();
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(userId).trim()) {
      sh.getRange(i + 1, 5).setValue(now); // last_seen_at
      return;
    }
  }
}
