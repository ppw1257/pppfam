// ════════════════════════════════════════════════════════════════════
// PPP SCHEDULE — Static JSON Cache Architecture
// Google Apps Script (Code.gs)
// ════════════════════════════════════════════════════════════════════

// ── ความปลอดภัย: ไม่เก็บ password plaintext ──
// ใส่ password ใน Script Properties แทน:
//   Apps Script → Project Settings → Script Properties
//   Key: ADMIN_PASSWORD   Value: ppp1257pondphuwinppw
//   Key: GITHUB_TOKEN     Value: ghp_xxxx
function getAdminPassword_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || '';
}

const TRIGGER_INTERVAL_MINUTES = 5;

// ── GitHub export config ──
const GITHUB_TOKEN  = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN') || '';
const GITHUB_OWNER  = 'ppw1257';
const GITHUB_REPO   = 'pppfam';
const GITHUB_BRANCH = 'main';

// ═══════════════════════════════════════════════════════════════════
// GITHUB PUSH
// ═══════════════════════════════════════════════════════════════════
function pushToGitHub_(filename, data) {
  if (!GITHUB_TOKEN) { Logger.log('GitHub push skipped: no GITHUB_TOKEN'); return false; }
  var url = 'https://api.github.com/repos/' + GITHUB_OWNER + '/' + GITHUB_REPO + '/contents/data/' + filename;
  var sha = '';
  try {
    var getRes = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'Authorization': 'token ' + GITHUB_TOKEN, 'Accept': 'application/vnd.github.v3+json' },
      muteHttpExceptions: true
    });
    if (getRes.getResponseCode() === 200) {
      sha = JSON.parse(getRes.getContentText()).sha || '';
    } else if (getRes.getResponseCode() === 401) {
      Logger.log('GitHub Token หมดอายุหรือ invalid — กรุณาสร้าง token ใหม่');
      return false;
    }
  } catch(e) { Logger.log('SHA fetch error: ' + e.message); }

  var body = { message: 'auto: update data/' + filename, branch: GITHUB_BRANCH };
  body.content = Utilities.base64Encode(JSON.stringify(data), Utilities.Charset.UTF_8);
  if (sha) body.sha = sha;

  try {
    var res = UrlFetchApp.fetch(url, {
      method: 'put',
      headers: {
        'Authorization': 'token ' + GITHUB_TOKEN,
        'Accept': 'application/vnd.github.v3+json',
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code === 200 || code === 201) {
      Logger.log('GitHub OK: ' + filename);
      return true;
    } else if (code === 422) {
      Logger.log('GitHub 422 (SHA conflict) [' + filename + '] — retrying without SHA...');
      // retry without SHA (conflict resolution)
      delete body.sha;
      var retry = UrlFetchApp.fetch(url, {
        method: 'put',
        headers: { 'Authorization': 'token ' + GITHUB_TOKEN, 'Accept': 'application/vnd.github.v3+json', 'Content-Type': 'application/json' },
        payload: JSON.stringify(body),
        muteHttpExceptions: true
      });
      Logger.log('Retry result: ' + retry.getResponseCode());
      return retry.getResponseCode() === 200 || retry.getResponseCode() === 201;
    } else {
      Logger.log('GitHub FAILED [' + filename + ']: ' + code + ' ' + res.getContentText().slice(0,200));
      return false;
    }
  } catch(e) {
    Logger.log('GitHub push exception [' + filename + ']: ' + e.message);
    return false;
  }
}

function pushAllToGitHub_(evData, wkData, bdData, anData) {
  if (!GITHUB_TOKEN) { Logger.log('GitHub push skipped: no GITHUB_TOKEN'); return; }
  var allOk = true;
  // push all.json ก่อน (สำคัญที่สุด — website ดึงไฟล์นี้)
  var ok = pushToGitHub_('all.json', {
    events:        evData,
    works:         wkData,
    birthdays:     bdData,
    anniversaries: anData,
    generated:     new Date().toISOString()
  });
  if (!ok) allOk = false;
  // push แต่ละไฟล์แยก (สำหรับ fallback)
  pushToGitHub_('events.json',        evData);
  pushToGitHub_('works.json',         wkData);
  pushToGitHub_('birthdays.json',     bdData);
  pushToGitHub_('anniversaries.json', anData);
  if (!allOk) Logger.log('WARNING: all.json push failed — website จะยังเห็นข้อมูลเก่า');
}

// ═══════════════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════════════
var WORKS_COLUMNS = [
  "id","category","title","artists","year",
  "image","type","role","brand","magazine","project",
  "release_date","contract_end","link","description"
];

function getSheetByName_(name) {
  var ss    = SpreadsheetApp.openById('1hf42HWcjwKuLx25O5CNbJm6X2-XOHwmHkaEKINRt5Uo');
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    if (name === "Works") {
      sheet = ss.insertSheet(name);
      sheet.getRange(1,1,1,WORKS_COLUMNS.length).setValues([WORKS_COLUMNS]);
      return sheet;
    }
    throw new Error("Sheet not found: " + name);
  }
  return sheet;
}

function generateId_() {
  return Utilities.getUuid().replace(/-/g, "").substring(0, 16);
}
function str_(v) {
  if (v === null || v === undefined) return "";
  var s = String(v).trim();
  return (s === "null" || s === "undefined") ? "" : s;
}
function safe_(v) {
  var s = str_(v);
  return /<[a-z][\s\S]*>/i.test(s) ? "" : s;
}
function cleanDate_(v) {
  if (!v) return "";
  if (v instanceof Date) {
    if (v.getFullYear() === 1899) return "";
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var s = str_(v);
  return s.startsWith("1899-12-30") ? "" : s;
}

function ensureEventColumns_(sheet) {
  var required = [
    "id","dateStart","dateEnd","startTime","endTime","timezone",
    "title","image","heroImage","poster","artists","country","city","location",
    "description","type","note","tv","map",
    "youtube","instagram","facebook","x","tiktok","website","whatsapp","shopee",
    "live","download","iq",
    "ticket1","ticket2","ticket3","ticketLive","ticketFanclub",
    "hashtags","keywords",
    "category","appearance","organizer","relatedWork","platform"
  ];
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){ return str_(h); });
  required.forEach(function(col) {
    if (headers.indexOf(col) === -1) {
      var nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(col);
      headers.push(col);
    }
  });
}

function ensureWorksColumns_(sheet) {
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){ return str_(h); });
  WORKS_COLUMNS.forEach(function(col) {
    if (headers.indexOf(col) === -1) {
      var nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue(col);
      headers.push(col);
    }
  });
}

function corsOutput_(json) {
  return ContentService
    .createTextOutput(JSON.stringify(json))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════════
// CACHE LAYER — PropertiesService
// ═══════════════════════════════════════════════════════════════════
function saveCache_(key, data) {
  var json = JSON.stringify(data);
  var CHUNK_SIZE = 8000;
  var chunks = [];
  for (var i = 0; i < json.length; i += CHUNK_SIZE) {
    chunks.push(json.slice(i, i + CHUNK_SIZE));
  }
  var props = PropertiesService.getScriptProperties();
  props.getKeys().filter(function(k){ return k.indexOf(key + '_chunk_') === 0; })
    .forEach(function(k){ props.deleteProperty(k); });
  chunks.forEach(function(chunk, idx) {
    props.setProperty(key + '_chunk_' + idx, chunk);
  });
  props.setProperty(key + '_chunks', String(chunks.length));
  props.setProperty(key + '_ts', String(Date.now()));
}

function loadCache_(key) {
  var props = PropertiesService.getScriptProperties();
  var chunkCount = parseInt(props.getProperty(key + '_chunks') || '0');
  if (!chunkCount) return null;
  var json = '';
  for (var i = 0; i < chunkCount; i++) {
    json += (props.getProperty(key + '_chunk_' + i) || '');
  }
  try { return JSON.parse(json); } catch(e) { return null; }
}

function getCacheTS_(key) {
  return parseInt(PropertiesService.getScriptProperties().getProperty(key + '_ts') || '0');
}

// ═══════════════════════════════════════════════════════════════════
// REGENERATE ALL CACHES + PUSH TO GITHUB
// ═══════════════════════════════════════════════════════════════════
function regenerateAllCaches() {
  // ── Lock flag: ป้องกัน trigger หลายตัวรันซ้อนกัน ──
  var props = PropertiesService.getScriptProperties();
  var lockTS = parseInt(props.getProperty('_regen_lock_ts') || '0');
  var now = Date.now();
  if (lockTS && (now - lockTS) < 4 * 60 * 1000) {
    // มี lock ภายใน 4 นาทีที่ผ่านมา — skip (อีก trigger กำลังรันอยู่)
    Logger.log('regenerateAllCaches() SKIPPED — lock active since ' + new Date(lockTS).toISOString());
    return;
  }
  props.setProperty('_regen_lock_ts', String(now));
  Logger.log('regenerateAllCaches() started');
  var evData = getEventsFromSheet_();
  var wkData = getWorksFromSheet_();
  var bdData = getBirthdaysFromSheet_();
  var anData = getAnniversariesFromSheet_();

  saveCache_('ppp_events',        { status:'ok', generated: new Date().toISOString(), events: evData });
  saveCache_('ppp_works',         { status:'ok', generated: new Date().toISOString(), works: wkData });
  saveCache_('ppp_birthdays',     { status:'ok', generated: new Date().toISOString(), birthdays: bdData });
  saveCache_('ppp_anniversaries', { status:'ok', generated: new Date().toISOString(), anniversaries: anData });
  saveCache_('ppp_all', {
    status:'ok', generated: new Date().toISOString(),
    events: evData, works: wkData, birthdays: bdData, anniversaries: anData
  });

  try {
    pushAllToGitHub_(evData, wkData, bdData, anData);
  } catch(e) {
    Logger.log('GitHub push error: ' + e.message);
  }

  // ปลด lock หลังเสร็จ
  props.deleteProperty('_regen_lock_ts');
  Logger.log('Done. events=' + evData.length + ' works=' + wkData.length +
             ' birthdays=' + bdData.length + ' anniversaries=' + anData.length);
}

// ═══════════════════════════════════════════════════════════════════
// WEB APP — doGet
// ═══════════════════════════════════════════════════════════════════
function doGet(e) {
  var action   = e && e.parameter && e.parameter.action ? e.parameter.action : "";
  var callback = e && e.parameter && e.parameter.callback ? e.parameter.callback : null;
  var result;

  try {
    switch (action) {
      case "getEvents":
        result = loadCache_('ppp_events');
        if (!result) { result = { status:'ok', events: getEventsFromSheet_() }; saveCache_('ppp_events', result); }
        break;
      case "getWorks":
        result = loadCache_('ppp_works');
        if (!result) { result = { status:'ok', works: getWorksFromSheet_() }; saveCache_('ppp_works', result); }
        break;
      case "getBirthdays":
        result = loadCache_('ppp_birthdays');
        if (!result) { result = { status:'ok', birthdays: getBirthdaysFromSheet_() }; saveCache_('ppp_birthdays', result); }
        break;
      case "getAnniversaries":
        result = loadCache_('ppp_anniversaries');
        if (!result) { result = { status:'ok', anniversaries: getAnniversariesFromSheet_() }; saveCache_('ppp_anniversaries', result); }
        break;
      case "getAll":
        // fresh=1: bypass PropertiesService cache — ดึงจาก Sheets โดยตรง
        // ใช้โดย _fetchFreshFromAppsScript หลัง admin save เพื่อป้องกันการ์ดขึ้นสองอัน
        var isFresh = e && e.parameter && e.parameter.fresh === '1';
        if (isFresh) {
          var evData2 = getEventsFromSheet_();
          var wkData2 = getWorksFromSheet_();
          var bdData2 = getBirthdaysFromSheet_();
          var anData2 = getAnniversariesFromSheet_();
          result = { status:'ok', generated: new Date().toISOString(),
            events: evData2, works: wkData2, birthdays: bdData2, anniversaries: anData2 };
          saveCache_('ppp_all', result);
        } else {
          result = loadCache_('ppp_all');
          if (!result) {
            var evData = getEventsFromSheet_();
            var wkData = getWorksFromSheet_();
            var bdData = getBirthdaysFromSheet_();
            var anData = getAnniversariesFromSheet_();
            result = { status:'ok', generated: new Date().toISOString(),
              events: evData, works: wkData, birthdays: bdData, anniversaries: anData };
            saveCache_('ppp_all', result);
          }
        }
        break;
      case "forceRegenerate":
        regenerateAllCaches();
        result = { status:'ok', message:'Cache regenerated' };
        break;
      default:
        result = { status:"ok", message:"PondPhuwinPermpoon API ready", generated: new Date().toISOString() };
    }
  } catch(err) {
    result = { status:"error", message: err.message };
  }

  if (!result) result = { status:'error', message:'No data' };

  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(result) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return corsOutput_(result);
}

// ═══════════════════════════════════════════════════════════════════
// WEB APP — doPost
// ═══════════════════════════════════════════════════════════════════
function doPost(e) {
  var result;
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action || "";

    // ── ตรวจ password จาก Script Properties แทน hardcode ──
    if (body.password && body.password !== getAdminPassword_()) {
      return corsOutput_({ status:"error", message:"Unauthorized" });
    }

    switch (action) {
      case "createEvent":       result = createEvent(body.event || {});              break;
      case "updateEvent":       result = updateEvent(body.id, body.event || {});     break;
      case "deleteEvent":       result = deleteEvent(body.id);                       break;
      case "deleteEventByRow":  result = deleteEventByRow(body.row);                 break;
      case "createBirthday":    result = createBirthday(body.birthday || {});        break;
      case "deleteBirthday":    result = deleteBirthday(body.id);                    break;
      case "createAnniversary": result = createAnniversary(body.anniversary || {});  break;
      case "deleteAnniversary": result = deleteAnniversary(body.id);                 break;
      case "createWork":        result = createWork(body.work || {});                break;
      case "updateWork":        result = updateWork(body.id, body.work || {});       break;
      case "deleteWork":        result = deleteWork(body.id);                        break;
      case "clearCache":        result = { status:"ok" };                            break;
      default: result = { status:"error", message:"Unknown action: " + action };
    }

    // regenerate cache + push GitHub ทุกครั้งที่ write สำเร็จ
    if (result && result.status === 'ok') {
      regenerateAllCaches();
    }

  } catch(err) {
    result = { status:"error", message: err.message };
  }
  return corsOutput_(result);
}

// ═══════════════════════════════════════════════════════════════════
// READ FROM SHEET
// ═══════════════════════════════════════════════════════════════════
function getEventsFromSheet_() {
  var sheet = getSheetByName_("Events");
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h){ return str_(h).toLowerCase().replace(/\s+/g,"_"); });
  var events = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i]; var obj = {};
    headers.forEach(function(h,j){ obj[h]=row[j]; });
    var ev = {
      id:            str_(obj.id||""),
      dateStart:     cleanDate_(obj.datestart||obj.date_start||obj.date||""),
      dateEnd:       cleanDate_(obj.dateend||obj.date_end||""),
      startTime:     str_(obj.starttime||obj.start_time||""),
      endTime:       str_(obj.endtime||obj.end_time||""),
      timezone:      safe_(obj.timezone||""),
      title:         safe_(obj.title||obj.event_name||""),
      image:         str_(obj.image||obj.image_url||""),
      heroImage:     str_(obj.heroimage||obj.hero_image||""),
      artists:       safe_(obj.artists||""),
      country:       safe_(obj.country||""),
      city:          safe_(obj.city||""),
      location:      safe_(obj.location||obj.location_name||""),
      description:   str_(obj.description||""),
      type:          safe_(obj.type||obj.event_type||""),
      note:          safe_(obj.note||""),
      tv:            safe_(obj.tv||obj.tv_channel||""),
      map:           str_(obj.map||obj.map_embed||""),
      youtube:       str_(obj.youtube||obj.youtube_link||""),
      instagram:     str_(obj.instagram||obj.instagram_link||""),
      facebook:      str_(obj.facebook||obj.facebook_link||""),
      x:             str_(obj.x||obj.x_link||""),
      tiktok:        str_(obj.tiktok||obj.tiktok_link||""),
      website:       str_(obj.website||""),
      whatsapp:      str_(obj.whatsapp||obj.whatsapp_link||""),
      shopee:        str_(obj.shopee||obj.shopee_link||""),
      live:          str_(obj.live||obj.live_link||""),
      download:      str_(obj.download||obj.file_link||""),
      iq:            str_(obj.iq||obj.iq_link||""),
      ticket1:       str_(obj.ticket1||obj.ticket_link1||""),
      ticket2:       str_(obj.ticket2||obj.ticket_link2||""),
      ticket3:       str_(obj.ticket3||obj.ticket_link3||""),
      ticketLive:    str_(obj.ticketlive||obj.ticket_live||""),
      ticketFanclub: str_(obj.ticketfanclub||obj.ticket_fanclub||""),
      hashtags:      safe_(obj.hashtags||obj.hashtag||""),
      keywords:      safe_(obj.keywords||obj.keyword||""),
      category:      safe_(obj.category||""),
      appearance:    safe_(obj.appearance||""),
      organizer:     safe_(obj.organizer||obj.brand||""),
      relatedWork:   safe_(obj.relatedwork||obj.related_work||""),
      poster:        str_(obj.poster||obj.poster_url||""),
      platform:      safe_(obj.platform||""),
    };
    if (ev.title) events.push(ev);
  }
  return events;
}

function getWorksFromSheet_() {
  var sheet = getSheetByName_("Works");
  ensureWorksColumns_(sheet);
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h){ return str_(h).toLowerCase().replace(/\s+/g,"_"); });
  var works = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i]; var obj = {};
    headers.forEach(function(h,j){ obj[h]=row[j]; });
    var w = {
      id:            str_(obj.id||""),
      category:      safe_(obj.category||""),
      title:         safe_(obj.title||""),
      artists:       safe_(obj.artists||""),
      year:          str_(obj.year||""),
      image:         str_(obj.image||""),
      description:   str_(obj.description||""),
      link:          str_(obj.link||""),
      type:          safe_(obj.type||""),
      role:          safe_(obj.role||""),
      brand:         safe_(obj.brand||""),
      magazine:      safe_(obj.magazine||""),
      project:       safe_(obj.project||""),
      release_date:  cleanDate_(obj.release_date||""),
      contract_end:  str_(obj.contract_end||""),
    };
    if (w.title) works.push(w);
  }
  return works;
}

function getBirthdaysFromSheet_() {
  var sheet = getSheetByName_("Birthdays");
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h){ return str_(h); });
  return data.slice(1).filter(function(r){ return r[0]; }).map(function(r) {
    var o={}; headers.forEach(function(h,j){ o[h]=r[j]; }); return o;
  });
}

function getAnniversariesFromSheet_() {
  var sheet = getSheetByName_("Anniversaries");
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h){ return str_(h); });
  return data.slice(1).filter(function(r){ return r[0]; }).map(function(r) {
    var o={}; headers.forEach(function(h,j){ o[h]=r[j]; }); return o;
  });
}

// ═══════════════════════════════════════════════════════════════════
// WRITE TO SHEET
// ═══════════════════════════════════════════════════════════════════
function createEvent(eventData) {
  var sheet   = getSheetByName_("Events");
  ensureEventColumns_(sheet);
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var id      = generateId_();
  var row = headers.map(function(h) {
    var key = str_(h).toLowerCase().replace(/\s+/g,"_");
    if (key === "id") return id;
    return str_(eventData[key] !== undefined ? eventData[key] : (eventData[str_(h)]||""));
  });
  sheet.appendRow(row);
  return { status:"ok", id:id };
}

function updateEvent(id, eventData) {
  var sheet   = getSheetByName_("Events");
  ensureEventColumns_(sheet);
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var idCol   = headers.map(function(h){ return str_(h).toLowerCase(); }).indexOf("id");
  if (idCol < 0) return { status:"error", message:"No id column" };
  for (var i = 1; i < data.length; i++) {
    if (str_(data[i][idCol]) === str_(id)) {
      var row = headers.map(function(h,j) {
        var key = str_(h).toLowerCase().replace(/\s+/g,"_");
        if (key === "id") return str_(id);
        var v = eventData[key] !== undefined ? eventData[key] : eventData[str_(h)];
        return v !== undefined ? str_(v) : str_(data[i][j]);
      });
      sheet.getRange(i+1,1,1,row.length).setValues([row]);
      return { status:"ok" };
    }
  }
  return { status:"error", message:"Not found: "+id };
}

function deleteEvent(id) {
  var sheet = getSheetByName_("Events");
  var data  = sheet.getDataRange().getValues();
  var idCol = data[0].map(function(h){ return str_(h).toLowerCase(); }).indexOf("id");
  if (idCol < 0) return { status:"error", message:"No id column" };
  for (var i = data.length-1; i >= 1; i--) {
    if (str_(data[i][idCol]) === str_(id)) { sheet.deleteRow(i+1); return { status:"ok" }; }
  }
  return { status:"error", message:"Not found" };
}

function deleteEventByRow(row) {
  var rowNum = parseInt(row);
  if (isNaN(rowNum) || rowNum < 2) return { status:"error", message:"Invalid row" };
  getSheetByName_("Events").deleteRow(rowNum);
  return { status:"ok" };
}

function createBirthday(data) {
  var id = generateId_();
  getSheetByName_("Birthdays").appendRow([id, str_(data.artist_name), str_(data.birthday_month), str_(data.birthday_day), str_(data.color||"#6b7280")]);
  return { status:"ok", id:id };
}
function deleteBirthday(id) {
  var sheet = getSheetByName_("Birthdays"); var data = sheet.getDataRange().getValues();
  for (var i=data.length-1; i>=1; i--) { if(str_(data[i][0])===str_(id)){ sheet.deleteRow(i+1); return {status:"ok"}; } }
  return { status:"error", message:"Not found" };
}

function createAnniversary(data) {
  var id = generateId_();
  getSheetByName_("Anniversaries").appendRow([id, str_(data.anniversary_name), str_(data.anniversary_month), str_(data.anniversary_day), str_(data.anniversary_year||""), str_(data.color||"#a855f7")]);
  return { status:"ok", id:id };
}
function deleteAnniversary(id) {
  var sheet = getSheetByName_("Anniversaries"); var data = sheet.getDataRange().getValues();
  for (var i=data.length-1; i>=1; i--) { if(str_(data[i][0])===str_(id)){ sheet.deleteRow(i+1); return {status:"ok"}; } }
  return { status:"error", message:"Not found" };
}

function createWork(data) {
  var sheet = getSheetByName_("Works");
  ensureWorksColumns_(sheet);
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var id = generateId_();
  var row = headers.map(function(h) {
    var key = str_(h).toLowerCase().replace(/\s+/g,"_");
    if (key === "id") return id;
    return str_(data[key] !== undefined ? data[key] : (data[str_(h)]||""));
  });
  sheet.appendRow(row);
  return { status:"ok", id:id };
}

function updateWork(id, data) {
  var sheet = getSheetByName_("Works");
  ensureWorksColumns_(sheet);
  var rows    = sheet.getDataRange().getValues();
  var headers = rows[0];
  var idCol   = headers.map(function(h){ return str_(h).toLowerCase(); }).indexOf("id");
  if (idCol < 0) return { status:"error", message:"No id column" };
  for (var i = 1; i < rows.length; i++) {
    if (str_(rows[i][idCol]) === str_(id)) {
      var row = headers.map(function(h,j) {
        var key = str_(h).toLowerCase().replace(/\s+/g,"_");
        if (key === "id") return str_(id);
        var v = data[key] !== undefined ? data[key] : data[str_(h)];
        return v !== undefined ? str_(v) : str_(rows[i][j]);
      });
      sheet.getRange(i+1,1,1,row.length).setValues([row]);
      return { status:"ok" };
    }
  }
  return { status:"error", message:"Not found: "+id };
}

function deleteWork(id) {
  var sheet = getSheetByName_("Works");
  var rows  = sheet.getDataRange().getValues();
  var idCol = rows[0].map(function(h){ return str_(h).toLowerCase(); }).indexOf("id");
  if (idCol < 0) return { status:"error", message:"No id column" };
  for (var i = rows.length-1; i >= 1; i--) {
    if (str_(rows[i][idCol]) === str_(id)) { sheet.deleteRow(i+1); return { status:"ok" }; }
  }
  return { status:"error", message:"Not found" };
}

// ═══════════════════════════════════════════════════════════════════
// TRIGGER SETUP — แก้ปัญหา trigger ซ้อนกัน (สาเหตุ 8,398 runs)
// ═══════════════════════════════════════════════════════════════════
function setupTriggers() {
  // ลบ trigger เก่าทั้งหมดก่อน (ป้องกัน trigger ซ้อนกัน)
  var triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === 'regenerateAllCaches') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log('Removed ' + removed + ' old triggers');

  // สร้าง trigger ใหม่แค่ 1 ตัวเท่านั้น
  ScriptApp.newTrigger('regenerateAllCaches')
    .timeBased()
    .everyMinutes(TRIGGER_INTERVAL_MINUTES)
    .create();
  Logger.log('Created 1 new trigger: every ' + TRIGGER_INTERVAL_MINUTES + ' min');

  // ตรวจสอบว่ามี trigger กี่ตัวตอนนี้
  var remaining = ScriptApp.getProjectTriggers().filter(function(t){
    return t.getHandlerFunction() === 'regenerateAllCaches';
  });
  Logger.log('Total regenerateAllCaches triggers now: ' + remaining.length);

  regenerateAllCaches();
  Logger.log('Initial cache + GitHub push done.');
}

function removeTriggers() {
  var count = 0;
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); count++; });
  Logger.log('Removed ' + count + ' triggers.');
}

function checkTriggerCount() {
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('Total triggers: ' + triggers.length);
  triggers.forEach(function(t){
    Logger.log('  - ' + t.getHandlerFunction() + ' | ' + t.getEventType());
  });
}

function checkCacheStatus() {
  ['ppp_events','ppp_works','ppp_birthdays','ppp_anniversaries','ppp_all'].forEach(function(key) {
    var ts = getCacheTS_(key);
    var age = ts ? ((Date.now()-ts)/60000).toFixed(1)+'min ago' : 'never';
    var data = loadCache_(key);
    Logger.log(key + ': ' + age + ', size=' + (data ? JSON.stringify(data).length : 0) + ' bytes');
  });
}

function checkGitHubToken() {
  if (!GITHUB_TOKEN) { Logger.log('ERROR: ไม่มี GITHUB_TOKEN ใน Script Properties'); return; }
  var res = UrlFetchApp.fetch('https://api.github.com/user', {
    headers: { 'Authorization': 'token ' + GITHUB_TOKEN },
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  if (code === 200) {
    var user = JSON.parse(res.getContentText());
    Logger.log('GitHub Token OK — logged in as: ' + user.login);
  } else if (code === 401) {
    Logger.log('GitHub Token หมดอายุหรือ invalid! กรุณาสร้าง token ใหม่');
  } else {
    Logger.log('GitHub Token check failed: ' + code);
  }
}

function authorizeAndTest() {
  var res = UrlFetchApp.fetch('https://api.github.com', { muteHttpExceptions: true });
  Logger.log('Auth test: ' + res.getResponseCode());
}

function authorizeMe() {
  UrlFetchApp.fetch('https://api.github.com', { muteHttpExceptions: true });
  Logger.log('Authorization OK');
}
