/**
 * ระบบวันลาบุคลากร — Google Apps Script (ฝั่งเซิร์ฟเวอร์)
 * ผู้ใช้งาน = ชื่อ-นามสกุล, รหัสผ่านเริ่มต้น = 123456 (DEFAULT_PW) หรือเลขบัตรประชาชนถ้า HR กรอกไว้
 * เลขบัตรและรหัสผ่านเก็บแบบเข้ารหัสทางเดียว (salted SHA-256) ไม่เก็บตัวเลขจริง
 */

var SH = {
  STAFF: 'บุคลากร',
  LEAVES: 'การลา',
  CARRY: 'ยอดยกมา',
  TYPES: 'สิทธิ์การลา',
  HOLIDAYS: 'วันหยุด',
  SETTINGS: 'ตั้งค่า',
  ORG: 'กลุ่มงาน' // ไม่บังคับ: รายการกลุ่มงาน/งานของหน่วยงาน
};
var STAFF_HEAD = ['id', 'ชื่อ-สกุล', 'ตำแหน่ง', 'ระดับ', 'กลุ่มงาน', 'งาน', 'ประเภทบุคลากร', 'บทบาท',
  'เลขบัตรประชาชน (กรอกชั่วคราว)', 'salt', 'hash', 'ควรเปลี่ยนรหัสผ่าน', 'เข้าใช้ล่าสุด'];
var LEAVE_HEAD = ['id', 'staffId', 'ชื่อ-สกุล', 'ประเภทการลา', 'ฉุกเฉิน', 'ตั้งแต่', 'ถึง', 'ช่วงเวลา',
  'วันทำการ', 'ปีงบประมาณ', 'หมายเหตุ', 'เลขที่ใบลา', 'บันทึกโดย', 'เวลาบันทึก',
  'สถานะ', 'เหตุผลการยกเลิก', 'ยกเลิกโดย', 'เวลายกเลิก']; // คอลัมน์ 15–18: ขอยกเลิกวันลา (ว่าง = ปกติ)
var CANCELLED = 'ยกเลิก';
var CARRY_HEAD = ['staffId', 'ชื่อ-สกุล', 'ปีงบประมาณ', 'ลาป่วยยกมา', 'ลากิจยกมา', 'ลาพักผ่อนยกมา'];
var TYPES_HEAD = ['รหัส', 'ประเภทบุคลากร', 'ลาป่วย/ปี', 'ลากิจ/ปี', 'ลาพักผ่อน/ปี', 'ลาพักผ่อนสะสมรวมไม่เกิน'];
var HOL_HEAD = ['วันที่ (ปปปป-ดด-วว ค.ศ.)', 'ชื่อวันหยุด'];

var DEFAULT_TYPES = [
  ['gov', 'ข้าราชการ', 60, 45, 10, 20],
  ['emp', 'พนักงานราชการ', 30, 10, 10, 15],
  ['moph', 'พนักงานกระทรวงสาธารณสุข', 30, 10, 10, 15],
  ['temp', 'ลูกจ้างชั่วคราว', 30, 10, 10, 15]
];
var DEFAULT_HOLIDAYS = []; // HR เพิ่มวันหยุดราชการเองในหน้าตั้งค่า หรือในชีต วันหยุด
var KIND_CODES = { sick: 1, personal: 1, vacation: 1 };
var KIND_ALIASES = { 'ลาป่วย': 'sick', 'ลากิจ': 'personal', 'ลากิจส่วนตัว': 'personal', 'ลาพักผ่อน': 'vacation' };
var DEFAULT_PW = '123456'; // รหัสผ่านเริ่มต้นของทุกคน (ไม่ใช้เลขบัตรประชาชน) — ระบบแนะนำให้เปลี่ยนหลังเข้าครั้งแรก
var SESSION_SECONDS = 6 * 60 * 60;
var MAX_FAILS = 5;

// ---------------------------------------------------------------- เมนูในชีต
function onOpen() {
  SpreadsheetApp.getUi().createMenu('ระบบวันลา')
    .addItem('1) ตั้งค่าเริ่มต้นระบบ (ทำครั้งแรกครั้งเดียว)', 'setupSystem')
    .addItem('2) ตั้งรหัสผ่านให้บุคลากรที่วางลงชีต (เริ่มต้น 123456)', 'menuHashPending')
    .addItem('3) นำเข้าเลขบัตรจากไฟล์รายชื่อบุคลากร', 'menuImportCids')
    .addSeparator()
    .addItem('ดูลิงก์เข้าใช้งานระบบ', 'menuShowUrl')
    .addSeparator()
    .addItem('ตั้งค่าอัปเดตอัตโนมัติจาก GitHub', 'menuSetupAutoUpdate')
    .addItem('อัปเดตจาก GitHub ตอนนี้', 'menuUpdateNow')
    .addToUi();
}

// ---------------------------------------------------------------- อัปเดตโค้ดอัตโนมัติจาก GitHub
// ทุก 10 นาที ตรวจ commit ล่าสุดของ branch main ถ้าเปลี่ยน → ดึง Code.gs / Index.html / appsscript.json
// มาแทนโค้ดในโปรเจกต์นี้ สร้างเวอร์ชันใหม่ และชี้เว็บแอป (ลิงก์ /exec เดิม) ไปที่เวอร์ชันนั้น
// ต้องเปิด Google Apps Script API ที่ https://script.google.com/home/usersettings และใส่ GitHub token (อ่านอย่างเดียว)
var GH_REPO = 'nongplasawai1-blip/leave-system';
var GH_BRANCH = 'main';
var GH_FILES = [
  { path: 'apps-script/Code.gs', name: 'Code', type: 'SERVER_JS' },
  { path: 'apps-script/Index.html', name: 'Index', type: 'HTML' },
  { path: 'apps-script/appsscript.json', name: 'appsscript', type: 'JSON' }
];

function menuSetupAutoUpdate() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('ตั้งค่าอัปเดตอัตโนมัติจาก GitHub',
    'GitHub token (ไม่บังคับ — repo เป็น public ใช้ได้โดยไม่ต้องมี token)\n' +
    'ถ้ามี: Fine-grained token, repo leave-system, สิทธิ์ Contents: Read-only\nเว้นว่างแล้วกด OK = ไม่ใช้ token / ใช้ token เดิม', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var t = String(r.getResponseText() || '').trim();
  var p = PropertiesService.getScriptProperties();
  if (t) p.setProperty('GH_TOKEN', t);
  ScriptApp.getProjectTriggers().forEach(function (tr) { if (tr.getHandlerFunction() === 'autoUpdateTick') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('autoUpdateTick').timeBased().everyMinutes(10).create();
  p.deleteProperty('GH_SHA'); // บังคับอัปเดตรอบแรกทันที
  ui.alert('เปิดอัปเดตอัตโนมัติแล้ว (ตรวจทุก 10 นาที)\n\n' + updateMessage_(autoUpdate_(true)));
}

function menuUpdateNow() {
  PropertiesService.getScriptProperties().deleteProperty('GH_SHA');
  SpreadsheetApp.getUi().alert(updateMessage_(autoUpdate_(true)));
}

// เรียกจากตัวตั้งเวลา (time-driven trigger)
function autoUpdateTick() { autoUpdate_(false); }

function updateMessage_(r) {
  if (r.error) return 'อัปเดตไม่สำเร็จ: ' + r.error;
  if (r.pending) return 'พบโค้ดใหม่ จะอัปเดตในรอบถัดไป (ประมาณ 10 นาที)';
  if (!r.updated) return 'โค้ดเป็นเวอร์ชันล่าสุดแล้ว (' + r.label + ')';
  return 'อัปเดตแล้ว: ' + r.label + ' → เวอร์ชัน ' + r.version + ', เว็บแอปที่อัปเดต ' + r.deployments + ' ลิงก์';
}

function autoUpdate_(manual) {
  var p = PropertiesService.getScriptProperties();
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { error: 'ระบบกำลังทำงานอื่นอยู่ ลองใหม่อีกครั้ง' };
  try {
    var token = p.getProperty('GH_TOKEN') || '';
    var gh = ghLatest_(token, p.getProperty('GH_SHA'));
    var label = gh.sha ? 'commit ' + gh.sha.slice(0, 7) : 'โค้ดล่าสุดบน ' + GH_BRANCH;
    if (gh.key === p.getProperty('GH_SHA')) return { updated: false, label: label };
    // อ่านไฟล์แบบไม่ระบุ commit (GitHub จำกัดจำนวนครั้ง) แคชของ GitHub อาจยังเป็นไฟล์ต่างรุ่นกันได้ราว 5 นาที
    // → รอให้เห็นชุดเดิมซ้ำอีกรอบก่อนอัปเดต (ยกเว้นกดอัปเดตเองจากเมนู)
    if (!gh.sha && !manual && p.getProperty('GH_PENDING') !== gh.key) { p.setProperty('GH_PENDING', gh.key); return { pending: true, label: label }; }
    var sha = gh.key;
    var fresh = gh.files;
    // กันโค้ดเสียจนอัปเดตครั้งต่อไปไม่ได้: ต้องยังมีตัวอัปเดตและสิทธิ์ที่จำเป็น
    if (fresh[0].source.indexOf('function autoUpdateTick') < 0) throw new Error('Code.gs บน GitHub ไม่มีตัวอัปเดตอัตโนมัติ จึงไม่อัปเดต');
    if (fresh[2].source.indexOf('auth/script.projects') < 0 || fresh[2].source.indexOf('auth/script.deployments') < 0) throw new Error('appsscript.json บน GitHub ไม่มีสิทธิ์ script.projects/script.deployments จึงไม่อัปเดต');
    JSON.parse(fresh[2].source);

    var id = ScriptApp.getScriptId();
    var cur = scriptApi_('get', id + '/content').files || [];
    var keep = cur.filter(function (f) { return !GH_FILES.some(function (g) { return g.name === f.name; }); }); // ไฟล์อื่นในโปรเจกต์เก็บไว้
    scriptApi_('put', id + '/content', { files: fresh.concat(keep) });
    var desc = 'GitHub ' + GH_BRANCH + ' ' + (gh.sha ? gh.sha.slice(0, 7) : sha.slice(4, 11));
    var ver = scriptApi_('post', id + '/versions', { description: desc }).versionNumber;
    var n = 0;
    (scriptApi_('get', id + '/deployments').deployments || []).forEach(function (d) {
      var c = d.deploymentConfig || {};
      var web = (d.entryPoints || []).some(function (e) { return e.entryPointType === 'WEB_APP'; });
      if (!web || !c.versionNumber) return; // ข้าม HEAD (ลิงก์ทดสอบ /dev)
      scriptApi_('put', id + '/deployments/' + d.deploymentId,
        { deploymentConfig: { scriptId: id, versionNumber: ver, manifestFileName: 'appsscript', description: desc } });
      n++;
    });
    p.setProperty('GH_SHA', sha);
    p.setProperty('GH_LAST', JSON.stringify({ at: new Date().toISOString(), sha: sha, version: ver, deployments: n }));
    p.deleteProperty('GH_ERROR');
    p.deleteProperty('GH_PENDING');
    return { updated: true, label: label, version: ver, deployments: n };
  } catch (e) {
    var msg = String((e && e.message) || e);
    p.setProperty('GH_ERROR', new Date().toISOString() + ' ' + msg);
    if (!manual) console.error('autoUpdate: ' + msg);
    return { error: msg };
  } finally { lock.releaseLock(); }
}

// คืน { key, sha, files } — key ใช้เทียบว่าเคยอัปเดตชุดนี้แล้วหรือยัง
function ghLatest_(token, knownKey) {
  var c = ghFetch_('https://api.github.com/repos/' + GH_REPO + '/commits/' + GH_BRANCH, token, 'application/vnd.github.sha');
  var sha = c.code === 200 ? c.text.trim() : '';
  if (c.code === 401) throw new Error('GitHub token ใช้ไม่ได้หรือหมดอายุ (HTTP 401) — ตั้งค่าใหม่ที่เมนู ตั้งค่าอัปเดตอัตโนมัติจาก GitHub');
  if (!/^[0-9a-f]{40}$/.test(sha)) sha = ''; // เช่น 403/429 เกินโควตา GitHub แบบไม่มี token → อ่านไฟล์จาก branch แทน
  if (sha && sha === knownKey) return { key: sha, sha: sha, files: null }; // commit เดิม ไม่ต้องโหลดไฟล์
  var files = GH_FILES.map(function (f) {
    var url = 'https://raw.githubusercontent.com/' + GH_REPO + '/' + (sha || GH_BRANCH) + '/' + f.path;
    var r = ghFetch_(url, token, '*/*');
    if (r.code !== 200) throw new Error('ดึง ' + f.path + ' จาก GitHub ไม่ได้ (HTTP ' + r.code + ')' + (r.code === 404 && !token ? ' — ถ้า repo เป็น private ต้องใส่ GitHub token' : ''));
    return { name: f.name, type: f.type, source: r.text };
  });
  if (sha) return { key: sha, sha: sha, files: files };
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, files.map(function (f) { return f.source; }).join('\u0000'), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  return { key: 'raw:' + d, sha: '', files: files };
}

function ghFetch_(url, token, accept) {
  var h = { Accept: accept };
  if (token) h.Authorization = 'Bearer ' + token;
  if (url.indexOf('https://api.github.com/') === 0) h['X-GitHub-Api-Version'] = '2022-11-28';
  var r = UrlFetchApp.fetch(url, { headers: h, muteHttpExceptions: true });
  return { code: r.getResponseCode(), text: r.getContentText('UTF-8') };
}

function scriptApi_(method, path, body) {
  var o = { method: method, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
  if (body) { o.contentType = 'application/json'; o.payload = JSON.stringify(body); }
  var r = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/' + path, o);
  var code = r.getResponseCode(), txt = r.getContentText();
  if (code !== 200) {
    var m = txt; try { m = JSON.parse(txt).error.message; } catch (e) {}
    if (code === 403) m = 'ยังไม่ได้เปิด Google Apps Script API ที่ https://script.google.com/home/usersettings หรือยังไม่ได้อนุญาตสิทธิ์ (' + m + ')';
    else if (/versions$/.test(path)) m += ' — ถ้าเวอร์ชันครบ 200 แล้ว ให้ลบเวอร์ชันเก่าใน Apps Script → ประวัติโปรเจกต์';
    throw new Error('Apps Script API ' + method.toUpperCase() + ' ' + path.replace(/^[^/]+/, '') + ' (HTTP ' + code + '): ' + m);
  }
  return txt ? JSON.parse(txt) : {};
}

function setupSystem() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SH.STAFF, STAFF_HEAD, [6, 7, 9]);
  ensureHead_(ensureSheet_(ss, SH.LEAVES, LEAVE_HEAD, [6, 7, 11, 12, 16]), LEAVE_HEAD);
  ensureSheet_(ss, SH.CARRY, CARRY_HEAD, []);
  var t = ensureSheet_(ss, SH.TYPES, TYPES_HEAD, []);
  if (t.getLastRow() < 2) t.getRange(2, 1, DEFAULT_TYPES.length, 6).setValues(DEFAULT_TYPES);
  var h = ensureSheet_(ss, SH.HOLIDAYS, HOL_HEAD, [1]);
  if (h.getLastRow() < 2 && DEFAULT_HOLIDAYS.length) h.getRange(2, 1, DEFAULT_HOLIDAYS.length, 2).setValues(DEFAULT_HOLIDAYS);
  var s = ensureSheet_(ss, SH.SETTINGS, ['ค่า', 'ข้อมูล'], []);
  if (s.getLastRow() < 2) s.getRange(2, 1, 1, 2).setValues([['startFY', 2570]]);
  pepper_();
  var sheet1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (sheet1 && ss.getSheets().length > 1 && sheet1.getLastRow() === 0) ss.deleteSheet(sheet1);

  if (staffRows_().filter(function (r) { return r.role === 'HR'; }).length === 0) {
    var n = ui.prompt('สร้างบัญชีผู้ดูแล (HR) คนแรก', 'กรอกชื่อ-นามสกุลของคุณ (เช่น นางสาวสุดา ใจงาม)', ui.ButtonSet.OK_CANCEL);
    if (n.getSelectedButton() !== ui.Button.OK) return;
    var c = ui.prompt('สร้างบัญชีผู้ดูแล (HR) คนแรก', 'รหัสผ่านเริ่มต้น: เว้นว่างแล้วกด OK = ' + DEFAULT_PW + ' (หรือกรอกเลขบัตรประชาชน 13 หลักเพื่อใช้แทน)', ui.ButtonSet.OK_CANCEL);
    if (c.getSelectedButton() !== ui.Button.OK) return;
    var cid = digits_(c.getResponseText());
    if (cid && !validCid_(cid)) { ui.alert('เลขบัตรประชาชนไม่ถูกต้อง กรุณาเรียกเมนูนี้ใหม่อีกครั้ง'); return; }
    var salt = Utilities.getUuid();
    appendStaff_({ name: n.getResponseText().trim(), position: '', level: '', group: 'กลุ่มงานบริหารทรัพยากรบุคคล', section: '', ptype: 'gov', role: 'HR' }, salt, hash_(salt, cid || DEFAULT_PW));
  }
  ui.alert('ตั้งค่าเสร็จแล้ว (รหัสผ่านเริ่มต้นของบุคลากร = ' + DEFAULT_PW + ')\n\nขั้นต่อไป: เมนู ส่วนขยาย → Apps Script → ปุ่ม "ทำให้ใช้งานได้" (Deploy) → การทำให้ใช้งานได้รายการใหม่ → เว็บแอป\n• เรียกใช้ในฐานะ: ฉัน\n• ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน');
}

function menuHashPending() {
  var n = withLock_(hashPending_);
  SpreadsheetApp.getUi().alert(n ? ('ตั้งรหัสผ่านเริ่มต้นแล้ว ' + n + ' คน\nคนที่ไม่ได้กรอกเลขบัตรใช้รหัสผ่าน ' + DEFAULT_PW + '\n(ถ้ามีการกรอกเลขบัตร ระบบลบออกจากชีตแล้ว)') : 'ไม่มีบุคลากรที่รอตั้งรหัสผ่าน');
}

// ไฟล์ "รายชื่อบุคลากร 39 คน (สำหรับนำเข้าระบบวันลา)" — เลขบัตรอ่านตรงจากไฟล์ ไม่ผ่านคนกลาง และไม่เก็บตัวเลขจริง
var CID_SOURCE_ID = '1exDRZW-j2zEnAFpQY2NRwFrKAdTjFh4T_Bm3OELkzd4';

function menuImportCids() {
  var ui = SpreadsheetApp.getUi();
  var p = ui.prompt('นำเข้าเลขบัตรประชาชน',
    'วางลิงก์ Google Sheet ที่มีชื่อ-สกุลและเลขบัตรประชาชน\n(เว้นว่าง = ไฟล์ "รายชื่อบุคลากร 39 คน")\n\n' +
    'รหัสผ่านเริ่มต้นจะเป็นเลขบัตรประชาชนตามไฟล์ ยกเว้นคนที่ตั้งรหัสผ่านเองแล้วจะไม่ถูกเปลี่ยน', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  var t = String(p.getResponseText() || '').trim();
  var m = t.match(/\/d\/([A-Za-z0-9_-]+)/);
  var src = SpreadsheetApp.openById(m ? m[1] : (t || CID_SOURCE_ID));
  var r = withLock_(function () { return importCids_(src.getSheets().map(function (s) { return s.getDataRange().getValues(); })); });
  if (!r.rows) {
    ui.alert('ไม่พบเลขบัตรประชาชน (ตัวเลข 13 หลัก) ในไฟล์ "' + src.getName() + '"\n\n' +
      'ตรวจสอบว่าเลือกไฟล์ถูก และมีคอลัมน์เลขบัตรประชาชนครบ 13 หลัก\nยังไม่มีการเปลี่ยนแปลงรหัสผ่านใด ๆ');
    return;
  }
  var msg = 'ไฟล์: ' + src.getName() + ' (พบเลข 13 หลัก ' + r.rows + ' แถว)\n\n' +
    'ตั้งรหัสผ่านเริ่มต้น (เลขบัตร) แล้ว ' + r.done + ' คน\n' +
    'ตั้งรหัสผ่านเองแล้ว (ไม่เปลี่ยน) ' + r.hasPw + ' คน';
  if (r.bad.length) msg += '\n\nเลขบัตรไม่ถูกต้อง ' + r.bad.length + ' คน:\n• ' + r.bad.join('\n• ');
  if (r.dup.length) msg += '\n\nพบชื่อซ้ำหลายแถวในไฟล์ (ไม่ได้ตั้ง) ' + r.dup.length + ' คน:\n• ' + r.dup.join('\n• ');
  if (r.noCid.length) msg += '\n\nไม่พบชื่อ/เลขบัตรในไฟล์ ' + r.noCid.length + ' คน:\n• ' + r.noCid.join('\n• ');
  if (r.unknown.length) msg += '\n\nชื่อในไฟล์ที่ไม่มีในระบบ ' + r.unknown.length + ' คน:\n• ' + r.unknown.join('\n• ');
  ui.alert(msg);
}

// tables = ค่าของแต่ละแท็บในไฟล์ต้นทาง
// 1) ถ้ามีหัวตาราง (ชื่อ / เลขบัตร) จับคู่ชื่อแบบตรงตัว
// 2) ไม่ว่าหัวตารางจะเขียนอย่างไร: หาแถวที่มีตัวเลข 13 หลัก แล้วจับคู่ชื่อบุคลากรที่ปรากฏในข้อความแถวนั้น
function importCids_(tables) {
  var found = {}, rows = [];
  tables.forEach(function (v) {
    var h = -1, c = null;
    for (var i = 0; i < Math.min(v.length, 10) && h < 0; i++) { c = cidCols_(v[i]); if (c) h = i; }
    for (var j = 0; j < v.length; j++) {
      var row = v[j], ci = -1;
      if (c && j > h) ci = c.cid;
      if (ci < 0 || digits_(row[ci]).length !== 13) {
        ci = -1;
        for (var k = 0; k < row.length && ci < 0; k++) if (digits_(row[k]).length === 13) ci = k;
      }
      if (ci < 0) continue;
      var cid = digits_(row[ci]);
      var text = row.filter(function (_, k) { return k !== ci; }).map(String).join(' ');
      var label = text.replace(/\s+/g, ' ').trim();
      if (c && j > h) {
        var col = function (k) { return k >= 0 ? String(row[k]).trim() : ''; };
        label = c.full >= 0 ? col(c.full) : (col(c.pre) + col(c.first) + ' ' + col(c.last));
        var key = normName_(label);
        if (key) found[key] = cid;
      }
      rows.push({ cid: cid, text: squash_(text), label: label.slice(0, 60), used: false });
    }
  });
  var sh = sheet_(SH.STAFF), staff = staffRows_();
  var r = { done: 0, hasPw: 0, bad: [], noCid: [], dup: [], unknown: [], rows: rows.length };
  staff.forEach(function (s) {
    var key = normName_(s.name), cid = found[key];
    var hits = rows.filter(function (x) { return x.cid === cid || (key.length >= 4 && x.text.indexOf(key) >= 0); });
    hits.forEach(function (x) { x.used = true; });
    if (s.hash && !s.mustChange) { r.hasPw++; return; } // ตั้งรหัสผ่านเองแล้ว ไม่เปลี่ยน
    if (!cid) {
      var ids = hits.map(function (x) { return x.cid; }).filter(function (x, i, a) { return a.indexOf(x) === i; });
      if (ids.length > 1) { r.dup.push(s.name); return; }
      cid = ids[0];
    }
    if (!cid) { r.noCid.push(s.name); return; }
    if (!validCid_(cid)) { r.bad.push(s.name); return; }
    var salt = Utilities.getUuid();
    sh.getRange(s.row, 9, 1, 4).setValues([['', salt, hash_(salt, cid), true]]);
    r.done++;
  });
  rows.forEach(function (x) { if (!x.used) r.unknown.push(x.label || '(ไม่มีชื่อ)'); });
  return r;
}
function squash_(s) {
  return String(s || '').normalize('NFKC').replace(/[​-‍﻿]/g, '').replace(/\s+/g, '').toLowerCase();
}
function cidCols_(row) {
  var c = { full: -1, pre: -1, first: -1, last: -1, cid: -1 };
  row.forEach(function (x, i) {
    var t = String(x).replace(/\s+/g, '');
    if (!t) return;
    if (/บัตร|ประชาชน|^cid$|^pid$/i.test(t)) { if (c.cid < 0) c.cid = i; }
    else if (/^(ชื่อ-?(นาม)?สกุล|ชื่อและ(นาม)?สกุล|ชื่อ\/สกุล)$/.test(t)) { if (c.full < 0) c.full = i; }
    else if (/^คำนำหน้า/.test(t)) c.pre = i;
    else if (/^(นาม)?สกุล$/.test(t)) c.last = i;
    else if (/^ชื่อ$/.test(t)) c.first = i;
  });
  if (c.cid < 0) return null;
  if (c.full < 0 && c.first < 0) return null;
  return c;
}

function menuShowUrl() {
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? ('ลิงก์ระบบ:\n' + url) : 'ยังไม่ได้ Deploy เป็นเว็บแอป (ดูคู่มือขั้นตอนที่ 4)');
}

// ---------------------------------------------------------------- เว็บแอป
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบวันลาบุคลากร')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------------------------------------------------------------- API: เข้าสู่ระบบ
function login(username, password) {
  var key = 'f_' + normName_(username).slice(0, 200);
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get(key) || 0);
  if (fails >= MAX_FAILS) throw new Error('ใส่รหัสผิดเกิน ' + MAX_FAILS + ' ครั้ง กรุณารอ 15 นาทีแล้วลองใหม่');
  if (staffRows_().some(function (r) { return r.pendingCid || !r.id || !r.hash; })) withLock_(hashPending_);
  repairLeaveIds_();
  var want = normName_(username);
  var pw = String(password || '');
  var cands = staffRows_().filter(function (r) { return normName_(r.name) === want && r.hash; });
  var hit = null;
  cands.forEach(function (r) { if (!hit && pwMatches_(r, pw)) hit = r; });
  if (!hit) {
    cache.put(key, String(fails + 1), 900);
    throw new Error('ชื่อ-นามสกุลหรือรหัสผ่านไม่ถูกต้อง');
  }
  cache.remove(key);
  var token = Utilities.getUuid() + Utilities.getUuid();
  var sess = { id: hit.id, role: hit.role, v: hit.salt };
  cache.put('s_' + token, JSON.stringify(sess), SESSION_SECONDS);
  withLock_(function () { setStaffCell_(findStaff_(hit.id).row, 13, new Date()); });
  return { token: token, mustChange: hit.mustChange, data: buildData_(sess) };
}

// รหัสเริ่มต้น (เลขบัตร) พิมพ์มีขีด/เว้นวรรคได้ แต่หลังเปลี่ยนรหัสแล้วต้องตรงทุกตัวอักษร
function pwMatches_(r, pw) {
  pw = String(pw || '');
  return hash_(r.salt, pw) === r.hash || (r.mustChange && hash_(r.salt, digits_(pw)) === r.hash);
}

function logout(token) { CacheService.getScriptCache().remove('s_' + token); return true; }

function bootstrap(token) { return { data: buildData_(session_(token)) }; }

function changePassword(token, oldPw, newPw) {
  var s = session_(token);
  newPw = String(newPw || '');
  if (newPw.length < 8) throw new Error('รหัสผ่านใหม่ต้องยาวอย่างน้อย 8 ตัวอักษร');
  if (/^\d{13}$/.test(newPw)) throw new Error('รหัสผ่านใหม่ห้ามเป็นตัวเลข 13 หลัก (เลขบัตรประชาชน)');
  var salt = Utilities.getUuid();
  withLock_(function () {
    var r = findStaff_(s.id); // หาแถวภายใน lock เพราะแถวอาจเลื่อนเมื่อมีการลบบุคลากร
    if (r.salt !== s.v) throw new Error('SESSION: รหัสผ่านถูกเปลี่ยนแล้ว กรุณาเข้าสู่ระบบใหม่');
    if (!pwMatches_(r, oldPw)) throw new Error('รหัสผ่านเดิมไม่ถูกต้อง');
    sheet_(SH.STAFF).getRange(r.row, 10, 1, 3).setValues([[salt, hash_(salt, newPw), false]]);
  });
  rebind_(token, s, salt);
  return { data: buildData_(s) };
}

// session ปัจจุบันใช้ต่อได้หลังเปลี่ยนรหัสของตัวเอง ส่วน session อื่นของบัญชีนี้หมดอายุ
function rebind_(token, s, salt) {
  s.v = salt;
  CacheService.getScriptCache().put('s_' + token, JSON.stringify(s), SESSION_SECONDS);
}

// ---------------------------------------------------------------- API: HR
function saveStaff(token, st) {
  var s = requireHR_(token);
  st = st || {};
  var name = String(st.name || '').trim();
  if (!name) throw new Error('กรอกชื่อ-สกุล');
  var ptype = typeCode_(st.ptype);
  var role = st.role === 'HR' ? 'HR' : 'บุคลากร';
  var cid = digits_(st.cid);
  if (cid && !validCid_(cid)) throw new Error('เลขบัตรประชาชนไม่ถูกต้อง ตรวจสอบอีกครั้ง');
  withLock_(function () {
    var rows = staffRows_();
    var dup = rows.filter(function (r) { return normName_(r.name) === normName_(name) && r.id !== st.id; });
    if (dup.length) throw new Error('มีชื่อ "' + name + '" อยู่ในระบบแล้ว ชื่อผู้ใช้ต้องไม่ซ้ำกัน');
    if (st.id) {
      var r = findStaff_(st.id, rows);
      if (r.id === s.id && role !== 'HR') throw new Error('เปลี่ยนบทบาทของตัวเองออกจาก HR ไม่ได้');
      var sh = sheet_(SH.STAFF);
      sh.getRange(r.row, 2, 1, 7).setValues([[name, st.position || '', st.level || '', st.group || '', st.section || '', ptype, role]]);
      if (cid) {
        var salt = Utilities.getUuid();
        sh.getRange(r.row, 10, 1, 3).setValues([[salt, hash_(salt, cid), true]]);
        if (r.id === s.id) rebind_(token, s, salt);
      }
      if (r.name !== name) renameInSheet_(SH.LEAVES, 2, 3, r.id, name), renameInSheet_(SH.CARRY, 1, 2, r.id, name);
    } else {
      var salt2 = Utilities.getUuid();
      var nid = appendStaff_({ name: name, position: st.position, level: st.level, group: st.group, section: st.section, ptype: ptype, role: role }, salt2, hash_(salt2, cid || DEFAULT_PW));
      // บุคลากรใหม่ไม่มีวันลาพักผ่อนสะสมจากปีก่อน: บันทึกยอดยกมา 0 ไม่ให้ระบบคำนวณย้อนไปถึงปีเริ่มต้น
      var fyNow = fyOf_(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd'));
      if (fyNow > (Number(getSetting_('startFY')) || 2570)) sheet_(SH.CARRY).appendRow([nid, name, fyNow, '', '', 0]);
    }
  });
  return { data: buildData_(s) };
}

function deleteStaff(token, id) {
  var s = requireHR_(token);
  if (id === s.id) throw new Error('ลบบัญชีของตัวเองไม่ได้');
  withLock_(function () {
    deleteWhere_(SH.LEAVES, 2, id);
    deleteWhere_(SH.CARRY, 1, id);
    deleteWhere_(SH.STAFF, 1, id);
  });
  return { data: buildData_(s) };
}

function resetPassword(token, id, cid) {
  var s = requireHR_(token);
  cid = digits_(cid); // ไม่ระบุ = กลับเป็นรหัสผ่านเริ่มต้น DEFAULT_PW
  if (cid && !validCid_(cid)) throw new Error('เลขบัตรประชาชนไม่ถูกต้อง');
  withLock_(function () {
    var r = findStaff_(id);
    var salt = Utilities.getUuid();
    sheet_(SH.STAFF).getRange(r.row, 10, 1, 3).setValues([[salt, hash_(salt, cid || DEFAULT_PW), true]]);
    if (r.id === s.id) rebind_(token, s, salt);
  });
  return { data: buildData_(s) };
}

function addLeave(token, l) {
  var s = requireHR_(token);
  if (!KIND_CODES[l.kind]) throw new Error('ประเภทการลาไม่ถูกต้อง');
  if (!isIso_(l.start) || !isIso_(l.end) || l.end < l.start) throw new Error('วันที่ลาไม่ถูกต้อง');
  var days = Number(l.days);
  if (!(days > 0) || days > 366) throw new Error('จำนวนวันทำการไม่ถูกต้อง');
  withLock_(function () {
    var r = findStaff_(l.staffId);
    var me = findStaff_(s.id);
    var sh = sheet_(SH.LEAVES), row = sh.getLastRow() + 1;
    if (row > sh.getMaxRows()) sh.insertRowAfter(sh.getMaxRows());
    // วันที่ เลขที่ใบลา และหมายเหตุ เก็บเป็นข้อความ (กัน 12/2569 กลายเป็นวันที่ หรือ = กลายเป็นสูตร)
    ['F', 'G', 'K', 'L'].forEach(function (c) { sh.getRange(c + row).setNumberFormat('@'); });
    sh.getRange(row, 1, 1, 14).setValues([[Utilities.getUuid(), r.id, r.name, l.kind, !!l.emergency, l.start, l.end,
      l.start === l.end ? (l.half || 'full') : 'full', days, fyOf_(l.start), String(l.note || ''), String(l.docNo || ''), me.name, new Date()]]);
  });
  return { data: buildData_(s) };
}

function deleteLeave(token, leaveId) {
  var s = requireHR_(token);
  withLock_(function () { deleteWhere_(SH.LEAVES, 1, leaveId); });
  return { data: buildData_(s) };
}

// ขอยกเลิกวันลา: รายการยังอยู่ในประวัติ แต่ไม่นับเป็นวันลาที่ใช้ไป
// from/to (ไม่บังคับ) = ยกเลิกเฉพาะบางวัน: แยกรายการเดิมเป็น ช่วงก่อน / ช่วงที่ยกเลิก / ช่วงหลัง
// ช่วงที่ไม่ได้ยกเลิกยังนับเป็นวันลาตามวันทำการ (ตัดเสาร์-อาทิตย์และวันหยุดราชการ)
function cancelLeave(token, leaveId, reason, from, to) {
  var s = requireHR_(token);
  reason = String(reason || '').trim();
  if (!reason) throw new Error('กรุณาระบุเหตุผลการยกเลิก');
  withLock_(function () {
    var me = findStaff_(s.id);
    var sh = sheet_(SH.LEAVES), row = leaveRow_(sh, leaveId);
    ensureHead_(sh, LEAVE_HEAD);
    var cur = sh.getRange(row, 1, 1, 18).getValues()[0];
    if (isCancelled_(cur[14])) throw new Error('รายการนี้ยกเลิกไปแล้ว');
    var start = isoCell_(cur[5]), end = isoCell_(cur[6]);
    from = from ? String(from) : start; to = to ? String(to) : end;
    if (!isIso_(from) || !isIso_(to) || to < from) throw new Error('วันที่ยกเลิกไม่ถูกต้อง');
    if (from < start || to > end) throw new Error('วันที่ยกเลิกต้องอยู่ในช่วงวันลา ' + start + ' ถึง ' + end);
    var cancelVals = [CANCELLED, reason.slice(0, 500), me.name, new Date()];
    sh.getRange(row, 16).setNumberFormat('@'); // เหตุผลเก็บเป็นข้อความ (กัน = กลายเป็นสูตร)
    if (from === start && to === end) { sh.getRange(row, 15, 1, 4).setValues([cancelVals]); return; }

    var hol = {};
    rowsOf_(SH.HOLIDAYS).forEach(function (x) { var d = isoCell_(x[0]); if (d) hol[d] = 1; });
    var before = from > start ? { a: start, b: addDays_(from, -1) } : null;
    var after = to < end ? { a: addDays_(to, 1), b: end } : null;
    [before, after].forEach(function (p) { if (p) p.days = workDays_(p.a, p.b, hol); });
    var keep = [before, after].filter(function (p) { return p && p.days > 0; });
    var total = Number(cur[8]) || 0;
    var cancelDays = Math.round((total - keep.reduce(function (n, p) { return n + p.days; }, 0)) * 2) / 2;
    if (!(cancelDays > 0)) throw new Error('ช่วงวันที่เลือกไม่มีวันทำการให้ยกเลิก');
    // แถวเดิม = ช่วงที่ยกเลิก
    sh.getRange(row, 6, 1, 5).setValues([[from, to, 'full', cancelDays, fyOf_(from)]]);
    sh.getRange(row, 15, 1, 4).setValues([cancelVals]);
    // ช่วงที่ยังลาอยู่ = รายการใหม่ (คัดลอกประเภท หมายเหตุ เลขที่ใบลา ผู้บันทึกเดิม)
    keep.forEach(function (p) {
      var r = sh.getLastRow() + 1;
      if (r > sh.getMaxRows()) sh.insertRowAfter(sh.getMaxRows());
      ['F', 'G', 'K', 'L'].forEach(function (c) { sh.getRange(c + r).setNumberFormat('@'); });
      sh.getRange(r, 1, 1, 14).setValues([[Utilities.getUuid(), cur[1], cur[2], cur[3], cur[4], p.a, p.b, 'full', p.days,
        fyOf_(p.a), cur[10], cur[11], cur[12], cur[13]]]);
    });
  });
  return { data: buildData_(s) };
}
function addDays_(iso, n) {
  var d = new Date(Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10) + n));
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}
function workDays_(a, b, hol) {
  var n = 0;
  for (var d = a; d <= b; d = addDays_(d, 1)) {
    var wd = new Date(Date.UTC(+d.slice(0, 4), +d.slice(5, 7) - 1, +d.slice(8, 10))).getUTCDay();
    if (wd !== 0 && wd !== 6 && !hol[d]) n++;
  }
  return n;
}

// คืนสถานะรายการที่ยกเลิกผิด กลับมานับเป็นวันลาตามเดิม
function restoreLeave(token, leaveId) {
  var s = requireHR_(token);
  withLock_(function () {
    var sh = sheet_(SH.LEAVES), row = leaveRow_(sh, leaveId);
    if (!isCancelled_(sh.getRange(row, 15).getValue())) throw new Error('รายการนี้ไม่ได้ถูกยกเลิก');
    sh.getRange(row, 15, 1, 4).setValues([['', '', '', '']]);
  });
  return { data: buildData_(s) };
}

function saveCarry(token, staffId, fy, vals) {
  var s = requireHR_(token);
  fy = Number(fy);
  withLock_(function () {
    var r = findStaff_(staffId);
    var sh = sheet_(SH.CARRY);
    var data = sh.getDataRange().getValues();
    var row = -1;
    for (var i = 1; i < data.length; i++) if (String(data[i][0]) === staffId && Number(data[i][2]) === fy) { row = i + 1; break; }
    var v = function (x) { return (x === '' || x == null) ? '' : Number(x); };
    var out = [staffId, r.name, fy, v(vals.sick), v(vals.personal), v(vals.vacation)];
    var empty = out[3] === '' && out[4] === '' && out[5] === '';
    if (row > 0) { if (empty) sh.deleteRow(row); else sh.getRange(row, 1, 1, 6).setValues([out]); }
    else if (!empty) sh.appendRow(out);
  });
  return { data: buildData_(s) };
}

function saveTypes(token, types, startFY) {
  var s = requireHR_(token);
  withLock_(function () {
    var sh = sheet_(SH.TYPES);
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      var t = types[String(data[i][0])];
      if (!t) continue;
      sh.getRange(i + 1, 3, 1, 4).setValues([[Number(t.sick) || 0, Number(t.personal) || 0, Number(t.vacation) || 0, Number(t.carryCap) || 0]]);
    }
    if (startFY) setSetting_('startFY', Number(startFY));
  });
  return { data: buildData_(s) };
}

function addHoliday(token, d, n) {
  var s = requireHR_(token);
  if (!isIso_(d) || !String(n || '').trim()) throw new Error('ระบุวันที่และชื่อวันหยุด');
  withLock_(function () { deleteWhere_(SH.HOLIDAYS, 1, d); sheet_(SH.HOLIDAYS).appendRow([d, String(n).trim()]); });
  return { data: buildData_(s) };
}

function deleteHoliday(token, d) {
  var s = requireHR_(token);
  withLock_(function () { deleteWhere_(SH.HOLIDAYS, 1, d); });
  return { data: buildData_(s) };
}

function exportSheet(token, title, rows) {
  var s = requireHR_(token);
  var name = String(title || 'สรุปวันลา').replace(/[\[\]\*\?\/\\:]/g, ' ').slice(0, 90);
  if (Object.keys(SH).some(function (k) { return SH[k] === name.trim(); })) throw new Error('ชื่อชีตซ้ำกับชีตของระบบ');
  if (!rows || !rows.length) throw new Error('ไม่มีข้อมูลให้ส่งออก');
  var width = rows[0].length;
  rows = rows.map(function (r) { var x = r.slice(0, width); while (x.length < width) x.push(''); return x; });
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1, rows.length, width).setValues(rows);
  sh.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#d3f2e8');
  sh.setFrozenRows(1);
  return { sheet: name, data: buildData_(s) };
}

// ---------------------------------------------------------------- สร้างข้อมูลส่งให้หน้าเว็บ
function buildData_(s) {
  var staff = staffRows_();
  var me = null;
  staff.forEach(function (r) { if (r.id === s.id) me = r; });
  if (!me) throw new Error('SESSION: ไม่พบบัญชีของคุณ กรุณาเข้าสู่ระบบใหม่');
  if (me.salt !== s.v) throw new Error('SESSION: รหัสผ่านถูกเปลี่ยน กรุณาเข้าสู่ระบบใหม่');
  var isHR = me.role === 'HR';
  var pub = function (r) {
    var o = { id: r.id, name: r.name, position: r.position, level: r.level, group: r.group, section: r.section, ptype: r.ptype, role: r.role };
    if (isHR) { o.hasPassword = !!r.hash; o.mustChange = r.mustChange; o.lastLogin = r.lastLogin; }
    return o;
  };
  var leaves = rowsOf_(SH.LEAVES).map(function (x) {
    var k = String(x[3]).trim();
    return { id: String(x[0]), staffId: String(x[1]), kind: KIND_ALIASES[k] || k, emergency: x[4] === true || String(x[4]).toUpperCase() === 'TRUE',
      start: isoCell_(x[5]), end: isoCell_(x[6]), half: String(x[7] || 'full'), days: Number(x[8]) || 0, fy: Number(x[9]) || fyOf_(isoCell_(x[5])),
      note: String(x[10] || ''), docNo: String(x[11] || ''), recordedBy: String(x[12] || ''),
      cancelled: isCancelled_(x[14]), cancelReason: String(x[15] || ''), cancelledBy: String(x[16] || ''),
      cancelledAt: x[17] instanceof Date ? Utilities.formatDate(x[17], tz_(), 'yyyy-MM-dd') : isoCell_(x[17]) };
  });
  var carry = rowsOf_(SH.CARRY).map(function (x) {
    return { staffId: String(x[0]), fy: Number(x[2]), sick: x[3] === '' ? null : Number(x[3]), personal: x[4] === '' ? null : Number(x[4]), vacation: x[5] === '' ? null : Number(x[5]) };
  });
  var types = {};
  rowsOf_(SH.TYPES).forEach(function (x) {
    types[String(x[0])] = { label: String(x[1]), sick: Number(x[2]) || 0, personal: Number(x[3]) || 0, vacation: Number(x[4]) || 0, carryCap: Number(x[5]) || 0 };
  });
  var holidays = rowsOf_(SH.HOLIDAYS).map(function (x) { return { d: isoCell_(x[0]), n: String(x[1] || '') }; })
    .filter(function (h) { return h.d; });
  if (!isHR) {
    staff = [me];
    leaves = leaves.filter(function (l) { return l.staffId === me.id; });
    carry = carry.filter(function (c) { return c.staffId === me.id; });
  }
  return { me: pub(me), isHR: isHR, staff: staff.map(pub), leaves: leaves, carry: carry,
    config: { types: types, holidays: holidays, startFY: Number(getSetting_('startFY')) || 2570, org: orgRows_() } };
}

// แท็บ "กลุ่มงาน" (ไม่บังคับ): คอลัมน์ A = กลุ่มงาน, B = งานในกลุ่ม คั่นด้วยจุลภาค
// ไม่มีแท็บหรือแท็บว่าง = หน้าเว็บใช้รายการกลุ่มงานตั้งต้นในโค้ด
function orgRows_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SH.ORG);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .map(function (x) {
      return { g: String(x[0] || '').trim(),
        s: String(x[1] || '').split(/[,\n]/).map(function (t) { return t.trim(); }).filter(String) };
    })
    .filter(function (o) { return o.g; });
}

// ---------------------------------------------------------------- ตัวช่วย
function session_(token) {
  var raw = token && CacheService.getScriptCache().get('s_' + token);
  if (!raw) throw new Error('SESSION: หมดเวลาการใช้งาน กรุณาเข้าสู่ระบบใหม่');
  CacheService.getScriptCache().put('s_' + token, raw, SESSION_SECONDS);
  return JSON.parse(raw);
}
function requireHR_(token) {
  var s = session_(token);
  var r = findStaff_(s.id);
  if (r.salt !== s.v) throw new Error('SESSION: รหัสผ่านถูกเปลี่ยน กรุณาเข้าสู่ระบบใหม่');
  if (r.role !== 'HR') throw new Error('เฉพาะเจ้าหน้าที่ HR เท่านั้น');
  return s;
}
function sheet_(name) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('ยังไม่ได้ตั้งค่าระบบ: เปิดชีตแล้วเลือกเมนู ระบบวันลา → ตั้งค่าเริ่มต้นระบบ');
  return sh;
}
function rowsOf_(name) {
  var v = sheet_(name).getDataRange().getValues();
  return v.slice(1).filter(function (r) { return String(r[0]).trim() !== ''; });
}
function staffRows_() {
  var v = sheet_(SH.STAFF).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var x = v[i];
    if (!String(x[1]).trim()) continue;
    out.push({ row: i + 1, id: String(x[0]), name: String(x[1]).trim(), position: String(x[2] || ''), level: String(x[3] || ''),
      group: String(x[4] || ''), section: String(x[5] || ''), ptype: typeCode_(x[6]), role: String(x[7]).trim().toUpperCase() === 'HR' ? 'HR' : 'บุคลากร',
      pendingCid: digits_(x[8]), salt: String(x[9] || ''), hash: String(x[10] || ''),
      mustChange: x[11] === true || String(x[11]).toUpperCase() === 'TRUE', lastLogin: x[12] instanceof Date ? x[12].toISOString() : String(x[12] || '') });
  }
  return out;
}
function findStaff_(id, rows) {
  rows = rows || staffRows_();
  for (var i = 0; i < rows.length; i++) if (rows[i].id === id) return rows[i];
  throw new Error('ไม่พบบุคลากรในระบบ');
}
function appendStaff_(st, salt, hash) {
  var sh = sheet_(SH.STAFF), id = 's' + Utilities.getUuid().slice(0, 8);
  sh.appendRow([id, st.name, st.position || '', st.level || '', st.group || '', st.section || '',
    typeCode_(st.ptype), st.role === 'HR' ? 'HR' : 'บุคลากร', '', salt, hash, true, '']);
  return id;
}
function setStaffCell_(row, col, val) { sheet_(SH.STAFF).getRange(row, col).setValue(val); }
function hashPending_() {
  var sh = sheet_(SH.STAFF);
  var v = sh.getDataRange().getValues();
  var n = 0;
  for (var i = 1; i < v.length; i++) {
    if (!String(v[i][1]).trim()) continue;
    if (!String(v[i][0]).trim()) { sh.getRange(i + 1, 1).setValue('s' + Utilities.getUuid().slice(0, 8)); }
    if (!String(v[i][7]).trim()) sh.getRange(i + 1, 8).setValue('บุคลากร');
    var cid = digits_(v[i][8]);
    if (!cid) {
      // วางรายชื่อลงชีตโดยไม่มีเลขบัตร → ตั้งรหัสผ่านเริ่มต้น DEFAULT_PW ให้
      if (!String(v[i][10]).trim()) {
        var s0 = Utilities.getUuid();
        sh.getRange(i + 1, 10, 1, 3).setValues([[s0, hash_(s0, DEFAULT_PW), true]]);
        n++;
      }
      continue;
    }
    if (!validCid_(cid)) { sh.getRange(i + 1, 9).setNote('เลขบัตรไม่ถูกต้อง ตรวจสอบอีกครั้ง'); continue; }
    var salt = Utilities.getUuid();
    sh.getRange(i + 1, 9, 1, 4).setValues([['', salt, hash_(salt, cid), true]]);
    sh.getRange(i + 1, 9).clearNote();
    n++;
  }
  return n;
}
// ซ่อมรายการลา/ยอดยกมาที่ staffId เสีย (เช่น ถูกวางข้อความทับในชีต): ผูกกลับด้วยชื่อ-สกุล
// เฉพาะกรณีชื่อตรงกับบุคลากรคนเดียว (ชื่อซ้ำหรือไม่พบ = ไม่แตะ)
function repairLeaveIds_() {
  var staff = staffRows_(), ids = {}, byName = {};
  staff.forEach(function (r) {
    if (!r.id) return;
    ids[r.id] = 1;
    var k = normName_(r.name);
    byName[k] = byName[k] === undefined ? r.id : null;
  });
  var fix = function (name, idCol, nameCol) {
    var sh = sheet_(name), v = sh.getDataRange().getValues(), n = 0;
    for (var i = 1; i < v.length; i++) {
      var id = String(v[i][idCol - 1]).trim();
      if (ids[id] || !String(v[i][nameCol - 1]).trim()) continue;
      var nid = byName[normName_(v[i][nameCol - 1])];
      if (nid) { sh.getRange(i + 1, idCol).setValue(nid); n++; }
    }
    return n;
  };
  var bad = function (name, idCol, nameCol) {
    return sheet_(name).getDataRange().getValues().slice(1).some(function (x) {
      return String(x[nameCol - 1]).trim() && !ids[String(x[idCol - 1]).trim()] && byName[normName_(x[nameCol - 1])];
    });
  };
  if (!bad(SH.LEAVES, 2, 3) && !bad(SH.CARRY, 1, 2)) return 0;
  return withLock_(function () { return fix(SH.LEAVES, 2, 3) + fix(SH.CARRY, 1, 2); });
}
function deleteWhere_(name, col, val) {
  var sh = sheet_(name);
  var v = sh.getDataRange().getValues();
  for (var i = v.length - 1; i >= 1; i--) if (String(isoCellOrRaw_(v[i][col - 1])) === String(val)) sh.deleteRow(i + 1);
}
function leaveRow_(sh, leaveId) {
  var ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (var i = 1; i < ids.length; i++) if (String(ids[i][0]) === String(leaveId)) return i + 1;
  throw new Error('ไม่พบรายการลา (อาจถูกลบไปแล้ว) กรุณาโหลดหน้าใหม่');
}
function isCancelled_(v) { return String(v == null ? '' : v).trim() === CANCELLED; }
// เติมหัวคอลัมน์ที่ยังไม่มี (ชีตที่สร้างก่อนเพิ่มคอลัมน์ใหม่) โดยไม่แตะข้อมูลเดิม
function ensureHead_(sh, head) {
  var cur = sh.getRange(1, 1, 1, head.length).getValues()[0];
  for (var c = 0; c < head.length; c++) if (String(cur[c] || '').trim() === '') sh.getRange(1, c + 1).setValue(head[c]);
  return sh;
}
function renameInSheet_(name, idCol, nameCol, id, newName) {
  var sh = sheet_(name);
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][idCol - 1]) === id) sh.getRange(i + 1, nameCol).setValue(newName);
}
function ensureSheet_(ss, name, head, textCols) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#d3f2e8');
    sh.setFrozenRows(1);
  }
  (textCols || []).forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
  return sh;
}
function getSetting_(k) {
  var v = rowsOf_(SH.SETTINGS);
  for (var i = 0; i < v.length; i++) if (String(v[i][0]) === k) return v[i][1];
  return null;
}
function setSetting_(k, val) {
  var sh = sheet_(SH.SETTINGS);
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) if (String(v[i][0]) === k) { sh.getRange(i + 1, 2).setValue(val); return; }
  sh.appendRow([k, val]);
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('ระบบกำลังบันทึกข้อมูลอื่นอยู่ กรุณาลองใหม่');
  try { return fn(); } finally { lock.releaseLock(); }
}
function pepper_() {
  var p = PropertiesService.getScriptProperties();
  var v = p.getProperty('PEPPER');
  if (!v) { v = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('PEPPER', v); }
  return v;
}
function hash_(salt, secret) {
  var h = salt + '|' + String(secret) + '|' + pepper_();
  for (var i = 0; i < 200; i++) {
    h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h, Utilities.Charset.UTF_8)
      .map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  }
  return h;
}
function digits_(s) { return String(s == null ? '' : s).replace(/\D/g, ''); }
function validCid_(id) {
  if (!/^\d{13}$/.test(id)) return false;
  var sum = 0;
  for (var i = 0; i < 12; i++) sum += Number(id.charAt(i)) * (13 - i);
  return (11 - (sum % 11)) % 10 === Number(id.charAt(12));
}
function normName_(n) {
  // NFKC ให้ "ำ" กับ "ํา" ตรงกัน และตัดอักขระล่องหน (zero-width) ที่ติดมากับการคัดลอก
  return String(n || '').normalize('NFKC').replace(/[​-‍﻿]/g, '').replace(/\s+/g, ' ').trim()
    .replace(/^(นางสาว|นาย|นาง|น\.ส\.|ด\.ช\.|ด\.ญ\.|ดร\.|ว่าที่\s?ร\.?ต\.?(หญิง)?)\s*/, '')
    .replace(/\s+/g, '').toLowerCase();
}
function typeCode_(v) {
  var s = String(v || '').replace(/ํา/g, 'ำ').trim(); // "ํา" สองตัว → "ำ"
  var map = { 'ข้าราชการ': 'gov', 'พนักงานราชการ': 'emp', 'พนักงานกระทรวงสาธารณสุข': 'moph', 'พกส.': 'moph', 'ลูกจ้างชั่วคราว': 'temp' };
  if (map[s]) return map[s];
  if (['gov', 'emp', 'moph', 'temp'].indexOf(s) >= 0) return s;
  if (/^พนักงานกระทรวง|^พกส/.test(s)) return 'moph';
  if (/^พนักงานราชการ|^พนง\.?\s?ราชการ/.test(s)) return 'emp';
  if (/^ลูกจ้างประจำ/.test(s)) return 'gov'; // สิทธิ์ลาตามระเบียบกระทรวงการคลังใกล้เคียงข้าราชการ
  if (/^ลูกจ้าง/.test(s)) return 'temp';
  return 'gov';
}
function isIso_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')); }
function isoCell_(v) {
  var s = v instanceof Date ? Utilities.formatDate(v, tz_(), 'yyyy-MM-dd') : String(v || '').trim();
  // รับ วว/ดด/ปปปป ที่พิมพ์เองในชีตด้วย (ค.ศ. หรือ พ.ศ.)
  var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/);
  if (m) s = m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  if (!isIso_(s)) return '';
  var y = Number(s.slice(0, 4));
  return y > 2400 ? (y - 543) + s.slice(4) : s;
}
var TZ_;
function tz_() { return TZ_ || (TZ_ = SpreadsheetApp.getActive().getSpreadsheetTimeZone()); }
function isoCellOrRaw_(v) { return isoCell_(v) || v; }
function fyOf_(iso) { var y = Number(iso.slice(0, 4)), m = Number(iso.slice(5, 7)); return (m >= 10 ? y + 1 : y) + 543; }


// ---------------------------------------------------------------- API สำหรับหน้าเว็บบน GitHub Pages
// หน้าเว็บภายนอกส่ง POST (Content-Type: text/plain) มาที่ลิงก์ /exec เดิม เช่น {"fn":"login","args":["ชื่อ","รหัส"]}
// ตอบกลับเป็น JSON {ok:true,result:...} หรือ {ok:false,error:"ข้อความ"}
// ลิงก์ /exec เมื่อเปิดในเบราว์เซอร์ยังเข้า doGet แสดงหน้าเว็บเดิมเหมือนเดิม
function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var fn = apiFn_(String(req.fn || ''));
    if (!fn) throw new Error('ไม่รู้จักคำสั่ง: ' + req.fn);
    out = { ok: true, result: fn.apply(null, Array.isArray(req.args) ? req.args : []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
// เปิดให้เรียกจากภายนอกได้เฉพาะฟังก์ชันที่หน้าเว็บใช้ (ฟังก์ชันที่ลงท้ายด้วย _ เรียกไม่ได้)
function apiFn_(name) {
  var api = {
    login: login, logout: logout, bootstrap: bootstrap, changePassword: changePassword,
    saveStaff: saveStaff, deleteStaff: deleteStaff, resetPassword: resetPassword,
    addLeave: addLeave, deleteLeave: deleteLeave, cancelLeave: cancelLeave, restoreLeave: restoreLeave, saveCarry: saveCarry, saveTypes: saveTypes,
    addHoliday: addHoliday, deleteHoliday: deleteHoliday, exportSheet: exportSheet
  };
  return Object.prototype.hasOwnProperty.call(api, name) ? api[name] : null;
}
