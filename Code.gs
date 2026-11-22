/**
 * हल्बा युवा शक्ति महोत्सव – 2026 · पुरस्कार सहयोग पंजीयन
 * Google Apps Script (Google Sheet से जुड़ा हुआ)
 *
 * यह स्क्रिप्ट वेबसाइट (index.html) से आई प्रविष्टियों को
 * "प्रविष्टियाँ" शीट में जोड़ती है और "पुरस्कार स्थिति" शीट में
 * चुने गए पुरस्कार पर सहयोगी का नाम लिखती है।
 */

var SHEET_ENTRIES = 'प्रविष्टियाँ';
var SHEET_SLOTS = 'पुरस्कार स्थिति';

/**
 * नई / खाली Google Sheet में पहली बार एक बार चलाएँ:
 * Apps Script में ऊपर फ़ंक्शन सूची से "setup" चुनें → ▶ Run
 * यह "प्रविष्टियाँ" और "पुरस्कार स्थिति" शीट अपने-आप बना देगा।
 * जो शीट पहले से हैं, उनका डेटा नहीं मिटेगा।
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Kolkata');
  var maroon = '#8a1c2b';

  var ent = ss.getSheetByName(SHEET_ENTRIES);
  if (!ent) {
    ent = ss.getSheets().length === 1 && ss.getSheets()[0].getLastRow() === 0
      ? ss.getSheets()[0].setName(SHEET_ENTRIES)
      : ss.insertSheet(SHEET_ENTRIES);
  }
  if (ent.getLastRow() === 0) {
    var heads = ["दिनांक-समय", "प्रविष्टि क्र.", "सहयोग का प्रकार", "नाम", "पद / व्यवसाय", "पिता / पति का नाम", "मोबाइल", "व्हाट्सऐप", "गाँव / शहर", "ब्लॉक / जिला", "पुरस्कार विवरण", "पुरस्कार कोड", "स्मृति में", "कुल राशि (₹)", "भुगतान का तरीका", "UTR / लेन-देन क्र.", "संपर्क पदाधिकारी", "नाम घोषणा की सहमति", "पूर्व सहयोग वर्ष", "राशि प्राप्त? (समिति भरे)"];
    ent.getRange(1, 1, 1, heads.length).setValues([heads])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground(maroon).setWrap(true);
    ent.setFrozenRows(1);
    ent.setColumnWidths(1, heads.length, 140);
    ent.setColumnWidth(11, 320);
    ent.getRange('T2:T1000').setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['हाँ', 'नहीं'], true).setAllowInvalid(true).build());
  }

  var sl = ss.getSheetByName(SHEET_SLOTS) || ss.insertSheet(SHEET_SLOTS);
  if (sl.getLastRow() === 0) {
    var rows = [
    ["s1-1", "खो-खो (बालक)", "प्रथम", 5001],
    ["s1-2", "खो-खो (बालक)", "द्वितीय", 3001],
    ["s1-3", "खो-खो (बालक)", "तृतीय", 2001],
    ["s2-1", "खो-खो (बालिका)", "प्रथम", 5001],
    ["s2-2", "खो-खो (बालिका)", "द्वितीय", 3001],
    ["s2-3", "खो-खो (बालिका)", "तृतीय", 2001],
    ["s3-1", "कबड्डी (पुरुष)", "प्रथम", 5001],
    ["s3-2", "कबड्डी (पुरुष)", "द्वितीय", 3001],
    ["s3-3", "कबड्डी (पुरुष)", "तृतीय", 2001],
    ["s4-1", "कबड्डी (महिला)", "प्रथम", 5001],
    ["s4-2", "कबड्डी (महिला)", "द्वितीय", 3001],
    ["s4-3", "कबड्डी (महिला)", "तृतीय", 2001],
    ["s5-1", "कबड्डी (बालक)", "प्रथम", 5001],
    ["s5-2", "कबड्डी (बालक)", "द्वितीय", 3001],
    ["s5-3", "कबड्डी (बालक)", "तृतीय", 2001],
    ["s6-1", "कबड्डी (बालिका)", "प्रथम", 5001],
    ["s6-2", "कबड्डी (बालिका)", "द्वितीय", 3001],
    ["s6-3", "कबड्डी (बालिका)", "तृतीय", 2001],
    ["s7-1", "रिले रेस (पुरुष)", "प्रथम", 3001],
    ["s7-2", "रिले रेस (पुरुष)", "द्वितीय", 2001],
    ["s7-3", "रिले रेस (पुरुष)", "तृतीय", 1501],
    ["s8-1", "रिले रेस (महिला)", "प्रथम", 3001],
    ["s8-2", "रिले रेस (महिला)", "द्वितीय", 2001],
    ["s8-3", "रिले रेस (महिला)", "तृतीय", 1501],
    ["s9-1", "रिले रेस (बालक)", "प्रथम", 3001],
    ["s9-2", "रिले रेस (बालक)", "द्वितीय", 2001],
    ["s9-3", "रिले रेस (बालक)", "तृतीय", 1501],
    ["s10-1", "रिले रेस (बालिका)", "प्रथम", 3001],
    ["s10-2", "रिले रेस (बालिका)", "द्वितीय", 2001],
    ["s10-3", "रिले रेस (बालिका)", "तृतीय", 1501],
    ["s11-1", "दौड़ 400 मीटर (पुरुष)", "प्रथम", 3001],
    ["s11-2", "दौड़ 400 मीटर (पुरुष)", "द्वितीय", 2001],
    ["s11-3", "दौड़ 400 मीटर (पुरुष)", "तृतीय", 1001],
    ["s12-1", "दौड़ 400 मीटर (महिला)", "प्रथम", 3001],
    ["s12-2", "दौड़ 400 मीटर (महिला)", "द्वितीय", 2001],
    ["s12-3", "दौड़ 400 मीटर (महिला)", "तृतीय", 1001],
    ["s13-1", "दौड़ 400 मीटर (बालक)", "प्रथम", 3001],
    ["s13-2", "दौड़ 400 मीटर (बालक)", "द्वितीय", 2001],
    ["s13-3", "दौड़ 400 मीटर (बालक)", "तृतीय", 1001],
    ["s14-1", "दौड़ 400 मीटर (बालिका)", "प्रथम", 3001],
    ["s14-2", "दौड़ 400 मीटर (बालिका)", "द्वितीय", 2001],
    ["s14-3", "दौड़ 400 मीटर (बालिका)", "तृतीय", 1001],
    ["c1-1", "सामूहिक नृत्य", "प्रथम", 5001],
    ["c1-2", "सामूहिक नृत्य", "द्वितीय", 3001],
    ["c1-3", "सामूहिक नृत्य", "तृतीय", 2001],
    ["c2-1", "सुवा नृत्य", "प्रथम", 5001],
    ["c2-2", "सुवा नृत्य", "द्वितीय", 3001],
    ["c2-3", "सुवा नृत्य", "तृतीय", 2001],
    ["c3-1", "एकल नृत्य", "प्रथम", 3001],
    ["c3-2", "एकल नृत्य", "द्वितीय", 2001],
    ["c3-3", "एकल नृत्य", "तृतीय", 1001],
    ["c4-1", "हल्बा जनजाति के गीत", "प्रथम", 3001],
    ["c4-2", "हल्बा जनजाति के गीत", "द्वितीय", 2001],
    ["c4-3", "हल्बा जनजाति के गीत", "तृतीय", 1001],
    ["c5-1", "रंगोली", "प्रथम", 3001],
    ["c5-2", "रंगोली", "द्वितीय", 2001],
    ["c5-3", "रंगोली", "तृतीय", 1001],
    ["m-1", "हाफ मैराथन (पुरुष)", "प्रथम", 5001],
    ["m-2", "हाफ मैराथन (पुरुष)", "द्वितीय", 3001],
    ["m-3", "हाफ मैराथन (पुरुष)", "तृतीय", 2001],
    ["m-4", "हाफ मैराथन (पुरुष)", "चतुर्थ", 1501],
    ["m-5", "हाफ मैराथन (पुरुष)", "पंचम", 1201],
    ["m-6", "हाफ मैराथन (पुरुष)", "षष्ठम", 1001],
    ["m-7", "हाफ मैराथन (पुरुष)", "सप्तम", 801],
    ["m-8", "हाफ मैराथन (पुरुष)", "अष्टम", 701],
    ["m-9", "हाफ मैराथन (पुरुष)", "नवम", 601],
    ["m-10", "हाफ मैराथन (पुरुष)", "दशम", 501],
    ["w-1", "हाफ मैराथन (महिला)", "प्रथम", 5001],
    ["w-2", "हाफ मैराथन (महिला)", "द्वितीय", 3001],
    ["w-3", "हाफ मैराथन (महिला)", "तृतीय", 2001],
    ["w-4", "हाफ मैराथन (महिला)", "चतुर्थ", 1501],
    ["w-5", "हाफ मैराथन (महिला)", "पंचम", 1201],
    ["w-6", "हाफ मैराथन (महिला)", "षष्ठम", 1001],
    ["w-7", "हाफ मैराथन (महिला)", "सप्तम", 801],
    ["w-8", "हाफ मैराथन (महिला)", "अष्टम", 701],
    ["w-9", "हाफ मैराथन (महिला)", "नवम", 601],
    ["w-10", "हाफ मैराथन (महिला)", "दशम", 501]
    ];
    sl.getRange(1, 1, 1, 7).setValues([['कोड', 'प्रतियोगिता', 'स्थान', 'राशि (₹)', 'सहयोगी का नाम', 'प्रविष्टि क्र.', 'दिनांक']])
      .setFontWeight('bold').setFontColor('#ffffff').setBackground(maroon);
    sl.getRange(2, 1, rows.length, 4).setValues(rows);
    sl.setFrozenRows(1);
    sl.setColumnWidth(2, 200);
    sl.setColumnWidth(5, 260);
    var rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$E2<>""').setBackground('#e1f2e8')
      .setRanges([sl.getRange(2, 1, rows.length, 7)]).build();
    sl.setConditionalFormatRules([rule]);
  }
  SpreadsheetApp.flush();
}

function doGet(e) {
  return json_({ ok: true, taken: readTaken_() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  }
  try {
    var d = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    var name = clean_(d.name, 120);
    var mobile = String(d.mobile || '').replace(/\D/g, '').slice(-10);
    if (name.length < 3 || !/^[6-9]\d{9}$/.test(mobile)) {
      return json_({ ok: false, error: 'invalid' });
    }
    var type = d.type === 'prize' ? 'prize' : 'support';

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var slotSheet = ss.getSheetByName(SHEET_SLOTS);
    var slotData = slotSheet.getRange(2, 1, slotSheet.getLastRow() - 1, 5).getValues();
    var rowOf = {}, amtOf = {}, labelOf = {};
    slotData.forEach(function (r, i) {
      rowOf[r[0]] = i + 2;
      amtOf[r[0]] = Number(r[3]) || 0;
      labelOf[r[0]] = r[1] + ' – ' + r[2] + ' – ₹' + r[3];
    });

    var slots = [];
    var total = 0;
    if (type === 'prize') {
      slots = (Array.isArray(d.slots) ? d.slots : []).filter(function (s) { return rowOf[s]; });
      if (!slots.length) return json_({ ok: false, error: 'invalid' });
      var taken = slots.filter(function (s) {
        return String(slotSheet.getRange(rowOf[s], 5).getValue()).trim() !== '';
      });
      if (taken.length) return json_({ ok: false, error: 'taken', slots: taken, taken: readTaken_() });
      slots.forEach(function (s) { total += amtOf[s]; });
    } else {
      total = Math.max(0, Math.round(Number(d.amount) || 0));
      if (!total) return json_({ ok: false, error: 'invalid' });
    }

    var now = new Date();
    var id = 'HYSM26-' + Utilities.formatDate(now, 'Asia/Kolkata', 'MMdd') + '-' +
      ('000' + (ss.getSheetByName(SHEET_ENTRIES).getLastRow())).slice(-3);
    var stamp = Utilities.formatDate(now, 'Asia/Kolkata', 'dd/MM/yyyy HH:mm');
    var desig = clean_(d.desig, 120);
    var publicName = d.publicName ? (name + (desig ? ', ' + desig : '')) : 'एक सहयोगी';

    ss.getSheetByName(SHEET_ENTRIES).appendRow([
      stamp,
      id,
      type === 'prize' ? 'पुरस्कार' : 'कार्यक्रम सहयोग',
      name,
      desig,
      clean_(d.father, 120),
      "'" + mobile,
      d.wa ? "'" + String(d.wa).replace(/\D/g, '').slice(-10) : '',
      clean_(d.place, 120),
      clean_(d.block, 120),
      slots.map(function (s) { return labelOf[s]; }).join('\n'),
      slots.join(', '),
      clean_(d.memory, 200),
      total,
      clean_(d.pay, 60),
      clean_(d.utr, 60),
      clean_(d.ref, 120),
      d.publicName ? 'हाँ' : 'नहीं',
      clean_((d.prevYears || []).join(', '), 120),
      ''
    ]);

    slots.forEach(function (s) {
      slotSheet.getRange(rowOf[s], 5, 1, 3).setValues([[publicName, id, stamp]]);
    });

    return json_({ ok: true, id: id, total: total, taken: readTaken_() });
  } catch (err) {
    return json_({ ok: false, error: 'server' });
  } finally {
    lock.releaseLock();
  }
}

function readTaken_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SLOTS);
  var last = sh.getLastRow();
  var out = {};
  if (last < 2) return out;
  sh.getRange(2, 1, last - 1, 5).getValues().forEach(function (r) {
    var n = String(r[4]).trim();
    if (r[0] && n) out[r[0]] = n;
  });
  return out;
}

/* शीट में फॉर्मूला न बने, इसलिए = + - @ से शुरू होने वाले टेक्स्ट को सुरक्षित करें */
function clean_(v, max) {
  var s = String(v == null ? '' : v).replace(/\s+/g, ' ').trim().slice(0, max || 200);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
