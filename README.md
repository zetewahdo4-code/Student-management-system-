/**
 * ራያ ዩኒቨርስቲ ግቢ ጉባኤ – የምዝገባ ሰርቨር (Google Apps Script)
 * ጽሑፍ (ሙሉ መረጃ) → Google Sheet (Excel ሆኖ ይወርዳል)
 * ፎቶ እና ፊርማ → Google Drive አቃፊ፤ መታወቂያቸው በአባሉ ረድፍ ውስጥ ይቀመጣል
 */
var SPREADSHEET_ID = "1xUaoz19ZQ3_82ME_Jmytm0iEFPOTc1gWfsyqi1QtNZQ";
var SHEET_STUDENTS = "Students";
var SHEET_GRADS    = "Graduates";
var PHOTO_FOLDER   = "Raya-Gubae-Photos";
var DEFAULT_PIN    = "1234";

var HEAD = {
  students: ["ቀን","ሙሉ ስም","የክርስትና ስም","ዲፓርትመንት","Social/Natural","ጾታ","ንስሓ አባት አለ?","የንስሓ አባት ስም","የአገልግሎት ክፍሎች","ስልክ","መታወቂያ ቁጥር","ዓመት","ማዕረግ","የትውልድ ከተማ","ፎቶ (Drive ID)","ፊርማ (Drive ID)","የሞያ ተሰጥዖ"],
  graduates:["ቀን","ሙሉ ስም","የክርስትና ስም","ዲፓርትመንት","ያገለገሉበት ክፍል","አድራሻ","ስልክ","የተመረቁበት ዓመት","ማዕረግ","የትውልድ ከተማ","Social/Natural","ጾታ","የተመረቁበት ወር","የአሁን ሥራ","ፎቶ (Drive ID)","ፊርማ (Drive ID)","የሞያ ተሰጥዖ"]
};
var NCOL = 17, PHOTO_COL = 14, SIGN_COL = 15, TALENT_COL = 16;
var TZ = "Africa/Addis_Ababa";
var MAX_PIN_FAILS = 30;   // በ15 ደቂቃ ውስጥ የሚፈቀድ የተሳሳተ ፒን ብዛት
var KEYS = {
  students: {"ስም":1,"ክርስትና":2,"ዲፓርትመንት":3,"ስቴም":4,"ጾታ":5,"ንስሓ":6,"አባት":7,"ክፍል":8,"ስልክ":9,"መታወቂያ":10,"ዓመት":11,"ማዕረግ":12,"ትውልድ":13,"ተሰጥዖ":16},
  graduates:{"ስም":1,"ክርስትና":2,"ዲፓርትመንት":3,"ያገለገሉበት":4,"አድራሻ":5,"ስልክ":6,"የተመረቁበት":7,"ማዕረግ":8,"ትውልድ":9,"ስቴም":10,"ጾታ":11,"ወር":12,"ስራ":13,"ተሰጥዖ":16}
};

/* ---------- አጋዥ ---------- */
function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function getPin(){ return PropertiesService.getScriptProperties().getProperty("PIN") || DEFAULT_PIN; }
function pinState(pin){
  var c = CacheService.getScriptCache(), n = Number(c.get("pinfail")||0);
  if(n >= MAX_PIN_FAILS) return "locked";
  if(String(pin).trim() === getPin()){ if(n) c.remove("pinfail"); return "ok"; }
  c.put("pinfail", String(n+1), 900);
  return "bad";
}
function tkey(t){ return t === "graduates" ? "graduates" : "students"; }
function normId(s){ return String(s||"").replace(/\s+/g,"").toUpperCase(); }
function pk(p){ return String(p||"").replace(/\D/g,"").slice(-9); }
function akey(s){
  return String(s||"").replace(/\s+/g," ").trim().toLowerCase().replace(/[\u1210-\u1217\u1220-\u1227\u1280-\u1287\u12D0-\u12D7\u1340-\u1347]/g,function(c){
    var n=c.charCodeAt(0);
    if(n<0x1218)n-=0x10;else if(n<0x1228)n+=0x10;else if(n<0x1288)n-=0x80;else if(n<0x12D8)n-=0x30;else n-=8;
    return String.fromCharCode(n);
  });
}
function cellStr(v){
  if(v===null||v===undefined) return "";
  if(v instanceof Date) return Utilities.formatDate(v, TZ, "yyyy-MM-dd HH:mm:ss");
  return String(v);
}
var _sheets = {};
function getSheet(t){
  t = tkey(t);
  if(_sheets[t]) return _sheets[t];          // በአንድ ጥሪ ውስጥ አንድ ጊዜ ብቻ ይከፈታል
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var name = t==="students" ? SHEET_STUDENTS : SHEET_GRADS;
  var sh = ss.getSheetByName(name), fresh = false;
  if(!sh){ sh = ss.insertSheet(name); fresh = true; }
  if(sh.getMaxColumns() < NCOL) sh.insertColumnsAfter(sh.getMaxColumns(), NCOL - sh.getMaxColumns());
  if(fresh){
    sh.getRange(1,1,sh.getMaxRows(),NCOL).setNumberFormat("@");   // ስልክ 0911… እንዳይበላሽ
    sh.getRange(1,1,1,NCOL).setValues([HEAD[t]]).setFontWeight("bold").setBackground("#7a1f2b").setFontColor("#ffffff");
    sh.setFrozenRows(1);
  } else if(!sh.getRange(1,NCOL).getValue()){      // የድሮ ሉህ፦ አዲሱን የ "ተሰጥዖ" አምድ ይጨምራል
    sh.getRange(1,NCOL,sh.getMaxRows(),1).setNumberFormat("@");
    sh.getRange(1,NCOL).setValue(HEAD[t][NCOL-1]).setFontWeight("bold").setBackground("#7a1f2b").setFontColor("#ffffff");
  }
  _sheets[t] = sh;
  return sh;
}
function readRows(t){
  var sh = getSheet(t), last = sh.getLastRow();
  if(last < 2) return [];
  return sh.getRange(2,1,last-1,NCOL).getValues().map(function(r){ return r.map(cellStr); });
}
function writeRow(sh,row,vals){
  if(row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows());
  sh.getRange(row,1,1,NCOL).setNumberFormat("@").setValues([vals]);
}
function findRow(sh, old){ // አዲሱን መጀመሪያ ይፈልጋል
  var last = sh.getLastRow(); if(last < 2) return 0;
  var pc = t_phone(sh), v = sh.getRange(2,1,last-1,pc+1).getValues(), nm = akey(old.n), ph = pk(old.p), best = 0;
  for(var i=v.length-1;i>=0;i--){
    if(akey(cellStr(v[i][1]))===nm && pk(cellStr(v[i][pc]))===ph){
      if(!old.d || cellStr(v[i][0])===String(old.d)) return i+2;
      if(!best) best = i+2;
    }
  }
  return best;
}
function t_phone(sh){ return sh.getName()===SHEET_STUDENTS ? 9 : 6; }
function dupCheck(t, vals, skipRow){
  var sh = getSheet(t), last = sh.getLastRow();
  if(last < 2) return "";
  var id = t==="students" ? normId(vals[10]) : "", nm = akey(vals[1]);
  var names = sh.getRange(2,2,last-1,1).getValues();
  var ids = (t==="students" && id) ? sh.getRange(2,11,last-1,1).getValues() : null;
  for(var i=0;i<names.length;i++){
    if(skipRow && i+2===skipRow) continue;
    if(ids && normId(cellStr(ids[i][0]))===id) return "dup";
    if(nm && akey(cellStr(names[i][0]))===nm) return "dupname";
  }
  return "";
}
function folder(){
  var it = DriveApp.getFoldersByName(PHOTO_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(PHOTO_FOLDER);
}
function saveImg(dataUrl, base){
  var m = /^data:(image\/[a-zA-Z+]+);base64,(.+)$/.exec(dataUrl||"");
  if(!m || m[2].length > 3500000) return "";
  var ext = m[1].indexOf("png")>-1 ? ".png" : ".jpg";
  var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], String(base).replace(/[\\\/:*?"<>|]/g,"_")+ext);
  return folder().createFile(blob).getId();
}
function trash(id){ if(id){ try{ DriveApp.getFileById(id).setTrashed(true); }catch(e){} } }

/* ---------- GET ---------- */
function doGet(e){
  var p = (e && e.parameter) || {};
  try{
    if(p.action==="ping") return out({ok:true});
    if(p.action==="reg")  return out(register(p));
    if(p.pin===undefined) return out({ok:false, err:"nopin"});
    var ps = pinState(p.pin);
    if(ps==="locked") return out({ok:false, err:"locked", msg:"ብዙ ጊዜ የተሳሳተ ፒን ገብቷል፤ ከ15 ደቂቃ በኋላ እንደገና ይሞክሩ።"});
    if(ps!=="ok") return out({ok:false, err:"pin"});

    if(p.newpin){
      PropertiesService.getScriptProperties().setProperty("PIN", String(p.newpin).trim());
      return out({ok:true, changed:true});
    }
    if(p.del)            return out(removeRows(JSON.parse(p.del)));
    if(p.action==="edit")   return out(editRow(p));
    if(p.action==="photos") return out(getPhotos(p.ids));

    return out({ok:true, students:{rows:readRows("students")}, graduates:{rows:readRows("graduates")}});
  }catch(err){
    return out({ok:false, err:"server", msg:String(err)});
  }
}

function register(p){
  var t = tkey(p.type), sh, vals = new Array(NCOL), k;
  for(var i=0;i<NCOL;i++) vals[i]="";
  vals[0] = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm:ss");
  for(k in KEYS[t]) if(p[k]!==undefined) vals[KEYS[t][k]] = String(p[k]).trim().slice(0,500);
  if(t==="students") vals[10] = normId(vals[10]);
  if(!vals[1] || !vals[t==="students"?9:6]) return {ok:false, err:"missing"};
  var rid = p.rid ? "rid_"+String(p.rid).slice(0,40) : "", cache = CacheService.getScriptCache();
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try{
    if(rid && cache.get(rid)) return {ok:true, again:true};   // ተመሳሳይ ጥያቄ ድጋሚ ቢመጣ አይደገምም
    var d = dupCheck(t, vals, 0);
    if(d) return {ok:false, err:d};
    sh = getSheet(t);
    writeRow(sh, sh.getLastRow()+1, vals);
    if(rid) cache.put(rid, "1", 1800);
    return {ok:true};
  } finally { lock.releaseLock(); }
}

function removeRows(list){
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try{
    var n = 0, by = {students:[], graduates:[]};
    list.forEach(function(x){
      var t = tkey(x.t), sh = getSheet(t), r = findRow(sh, x);
      if(r && by[t].indexOf(r)<0) by[t].push(r);
    });
    ["students","graduates"].forEach(function(t){
      var sh = getSheet(t);
      by[t].sort(function(a,b){return b-a}).forEach(function(r){
        var ids = sh.getRange(r,PHOTO_COL+1,1,2).getValues()[0];
        trash(ids[0]); trash(ids[1]);
        sh.deleteRow(r); n++;
      });
    });
    return {ok:true, deleted:n};
  } finally { lock.releaseLock(); }
}

function editRow(p){
  var t = tkey(p.type), old = JSON.parse(p.old), d = JSON.parse(p.data);
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try{
    var sh = getSheet(t), r = findRow(sh, old);
    if(!r) return {ok:false, err:"notfound"};
    var vals = sh.getRange(r,1,1,NCOL).getValues()[0].map(cellStr), k;
    for(k in KEYS[t]) if(d[k]!==undefined) vals[KEYS[t][k]] = String(d[k]).trim().slice(0,500);
    if(t==="students") vals[10] = normId(vals[10]);
    var x = dupCheck(t, vals, r);
    if(x) return {ok:false, err:x};
    writeRow(sh, r, vals);
    return {ok:true};
  } finally { lock.releaseLock(); }
}

function getPhotos(ids){
  var allow = {};
  ["students","graduates"].forEach(function(t){
    var sh = getSheet(t), last = sh.getLastRow();
    if(last < 2) return;
    sh.getRange(2,PHOTO_COL+1,last-1,2).getValues().forEach(function(r){ var a=cellStr(r[0]), b=cellStr(r[1]); if(a) allow[a]=1; if(b) allow[b]=1; });
  });
  var photos = {};
  String(ids||"").split(",").slice(0,10).forEach(function(id){
    id = id.trim();
    if(!id || !allow[id]) return;
    try{
      var b = DriveApp.getFileById(id).getBlob();
      photos[id] = "data:"+b.getContentType()+";base64,"+Utilities.base64Encode(b.getBytes());
    }catch(e){}
  });
  return {ok:true, photos:photos};
}

/* ---------- POST (ፎቶ እና ፊርማ) ---------- */
function doPost(e){
  try{
    var b = JSON.parse(e.postData.contents);
    if(b.action!=="media") return out({ok:false});
    var lock = LockService.getScriptLock(); lock.waitLock(25000);
    try{
      var t = tkey(b.type), sh = getSheet(t), r = findRow(sh, {n:b.name, p:b.phone});
      if(!r) return out({ok:false, err:"notfound"});
      var cur = sh.getRange(r,PHOTO_COL+1,1,2).getValues()[0].map(cellStr);
      var base = (b.name||"member")+"_"+pk(b.phone);
      if(b.photo){ var a = saveImg(b.photo, base+"_photo"); if(a){ trash(cur[0]); cur[0]=a; } }
      if(b.sign){  var s = saveImg(b.sign,  base+"_sign");  if(s){ trash(cur[1]); cur[1]=s; } }
      sh.getRange(r,PHOTO_COL+1,1,2).setNumberFormat("@").setValues([cur]);
      return out({ok:true});
    } finally { lock.releaseLock(); }
  }catch(err){
    return out({ok:false, err:"server", msg:String(err)});
  }
}

/* አንድ ጊዜ ብቻ Run አድርግ — ፈቃድ ለመስጠት እና ሉሆቹን/አቃፊውን ለመፍጠር */
function setup(){ getSheet("students"); getSheet("graduates"); folder(); }
