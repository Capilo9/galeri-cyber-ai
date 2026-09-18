/** Galeri Cyber AI — bind this project to the target Google Spreadsheet. */
var CYBER_API_VERSION = '2026-09-18-thumbnail-v2';
var SUBJECTS = ['IPAS','Matematika','Bahasa Indonesia','Pendidikan Pancasila','Bahasa Inggris','PJOK','Seni','Coding dan Kecerdasan Artifisial','Bahasa Jawa','Lainnya'];
var HEADERS = ['ID','Judul Materi','Deskripsi','Thumbnail URL','Link Materi','Nama Guru atau Pembuat','Sekolah atau Instansi','Kategori Mapel SD','Kelas SD','Tags Pisahkan koma','Tanggal Dibuat','Jumlah Dilihat','Jumlah Suka','Unggulan TRUE FALSE','Share Link AI Studio Fullscreen','File Lembar Literasi QR Code','Nama File Lembar Literasi'];
function json_(obj){obj.apiVersion=CYBER_API_VERSION;return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);}
function failure_(code,message,errors){var e=new Error(message);e.code=code;e.fields=errors;throw e;}
function clean_(s){return typeof s==='string'?s.replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g,'').trim():'';}
function sheetText_(s){s=clean_(s);return /^[=+@-]/.test(s)?"'"+s:s;}
function https_(value){
  if(typeof value!=='string'||value.length>2048)return null;
  var s=value.trim();if(/[\s\\\u0000-\u001f]/.test(s))return null;
  var m=s.match(/^https:\/\/([^/?#]+)([^\s]*)$/i);if(!m)return null;
  var authority=m[1];if(authority.indexOf('@')>=0)return null;
  var a=authority.match(/^([a-zA-Z0-9.-]+)(?::([0-9]{1,5}))?$/);if(!a||a[2]&&Number(a[2])>65535)return null;
  var host=a[1].toLowerCase();if(host.indexOf('.')<0||host.endsWith('.')||/(^|\.)(localhost|local|internal|test|invalid)$/.test(host)||!host.split('.').every(function(label){return /^[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?$/.test(label)}))return null;
  // Reject alternate numeric IPv4 spellings and non-public IPv4 ranges.
  if(/^(?:0x[0-9a-f]+|[0-9]+)(?:\.(?:0x[0-9a-f]+|[0-9]+))*$/i.test(host)){
    if(!/^(?:0|[1-9][0-9]{0,2})(?:\.(?:0|[1-9][0-9]{0,2})){3}$/.test(host))return null;
    var ip=host.split('.').map(Number),x=ip[0],y=ip[1];if(ip.some(function(n){return n>255})||x===0||x===10||x===127||x>=224||x===169&&y===254||x===172&&y>=16&&y<=31||x===192&&y===168||x===100&&y>=64&&y<=127||x===198&&(y===18||y===19))return null;
  }
  return {url:s,host:host,path:m[2]||'/'};
}
function driveId_(s){var u=https_(s);if(!u||u.host!=='drive.google.com')return '';var m=u.path.match(/^\/file\/d\/([\w-]{10,})(?:[/?#]|$)/);if(m)return m[1];if(!/^\/(open|uc|thumbnail)\?/.test(u.path))return '';m=u.path.match(/[?&]id=([\w-]{10,})(?:[&#]|$)/);return m?m[1]:'';}
function thumbnail_(value){
  if(typeof value!=='string')return '';
  // Pasted links may contain invisible marks, escaped ampersands or quotes.
  value=value.replace(/[\u200B-\u200D\u2060\uFEFF]/g,'').trim().replace(/&amp;/gi,'&');
  if((value.charAt(0)==='"'&&value.slice(-1)==='"')||(value.charAt(0)==="'"&&value.slice(-1)==="'"))value=value.slice(1,-1).trim();
  if((value.match(/https?:\/\//gi)||[]).length!==1)return '';
  var u=https_(value);if(!u||/https?:\/\//i.test(u.path))return '';
  var id=driveId_(value);
  if(/^lh[3-6]\.googleusercontent\.com$/.test(u.host)){
    var m=u.path.match(/^\/d\/([\w-]{10,})(?:=|[/?#]|$)/);
    if(m)id=m[1];
  }
  if(id)return 'https://drive.google.com/thumbnail?id='+id+'&sz=w1200';
  if(u.host==='drive.google.com'||u.host==='docs.google.com'||u.host==='freeimage.host'&&/^\/i\//.test(u.path))return '';
  return u.url;
}

// Read only thumbnail column D. Never inspect or expose internal column O.
function readThumbnailRows_(sheet,rows){
  if(!rows.length)return rows;
  var range=sheet.getRange(2,4,rows.length,1);
  var rich=typeof range.getRichTextValues==='function'?range.getRichTextValues():[];
  var formulas=typeof range.getFormulas==='function'?range.getFormulas():[];
  return rows.map(function(row,i){
    var copy=row.slice(),candidate=thumbnail_(String(copy[3]||''));
    var text=rich[i]&&rich[i][0];
    if(!candidate&&text){
      candidate=thumbnail_(text.getLinkUrl()||'');
      if(!candidate){var runs=text.getRuns();for(var n=0;n<runs.length;n++){candidate=thumbnail_(runs[n].getLinkUrl()||'');if(candidate)break;}}
    }
    if(!candidate){
      // Only literal URL arguments; never execute a formula or follow cell references.
      var formula=formulas[i]&&formulas[i][0]||'';
      var match=formula.match(/^\s*=\s*(?:IMAGE|HYPERLINK)\s*\(\s*"((?:[^"]|"")*)"\s*(?:[,;)]|$)/i);
      if(match)candidate=thumbnail_(match[1].replace(/""/g,'"'));
    }
    copy[3]=candidate;return copy;
  });
}

function diagnoseThumbnails(){
  var s=getOrCreateSheet();
  var raw=s.getLastRow()>1?s.getRange(2,1,s.getLastRow()-1,17).getValues():[];
  var rows=readThumbnailRows_(s,raw);
  var report={apiVersion:CYBER_API_VERSION,total:rows.length,recognized:0,missingOrInvalid:[]};
  rows.forEach(function(r,i){if(thumbnail_(String(r[3]||'')))report.recognized++;else report.missingOrInvalid.push({row:i+2,id:String(r[0]||'')});});
  console.log(JSON.stringify(report));return report;
}
function validate_(raw){var v={},errors={};['title','description','creatorName','creatorSchool','category','gradeClass','link','aiStudioLink','thumbnail','literacySheetFile','literacySheetName'].forEach(function(k){v[k]=clean_(raw[k])});v.tags=Array.isArray(raw.tags)?raw.tags.map(clean_):[];
  [['title',5,120],['creatorName',2,100],['creatorSchool',2,150],['literacySheetName',3,120]].forEach(function(r){if(v[r[0]].length<r[1]||v[r[0]].length>r[2])errors[r[0]]='Isi '+r[1]+'–'+r[2]+' karakter.'});
  if(v.description.length<30)errors.description='Isi minimal 30 karakter.';
  if(SUBJECTS.indexOf(v.category)<0)errors.category='Pilih mata pelajaran.';
  if(!/^Kelas [1-6]$/.test(v.gradeClass))errors.gradeClass='Pilih kelas.';
  if(v.tags.length<1||v.tags.length>8||v.tags.some(function(t){return t.length<2||t.length>30||t.indexOf(',')>=0}))errors.tags='Isi 1–8 tag, masing-masing 2–30 karakter.';
  if(!https_(v.link))errors.link='Gunakan URL HTTPS publik yang valid.';
  var internal=https_(v.aiStudioLink);if(!internal||internal.host!=='aistudio.google.com')errors.aiStudioLink='Gunakan link HTTPS dari aistudio.google.com.';
  if(!thumbnail_(v.thumbnail))errors.thumbnail='Gunakan link file Drive atau URL gambar HTTPS langsung.';if(!driveId_(v.literacySheetFile))errors.literacySheetFile='Gunakan link file Google Drive yang valid.';
  if(Object.keys(errors).length)failure_('VALIDATION_ERROR','Data belum valid.',errors);
  v.thumbnail=thumbnail_(v.thumbnail);return v;
}
function getOrCreateSheet(){var props=PropertiesService.getScriptProperties();var id=props.getProperty('SPREADSHEET_ID');var book=id?SpreadsheetApp.openById(id):SpreadsheetApp.getActiveSpreadsheet();if(!book)throw new Error('Spreadsheet is not configured');var s=book.getSheetByName('MateriGuruSD');if(!s)s=book.insertSheet('MateriGuruSD');if(s.getMaxColumns()<17)s.insertColumnsAfter(s.getMaxColumns(),17-s.getMaxColumns());var current=s.getRange(1,1,1,17).getValues()[0];for(var i=0;i<17;i++){if(current[i]!==''&&current[i]!==HEADERS[i])throw new Error('Sheet header mismatch');if(current[i]===''){if(s.getLastRow()>1&&s.getRange(2,i+1,s.getLastRow()-1,1).getValues().some(function(r){return r[0]!==''}))throw new Error('Unlabeled populated column');s.getRange(1,i+1).setValue(HEADERS[i]);}}return s;}
function setupDatabaseManual(){var lock=LockService.getScriptLock();lock.waitLock(10000);try{var s=getOrCreateSheet();s.setFrozenRows(1);return {sheet:s.getName(),columns:17};}finally{lock.releaseLock();}}
function date_(v){if(v instanceof Date)return Utilities.formatDate(v,'Asia/Jakarta','yyyy-MM-dd');var str=String(v||'');return /^\d{4}-\d{2}-\d{2}/.test(str)?str.slice(0,10):'';}
function count_(v){var n=Number(v);return isFinite(n)&&n>=0?Math.floor(n):0;}
function publicRow_(r){return {id:String(r[0]),title:String(r[1]),description:String(r[2]||''),thumbnail:thumbnail_(String(r[3]||'')),link:https_(String(r[4]||''))?String(r[4]):'',creatorName:String(r[5]||''),creatorSchool:String(r[6]||''),category:String(r[7]||''),gradeClass:String(r[8]||''),gradeLevel:'SD',tags:String(r[9]||'').split(',').map(function(t){return t.trim()}).filter(Boolean),dateAdded:date_(r[10]),views:count_(r[11]),likes:count_(r[12]),isFeatured:r[13]===true||String(r[13]).toUpperCase()==='TRUE',literacySheetFile:driveId_(String(r[15]||''))?String(r[15]):'',literacySheetName:String(r[16]||'')};}
function doGet(){try{var s=getOrCreateSheet();var rows=s.getLastRow()>1?s.getRange(2,1,s.getLastRow()-1,17).getValues():[];rows=readThumbnailRows_(s,rows);var seen={};var data=rows.filter(function(r){if(!r[0]||!String(r[1]||'').trim()||seen[r[0]])return false;seen[r[0]]=true;return true}).map(publicRow_);return json_({status:'success',count:data.length,data:data});}catch(e){console.log(JSON.stringify({action:'get',result:'error',code:'SERVER_ERROR',time:new Date().toISOString()}));return json_({status:'error',code:'SERVER_ERROR',message:'Data belum dapat dimuat. Silakan coba lagi.'});}}
function hash_(s){return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s)).slice(0,36);}
function rate_(cache,key,max,ttl){var n=Number(cache.get(key)||0);if(n>=max)failure_('RATE_LIMITED','Terlalu banyak permintaan. Silakan coba lagi nanti.');cache.put(key,String(n+1),ttl);}
function doPost(e){var action='unknown',lock=null;try{
  var body=e&&e.postData&&e.postData.contents;if(typeof body!=='string')failure_('VALIDATION_ERROR','Data kiriman tidak valid.');var p;try{p=JSON.parse(body)}catch(ex){failure_('VALIDATION_ERROR','Format kiriman tidak valid.');}if(!p||Array.isArray(p)||typeof p!=='object')failure_('VALIDATION_ERROR','Format kiriman tidak valid.');var boundedPayload=Object.assign({},p);if(p.action==='create')delete boundedPayload.description;if(Utilities.newBlob(JSON.stringify(boundedPayload)).getBytes().length>20480)failure_('VALIDATION_ERROR','Data kiriman tidak valid.');action=p.action;if(['create','view','like'].indexOf(action)<0)failure_('INVALID_ACTION','Aksi tidak dikenali.');
  if(action==='create'&&p.website)failure_('VALIDATION_ERROR','Data kiriman tidak valid.');var v=action==='create'?validate_(p):null;if(action!=='create'&&(typeof p.id!=='string'||p.id.length>150||!p.id))failure_('VALIDATION_ERROR','ID karya tidak valid.');
  lock=LockService.getScriptLock();if(!lock.tryLock(10000))failure_('RATE_LIMITED','Sistem sedang sibuk. Silakan coba lagi.');var s=getOrCreateSheet(),cache=CacheService.getScriptCache(),rows=s.getLastRow()>1?s.getRange(2,1,s.getLastRow()-1,17).getValues():[],result;
  if(action==='create'){
    var fingerprint=hash_(v.link+'|'+v.creatorName.toLowerCase()+'|'+v.title.toLowerCase());var dup=cache.get('dup:'+fingerprint);if(dup)return json_({status:'success',id:dup,message:'Karya ini sudah tersimpan.'});
    var now=Date.now();var recent=rows.find(function(r){return String(r[4]).trim()===v.link&&String(r[5]).replace(/^'/,'').toLowerCase()===v.creatorName.toLowerCase()&&String(r[1]).replace(/^'/,'').toLowerCase()===v.title.toLowerCase()&&String(r[0]).match(/^mat-sd-(\d+)-/)&&now-Number(String(r[0]).match(/^mat-sd-(\d+)-/)[1])<600000});if(recent)return json_({status:'success',id:String(recent[0]),message:'Karya ini sudah tersimpan.'});
    rate_(cache,'create:'+hash_(v.creatorName.toLowerCase()+'|'+v.creatorSchool.toLowerCase()),5,3600);rate_(cache,'global:create',100,3600);
    var id='mat-sd-'+now+'-'+Utilities.getUuid().slice(0,8);var row=[id,v.title,v.description,v.thumbnail,v.link,v.creatorName,v.creatorSchool,v.category,v.gradeClass,v.tags.join(', '),Utilities.formatDate(new Date(),'Asia/Jakarta','yyyy-MM-dd'),0,0,false,v.aiStudioLink,v.literacySheetFile,v.literacySheetName].map(function(value){return typeof value==='string'?sheetText_(value):value});s.getRange(s.getLastRow()+1,1,1,17).setValues([row]);SpreadsheetApp.flush();cache.put('dup:'+fingerprint,id,600);result={status:'success',id:id,message:'Bahan ajar berhasil ditambahkan.'};
  }else{
    var matches=[];rows.forEach(function(r,i){if(String(r[0])===p.id)matches.push(i)});if(!matches.length)failure_('NOT_FOUND','Karya tidak ditemukan.');if(matches.length>1)failure_('SERVER_ERROR','Karya belum dapat diperbarui.');rate_(cache,'counter:'+action+':'+hash_(p.id),action==='like'?120:600,3600);var col=action==='view'?12:13;var cell=s.getRange(matches[0]+2,col),n=count_(cell.getValue())+1;cell.setValue(n);SpreadsheetApp.flush();result={status:'success',id:p.id};result[action==='view'?'views':'likes']=n;
  }
  console.log(JSON.stringify({action:action,result:'success',time:new Date().toISOString()}));return json_(result);
}catch(ex){var code=ex.code||'SERVER_ERROR';console.log(JSON.stringify({action:['create','view','like'].indexOf(action)>=0?action:'unknown',result:'error',code:code,time:new Date().toISOString()}));return json_({status:'error',code:code,message:ex.code?ex.message:'Permintaan belum berhasil. Silakan coba lagi.',errors:ex.fields||undefined});}finally{if(lock)lock.releaseLock();}}
