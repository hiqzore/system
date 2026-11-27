const SPREADSHEET_ID = '1J5sBq0de2-VcgqvHMZq6yd-kkjPZdeMYKr-Apr1pumg';
function getSS(){
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}
function doGet(){
  return HtmlService.createHtmlOutputFromFile('Index').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function getDashboardPage(){
  return HtmlService.createTemplateFromFile('Dashboard').evaluate().getContent();
}


//============================================================== CHECK LOGIN DATA =====================================================


function checkLogin(e, p){
  try{
    var s = getSS().getSheetByName('Users');
    if(!s) return { status:'fail', message:'Sheet Users tidak ada' };
    var d = s.getDataRange().getValues();
    e = String(e).trim().toLowerCase();
    p = String(p).trim();
    for(var i = 1; i < d.length; i++){
      if(String(d[i][0]).trim().toLowerCase() === e && String(d[i][1]).trim() === p){
        return { status:'ok' };
      }
    }
    return { status:'fail', message:'Email / Password salah' };
  }catch(err){
    return { status:'fail', message:'ERROR: ' + err.message };
  }
}
function resolveSheetName(n){
  return n === 'Kontrak Host' ? 'Kontrak Host < 30 Day' : n;
}
function getSheetData(n){
  try{
    var r = resolveSheetName(n);
    var s = getSS().getSheetByName(r);
    if(!s) return { status:'notfound', real:r };
    return { status:'ok', data:s.getDataRange().getValues() };
  }catch(e){
    return { status:'error', message:e.message };
  }
}

// =================================================
// Penjelasan: DASHBOARD - Ambil Dashboard A1:V30, kembalikan box1-7 + data1-4 untuk chart
// =================================================

//======= Penjelasan: getDashboardAll() - Fungsi utama Dashboard, baca A1:V30, olah jadi box + grafik =======
//======= Panggilan: Index.html initDash() -> getDashboardAll() =======
function getDashboardAll(){
  try{
    //======= Penjelasan: Pakai getSS() biar konsisten ke SPREADSHEET_ID, bukan getActiveSpreadsheet() =======
    var ss = getSS();
    var sh = ss.getSheetByName('Dashboard');
    if(!sh) return {status:'error', message:'Sheet Dashboard tidak ketemu'};
    var vals = sh.getRange('A1:V30').getValues();
    //======= Penjelasan: val(c) - Ambil nilai dari alamat sel, misal val('C3'), ubah huruf kolom jadi angka =======
    function val(c){
      var m = c.match(/^([A-Z]+)(\d+)$/);
      if(!m) return '';
      var col = 0;
      for(var i=0;i<m[1].length;i++){ col = col*26 + (m[1].charCodeAt(i)-64); }
      var row = parseInt(m[2],10);
      if(row<1 || row>vals.length || col<1 || col>vals[0].length) return '';
      return vals[row-1][col-1];
    }
    //======= Penjelasan: toNum(v) - Ubah jadi angka, kalau 0-1 dianggap persen jadi dikali 100 =======
    function toNum(v){ var n=Number(v)||0; if(n!==0 && Math.abs(n)<=1) n=n*100; return n; }
    //======= Penjelasan: fmtDate - Format Date jadi dd MMM, misal 15 Nov =======
    function fmtDate(v){ if(v instanceof Date){ return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd MMM'); } return String(v==null?'':v); }
    //======= Penjelasan: fmtDateLong - Format Date jadi dd MMM yyyy HH:mm =======
    function fmtDateLong(v){ if(v instanceof Date){ return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd MMM yyyy HH:mm'); } return String(v==null?'':v); }
    //======= Penjelasan: fmtDateOnly - Format Date jadi dd MMM yyyy saja =======
    function fmtDateOnly(v){ if(v instanceof Date){ return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd MMM yyyy'); } return String(v==null?'':v); }
    //======= Penjelasan: fmtNum - Format angka pakai id-ID, kalau Date pakai fmtDateLong =======
    function fmtNum(v){ if(v==null||v==='') return '-'; if(v instanceof Date) return fmtDateLong(v); if(typeof v==='number') return v.toLocaleString('id-ID'); return String(v); }
    //======= Penjelasan: label(cell,def) - Ambil label dari sel, kalau kosong pakai def =======
    function label(cell, def){ var v=String(val(cell)||'').trim(); return v? v : def; }
    //======= Penjelasan: fmtRp - Format jadi Rp, misal Rp 1.000.000 =======
    function fmtRp(v){
      if(v==null||v==='') return '-';
      if(typeof v==='number') return 'Rp ' + Math.round(v).toLocaleString('id-ID');
      var n=Number(String(v).replace(/\./g,'').replace(',','.'))||0;
      return 'Rp ' + Math.round(n).toLocaleString('id-ID');
    }
    //======= Penjelasan: fmtDiamond - Format diamond jadi angka ribuan id-ID =======
    function fmtDiamond(v){
      if(v==null||v==='') return '-';
      if(typeof v==='number') return Math.round(v).toLocaleString('id-ID');
      var n=Number(String(v).replace(/\./g,'').replace(',','.'))||0;
      return Math.round(n).toLocaleString('id-ID');
    }
    //======= Penjelasan: box1 - Card Periode, B3-B6 + C3-C6 =======
    var box1={title:'Periode', icon:'🗓️', color:'#3B82F6', bg:'#EFF6FF', rows:[
      {label:label('B3','Bulan'), value:fmtNum(val('C3'))},
      {label:label('B4','Total Hari'), value:fmtNum(val('C4'))},
      {label:label('B5','Sisa Hari'), value:fmtNum(val('C5'))},
      {label:label('B6','Last Update'), value:fmtDateOnly(val('C6'))}
    ]};
    //======= Penjelasan: box2 - Card Target Bonus 2%,4%,7%,10%, Rp dari C21-C18, Diamond dari C14-C11 =======
    var box2={title:'Target Bonus', icon:'📊', color:'#10B981', bg:'#ECFDF5', rows:[
      {label:'2%', mid:fmtRp(val('C21')), value:fmtDiamond(val('C14'))},
      {label:'4%', mid:fmtRp(val('C20')), value:fmtDiamond(val('C13'))},
      {label:'7%', mid:fmtRp(val('C19')), value:fmtDiamond(val('C12'))},
      {label:'10%', mid:fmtRp(val('C18')), value:fmtDiamond(val('C11'))}
    ]};
    //======= Penjelasan: box3 - Card Rata-rata Diamond / Hari, C23 + potensi C24 + akhir bulan C25 =======
    var box3={title:'Rata-rata Diamond / Hari', icon:'💎', color:'#8B5CF6', bg:'#F5F3FF', big:fmtDiamond(val('C23')), sub:label('B23','Rata Rata Diamond per hari'), potensial:fmtDiamond(val('C24')), potensialLabel:label('B24','Potensi Sisa Hari'), fullMonth:fmtDiamond(val('C25')), fullMonthLabel:label('B25','Akhir Bulan')};
    //======= Penjelasan: box4 - Card Revenue, Rp dari V25, sub dari U25 =======
    var box4={title:'Revenue', icon:'💰', color:'#F59E0B', bg:'#FFFBEB', big:fmtRp(val('V25')), sub:String(val('U25')||'Revenue perusahaan tahun')};
    //======= Penjelasan: makeProgress - Hitung persen progress, cur / max * 100 =======
    function makeProgress(title, sub, maxCell, curCell){
      var mx = Number(val(maxCell))||0;
      var cr = Number(val(curCell))||0;
      var pct = mx>0 ? Math.min(100, Math.max(0, (cr/mx)*100)) : 0;
      return {
        title:title, sub:sub,
        max: fmtDiamond(mx),
        cur: fmtDiamond(cr),
        pct: Math.round(pct*10)/10,
        maxRaw: mx, curRaw: cr
      };
    }
    //======= Penjelasan: box5 - Progress Recruitment Target, max C16, cur C17 =======
    var box5 = makeProgress('Recruitment Target', label('B16','Recruitment Target') + ' / ' + label('B17','Tercapai'), 'C16', 'C17');
    //======= Penjelasan: Hitung Berlian - Total C9, PK C10, Multi C26, Lain = Total - PK - Multi =======
    var totalDiamond = Number(val('C9'))||0;
    var pkDiamond = Number(val('C10'))||0;
    var multiDiamond = Number(val('C26'))||0;
    var lainDiamond = Math.max(0, totalDiamond - pkDiamond - multiDiamond);
    var pctPK = totalDiamond>0 ? (pkDiamond/totalDiamond)*100 : 0;
    var pctMulti = totalDiamond>0 ? (multiDiamond/totalDiamond)*100 : 0;
    var pctLain = totalDiamond>0 ? (lainDiamond/totalDiamond)*100 : 0;
    //======= Penjelasan: box6 - Card Perbandingan Berlian PK vs Multi vs Lain =======
    var box6 = {
      title:'Perbandingan (Berlian Match, Multi-Guest, Berlian)',
      sub:'Berlian PK VS Berlian Multi-Guest VS Berlian Lainnya',
      max: fmtDiamond(totalDiamond),
      cur: fmtDiamond(pkDiamond),
      pct: Math.round(pctPK*10)/10,
      total: fmtDiamond(totalDiamond),
      pk: fmtDiamond(pkDiamond),
      multi: fmtDiamond(multiDiamond),
      lain: fmtDiamond(lainDiamond),
      pctPK: Math.round(pctPK*10)/10,
      pctMulti: Math.round(pctMulti*10)/10,
      pctLain: Math.round(pctLain*10)/10
    };
    //======= Penjelasan: box7 - Progress Berlian Match Target, max C15, cur C10 =======
     var box7 = makeProgress('Berlian Match Target', label('B15','Berlian Match Target') + ' / ' + label('B10','Tercapai'), 'C15', 'C10');
    //======= Penjelasan: detectUnit - Tebak satuan grafik, % kalau header ada %, Rp kalau ada Rp / Diamond / Revenue, dll =======
    function detectUnit(headers, all){
      var h = headers.join(' ').toLowerCase();
      if(h.indexOf('%')>-1 || h.indexOf('persen')>-1 || h.indexOf('percent')>-1) return '%';
      if(h.indexOf('rp')>-1 || h.indexOf('rupiah')>-1 || h.indexOf('revenue')>-1 || h.indexOf('diamond')>-1 || h.indexOf('berlian')>-1 || h.indexOf('idr')>-1 || h.indexOf('omset')>-1) return 'Rp';
      var max=0; for(var i=0;i<all.length;i++) for(var j=0;j<all[i].length;j++){ var n=Number(all[i][j])||0; if(Math.abs(n)>max) max=Math.abs(n); }
      if(max>1000) return 'Rp'; if(max>0 && max<=100) return '%'; return 'Rp';
    }
    //======= Penjelasan: rangeVals - Ambil nilai range, misal A1:C5 jadi array 2D =======
    function rangeVals(xRange){
      var m = xRange.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
      var c1=0, c2=0;
      for(var i=0;i<m[1].length;i++) c1=c1*26+(m[1].charCodeAt(i)-64);
      for(var j=0;j<m[3].length;j++) c2=c2*26+(m[3].charCodeAt(j)-64);
      var r1=parseInt(m[2],10), r2=parseInt(m[4],10);
      var out=[];
      for(var r=r1;r<=r2;r++){ var row=[]; for(var c=c1;c<=c2;c++) row.push(vals[r-1][c-1]); out.push(row); }
      return out;
    }
    //======= Penjelasan: getLine - Olah data untuk grafik garis, X + header + data =======
    function getLine(xRange, headerRange, dataRange){
      var xs = rangeVals(xRange).map(function(r){ return r[0]; });
      var headers = rangeVals(headerRange)[0];
      var all = rangeVals(dataRange);
      var labels=[]; var series=[];
      for(var c=0;c<headers.length;c++) series.push({label:String(headers[c]||('Data '+(c+1))), data:[]});
      for(var r=0;r<all.length;r++){ if(!xs[r] && !all[r][0]) continue; labels.push(fmtDate(xs[r])); for(var cc=0;cc<all[r].length;cc++) series[cc].data.push(toNum(all[r][cc])); }
      return {labels:labels, series:series, unit:detectUnit(headers, all)};
    }
    //======= Penjelasan: getBar - Olah data untuk grafik batang, sama pola kayak getLine =======
    function getBar(xRange, headerRange, dataRange){
      var xs = rangeVals(xRange).map(function(r){ return r[0]; });
      var headers = rangeVals(headerRange)[0];
      var all = rangeVals(dataRange);
      var labels=[]; var series=[];
      for(var c=0;c<headers.length;c++) series.push({label:String(headers[c]||('Data '+(c+1))), data:[]});
      for(var r=0;r<all.length;r++){ labels.push(fmtDate(xs[r])); for(var cc=0;cc<all[r].length;cc++) series[cc].data.push(toNum(all[r][cc])); }
      return {labels:labels, series:series, unit:detectUnit(headers, all)};
    }
    //======= Penjelasan: getDonut - Olah data untuk grafik donat, header + 1 baris data + center =======
    function getDonut(headerRange, dataRange, centerCell){
      var headers = rangeVals(headerRange)[0];
      var valsRow = rangeVals(dataRange)[0];
      var labels=[], data=[];
      for(var i=0;i<valsRow.length;i++){ if(valsRow[i]==null||valsRow[i]==='') continue; labels.push(String(headers[i]||('Data '+(i+1)))); data.push(toNum(valsRow[i])); }
      return {labels:labels, data:data, center:String(val(centerCell)||''), unit:detectUnit(headers, [valsRow])};
    }
    //======= Penjelasan: d1-d4 - 4 set data grafik, tiap set ada g1 line, g2 bar, g3 donut =======
    var d1={g1:getLine('E3:E14','F2:I2','F3:I14'), g2:getBar('E18:E21','F17:I17','F18:I21'), g3:getDonut('F23:I23','F24:I24','E24')};
    var d2={g1:getLine('K3:K14','L2:N2','L3:N14'), g2:getBar('K18:K21','L17:N17','L18:N21'), g3:getDonut('L23:N23','L24:N24','K24')};
    var d3={g1:getLine('P3:P14','Q2:S2','Q3:S14'), g2:getBar('P18:P21','Q17:S17','Q18:S21'), g3:getDonut('Q23:S23','Q24:S24','P24')};
    var d4={g1:getLine('U3:U14','V2:V2','V3:V14'), g2:getBar('U18:U21','V17:V17','V18:V21'), g3:getDonut('V23:V23','V24:V24','U24')};
    //======= Penjelasan: Return Dashboard - Kirim semua box + data1-4 ke Index.html =======
    return {status:'ok', box1:box1, box2:box2, box3:box3, box4:box4, box5:box5, box6:box6, box7:box7, data1:d1, data2:d2, data3:d3, data4:d4};
  }catch(e){ return {status:'error', message:e.toString()}; }
}

//======= Penjelasan: Sub Menu Data Kreator - Sheet Data Kreator, kolom A-H =======
//======= Kolom: A Agency, B ID Kreator, C Username, D Bulan Ini, E Bulan Lalu, F WA, G Kontrak Berakhir, H Masa =======





//===================================================  DATA KREATOR ===================================================



function _getDataKreatorSheet_(){
  var s=SpreadsheetApp.openById('1J5sBq0de2-VcgqvHMZq6yd-kkjPZdeMYKr-Apr1pumg');
  var h=s.getSheetByName('Data Kreator');
  if(h) return h;
  var a=s.getSheets();
  for(var i=0;i<a.length;i++){
    if(a[i].getName().toLowerCase().includes('data kreator')) return a[i];
  }
  return null;
}
function _formatTanggalGS_(v){
  if(v==null||v==='') return '-';
  try{
    if(v instanceof Date){
      if(isNaN(v.getTime())) return '-';
      return ('0'+v.getDate()).slice(-2)+'/'+('0'+(v.getMonth()+1)).slice(-2)+'/'+v.getFullYear();
    }
    var s=String(v).trim();
    if(!s||s==='-') return '-';
    var b={jan:0,januari:0,january:0,feb:1,februari:1,february:1,mar:2,maret:2,march:2,apr:3,april:3,mei:4,may:4,jun:5,juni:5,june:5,jul:6,juli:6,july:6,agu:7,agustus:7,aug:7,august:7,sep:8,september:8,sept:8,okt:9,oktober:9,oct:9,october:9,nov:10,november:10,des:11,desember:11,dec:11,december:11};
    if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)){
      var p=s.split('/');
      return ('0'+p[0]).slice(-2)+'/'+('0'+p[1]).slice(-2)+'/'+p[2];
    }
    var mY=s.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/);
    if(mY) return ('0'+mY[3]).slice(-2)+'/'+('0'+mY[2]).slice(-2)+'/'+mY[1];
    var mD=s.match(/^(\d{1,2})[\-\.](\d{1,2})[\-\.](\d{4})/);
    if(mD) return ('0'+mD[1]).slice(-2)+'/'+('0'+mD[2]).slice(-2)+'/'+mD[3];
    var m1=s.toLowerCase().match(/^(\d{1,2})[\s\-\/,]+([a-z]+)[\s\-\/,]+(\d{4})/);
    if(m1&&b[m1[2]]!=null) return ('0'+m1[1]).slice(-2)+'/'+('0'+(b[m1[2]]+1)).slice(-2)+'/'+m1[3];
    var m2=s.toLowerCase().match(/^([a-z]+)[\s\-\/,]+(\d{1,2})[\s\-\/,]+(\d{4})/);
    if(m2&&b[m2[1]]!=null) return ('0'+m2[2]).slice(-2)+'/'+('0'+(b[m2[1]]+1)).slice(-2)+'/'+m2[3];
    var m3=s.toLowerCase().match(/^([a-z]+)[\s\-\/,]+(\d{4})$/);
    if(m3&&b[m3[1]]!=null) return '01/'+('0'+(b[m3[1]]+1)).slice(-2)+'/'+m3[2];
    var d=new Date(s);
    if(!isNaN(d.getTime())) return ('0'+d.getDate()).slice(-2)+'/'+('0'+(d.getMonth()+1)).slice(-2)+'/'+d.getFullYear();
    return s;
  }catch(e){ return String(v); }
}
function searchDataKreator(q){
  try{
    var s=_getDataKreatorSheet_(); if(!s) return [];
    var last=s.getLastRow(); if(last<2) return [];
    var d=s.getRange(2,1,last-1,9).getDisplayValues();
    q=(q||'').toLowerCase().trim(); if(q.length<2) return [];
    var r=[];
    for(var i=0;i<d.length;i++){
      var id=(d[i][1]||'').trim(),u=(d[i][2]||'').trim();
      if(!id&&!u) continue;
      if(id.toLowerCase().includes(q)||u.toLowerCase().includes(q)){
        r.push({row:i+2,agency:d[i][0],idKreator:id,username:u,bulanIni:d[i][3],bulanLalu:d[i][4],wa:d[i][5],kontrakBerakhir:_formatTanggalGS_(d[i][6]),masaBerakhir:d[i][7],gabung:_formatTanggalGS_(d[i][8])});
        if(r.length>=10) break;
      }
    }
    return r;
  }catch(e){ return [{username:'ERROR:'+e.message}]; }
}
function getDataKreatorDetail(k){
  try{
    var s=_getDataKreatorSheet_(); if(!s) return null;
    var last=s.getLastRow(); if(last<2) return null;
    var d=s.getRange(2,1,last-1,9).getDisplayValues();
    k=(k||'').toLowerCase().trim();
    for(var i=0;i<d.length;i++){
      if((d[i][1]||'').toLowerCase().trim()==k||(d[i][2]||'').toLowerCase().trim()==k){
        return {row:i+2,agency:d[i][0],idKreator:d[i][1],username:d[i][2],bulanIni:d[i][3],bulanLalu:d[i][4],wa:d[i][5],kontrakBerakhir:_formatTanggalGS_(d[i][6]),masaBerakhir:d[i][7],gabung:_formatTanggalGS_(d[i][8])};
      }
    }
    return null;
  }catch(e){ return null; }
}
function getDataKreatorTable50(){
  try{
    var s=_getDataKreatorSheet_(); if(!s) return {status:'error',message:'not found'};
    var l=Math.min(s.getLastRow(),51);
    if(l<2) return {status:'ok',data:[]};
    return {status:'ok',data:s.getRange(1,1,l,9).getDisplayValues()};
  }catch(e){ return {status:'error',message:e.message}; }
}
function getDataKreatorTableAll(){
  try{
    var s=_getDataKreatorSheet_(); if(!s) return {status:'error',message:'not found'};
    var l=s.getLastRow(); if(l<2) return {status:'ok',data:[]};
    var raw=s.getRange(1,1,l,9).getDisplayValues(),out=[];
    out.push(raw[0]);
    for(var i=1;i<raw.length;i++){
      var b=String(raw[i][1]==null?'':raw[i][1]).trim();
      if(!b) continue;
      var row=raw[i].slice();
      row[6]=_formatTanggalGS_(row[6]);
      row[7]=String(row[7]==null?'':row[7]).trim();
      row[8]=_formatTanggalGS_(row[8]);
      out.push(row);
    }
    out.sort(function(a,b){ return String(a[2]||'').localeCompare(String(b[2]||'')); });
    return {status:'ok',data:out};
  }catch(e){ return {status:'error',message:e.message}; }
}
function getDkSuggestAll(){
  try{
    var s=_getDataKreatorSheet_(); if(!s) return [];
    var last=s.getLastRow(); if(last<2) return [];
    var d=s.getRange(2,1,last-1,9).getDisplayValues(),r=[];
    for(var i=0;i<d.length;i++){
      var id=(d[i][1]||'').trim(),u=(d[i][2]||'').trim();
      if(!id&&!u) continue;
      r.push({row:i+2,agency:d[i][0],idKreator:id,username:u,nama:u,bulanIni:d[i][3],bulanLalu:d[i][4],wa:d[i][5],kontrakBerakhir:_formatTanggalGS_(d[i][6]),masaBerakhir:d[i][7],gabung:_formatTanggalGS_(d[i][8])});
    }
    return r;
  }catch(e){ return []; }
}
function getAvatarTikTok(username){
  try{
    username=String(username||'').trim().replace(/^@+/,'');
    if(!username) return '';
    var cache=CacheService.getScriptCache();
    var key='avatar_'+username.toLowerCase();
    var cached=cache.get(key);
    if(cached) return cached;
    var avatar='';
    try{
      var res=UrlFetchApp.fetch('https://www.tiktok.com/@'+encodeURIComponent(username),{muteHttpExceptions:true,followRedirects:true,headers:{'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36','Accept-Language':'en-US,en;q=0.9,id;q=0.8'}});
      var html=res.getContentText();
      var pats=[/"avatarLarger"\s*:\s*"([^"]+)"/,/"avatarMedium"\s*:\s*"([^"]+)"/,/"avatarThumb"\s*:\s*"([^"]+)"/,/<meta property="og:image" content="([^"]+)"/];
      for(var i=0;i<pats.length;i++){
        var m=html.match(pats[i]);
        if(m&&m[1]&&m[1].indexOf('http')>-1){ avatar=m[1].replace(/\\u002F/g,'/').replace(/\\/g,''); break; }
      }
    }catch(e){}
    if(!avatar){
      try{
        var u2='https://unavatar.io/tiktok/'+encodeURIComponent(username);
        var r2=UrlFetchApp.fetch(u2,{muteHttpExceptions:true,followRedirects:true});
        if(r2.getResponseCode()===200){
          var ct=String(r2.getHeaders()['Content-Type']||'');
          if(ct.indexOf('image')>-1) avatar=u2;
        }
      }catch(e){}
    }
    if(!avatar) avatar='https://unavatar.io/tiktok/'+encodeURIComponent(username);
    if(avatar){ try{ cache.put(key,avatar,21600); }catch(e){} }
    return avatar;
  }catch(e){ return ''; }
}
function clearAvatarCache(){
  CacheService.getScriptCache().removeAll(['avatar_highcore_corporation']);
  return 'ok';
}


// =========================================================== DATA BASE ========================================================================

function getSS(){ return SpreadsheetApp.getActiveSpreadsheet(); }

// Semua data sekarang di 1 Sheet "Data Base"
function GSH(){ var s=getSS(); return s.getSheetByName('Data Base') || s.getSheets()[0]; }

// ---- KONTAK CREATOR : C-D-E ----
function getDataBaseKontakTable30(){ try{ var h=GSH(),l=h.getLastRow(); if(l<2) return{sheetName:h.getName(),total:0,data:[]}; var v=h.getRange(2,3,Math.min(l-1,1999),3).getValues(),d=[]; for(var i=v.length-1;i>=0;i--) if(v[i][0]||v[i][2]){ d.push(['','',v[i][0],v[i][1],v[i][2],'']); if(d.length>=30) break; } return{sheetName:h.getName(),total:l-1,data:d}; }catch(e){ return{sheetName:'ERROR',total:0,data:[]}; } }
function getDataBaseKontakSearch(q){ try{ var h=GSH(),l=h.getLastRow(); if(l<2) return{result:null}; var v=h.getRange(2,3,l-1,3).getValues(); q=String(q).toLowerCase().trim(); for(var i=0;i<v.length;i++) if(String(v[i][0]).toLowerCase().includes(q)||String(v[i][1]).toLowerCase().includes(q)) return{result:{id:v[i][0],username:v[i][1],wa:v[i][2]}}; return{result:null}; }catch(e){ return{result:null}; } }
function getDataBaseKontakAll(){ try{ var h=GSH(),l=h.getLastRow(); if(l<2) return{sheetName:h.getName(),total:0,data:[]}; var v=h.getRange(2,3,l-1,3).getValues(),d=[]; for(var i=v.length-1;i>=0;i--) if(v[i][0]||v[i][2]){ d.push(['','',v[i][0],v[i][1],v[i][2],'']); } return{sheetName:h.getName(),total:l-1,data:d}; }catch(e){ return{sheetName:'ERROR',total:0,data:[]}; } }
function getDataBaseKontakSearchAll(q){ try{ var h=GSH(),l=h.getLastRow(); if(l<2) return{sheetName:h.getName(),total:0,data:[]}; var v=h.getRange(2,3,l-1,3).getValues(),d=[]; q=String(q||'').toLowerCase().trim(); if(!q) return getDataBaseKontakAll(); for(var i=v.length-1;i>=0;i--) if(String(v[i][0]).toLowerCase().includes(q)||String(v[i][1]).toLowerCase().includes(q)) d.push(['','',v[i][0],v[i][1],v[i][2],'']); return{sheetName:h.getName(),total:d.length,data:d}; }catch(e){ return{sheetName:'ERROR',total:0,data:[]}; } }
function saveDataBaseKontakBatch(r){ try{ var h=GSH(),c=h.getRange('E2:E2000').getValues(),a=-1; for(var i=c.length-1;i>=0;i--) if(c[i][0]&&String(c[i][0]).trim()!=''){ a=i; break; } var s=a==-1?2:a+3,o=0; for(var k=0;k<r.length;k++){ var id=String(r[k][0]||'').trim(),wa=String(r[k][1]||'').trim(); if(!id&&!wa) continue; h.getRange(s,3).setValue(id); h.getRange(s,5).setValue(wa); s++; o++; } return{status:'ok',saved:o}; }catch(e){ return{status:'error',message:e.message}; } }

// ---- AGENCY : R-S-T-U-V-W-X (kolom 18-24) ----
const A_=()=>GSH();
function searchAgencyContact(q){ try{ q=(q||'').toLowerCase().trim(); if(!q) return[]; let sh=A_(); if(!sh) return[]; let v=sh.getRange('R2:X2000').getValues(),o=[]; for(let r of v){ let a=r[1]+'',n=r[3]+''; if((a||n)&&(a.toLowerCase().includes(q)||n.toLowerCase().includes(q))){ o.push({filter:r[0],agency:a,email:r[2],nama:n,wa:r[4],management:a}); if(o.length>=10) break; } } return o; }catch(e){ return[]; } }
function getAgencyContactDetail(q) {
  try{
  q = (q || '').toLowerCase().trim();
  var sh=A_(); if(!sh) return { found: false };
  var vals = sh.getRange('R2:X2000').getValues();
  for (var i = 0; i < vals.length; i++) {
    var r = vals[i];
    var a = r[1] + '';
    var n = r[3] + '';
    if (a.toLowerCase().includes(q) || n.toLowerCase().includes(q)) {
      var out = {};
      out.filter = r[0];
      out.agency = a;
      out.email = r[2];
      out.nama = n;
      out.wa = r[4];
      out.bank = r[5];
      out.norek = r[6];
      return { found: true, data: out };
    }
  }
  return { found: false };
  }catch(e){ return { found: false }; }
}
function saveAgencyContactBatch(r){ try{ if(!r?.length) return{status:'error',message:'No data'}; let s=A_(); if(!s) return{status:'error',message:'Sheet Data Base tidak ketemu'}; let c=s.getRange('T2:T2000').getValues(),l=0; c.forEach((x,i)=>{ if((x[0]+'').trim()) l=i+1; }); let sr=l+2; s.getRange(sr,18,r.length,7).setValues(r); return{status:'ok',saved:r.length,startRow:sr}; }catch(e){ return{status:'error',message:e.message}; } }
function getAgencyLast5(){ try{ let sh=A_(); if(!sh) return{data:[]}; let v=sh.getRange('R2:X2000').getValues().filter(x=>(x[2]+'').trim()); return{data:v.slice(-5).reverse()}; }catch(e){ return{data:[]}; } }
function getAgencyAll() {
  try{
  var sh=A_(); if(!sh) return { data: [] };
  var v = sh.getRange('R2:X2000').getValues();
  var d = [];
  for (var i = v.length - 1; i >= 0; i--) {
    if ((v[i][2] + '').trim()) {
      d.push(v[i]);
    }
  }
  return { data: d };
  }catch(e){ return { data: [] }; }
}

// ---- REKENING : H-I-J-K-L-M-N-O (kolom 8-15) ----
function getRekSheet_SIMPLE(){ return GSH(); }
function getDefaultRow_AFNICHAN(){ return[100079,'Finance & Accounting','afnichan@highcore.id','Afni Chan','BCA','1234567890','6281234567890','Staff Finance']; }
function getRekeningLast5(){ try{ let sh=getRekSheet_SIMPLE(); if(!sh) return{data:[getDefaultRow_AFNICHAN()]}; let v=sh.getRange('H2:O2000').getValues(),f=v.filter(r=>r[1]||r[2]||r[3]); if(!f.length){ let a=v.find(r=>String(r[2]||'').toLowerCase()=='afnichan@highcore.id'); return{data:[a||getDefaultRow_AFNICHAN()]}; } return{data:f.slice(-5).reverse()}; }catch(e){ return{data:[getDefaultRow_AFNICHAN()]}; } }
function getRekeningAllData() {
  try {
    var sh = getRekSheet_SIMPLE();
    if (!sh) return { data: [getDefaultRow_AFNICHAN()] };
    var v = sh.getRange('H2:O2000').getValues();
    var d = [];
    for (var i = v.length - 1; i >= 0; i--) {
      if (v[i][1] || v[i][2] || v[i][3]) {
        d.push(v[i]);
      }
    }
    return { data: d };
  } catch (e) {
    return { data: [getDefaultRow_AFNICHAN()] };
  }
}
function searchRekeningContact(q){ try{ let sh=getRekSheet_SIMPLE(); if(!sh) return[{ket:'Finance & Accounting',email:'afnichan@highcore.id',nama:'Afni Chan',nomer:100079}]; let v=sh.getRange('H2:O2000').getValues(); q=(q||'').toLowerCase().trim(); if(q.length<2) return[{ket:'Finance & Accounting',email:'afnichan@highcore.id',nama:'Afni Chan',nomer:100079,bank:'BCA',norek:'1234567890',wa:'6281234567890',jabatan:'Staff Finance'}]; let o=[]; for(let r of v){ if(!r[1]&&!r[2]&&!r[3]) continue; let b=(r[1]+' '+r[2]+' '+r[3]).toLowerCase(); if(b.includes(q)){ o.push({ket:r[1],email:r[2],nama:r[3],nomer:r[0],bank:r[4],norek:r[5],wa:r[6],jabatan:r[7]}); if(o.length>=8) break; } } if(!o.length&&'afnichan@highcore.id'.includes(q)) o.push({ket:'Finance & Accounting',email:'afnichan@highcore.id',nama:'Afni Chan',nomer:100079}); return o; }catch{ return[]; } }
function getRekeningDetail(q){ try{ let sh=getRekSheet_SIMPLE(),d=getDefaultRow_AFNICHAN(); if(!sh) return{found:1,row:d}; let v=sh.getRange('H2:O2000').getValues(); q=(q||'').toLowerCase().trim(); if(!q) return{found:1,row:d}; for(let r of v){ if(!r[1]&&!r[2]&&!r[3]) continue; if(String(r[1]).toLowerCase().trim()==q||String(r[2]).toLowerCase().trim()==q||String(r[3]).toLowerCase().trim()==q) return{found:1,row:r}; } for(let r of v){ if(!r[1]&&!r[2]&&!r[3]) continue; if((r[1]+' '+r[2]+' '+r[3]).toLowerCase().includes(q)) return{found:1,row:r}; } let a=v.find(r=>String(r[2]||'').toLowerCase()=='afnichan@highcore.id'); return{found:1,row:a||d}; }catch{ return{found:1,row:getDefaultRow_AFNICHAN()}; } }
function getRekeningByEmail(q){ return getRekeningDetail(q||'afnichan@highcore.id'); }
function saveRekeningBatch(r){ try{ let sh=getRekSheet_SIMPLE(); if(!sh) return{status:'error'}; let last=0,max=100078; sh.getRange('H2:H2000').getValues().forEach(a=>{ let n=parseInt(a[0]); if(!isNaN(n)&&n>max) max=n; }); sh.getRange('J2:J2000').getValues().forEach((a,i)=>{ if((a[0]+'').trim()) last=i+1; }); let s=r.map(x=>[++max,x[0],x[1],x[2],x[3],x[4],x[5],x[6]]); if(s.length) sh.getRange(last+2,8,s.length,8).setValues(s); return{status:'ok',saved:s.length}; }catch(e){ return{status:'error',message:e.message}; } }
function getRekeningAll(){ try{ let sh=GSH(); if(!sh) return{data:[]}; let v=sh.getRange('H2:O2000').getValues().filter(r=>r[1]||r[2]||r[3]); return{data:v}; }catch(e){ return{data:[]}; } }

//=============================================================== Lembar Kerja ========================================================

function getSS_LK(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function GSH_LK(){ var s=getSS_LK(); return s.getSheetByName('Data Base')||s.getSheets()[0]; }
function cariBarisKosongPertama_LK(sh,col,startRow){ startRow=startRow||2; var last=sh.getLastRow(); if(last<startRow) return startRow; var vals=sh.getRange(startRow,col,last-startRow+1,1).getValues(); for(var i=0;i<vals.length;i++){ var v=String(vals[i][0]||'').trim(); if(v==='') return startRow+i; } return last+1; }
function saveLembarKontakBatch(rows){ try{ var sh=GSH_LK(),saved=0; for(var i=0;i<rows.length;i++){ var id=String(rows[i][0]||'').trim(),wa=String(rows[i][1]||'').trim(); if(!id&&!wa) continue; var r=cariBarisKosongPertama_LK(sh,4,2); if(id) sh.getRange(r,3).setValue(id); if(wa) sh.getRange(r,5).setValue(wa); saved++; } return{status:'ok',saved:saved}; }catch(e){ return{status:'error',message:e.message}; } }
function saveLembarAgencyBatch(rows){ try{ var sh=GSH_LK(),saved=0,startRow=0; for(var i=0;i<rows.length;i++){ var f=rows[i]; if(!f.some(function(x){ return String(x||'').trim()!==''; })) continue; var r=cariBarisKosongPertama_LK(sh,18,2); for(var c=0;c<7;c++){ if(String(f[c]||'').trim()!=='') sh.getRange(r,18+c).setValue(f[c]); } if(!startRow) startRow=r; saved++; } return{status:'ok',saved:saved,startRow:startRow}; }catch(e){ return{status:'error',message:e.message}; } }
function saveLembarRekeningBatch(rows){ try{ var sh=GSH_LK(),saved=0; for(var i=0;i<rows.length;i++){ var v=rows[i],data=[v[1],v[2],v[3],v[4],v[5],v[6],v[7]]; if(!data.some(function(x){ return String(x||'').trim()!==''; })) continue; var r=cariBarisKosongPertama_LK(sh,9,2); for(var c=0;c<7;c++){ if(String(data[c]||'').trim()!=='') sh.getRange(r,9+c).setValue(data[c]); } saved++; } return{status:'ok',saved:saved}; }catch(e){ return{status:'error',message:e.message}; } }


//========================================================== KAS PIC LISTY =================================================================


function getKasPICData(){
  try{
    var s = getSS().getSheetByName('Kas PIC Listy');
    if(!s) return {totalSaldo:'-', pengeluaranBulanIni:'-', pengeluaranBulanLalu:'-', bulan:[], kuartal:[], tahun:[]};
    function fil(r){ return String(r[0]).trim()!=''; }
    return {
      totalSaldo: s.getRange('U2').getDisplayValue(),
      pengeluaranBulanIni: s.getRange('V2').getDisplayValue(),
      pengeluaranBulanLalu: s.getRange('W2').getDisplayValue(),
      bulan: s.getRange('L2:M13').getDisplayValues().filter(fil),
      kuartal: s.getRange('O2:P13').getDisplayValues().filter(fil),
      tahun: s.getRange('R2:S13').getDisplayValues().filter(fil)
    };
  }catch(e){ return {totalSaldo:'Error', pengeluaranBulanIni:'-', pengeluaranBulanLalu:'-', bulan:[], kuartal:[], tahun:[]}; }
}

function getKasRecentData(opt){
  try{
    var s = getSS().getSheetByName('Kas PIC Listy');
    if(!s) return {status:'fail'};
    var v = s.getRange('A1:F'+s.getLastRow()).getDisplayValues(), r=[];
    for(var i=1;i<v.length;i++) if(String(v[i][0]).trim()!='' || String(v[i][1]).trim()!='') r.push({row:i+1, data:v[i]});
    var viewAll = opt && opt.viewAll;
    var rows = viewAll ? r.reverse() : r.slice(-6).reverse();
    return {status:'ok', rows:rows};
  }catch(e){ return {status:'fail'}; }
}

function addKasTransaksi(o){
  try{
    var s = getSS().getSheetByName('Kas PIC Listy');
    if(!s) return {status:'fail', message:'Sheet Kas PIC Listy tidak ketemu'};
    var last = s.getLastRow() + 1;
    var vals = s.getRange('A2:B2000').getDisplayValues();
    for(var i=0;i<vals.length;i++){ if(String(vals[i][0]).trim()=='' && String(vals[i][1]).trim()==''){ last = i+2; break; } }
    var debit = Number(o.debit||0), kredit = Number(o.kredit||0);
    s.getRange(last,1,1,5).setValues([[o.tanggal||'', o.keterangan||'', o.jenis||'', debit||'', kredit||'']]);
    return {status:'ok', row:last};
  }catch(e){ return {status:'fail', message:e.message}; }
}

function editKasTransaksi(data){
  try{
    var s = getSS().getSheetByName('Kas PIC Listy');
    if(!s) return {status:'fail', message:'Sheet Kas PIC Listy tidak ketemu'};
    if(!data || !data.row) return {status:'fail', message:'Row tidak ketemu'};
    var debit = Number(data.debit||0), kredit = Number(data.kredit||0);
    s.getRange(data.row,1,1,5).setValues([[data.tanggal||'', data.keterangan||'', data.jenis||'', debit||'', kredit||'']]);
    return {status:'ok', row:data.row};
  }catch(e){ return {status:'fail', message:e.message}; }
}

function deleteKasRow(r){
  try{ getSS().getSheetByName('Kas PIC Listy').getRange('A'+r+':E'+r).clearContent(); return {status:'ok'}; }
  catch(e){ return {status:'fail', message:e.message}; }
}

function findKontrakHostSheet(){
  var s = getSS().getSheets();
  for(var i=0;i<s.length;i++) if(s[i].getName().toLowerCase().indexOf('kontrak host')>-1) return s[i];
  return null;
}

function getKontrakHostSummary(){
  try{
    var sh = findKontrakHostSheet(); if(!sh) return {status:'fail'};
    var v = sh.getRange('N2:Q2').getDisplayValues()[0];
    var kurang = v[0] === '' ? 0 : Number(v[0]);
    var perpanjang = v[1] === '' ? 0 : Number(v[1]);
    var pemutusan = v[2] === '' ? 0 : Number(v[2]);
    var colI = sh.getRange('I2:I2000').getDisplayValues();
    var terisiI = 0;
    for(var i=0;i<colI.length;i++){
      if(String(colI[i][0]||'').trim() !== '') terisiI++;
    }
    var belum = kurang - terisiI;
    if(belum < 0) belum = 0;
    return {status:'ok', kurang30:kurang, perpanjang:perpanjang, pemutusan:pemutusan, belum:belum};
  }catch(e){ return {status:'fail'}; }
}

//====================================================== Kontrak Host =============================================================


function getSS(){
  return SpreadsheetApp.getActiveSpreadsheet();
}

function findKontrakHostSheet(){
  var ss = getSS();
  var sh = ss.getSheetByName('Kontrak Host');
  if(sh) return sh;
  var all = ss.getSheets();
  for(var i=0;i<all.length;i++){
    var nm = String(all[i].getName()||'').toLowerCase();
    if(nm.indexOf('kontrak')>-1 && nm.indexOf('host')>-1) return all[i];
  }
  for(var j=0;j<all.length;j++){
    var nm2 = String(all[j].getName()||'').toLowerCase();
    if(nm2.indexOf('kontrak')>-1) return all[j];
  }
  return null;
}

function getKontrakHostSummary(){
  try{
    var sh = findKontrakHostSheet(); if(!sh) return {status:'fail'};
    var data = sh.getRange('A2:L2000').getDisplayValues();
    var kurang30 = 0, perpanjang = 0, pemutusan = 0;
    for(var i=0;i<data.length;i++){
      if(String(data[i].join('')).trim()=='') continue;
      var h = String(data[i][7]||'').trim().toLowerCase();
      var ket = String(data[i][8]||'').trim();
      var ang = h.match(/(\d+)/);
      if(ang && Number(ang[0]) < 30 && h.indexOf('expired')===-1) kurang30++;
      if(ket==='Perpanjang Kontrak') perpanjang++;
      else if(ket==='Pemutusan Kontrak') pemutusan++;
    }
    return {status:'ok', kurang30:kurang30, perpanjang:perpanjang, pemutusan:pemutusan};
  }catch(e){ return {status:'fail'}; }
}

function arsipkanKontrak6Hari(){
  try{
    var ss = getSS();
    var sh = findKontrakHostSheet(); if(!sh) return 0;
    var dst = ss.getSheetByName('Arsip Data');
    if(!dst){
      dst = ss.insertSheet('Arsip Data');
      dst.getRange(1,1,1,12).setValues(sh.getRange(1,1,1,12).getDisplayValues());
    }
    var data = sh.getRange('A2:L2000').getDisplayValues();
    var arsip = dst.getRange('A2:L2000').getDisplayValues();
    var sudahAda = {};
    for(var a=0;a<arsip.length;a++){
      if(String(arsip[a][0]||'').trim()=='') continue;
      var keyA = String(arsip[a][1]||'').trim() + '|' + String(arsip[a][2]||'').trim() + '|' + String(arsip[a][6]||'').trim();
      sudahAda[keyA] = true;
    }
    var mauCopy = [];
    for(var i=0;i<data.length;i++){
      var h = String(data[i][7]||'').trim().toLowerCase();
      if(h !== '6 hari lagi') continue;
      if(String(data[i][0]||'').trim()=='' && String(data[i][1]||'').trim()=='') continue;
      var k = String(data[i][1]||'').trim() + '|' + String(data[i][2]||'').trim() + '|' + String(data[i][6]||'').trim();
      if(!sudahAda[k]){
        mauCopy.push(data[i]);
        sudahAda[k] = true;
      }
    }
    if(mauCopy.length===0) return 0;
    var targetRow = 2;
    var colA = dst.getRange('A2:A2000').getDisplayValues();
    for(var r=0;r<colA.length;r++){
      if(String(colA[r][0]||'').trim()==''){ targetRow = r + 2; break; }
      targetRow = r + 3;
    }
    dst.getRange(targetRow,1,mauCopy.length,12).setValues(mauCopy);
    return mauCopy.length;
  }catch(e){ return 0; }
}

function getKontrakHostTable(){
  try{
    arsipkanKontrak6Hari();
    var sh = findKontrakHostSheet(); if(!sh) return {status:'fail'};
    var data = sh.getRange(1,1,sh.getLastRow(),12).getDisplayValues();
    return {status:'ok', data:data};
  }catch(e){ return {status:'fail'}; }
}

function saveKontrakHostManual(row, i, j, k){
  try{
    var sh = findKontrakHostSheet(); if(!sh) return {status:'fail'};
    i = String(i||''); j = String(j||''); k = String(k||'');
    if(i=='' && j=='' && k==''){ sh.getRange(row,9,1,3).clearContent(); return {status:'cleared'}; }
    sh.getRange(row,9,1,3).setValues([[i,j,k]]);
    return {status:'ok'};
  }catch(e){ return {status:'fail'}; }
}

function getSistemTable(){
  try{
    var sh = findKontrakHostSheet(); if(!sh) return {status:'fail'};
    var data = sh.getRange(2, 19, 1999, 8).getDisplayValues();
    var today = new Date(); today.setHours(0,0,0,0);
    var bln = {jan:0,januari:0,january:0,feb:1,februari:1,february:1,mar:2,maret:2,march:2,apr:3,april:3,mei:4,may:4,jun:5,juni:5,june:5,jul:6,juli:6,july:6,agu:7,agustus:7,august:7,sep:8,september:8,okt:9,oktober:9,october:9,nov:10,november:10,des:11,desember:11,december:11,dec:11};
    for(var i=0;i<data.length;i++){
      var yStr = String(data[i][6]||'').trim();
      if(!yStr) continue;
      var d = null;
      var s = yStr.split('T')[0].trim();
      var p = s.split('-');
      if(p.length===3 && p[0].length===4){
        d = new Date(Number(p[0]), Number(p[1])-1, Number(p[2]));
      } else {
        var q = s.split('/');
        if(q.length===3){
          if(Number(q[0])>12) d = new Date(Number(q[2]), Number(q[1])-1, Number(q[0]));
          else d = new Date(Number(q[2]), Number(q[0])-1, Number(q[1]));
        } else {
          var m = s.match(/(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})/);
          if(m && bln[m[2].toLowerCase()]!==undefined){
            d = new Date(Number(m[3]), bln[m[2].toLowerCase()], Number(m[1]));
          } else {
            d = new Date(s);
          }
        }
      }
      if(!d || isNaN(d.getTime())) continue;
      d.setHours(0,0,0,0);
      var sisa = Math.round((d.getTime() - today.getTime()) / 86400000);
      data[i][7] = sisa < 0 ? 'Expired' : sisa + ' hari lagi';
    }
    return {status:'ok', data:data};
  }catch(e){ return {status:'fail', msg:e.toString()}; }
}

function tambahDocumentBaru(d){
  try{
    var sh = findKontrakHostSheet(); if(!sh) return {status:'error', message:'Sheet Kontrak Host tidak ketemu'};
    d = d || {};
    var bln = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
    function fmtTanggal(v){
      if(!v) return '';
      var s = String(v).trim();
      var p = s.split('-');
      if(p.length===3 && p[0].length===4){
        return Number(p[2]) + '-' + bln[Number(p[1])-1] + '-' + p[0];
      }
      return s;
    }
    var vals = sh.getRange(2, 19, 1999, 7).getDisplayValues();
    var target = 0;
    for(var i=0;i<vals.length;i++){
      var kosong = true;
      for(var j=0;j<7;j++){ if(String(vals[i][j]||'').trim()!==''){ kosong=false; break; } }
      if(kosong){ target = i + 2; break; }
    }
    if(!target) return {status:'error', message:'S2:Y2000 sudah penuh'};
    sh.getRange(target, 19, 1, 7).setValues([[
      fmtTanggal(d.s),
      String(d.t||''),
      String(d.u||''),
      String(d.v||''),
      String(d.w||''),
      fmtTanggal(d.x),
      fmtTanggal(d.y)
    ]]);
    return {status:'ok', row:target};
  }catch(e){
    return {status:'error', message:e.toString()};
  }
}

function editDocumentBaru(d){
  try{
    var sh = findKontrakHostSheet(); if(!sh) return {status:'error', message:'Sheet Kontrak Host tidak ketemu'};
    d = d || {};
    var row = Number(d.row||0);
    if(!row || row < 2 || row > 2000) return {status:'error', message:'Pilih document dulu lewat pencarian'};
    var bln = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
    function fmtTanggal(v){
      if(!v) return '';
      var s = String(v).trim();
      var p = s.split('-');
      if(p.length===3 && p[0].length===4){
        return Number(p[2]) + '-' + bln[Number(p[1])-1] + '-' + p[0];
      }
      return s;
    }
    sh.getRange(row, 19, 1, 7).setValues([[
      fmtTanggal(d.s),
      String(d.t||''),
      String(d.u||''),
      String(d.v||''),
      String(d.w||''),
      fmtTanggal(d.x),
      fmtTanggal(d.y)
    ]]);
    return {status:'ok', row:row};
  }catch(e){
    return {status:'error', message:e.toString()};
  }
}

//============================================== MASTER MATCH ========================================================


function findMasterMatchSheet(){
  var s = getSS().getSheets();
  for(var i = 0; i < s.length; i++){
    if(s[i].getName().toLowerCase().indexOf('master match') > -1) return s[i];
  }
  return null;
}

function isAkhirBulan(up){
  try{
    var s = String(up || '');
    var m = s.match(/(\d{1,2})\s+(Januari|Februari|Maret|April|Mei|Juni|Juli|Agustus|September|Oktober|November|Desember|January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})/i);
    if(!m) return false;
    var day = parseInt(m[1], 10);
    var monName = m[2].toLowerCase();
    var year = parseInt(m[3], 10);
    var map = {januari:0, january:0, februari:1, february:1, maret:2, march:2, april:3, mei:4, may:4, juni:5, june:5, juli:6, july:6, agustus:7, august:7, september:8, oktober:9, october:9, november:10, desember:11, december:11};
    var mon = map[monName];
    if(mon === undefined) return false;
    var last = new Date(year, mon + 1, 0).getDate();
    return day === last;
  }catch(e){
    return false;
  }
}

function getMasterMatchData(){
  try{
    var sh = findMasterMatchSheet();
    if(!sh) return {status:'fail'};
    var v = sh.getRange('A2:C' + sh.getLastRow()).getDisplayValues();
    var data = [];
    for(var i = 0; i < v.length; i++){
      if(String(v[i][0] + v[i][1]).trim() != '') data.push([v[i][0], v[i][1], v[i][2]]);
    }
    data.sort(function(a,b){ return (parseInt(String(b[2]).replace(/[^0-9]/g,''))||0)-(parseInt(String(a[2]).replace(/[^0-9]/g,''))||0); });
    var up = '';
    try{
      up = sh.getRange('D3').getDisplayValue() || '';
    }catch(e){}
    var props = PropertiesService.getScriptProperties();
    var KEY = 'MM_FINAL_POSTER_DATA';
    var liveData = data.slice(0,10);
    var finalData = [];
    var finalUpdate = up;
    if(isAkhirBulan(up)){
      finalData = data.slice(0,4);
      try{
        props.setProperty(KEY, JSON.stringify({data:finalData, update:up}));
      }catch(e){}
    }else{
      try{
        var raw = props.getProperty(KEY);
        if(raw){
          var obj = JSON.parse(raw);
          if(obj && obj.data && obj.data.length >= 4){ finalData = obj.data; finalUpdate = obj.update || up; }
        }
      }catch(e){}
    }
    return {status:'ok', data:liveData, liveData:liveData, posterData:finalData, finalData:finalData, updateData:up, finalUpdate:finalUpdate, isAkhirBulan:isAkhirBulan(up)};
  }catch(e){
    return {status:'fail'};
  }
}
