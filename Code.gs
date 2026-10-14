const PRODUCT_HEADERS = ['id','sku','name','category','size','color','location','cost','price','reorder','active','barcode'];
const LOG_HEADERS = ['id','timestamp','type','operator','details'];
// Google Sheets cells hold at most 50,000 characters. Long ledger entries
// continue in the columns to the right; continuation cells start with '~'.
const CELL_CHUNK = 45000;
// Shared limit on incorrect access codes from ANYONE, per time window.
const LOGIN_MAX_FAILURES = 25;
const LOGIN_WINDOW_MS = 15*60*1000;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('ASB Student Store')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

// Run once from the editor. Existing inventory is never overwritten.
function setupStore() {
  const p = PropertiesService.getScriptProperties();
  let id = p.getProperty('STORE_SHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('ASB Student Store Inventory');
  p.setProperty('STORE_SHEET_ID', ss.getId());
  [['Products',PRODUCT_HEADERS],['Activity',LOG_HEADERS]].forEach(([name,headers])=>{
    let s=ss.getSheetByName(name);
    if(!s) s=ss.insertSheet(name);
    if(name==='Products' && s.getLastRow() && s.getRange(1,1,1,PRODUCT_HEADERS.length).getValues()[0].indexOf('barcode')<0) s.getRange(1,PRODUCT_HEADERS.length).setValue('barcode');
    if(!s.getLastRow()) { s.appendRow(headers); s.setFrozenRows(1); s.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#dceee7'); s.autoResizeColumns(1,headers.length); }
  });
  ['student','adviser'].forEach(role=>{
    if(!p.getProperty(role+'_hash')) {
      const pin=Utilities.getUuid().replace(/-/g,'').slice(0,12);
      p.setProperty(role+'_hash',hash_(pin));
      console.log(role+' access code (save now): '+pin);
    }
  });
  console.log('Inventory spreadsheet: '+ss.getUrl());
}

// To rotate access codes, add NEW_STUDENT_CODE / NEW_ADVISER_CODE in Project
// Settings > Script properties, run this, then log in again.
function rotateAccessCodes() {
  const p=PropertiesService.getScriptProperties(), found=[];
  const codes={student:p.getProperty('NEW_STUDENT_CODE'),adviser:p.getProperty('NEW_ADVISER_CODE')};
  Object.keys(codes).forEach(role=>{ if(codes[role]!=null) { codes[role]=code_(codes[role]); if(codes[role].length<10) throw Error('NEW_'+role.toUpperCase()+'_CODE needs at least 10 characters. Nothing was changed.'); } });
  if(codes.student==null&&codes.adviser==null) throw Error('No NEW_STUDENT_CODE or NEW_ADVISER_CODE found. In Project Settings > Script properties, add them, click "Save script properties", then run this again. Nothing was changed.');
  const s=codes.student!=null?hash_(codes.student):p.getProperty('student_hash'), a=codes.adviser!=null?hash_(codes.adviser):p.getProperty('adviser_hash');
  if(s===a) throw Error('Student and adviser codes must be different. Nothing was changed.');
  Object.keys(codes).forEach(role=>{ if(codes[role]!=null) { p.setProperty(role+'_hash',hash_(codes[role])); p.deleteProperty('NEW_'+role.toUpperCase()+'_CODE'); found.push(role); } });
  p.setProperty('auth_epoch',Utilities.getUuid());
  console.log('Updated '+found.join(' and ')+' code. Script ID: '+ScriptApp.getScriptId());
  p.deleteProperty('login_failures');
}

// Troubleshooting: add a script property CHECK_CODE holding the code you are
// typing, run this, read the log. The property is deleted afterwards.
function checkAccessCode() {
  const p=PropertiesService.getScriptProperties(), raw=p.getProperty('CHECK_CODE');
  if(raw==null) throw Error('Add a CHECK_CODE script property first and click "Save script properties".');
  p.deleteProperty('CHECK_CODE');
  const role=['adviser','student'].find(r=>p.getProperty(r+'_hash')===hash_(code_(raw)));
  console.log('Script ID: '+ScriptApp.getScriptId());
  console.log('Codes saved: student '+(p.getProperty('student_hash')?'yes':'NO')+', adviser '+(p.getProperty('adviser_hash')?'yes':'NO'));
  console.log(role?'MATCH: this is the '+role+' code. If the store still rejects it, the /exec link belongs to a different project or deployment.':'NO MATCH: this is not the current student or adviser code.');
  if(raw!==raw.trim()) console.log('Note: the value had spaces at the start or end; they are ignored.');
}

// Run from the editor if sign-in is paused after too many incorrect codes.
function clearLoginLockout() { PropertiesService.getScriptProperties().deleteProperty('login_failures'); }

function code_(v) { return String(v==null?'':v).trim(); }
function hash_(s) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s))); }
function db_() {
  const id=PropertiesService.getScriptProperties().getProperty('STORE_SHEET_ID');
  if(!id) throw Error('Run setupStore in the Apps Script editor first.');
  return SpreadsheetApp.openById(id);
}
function objects_(sheet) {
  const rows=sheet.getDataRange().getValues(), keys=rows.shift();
  return rows.filter(r=>r[0]).map(r=>Object.fromEntries(keys.map((k,i)=>[k,r[i]])));
}
function auth_(token,adviser) {
  const raw=CacheService.getScriptCache().get('session:'+String(token));
  if(!raw) throw Error('Session expired. Sign in again.');
  const s=JSON.parse(raw), p=PropertiesService.getScriptProperties();
  if(s.epoch!==(p.getProperty('auth_epoch')||'')) throw Error('Access codes changed. Sign in again.');
  if(adviser&&s.role!=='adviser') throw Error('Adviser access required.');
  return s;
}
function login(code,operator) {
  operator=text_(operator,80); if(!operator) throw Error('Enter your name or shift team.');
  code=code_(code); if(!code) throw Error('Enter the access code.');
  const p=PropertiesService.getScriptProperties(), lock=LockService.getScriptLock();
  let role, failed=false; lock.waitLock(30000);
  try {
    // Counts every failed attempt, whatever code was guessed, so guessing many
    // different codes is limited. Stored in properties (cache can be evicted).
    const now=Date.now(); let f=JSON.parse(p.getProperty('login_failures')||'{"count":0,"start":0}');
    if(now-f.start>LOGIN_WINDOW_MS) f={count:0,start:now};
    if(f.count>=LOGIN_MAX_FAILURES) throw Error('Sign-in is paused after too many incorrect codes. Try again in '+Math.ceil((f.start+LOGIN_WINDOW_MS-now)/60000)+' minutes.');
    const h=hash_(code); role=['adviser','student'].find(r=>p.getProperty(r+'_hash')===h);
    if(!role) {f.count++; p.setProperty('login_failures',JSON.stringify(f)); failed=true;}
  } finally {lock.releaseLock();}
  if(failed) {Utilities.sleep(1500); throw Error('Incorrect access code.');}
  const cache=CacheService.getScriptCache();
  const token=Utilities.getUuid()+Utilities.getUuid();
  cache.put('session:'+token,JSON.stringify({role,operator,epoch:p.getProperty('auth_epoch')||''}),21600);
  return {token,role,operator};
}
function logout(token) {CacheService.getScriptCache().remove('session:'+String(token));}
function text_(v,max) { return String(v==null?'':v).trim().slice(0,max); }
function n_(v,integer) {const n=Number(v); if(!Number.isFinite(n)||n<0||(integer&&!Number.isSafeInteger(n))) throw Error('Enter valid nonnegative '+(integer?'whole numbers.':'numbers.')); return n;}
function money_(v) {const n=n_(v,false); if(n>1000000) throw Error('Amount is too large.'); return Math.round(n*100)/100;}
function safe_(v) {return typeof v==='string'&&/^[=+\-@]/.test(v)?"'"+v:v;}
function state_() {
  const ss=db_(), products=objects_(ss.getSheetByName('Products')).map(p=>({...p,active:p.active===true||p.active==='TRUE',stock:0}));
  const byId=Object.fromEntries(products.map(p=>[p.id,p]));
  const logs=ss.getSheetByName('Activity').getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({
    id:String(r[0]),timestamp:r[1] instanceof Date?r[1].toISOString():String(r[1]),type:String(r[2]),operator:String(r[3]),
    details:JSON.parse(r.slice(4).map((c,i)=>{c=String(c);return i&&c[0]==='~'?c.slice(1):c;}).join(''))}));
  logs.forEach(l=>(l.details.lines||[]).forEach(x=>{if(byId[x.id]) byId[x.id].stock+=x.delta;}));
  return {products,logs};
}
function getStore(token) {
  const s=auth_(token), state=state_();
  return {role:s.role,operator:s.operator,products:state.products,logs:state.logs.slice(-200).reverse(),
    totals:state.logs.filter(l=>l.type==='sale').reduce((a,l)=>({units:a.units+l.details.lines.reduce((v,x)=>v-x.delta,0),revenue:Math.round((a.revenue+l.details.revenue)*100)/100}),{units:0,revenue:0})};
}
function saveProduct(token,input) {
  const session=auth_(token,true), lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    auth_(token,true);
    const sheet=db_().getSheetByName('Products'), rows=objects_(sheet);
    const id=input.id||Utilities.getUuid(), old=rows.find(x=>x.id===id);
    if(input.id&&!old) throw Error('Product no longer exists.');
    const item={id,sku:old?old.sku:'ASB-'+Utilities.getUuid().slice(0,8).toUpperCase(),name:text_(input.name,100),category:text_(input.category,60),size:text_(input.size,30),color:text_(input.color,40),location:text_(input.location,80),cost:money_(input.cost),price:money_(input.price),reorder:n_(input.reorder,true),active:input.active!==false,barcode:barcode_(input.barcode)};
    if(!item.name||!item.category) throw Error('Name and category are required.');
    const keys=[item.sku,item.barcode].filter(Boolean).map(barcodeKey_);
    if(rows.some(x=>x.id!==id&&[x.sku,x.barcode].filter(Boolean).some(b=>keys.includes(barcodeKey_(b))))) throw Error('This barcode is already assigned to another product variant.');
    sheet.getRange(2,PRODUCT_HEADERS.length,Math.max(sheet.getLastRow(),1),1).setNumberFormat('@');
    const values=PRODUCT_HEADERS.map(k=>safe_(item[k]));
    if(old) {
      // Match by id against raw rows so blank rows in the sheet can't shift the target.
      const rowNumber=sheet.getRange(1,1,sheet.getLastRow(),1).getValues().findIndex(r=>String(r[0])===String(id))+1;
      if(rowNumber<2) throw Error('Product no longer exists.');
      sheet.getRange(rowNumber,1,1,values.length).setValues([values]);
    } else sheet.appendRow(values);
    logRow_(Utilities.getUuid(),'product',session.operator,{product:id,name:item.name,before:old||null,after:item});
    SpreadsheetApp.flush(); return {ok:true};
  } finally {lock.releaseLock();}
}
function logRow_(id,type,operator,details) {
  const json=JSON.stringify(details), parts=[];
  for(let i=0;i<json.length;i+=CELL_CHUNK) parts.push((i?'~':'')+json.slice(i,i+CELL_CHUNK));
  const sheet=db_().getSheetByName('Activity'), row=[id,new Date().toISOString(),type,safe_(operator),...parts];
  if(row.length>sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(),row.length-sheet.getMaxColumns());
  sheet.appendRow(row);
}
function submitMovement(token,input) {
  const session=auth_(token,input.type!=='sale'), lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    auth_(token,input.type!=='sale');
    if(!['sale','restock','adjustment'].includes(input.type)) throw Error('Unknown operation.');
    const request=text_(input.request,100); if(!request) throw Error('Missing request identifier.');
    const state=state_(); if(state.logs.some(l=>l.id===request)) return {ok:true,duplicate:true};
    const date=text_(input.date,10), shift=text_(input.shift,80), note=text_(input.note,500);
    if(!/^\d{4}-\d{2}-\d{2}$/.test(date)||!shift) throw Error('Date and shift or delivery label are required.');
    if(input.type==='sale'&&state.logs.some(l=>l.type==='sale'&&l.details.date===date&&l.details.shift.toLowerCase()===shift.toLowerCase())) throw Error('This shift already has a sales entry. An adviser can correct quantities through Stock changes.');
    if(input.type==='adjustment'&&!note) throw Error('Explain the inventory correction.');
    if(!Array.isArray(input.lines)||!input.lines.length||input.lines.length>500) throw Error('Enter at least one quantity.');
    const seen={}, lines=input.lines.map(x=>{
      const product=state.products.find(p=>p.id===x.id);
      if(!product||!product.active) throw Error('An item is no longer active. Refresh inventory.');
      if(seen[x.id]) throw Error('Duplicate product.'); seen[x.id]=true;
      let qty;
      if(input.type==='adjustment') {qty=Number(x.qty); if(!Number.isSafeInteger(qty)) throw Error('Corrections must be whole numbers.');}
      else qty=n_(x.qty,true);
      if(!qty) throw Error('Remove zero quantity rows.');
      const delta=input.type==='sale'?-qty:qty;
      if(product.stock+delta<0) throw Error(product.name+' '+product.size+': only '+product.stock+' available.');
      return {id:product.id,sku:product.sku,name:product.name,size:product.size,color:product.color,delta,price:product.price,cost:product.cost};
    });
    const revenue=input.type==='sale'?Math.round(lines.reduce((s,x)=>s-x.delta*x.price,0)*100)/100:0;
    const cash=input.type==='sale'?money_(input.cash):0, electronic=input.type==='sale'?money_(input.electronic):0;
    const details={date,shift,note,lines,revenue,cash,electronic,variance:Math.round((cash+electronic-revenue)*100)/100};
    // One append commits the complete shift. Stock is derived from this ledger.
    logRow_(request,input.type,session.operator,details);
    SpreadsheetApp.flush(); return {ok:true,revenue};
  } finally {lock.releaseLock();}
}

// Full physical count. Each line carries the system stock shown when that item
// was counted (expected). The correction is counted - expected, so sales or
// deliveries recorded after an item was counted still apply on top of it.
function submitCount(token,input) {
  const session=auth_(token,true), lock=LockService.getScriptLock(); lock.waitLock(30000);
  try {
    auth_(token,true);
    const request=text_(input.request,100); if(!request) throw Error('Missing request identifier.');
    const state=state_(); if(state.logs.some(l=>l.id===request)) return {ok:true,duplicate:true};
    const date=text_(input.date,10), shift=text_(input.shift,80), note=text_(input.note,500);
    if(!/^\d{4}-\d{2}-\d{2}$/.test(date)||!shift) throw Error('Date and count label are required.');
    if(!Array.isArray(input.counts)||!input.counts.length||input.counts.length>2000) throw Error('Count at least one item.');
    const seen={}, lines=input.counts.map(x=>{
      const product=state.products.find(p=>p.id===x.id);
      if(!product||!product.active) throw Error('An item is no longer active. Refresh inventory.');
      if(seen[x.id]) throw Error('Duplicate product.'); seen[x.id]=true;
      const counted=n_(x.counted,true), expected=Number(x.expected);
      if(counted>100000) throw Error('Count is too large.');
      if(!Number.isSafeInteger(expected)) throw Error('Count data is incomplete. Refresh and re-enter the count.');
      const delta=counted-expected, label=[product.name,product.size,product.color].filter(Boolean).join(' ');
      if(product.stock+delta<0) throw Error(label+': stock changed after it was counted. Recount it.');
      return {id:product.id,sku:product.sku,name:product.name,size:product.size,color:product.color,expected,counted,delta,price:product.price,cost:product.cost,laterChange:product.stock-expected};
    });
    const details={date,shift,note,lines,revenue:0,uncounted:state.products.filter(p=>p.active&&!seen[p.id]).length};
    logRow_(request,'count',session.operator,details);
    SpreadsheetApp.flush();
    return {ok:true,changed:lines.filter(x=>x.delta).length,net:lines.reduce((s,x)=>s+x.delta,0)};
  } finally {lock.releaseLock();}
}

function changeAccessCodes(token,input) {
  const session=auth_(token,true), student=code_(input.student), adviser=code_(input.adviser);
  if(!student&&!adviser) throw Error('Enter at least one new code.');
  [student,adviser].forEach(code=>{if(code&&code.length<10) throw Error('Each code needs at least 10 characters.');});
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try {
    auth_(token,true);
    const p=PropertiesService.getScriptProperties();
    const s=student?hash_(student):p.getProperty('student_hash'), a=adviser?hash_(adviser):p.getProperty('adviser_hash');
    if(s===a) throw Error('Student and adviser codes must be different.');
    const values={student_hash:s,adviser_hash:a,auth_epoch:Utilities.getUuid()};
    p.setProperties(values);
    return {ok:true};
  } finally {lock.releaseLock();}
}
function barcode_(value) {
  const code=text_(value,120);
  if(code&&!/^[A-Za-z0-9 ._\/-]+$/.test(code)) throw Error('Barcodes must contain letters, digits, spaces, or . _ / -');
  return code;
}
function barcodeKey_(value) {
  const s=String(value||'').trim();
  // UPC-A and zero-prefixed EAN-13 represent the same retail code.
  return /^0\d{12}$/.test(s)?s.slice(1):s;
}
