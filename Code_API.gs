const CONFIG = {
  APP_NAME: 'Gestion des gains',
  SPREADSHEET_ID: '10MY1Jk4um82QOrPwjmJ5kydG_c7sEUHMEc0ht9wyP_I',
  ADMIN_USERNAME: 'admin',
  SESSION_MINUTES: 60 * 24 * 365,
  MAX_DEVICES_PER_USER: 8
};
const WORK_DAY_AMOUNT = 110;
const RECHARGE_AMOUNT = 10;
const EMP_HEADERS = ['ID','Nom','Identifiant','MotDePasseHash','Actif','DateCreation','DerniereConnexion'];
const ENTRY_HEADERS = ['ID','EmployeID','Date','Gains','Pourboires','Recharge','Bonus','AutresGains','Depenses','Net','Notes','DateCreation','Statut'];
const ALLOWED_ACTIONS = new Set(['login','logout','getAppData','getEntries','getMonthlySummary','saveEntry','deleteEntry','getEmployees','createEmployee','toggleEmployee','resetEmployeePassword']);

function doGet() { return jsonOutput({ok:true,app:CONFIG.APP_NAME}); }
function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(payload.action || '');
    if (!ALLOWED_ACTIONS.has(action)) throw new Error('Action inconnue.');
    const args = Array.isArray(payload.args) ? payload.args : [];
    const result = dispatch_(action, args);
    return jsonOutput({ok:true,result:result});
  } catch (err) {
    return jsonOutput({ok:false,error:safeError_(err)});
  }
}
function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function safeError_(err) {
  const message = String(err && err.message || 'Erreur serveur.');
  if (/Session expirée|Identifiant ou mot de passe incorrect|Action inconnue|administrateur|autorisé|minimum|existe déjà|requis|introuvable|invalide|configuré|limite/i.test(message)) return message;
  console.error(err);
  return 'Une erreur serveur est survenue.';
}
function dispatch_(action, a) {
  switch (action) {
    case 'login': return login(a[0], a[1]);
    case 'logout': return logout(a[0]);
    case 'getAppData': return getAppData(a[0], a[1], a[2]);
    case 'getEntries': return getEntries(a[0], a[1], a[2]);
    case 'getMonthlySummary': return getMonthlySummary(a[0], a[1], a[2]);
    case 'saveEntry': return saveEntry(a[0], a[1]);
    case 'deleteEntry': return deleteEntry(a[0], a[1]);
    case 'getEmployees': return getEmployees(a[0]);
    case 'createEmployee': return createEmployee(a[0], a[1], a[2], a[3]);
    case 'toggleEmployee': return toggleEmployee(a[0], a[1], a[2]);
    case 'resetEmployeePassword': return resetEmployeePassword(a[0], a[1], a[2]);
  }
}
function setup() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  ensureSheet_(ss, 'Employes', EMP_HEADERS);
  ensureSheet_(ss, 'Entries', ENTRY_HEADERS);
  const entries = ss.getSheetByName('Entries');
  entries.getRange('C:C').setNumberFormat('@');
  SpreadsheetApp.getUi().alert('Configuration terminée. Les onglets sont prêts.');
}
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  else {
    const existing = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getDisplayValues()[0];
    headers.forEach((h,i)=>{ if (existing.indexOf(h) < 0) sh.getRange(1, sh.getLastColumn()+1).setValue(h); });
  }
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,sh.getLastColumn()).setFontWeight('bold');
}
function onOpen() { SpreadsheetApp.getUi().createMenu('Gestion des gains').addItem('Initialiser l’application','setup').addItem('Définir le mot de passe admin','setAdminPasswordFromPrompt').addToUi(); }
function setAdminPasswordFromPrompt() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Mot de passe administrateur','Choisissez un mot de passe long et unique (12 caractères minimum).',ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const password = response.getResponseText();
  if (password.length < 12) { ui.alert('Le mot de passe doit contenir au moins 12 caractères.'); return; }
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', password);
  ui.alert('Mot de passe admin enregistré dans les propriétés du script.');
}
function login(username, password) {
  username = String(username || '').trim(); password = String(password || '');
  if (!username || !password) return {success:false};
  const props = PropertiesService.getScriptProperties();
  const adminPassword = props.getProperty('ADMIN_PASSWORD');
  if (username.toLowerCase() === CONFIG.ADMIN_USERNAME.toLowerCase() && adminPassword && constantTimeEquals_(password, adminPassword)) {
    return createSession_({type:'admin',userId:'admin',username:CONFIG.ADMIN_USERNAME,name:'Administrateur'});
  }
  const users = readRows_('Employes');
  const user = users.find(r=>String(r.Identifiant||'').toLowerCase()===username.toLowerCase() && truthy_(r.Actif));
  if (!user || !verifyPassword_(String(user.ID), password, String(user.MotDePasseHash||''))) throw new Error('Identifiant ou mot de passe incorrect.');
  updateRow_('Employes',user._row,{DerniereConnexion:new Date()});
  return createSession_({type:'employee',userId:String(user.ID),username:String(user.Identifiant),name:String(user.Nom)});
}
function constantTimeEquals_(a,b) { a=String(a);b=String(b);if(a.length!==b.length)return false;let diff=0;for(let i=0;i<a.length;i++)diff|=a.charCodeAt(i)^b.charCodeAt(i);return diff===0; }
function createSession_(u) {
  const props=PropertiesService.getScriptProperties(), now=Date.now(), expires=now+CONFIG.SESSION_MINUTES*60000, token=Utilities.getUuid();
  const key='SESSION_'+token; props.setProperty(key,JSON.stringify({type:u.type,userId:u.userId,username:u.username,name:u.name,createdAt:now,expiresAt:expires}));
  const all=props.getProperties(), sessions=[];
  Object.keys(all).forEach(k=>{if(k.indexOf('SESSION_')===0){try{const s=JSON.parse(all[k]);if(s.expiresAt<=now)props.deleteProperty(k);else if(s.userId===u.userId)sessions.push({key:k,createdAt:s.createdAt});}catch(e){props.deleteProperty(k);}}});
  sessions.sort((x,y)=>x.createdAt-y.createdAt);while(sessions.length>CONFIG.MAX_DEVICES_PER_USER){props.deleteProperty(sessions.shift().key);}
  return {success:true,token:token,type:u.type,userId:u.userId,username:u.username,name:u.name};
}
function logout(token) { if(token)PropertiesService.getScriptProperties().deleteProperty('SESSION_'+String(token));return {success:true}; }
function requireSession_(token) {
  if(!token)throw new Error('Session expirée.');const props=PropertiesService.getScriptProperties(),key='SESSION_'+String(token),raw=props.getProperty(key);
  if(!raw)throw new Error('Session expirée.');let s;try{s=JSON.parse(raw);}catch(e){props.deleteProperty(key);throw new Error('Session expirée.');}
  if(s.expiresAt<=Date.now()){props.deleteProperty(key);throw new Error('Session expirée.');}
  if(s.type==='employee'){const user=readRows_('Employes').find(r=>String(r.ID)===String(s.userId)&&truthy_(r.Actif));if(!user){props.deleteProperty(key);throw new Error('Session expirée.');}s.name=String(user.Nom);s.username=String(user.Identifiant);}
  return s;
}
function requireAdmin_(token){const s=requireSession_(token);if(s.type!=='admin')throw new Error('Accès réservé à l’administrateur.');return s;}
function requireEmployee_(token){return requireSession_(token);}
function getAppData(token, month, employeeId) {
  const user=requireEmployee_(token), m=/^\d{4}-\d{2}$/.test(String(month||''))?String(month):Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM');
  const allEntries=readEntries_(), allUsers=readRows_('Employes');
  let selected = allEntries;
  if(user.type!=='admin')selected=selected.filter(e=>String(e.employeeId)===String(user.userId));
  else if(employeeId && employeeId!=='all')selected=selected.filter(e=>String(e.employeeId)===String(employeeId));
  const monthly=selected.filter(e=>e.date.slice(0,7)===m).sort((a,b)=>b.date.localeCompare(a.date));
  const summary=summarize_(monthly), months=monthlySummaries_(selected);
  let overview=null, employees=null;
  if(user.type==='admin'){
    employees=allUsers.map(u=>({id:String(u.ID),name:String(u.Nom),username:String(u.Identifiant),active:truthy_(u.Actif),lastLogin:u.DerniereConnexion?String(u.DerniereConnexion):''}));
    overview=allUsers.map(u=>{const ee=allEntries.filter(e=>String(e.employeeId)===String(u.ID));return {id:String(u.ID),name:String(u.Nom),active:truthy_(u.Actif),summary:summarize_(ee.filter(e=>e.date.slice(0,7)===m)),months:monthlySummaries_(ee)};});
    const adminEntries=allEntries.filter(e=>String(e.employeeId)==='admin');
    if(adminEntries.length)overview.unshift({id:'admin',name:'Administrateur',active:true,summary:summarize_(adminEntries.filter(e=>e.date.slice(0,7)===m)),months:monthlySummaries_(adminEntries)});
  }
  return {user:{type:user.type,name:user.name,username:user.username,userId:user.userId},summary:summary,months:months,overview:overview,employees:employees,entries:monthly};
}
function getEntries(token, month, employeeId) { return getAppData(token,month,employeeId).entries; }
function getMonthlySummary(token, month, employeeId) { return getAppData(token,month,employeeId).summary; }
function saveEntry(token, entry) {
  const user=requireEmployee_(token);if(!entry||!entry.date||!/^\d{4}-\d{2}-\d{2}$/.test(String(entry.date)))throw new Error('Date invalide.');
  const status=entry.status==='off'?'off':'worked';const employeeId=user.type==='admin'?String(entry.employeeId||'admin'):String(user.userId);
  if(user.type!=='admin'&&entry.employeeId&&String(entry.employeeId)!==String(user.userId))throw new Error('Accès non autorisé.');
  if(user.type==='admin'&&employeeId!=='admin'&&!readRows_('Employes').some(r=>String(r.ID)===employeeId))throw new Error('Employé introuvable.');
  const id=entry.id?String(entry.id):'ENTRY_'+Utilities.getUuid().replace(/-/g,'').slice(0,16);
  const sh=sheet_('Entries'), rows=readEntries_(), old=entry.id?rows.find(r=>r.id===id):null;
  if(entry.id&&!old)throw new Error('Journée introuvable.');
  if(old&&user.type!=='admin'&&String(old.employeeId)!==String(user.userId))throw new Error('Accès non autorisé.');
  if(old&&user.type==='admin'&&employeeId!==String(old.employeeId)){} // admin may reassign the record
  const tips=status==='worked'?Math.max(0,Number(entry.tips)||0):0, recharge=status==='worked'&&entry.recharge?RECHARGE_AMOUNT:0;
  const gains=status==='worked'?WORK_DAY_AMOUNT:0, net=gains+tips+recharge;
  const obj={ID:id,EmployeID:employeeId,Date:String(entry.date),Gains:gains,Pourboires:tips,Recharge:recharge,Bonus:0,AutresGains:0,Depenses:0,Net:net,Notes:String(entry.notes||'').slice(0,500),DateCreation:old?old.createdAt:new Date(),Statut:status};
  if(old)writeObjectRow_('Entries',old._row,obj);else appendObjectRow_('Entries',obj);
  return {success:true,id:id};
}
function deleteEntry(token,id) {
  const user=requireEmployee_(token), rows=readEntries_(), entry=rows.find(r=>r.id===String(id));if(!entry)throw new Error('Journée introuvable.');
  if(user.type!=='admin'&&String(entry.employeeId)!==String(user.userId))throw new Error('Accès non autorisé.');
  sheet_('Entries').deleteRow(entry._row);return {success:true};
}
function getEmployees(token){requireAdmin_(token);return readRows_('Employes').map(r=>({id:String(r.ID),name:String(r.Nom),username:String(r.Identifiant),active:truthy_(r.Actif),lastLogin:r.DerniereConnexion?String(r.DerniereConnexion):''}));}
function createEmployee(token,name,username,password) {
  requireAdmin_(token);name=String(name||'').trim();username=String(username||'').trim();password=String(password||'');
  if(!name||!username||password.length<4)throw new Error('Nom, identifiant et mot de passe (4 caractères minimum) sont requis.');
  if(username.toLowerCase()==='admin')throw new Error('Cet identifiant est réservé à l’administrateur.');
  const rows=readRows_('Employes');if(rows.some(r=>String(r.Identifiant||'').toLowerCase()===username.toLowerCase()))throw new Error('Cet identifiant existe déjà.');
  const id='EMP_'+Utilities.getUuid().replace(/-/g,'').slice(0,16),salt=Utilities.getUuid();PropertiesService.getScriptProperties().setProperty('SALT_'+id,salt);
  appendObjectRow_('Employes',{ID:id,Nom:name,Identifiant:username,MotDePasseHash:hashPassword_(password,salt),Actif:true,DateCreation:new Date(),DerniereConnexion:''});
  return {success:true,id:id};
}
function toggleEmployee(token,id,active) {
  requireAdmin_(token);id=String(id);const row=readRows_('Employes').find(r=>String(r.ID)===id);if(!row)throw new Error('Employé introuvable.');
  updateRow_('Employes',row._row,{Actif:!!active});if(!active)invalidateUserSessions_(id);return {success:true};
}
function resetEmployeePassword(token,id,password) {
  requireAdmin_(token);id=String(id);password=String(password||'');if(password.length<4)throw new Error('Mot de passe : 4 caractères minimum.');
  const row=readRows_('Employes').find(r=>String(r.ID)===id);if(!row)throw new Error('Employé introuvable.');
  const salt=Utilities.getUuid();PropertiesService.getScriptProperties().setProperty('SALT_'+id,salt);updateRow_('Employes',row._row,{MotDePasseHash:hashPassword_(password,salt)});invalidateUserSessions_(id);return {success:true};
}
function invalidateUserSessions_(id) { const p=PropertiesService.getScriptProperties(),all=p.getProperties();Object.keys(all).forEach(k=>{if(k.indexOf('SESSION_')===0){try{if(JSON.parse(all[k]).userId===id)p.deleteProperty(k);}catch(e){p.deleteProperty(k);}}}); }
function hashPassword_(password,salt){const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(password)+String(salt));return bytes.map(b=>('0'+((b+256)%256).toString(16)).slice(-2)).join('');}
function verifyPassword_(id,password,hash){const salt=PropertiesService.getScriptProperties().getProperty('SALT_'+id);return !!salt&&!!hash&&constantTimeEquals_(hashPassword_(password,salt),hash);}
function summarize_(entries) {
  const worked=entries.filter(e=>e.status!=='off'),off=entries.filter(e=>e.status==='off');
  const sum=k=>entries.reduce((a,e)=>a+(Number(e[k])||0),0);
  return {month:entries.length?entries[0].date.slice(0,7):'',daysWorked:worked.length,daysOff:off.length,gains:sum('gains'),tips:sum('tips'),recharge:sum('recharge'),rechargeCount:worked.filter(e=>Number(e.recharge)>0).length,bonus:sum('bonus'),otherGains:sum('otherGains'),expenses:sum('expenses'),net:sum('net'),gross:sum('net')};
}
function monthlySummaries_(entries) {
  const grouped={};entries.filter(e=>e.status!=='off').forEach(e=>{const m=e.date.slice(0,7);if(!grouped[m])grouped[m]={month:m,net:0,tips:0,recharge:0,days:0};grouped[m].net+=Number(e.net)||0;grouped[m].tips+=Number(e.tips)||0;grouped[m].recharge+=Number(e.recharge)||0;grouped[m].days++;});
  return Object.keys(grouped).sort().reverse().map(k=>grouped[k]);
}
function readEntries_() {
  return readRows_('Entries').map(r=>({id:String(r.ID||''),employeeId:String(r.EmployeID||''),date:normalizeDate_(r.Date),gains:Number(r.Gains)||0,tips:Number(r.Pourboires)||0,recharge:Number(r.Recharge)||0,bonus:Number(r.Bonus)||0,otherGains:Number(r.AutresGains)||0,expenses:Number(r.Depenses)||0,net:Number(r.Net)||0,notes:String(r.Notes||''),status:String(r.Statut||'worked')==='off'?'off':'worked',createdAt:r.DateCreation,_row:r._row,employeeName:employeeName_(String(r.EmployeID||''))})).filter(r=>r.id&&r.date);
}
function employeeName_(id){if(id==='admin')return 'Administrateur';const u=readRows_('Employes').find(r=>String(r.ID)===id);return u?String(u.Nom):'Compte supprimé';}
function normalizeDate_(v){if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');return String(v||'').slice(0,10);}
function readRows_(name) {
  const sh=sheet_(name),last=sh.getLastRow();if(last<2)return [];
  const values=sh.getRange(1,1,last,sh.getLastColumn()).getValues(),headers=values[0].map(String);
  return values.slice(1).map((row,i)=>{const obj={_row:i+2};headers.forEach((h,j)=>obj[h]=row[j]);return obj;}).filter(o=>Object.keys(o).some(k=>k!=='_row'&&o[k]!==''&&o[k]!==null));
}
function sheet_(name){const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID),sh=ss.getSheetByName(name);if(!sh)throw new Error('Onglet '+name+' introuvable. Lancez setup().');return sh;}
function appendObjectRow_(name,obj){const sh=sheet_(name),headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];sh.appendRow(headers.map(h=>Object.prototype.hasOwnProperty.call(obj,h)?obj[h]:''));}
function writeObjectRow_(name,row,obj){const sh=sheet_(name),headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];sh.getRange(row,1,1,headers.length).setValues([headers.map(h=>Object.prototype.hasOwnProperty.call(obj,h)?obj[h]:'')]);}
function updateRow_(name,row,updates){const sh=sheet_(name),headers=sh.getRange(1,1,1,sh.getLastColumn()).getDisplayValues()[0];Object.keys(updates).forEach(k=>{const i=headers.indexOf(k);if(i>=0)sh.getRange(row,i+1).setValue(updates[k]);});}
function truthy_(v){return v===true||String(v).toLowerCase()==='true'||String(v)==='1'||String(v).toLowerCase()==='oui';}
