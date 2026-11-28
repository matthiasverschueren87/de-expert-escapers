// De Expert Escapers — alleen servercode; accounts en privégegevens blijven in de eigen Sheet.
const EE_SHEET='1-7h0RlVXqfZlLBdONIKchTlnU0ZGofux-XV3wyencmQ';
const EE_URL='https://matthiasverschueren87.github.io/de-expert-escapers/';
const EE_TABLES=['Members','Sessions','LoginCodes','RoomMeta','Opinions','Comments','Plans','Receipts','MailLog'];
const EE_ROOM_HEADERS=['ID','Naam','Aanbieder','Locatie','Website','Bucket-ID','Spelers','Prijs totaal (€)','Duur (min)','Thema','Datum','Notities','Favoriet','Volgorde'];
function props_(){return PropertiesService.getScriptProperties()}
let EE_DB_CACHE=null,EE_RECORD_CACHE={};
function db_(){return EE_DB_CACHE||(EE_DB_CACHE=SpreadsheetApp.openById(EE_SHEET))}
function uid_(){return Utilities.getUuid()}
function now_(){return new Date().toISOString()}
function fail_(message,code){const e=new Error(message);e.code=code||'INVALID';throw e}
function str_(v,max){if(typeof v!=='string'||v.length>max)fail_('Ongeldige tekst.');return v.trim()}
function email_(v){v=str_(v,180).toLowerCase();if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v))fail_('Vul een geldig e-mailadres in.');return v}
function digest_(v){return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(String(v),props_().getProperty('AUTH_PEPPER')))}
function safeEqual_(a,b){a=String(a||'');b=String(b||'');let n=a.length^b.length;for(let i=0;i<Math.max(a.length,b.length);i++)n|=(a.charCodeAt(i)||0)^(b.charCodeAt(i)||0);return n===0}
function sheet_(name){const s=db_().getSheetByName('EE_'+name);if(!s)fail_('Accounts worden nog geactiveerd.','NOT_READY');return s}
function records_(name){if(EE_RECORD_CACHE[name])return EE_RECORD_CACHE[name];const s=sheet_(name);if(s.getLastRow()<2)return EE_RECORD_CACHE[name]=[];if(s.getLastRow()>12000)fail_('De opslaglimiet is bereikt.');return EE_RECORD_CACHE[name]=s.getRange(2,1,s.getLastRow()-1,2).getValues().filter(r=>r[0]).map(r=>({id:String(r[0]),...JSON.parse(String(r[1]))}))}
function get_(name,id){return records_(name).find(r=>r.id===id)}
function put_(name,id,value){const s=sheet_(name),rows=records_(name),i=rows.findIndex(r=>r.id===id),row=i<0?s.getLastRow()+1:i+2;if(row>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),100);s.getRange(row,1,1,3).setValues([[id,JSON.stringify(value),now_()]]);const v={...value,id};if(i<0)rows.push(v);else rows[i]=v}
function cleanMember_(m){return {id:m.id,name:m.name}}
function member_(id){const m=get_('Members',id);if(!m||m.disabled)fail_('Dit teamlid is niet actief.');return m}
function admin_(m){return m.role==='admin'}
function assertAdmin_(m){if(!admin_(m))fail_('Alleen de beheerder kan dit.','FORBIDDEN')}
function setupAccounts(){
 const p=props_();const owner=email_(p.getProperty('OWNER_EMAIL')||'');
 if(!p.getProperty('AUTH_PEPPER'))p.setProperty('AUTH_PEPPER',uid_()+uid_());
 if(!p.getProperty('TEAM_INVITE'))p.setProperty('TEAM_INVITE',uid_()+uid_());
 EE_TABLES.forEach(n=>{let s=db_().getSheetByName('EE_'+n);if(!s){s=db_().insertSheet('EE_'+n);s.getRange(1,1,1,3).setValues([['ID','Gegevens','Bijgewerkt']]);s.setFrozenRows(1)}});
 if(!records_('Members').some(m=>m.email===owner))put_('Members','owner',{email:owner,name:'Matthias',role:'admin',verifiedAt:'',createdAt:now_(),notifications:true});
 p.setProperty('REMINDER_DAYS','2');
 console.log('Accounts voorbereid. Mails worden verzonden als '+Session.getEffectiveUser().getEmail()+'.');
 // Laat Google hier de beperkte mail-scope aanvragen, zonder een mail te versturen.
 console.log('Beschikbare mailontvangers vandaag: '+MailApp.getRemainingDailyQuota());
}
function enablePaymentReminders(){
 if(!props_().getProperty('OWNER_EMAIL'))fail_('Beheerder ontbreekt.');
 const existing=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='sendPaymentReminders');
 if(!existing.length)ScriptApp.newTrigger('sendPaymentReminders').timeBased().atHour(9).everyDays(1).inTimezone('Europe/Brussels').create();
 props_().setProperty('REMINDERS_ENABLED','true');console.log('Betaalherinneringen actief: 2 dagen vooraf, dagelijks rond 09:00 Belgische tijd.');
}
function authenticate_(x){
 if(typeof x.token!=='string'||x.token.length<40||x.token.length>200)fail_('Log eerst in.','AUTH_REQUIRED');
 const ses=get_('Sessions',digest_(x.token));if(!ses||ses.revoked||Date.parse(ses.expiresAt)<Date.now())fail_('Je sessie is verlopen. Log opnieuw in.','AUTH_REQUIRED');
 return member_(ses.userId);
}
function requestCode_(x){
 const email=email_(x.email),members=records_('Members'),existing=members.find(m=>m.email===email&&!m.disabled);
 if(!existing&&!safeEqual_(str_(x.invite||'',150),props_().getProperty('TEAM_INVITE')))fail_('Gebruik de uitnodigingslink van je team om een account aan te maken.');
 const name=existing?existing.name:str_(x.name||'',80);if(name.length<2)fail_('Vul je naam in.');
 const id=digest_('email:'+email),prev=get_('LoginCodes',id);
 if(prev&&Date.parse(prev.sentAt)>Date.now()-60000)fail_('Wacht een minuut voordat je een nieuwe code aanvraagt.','RATE_LIMIT');
 const day=Utilities.formatDate(new Date(),'Europe/Brussels','yyyy-MM-dd'),quotaKey='LOGIN_SENDS_'+day,count=Number(props_().getProperty(quotaKey)||0);
 if(count>=70||MailApp.getRemainingDailyQuota()<6)fail_('De dagelijkse maillimiet is bereikt. Probeer het later opnieuw.','RATE_LIMIT');
 if(prev&&prev.day===day&&Number(prev.sends)>=8)fail_('Te veel codes aangevraagd voor dit adres. Probeer morgen opnieuw.','RATE_LIMIT');
 const hex=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,uid_()+uid_()).slice(0,4).map(v=>('0'+((v+256)%256).toString(16)).slice(-2)).join('');
 const code=String(100000+(parseInt(hex,16)%900000));
 put_('LoginCodes',id,{email,name,codeHash:digest_('otp:'+email+':'+code),expiresAt:new Date(Date.now()+600000).toISOString(),attempts:0,sentAt:now_(),day,sends:prev&&prev.day===day?Number(prev.sends)+1:1,allowJoin:!existing});
 props_().setProperty(quotaKey,String(count+1));
 sendMail_(email,'Je inlogcode voor De Expert Escapers',mailFrame_('Het avontuur wacht.',`<p>Hallo ${html_(name)},</p><p>Gebruik deze code om veilig in te loggen:</p><p style="font-size:36px;letter-spacing:9px;color:#f4eddc;font-weight:bold">${code}</p><p>De code is 10 minuten geldig. Deel hem met niemand.</p>`),'Je inlogcode is '+code+'. Deze code is 10 minuten geldig.');
 return {ok:true,message:'De inlogcode is verstuurd. Controleer ook je spamfolder.'};
}
function verifyCode_(x){
 const email=email_(x.email),id=digest_('email:'+email),v=get_('LoginCodes',id);
 if(!v||v.used||Date.parse(v.expiresAt)<Date.now()||v.attempts>=5)fail_('Deze code is verlopen of geblokkeerd. Vraag een nieuwe code aan.');
 v.attempts++;put_('LoginCodes',id,v);
 if(!/^\d{6}$/.test(String(x.code))||!safeEqual_(v.codeHash,digest_('otp:'+email+':'+x.code)))fail_('De code klopt niet.');
 let m=records_('Members').find(a=>a.email===email&&!a.disabled);
 if(!m){if(!v.allowJoin)fail_('Deze uitnodiging is niet meer geldig.');m={id:uid_(),email,name:v.name,role:'member',createdAt:now_(),verifiedAt:now_(),notifications:true}}
 m.verifiedAt=now_();put_('Members',m.id,m);v.used=true;put_('LoginCodes',id,v);
 const token=uid_()+uid_()+uid_();put_('Sessions',digest_(token),{userId:m.id,createdAt:now_(),expiresAt:new Date(Date.now()+30*86400000).toISOString()});
 return {ok:true,token,data:loadData_(m)};
}
function legacyRooms_(){const s=db_().getSheetByName('Escaperooms');if(!s||s.getLastRow()<2)return [];if(s.getLastRow()>1001)fail_('Maximaal 1000 kamers.');const rows=s.getRange(2,1,s.getLastRow()-1,14).getValues();return rows.filter(r=>r[0]).map((r,i)=>({id:String(r[0]),name:String(r[1]),provider:String(r[2]),location:String(r[3]),website:String(r[4]),bucketId:String(r[5]),players:String(r[6]),price:r[7],duration:r[8],theme:String(r[9]),date:r[10] instanceof Date?Utilities.formatDate(r[10],'Europe/Brussels','yyyy-MM-dd'):String(r[10]),notes:String(r[11]),favorite:r[12]===true,order:i}))}
function version_(v){return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(v)))}
function room_(id){const r=legacyRooms_().find(r=>r.id===id),meta=get_('RoomMeta',id);if(!r||meta&&meta.archived)fail_('Deze kamer is niet beschikbaar.');return {...r,meta:meta||{},version:version_({...r,meta:meta||{}})}}
function loadData_(m){
 const metas=records_('RoomMeta'),rooms=legacyRooms_().filter(r=>!metas.some(t=>t.id===r.id&&t.archived)).map(r=>{const meta=metas.find(t=>t.id===r.id)||{};return {...r,meta,version:version_({...r,meta})}});
 return {me:{...cleanMember_(m),email:m.email,role:m.role,notifications:m.notifications!==false},members:records_('Members').filter(a=>a.verifiedAt&&!a.disabled).map(cleanMember_),rooms,opinions:records_('Opinions'),comments:records_('Comments').filter(c=>!c.deleted),plans:records_('Plans'),remindersEnabled:props_().getProperty('REMINDERS_ENABLED')==='true',invite:admin_(m)?props_().getProperty('TEAM_INVITE'):null};
}
function cleanRoom_(v){const r={};for(const k of ['name','provider','location','website','players','theme','date','notes'])r[k]=str_(v[k]||'',k==='notes'?3000:k==='website'?800:180);if(!r.name)fail_('Een kamer heeft een naam nodig.');if(r.website&&!/^https?:\/\//i.test(r.website))fail_('Gebruik een geldige website.');if(!['wishlist','booked','played'].includes(v.bucketId))fail_('Ongeldige status.');r.bucketId=v.bucketId;for(const k of ['price','duration']){r[k]=v[k]===''||v[k]==null?'':Number(v[k]);if(r[k]!==''&&(!Number.isFinite(r[k])||r[k]<0||r[k]>100000))fail_('Ongeldig bedrag of aantal.')}if(r.date&&!validDate_(r.date))fail_('Ongeldige datum.');return r}
function saveRoom_(x,m){
 const value=cleanRoom_(x.room||{}),id=x.room.id?str_(x.room.id,250):uid_(),s=db_().getSheetByName('Escaperooms'),all=legacyRooms_(),old=all.find(r=>r.id===id);let row=s.getLastRow()+1;
 if(old){const live=room_(id);if(x.version!==live.version)fail_('Deze kamer is intussen gewijzigd. Vernieuw de lijst en probeer opnieuw.','CONFLICT');row=s.getRange(2,1,s.getLastRow()-1,1).getValues().findIndex(r=>String(r[0])===id)+2}
 else if(all.length>=1000)fail_('De lijst is vol.');
 if(row>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),100);
 const r={id,...value};s.getRange(row,11).setNumberFormat('@');s.getRange(row,1,1,14).setValues([[id,r.name,r.provider,r.location,r.website,r.bucketId,r.players,r.price,r.duration,r.theme,r.date,r.notes,old?old.favorite:false,old?old.order:all.length]]);
 const meta=get_('RoomMeta',id)||{createdBy:old?null:m.id,createdAt:old?'':now_()};put_('RoomMeta',id,{...meta,updatedBy:m.id,updatedAt:now_()});
}
function archiveRoom_(x,m){const r=room_(str_(x.roomId,250));if(!admin_(m)&&r.meta.createdBy!==m.id)fail_('Alleen de toevoeger of beheerder kan deze kamer archiveren.','FORBIDDEN');if(records_('Plans').some(p=>p.roomId===r.id&&p.status!=='cancelled'&&p.date>=today_()))fail_('Annuleer eerst de geplande sessie.');put_('RoomMeta',r.id,{...r.meta,archived:true,updatedBy:m.id,updatedAt:now_()})}
function stars_(v){v=Number(v||0);if(!Number.isInteger(v)||v<0||v>5)fail_('Kies 0 tot 5 sterren.');return v}
function saveOpinion_(x,m){const r=room_(str_(x.roomId,250)),played=x.played===true,playedStars=stars_(x.playedStars),wishStars=stars_(x.wishStars);if(playedStars&&!played)fail_('Markeer de kamer eerst als gespeeld.');put_('Opinions',r.id+'::'+m.id,{roomId:r.id,userId:m.id,played,playedStars,wishStars,favorite:x.favorite===true,updatedAt:now_()})}
function addComment_(x,m){const r=room_(str_(x.roomId,250)),body=str_(x.body,1500);if(!body)fail_('Schrijf eerst een reactie.');if(records_('Comments').length>=6000)fail_('Het reactielimiet is bereikt.');put_('Comments',uid_(),{roomId:r.id,userId:m.id,body,createdAt:now_()})}
function deleteComment_(x,m){const c=get_('Comments',str_(x.commentId,80));if(!c)fail_('Reactie niet gevonden.');if(c.userId!==m.id&&!admin_(m))fail_('Je kunt alleen je eigen reacties verwijderen.','FORBIDDEN');put_('Comments',c.id,{...c,deleted:true,deletedAt:now_()})}
function cents_(n){if(!Number.isSafeInteger(n)||n<0||n>20000000)fail_('Ongeldig bedrag.');return n}
function splitCents_(total,ids){if(!ids.length)fail_('Kies minstens één speler.');const base=Math.floor(total/ids.length),extra=total%ids.length;return ids.map((id,i)=>({userId:id,amountCents:base+(i<extra?1:0)}))}
function validDate_(v){if(!/^\d{4}-\d{2}-\d{2}$/.test(v))return false;const d=new Date(v+'T12:00:00Z');return !isNaN(d)&&d.toISOString().slice(0,10)===v}
function canManagePlan_(p,m){return admin_(m)||p.createdBy===m.id||p.paidBy===m.id}
function savePlan_(x,m){
 const r=room_(str_(x.plan.roomId,250)),v=x.plan,id=v.id?str_(v.id,80):uid_(),old=get_('Plans',id);
 if(old&&!canManagePlan_(old,m))fail_('Alleen de organisator of voorschieter kan deze planning aanpassen.','FORBIDDEN');
 if(old&&x.version!==old.updatedAt)fail_('Deze planning is intussen gewijzigd. Vernieuw en probeer opnieuw.','CONFLICT');
 const date=str_(v.date,10),time=str_(v.time||'',5);if(!validDate_(date)||(time&&!/^([01]\d|2[0-3]):[0-5]\d$/.test(time)))fail_('Kies een geldige datum en tijd.');
 const totalCents=cents_(v.totalCents),paidBy=str_(v.paidBy,80);member_(paidBy);
 if(!Array.isArray(v.participants)||!v.participants.length||v.participants.length>30)fail_('Kies 1 tot 30 deelnemers.');
 const ids=new Set();const participants=v.participants.map(a=>{if(ids.has(a.userId))fail_('Dubbele deelnemer.');ids.add(a.userId);const member=member_(a.userId);if(!member.verifiedAt)fail_('Alle spelers moeten hun account activeren.');return {userId:member.id,amountCents:cents_(a.amountCents)}});
 if(participants.reduce((n,p)=>n+p.amountCents,0)!==totalCents)fail_('De spelersbedragen moeten samen exact het totaal zijn.');
 const dueDate=str_(v.dueDate||'',10);if(dueDate&&(!validDate_(dueDate)||dueDate>date))fail_('De betaaldatum moet op of vóór de speeldatum liggen.');
 const p={roomId:r.id,date,time,totalCents,paidBy,dueDate,notes:str_(v.notes||'',1500),status:'active',createdBy:old?old.createdBy:m.id,createdAt:old?old.createdAt:now_(),updatedAt:now_(),participants:participants.map(a=>{const prior=old&&old.participants.find(b=>b.userId===a.userId&&b.amountCents===a.amountCents);return {...a,status:prior?prior.status:a.userId===paidBy?'paid':'unpaid',paidAt:prior?prior.paidAt||'':a.userId===paidBy?now_():'',markedBy:prior?prior.markedBy||'':a.userId===paidBy?m.id:''}})};
 put_('Plans',id,p);
}
function payment_(x,m){const p=get_('Plans',str_(x.planId,80));if(!p||p.status==='cancelled')fail_('Deze planning is niet actief.');const share=p.participants.find(a=>a.userId===x.userId);if(!share)fail_('Speler niet gevonden.');const manage=canManagePlan_(p,m);if(!manage&&(m.id!==share.userId||x.status!=='reported'))fail_('Je mag alleen je eigen betaling melden.','FORBIDDEN');if(!['unpaid','reported','paid'].includes(x.status))fail_('Ongeldige betaalstatus.');if(share.status==='paid'&&!manage)fail_('Deze betaling is al bevestigd.');share.status=x.status;share.paidAt=x.status==='paid'?now_():'';share.markedBy=m.id;p.updatedAt=now_();put_('Plans',p.id,p)}
function cancelPlan_(x,m){const p=get_('Plans',str_(x.planId,80));if(!p||!canManagePlan_(p,m))fail_('Je kunt deze planning niet annuleren.','FORBIDDEN');p.status='cancelled';p.updatedAt=now_();put_('Plans',p.id,p)}
function today_(){return Utilities.formatDate(new Date(),'Europe/Brussels','yyyy-MM-dd')}
function daysAway_(date,today){return Math.round((Date.parse(date+'T12:00:00Z')-Date.parse(today+'T12:00:00Z'))/86400000)}
function reminderNeeded_(p,a,today){return p.status==='active'&&a.status==='unpaid'&&a.amountCents>0&&daysAway_(p.date,today)>=0&&daysAway_(p.date,today)<=2}
function html_(v){return String(v||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function eur_(v){return '€ '+(v/100).toFixed(2).replace('.',',')}
function mailFrame_(title,body){return '<!doctype html><html lang="nl"><head><meta charset="utf-8"></head><body style="margin:0;padding:24px 12px;background:#002b44;font-family:Arial,sans-serif;color:#f4eddc"><table role="presentation" style="max-width:580px;width:100%;margin:auto;background:#004369;border:1px solid #326981;border-radius:12px"><tr><td style="padding:32px"><p style="font-size:12px;letter-spacing:3px;color:#68c9c4">DE EXPERT ESCAPERS ↗</p><h1 style="font-size:30px;line-height:1.2;color:#f4eddc">'+html_(title)+'</h1><div style="font-size:16px;line-height:1.7;color:#d4e4e9">'+body+'</div><p style="padding-top:25px;border-top:1px solid #326981;font-size:12px;color:#b0cad6">Verzamel avonturen. Samen achter de volgende deur.</p></td></tr></table></body></html>'}
function sendMail_(to,subject,html,plain){MailApp.sendEmail({to,subject,name:'De Expert Escapers',replyTo:props_().getProperty('OWNER_EMAIL'),body:plain,htmlBody:html})}
function sendPaymentReminders(){
 const lock=LockService.getScriptLock();if(!lock.tryLock(15000))return;
 try{if(props_().getProperty('REMINDERS_ENABLED')!=='true')return;const today=today_(),plans=records_('Plans'),members=records_('Members'),log=records_('MailLog'),rooms=legacyRooms_();
 for(const p of plans)for(const a of p.participants){if(!reminderNeeded_(p,a,today))continue;const u=members.find(m=>m.id===a.userId),payer=members.find(m=>m.id===p.paidBy),room=rooms.find(r=>r.id===p.roomId);if(!u||!u.verifiedAt||u.disabled||u.notifications===false||!room)continue;
 const id=version_([p.id,a.userId,p.date,a.amountCents]);if(log.some(l=>l.id===id))continue;if(MailApp.getRemainingDailyQuota()<1)return;
 // Een vooraf geregistreerde poging voorkomt een dubbele mail na een onzekere providerrespons.
 put_('MailLog',id,{planId:p.id,userId:u.id,status:'attempted',at:now_()});
 const date=p.date.split('-').reverse().join('/'),body='<p>Hallo '+html_(u.name)+',</p><p>Ons avontuur <strong>'+html_(room.name)+'</strong> staat gepland op <strong>'+date+'</strong>.</p><p>Jouw openstaande bijdrage:</p><p style="font-size:34px;color:#f4eddc;font-weight:bold">'+eur_(a.amountCents)+'</p><p>'+html_(payer?payer.name:'De organisator')+' heeft voorgeschoten. Heb je al betaald? Meld dat in de tool; de organisator kan je betaling bevestigen.</p><p><a href="'+EE_URL+'?room='+encodeURIComponent(p.roomId)+'" style="display:inline-block;padding:14px 22px;background:#db1f48;color:#ffffff;text-decoration:none;border-radius:6px;font-weight:bold">Bekijk mijn betaling ↗</a></p>';
 try{sendMail_(u.email,'Je bijdrage voor '+room.name+' · De Expert Escapers',mailFrame_('Bijna tijd voor avontuur.',body),'Hallo '+u.name+', je openstaande bijdrage voor '+room.name+' op '+date+' is '+eur_(a.amountCents)+'. Voorschieter: '+(payer?payer.name:'organisator')+'. Bekijk je betaling: '+EE_URL);put_('MailLog',id,{planId:p.id,userId:u.id,status:'sent',at:now_()})}catch(err){put_('MailLog',id,{planId:p.id,userId:u.id,status:'needs_review',at:now_(),error:String(err.message).slice(0,300)})}
 }}finally{lock.releaseLock()}
}
function doPost(e){EE_DB_CACHE=null;EE_RECORD_CACHE={};let lock;try{
 if(!e.postData||e.postData.contents.length>120000)fail_('Ongeldig verzoek.');const x=JSON.parse(e.postData.contents);if(!props_().getProperty('AUTH_PEPPER'))fail_('Accounts worden nog geactiveerd.','NOT_READY');
 lock=LockService.getScriptLock();lock.waitLock(20000);
 if(x.action==='requestCode')return json_3_(requestCode_(x));if(x.action==='verifyCode')return json_3_(verifyCode_(x));
 const m=authenticate_(x);
 if(x.action==='load')return json_3_({ok:true,data:loadData_(m)});
 if(x.action==='logout'){const t=get_('Sessions',digest_(x.token));t.revoked=true;put_('Sessions',digest_(x.token),t);return json_3_({ok:true})}
 if(!/^[a-zA-Z0-9-]{16,100}$/.test(x.requestId||''))fail_('Ontbrekende aanvraagcode.');
 const receiptId=m.id+':'+x.requestId,receipt=get_('Receipts',receiptId);if(receipt)return json_3_({ok:true,data:loadData_(m)});
 switch(x.action){case 'saveRoom':saveRoom_(x,m);break;case 'archiveRoom':archiveRoom_(x,m);break;case 'saveOpinion':saveOpinion_(x,m);break;case 'addComment':addComment_(x,m);break;case 'deleteComment':deleteComment_(x,m);break;case 'savePlan':savePlan_(x,m);break;case 'payment':payment_(x,m);break;case 'cancelPlan':cancelPlan_(x,m);break;case 'profile':m.name=str_(x.name,80);if(m.name.length<2)fail_('Vul je naam in.');m.notifications=x.notifications!==false;put_('Members',m.id,m);break;default:fail_('Onbekende actie.');}
 put_('Receipts',receiptId,{userId:m.id,action:x.action,at:now_()});return json_3_({ok:true,data:loadData_(m)});
 }catch(err){return json_3_({ok:false,error:String(err.message||err),code:err.code||'ERROR'})}finally{if(lock)lock.releaseLock()}}
function doGet(){return json_3_({ok:true,version:3,message:'De Expert Escapers · accounts vereist'})}
function json_3_(v){return ContentService.createTextOutput(JSON.stringify(v)).setMimeType(ContentService.MimeType.JSON)}
