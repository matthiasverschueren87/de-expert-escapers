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
 props_().setProperty('REMINDERS_ENABLED','true');console.log('Speel- en betaalherinneringen actief: 2 dagen vooraf, dagelijks rond 09:00 Belgische tijd.');
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
 return {me:{...cleanMember_(m),email:m.email,role:m.role,notifications:m.notifications!==false,gameNotifications:m.gameNotifications!==false},members:records_('Members').filter(a=>a.verifiedAt&&!a.disabled).map(cleanMember_),rooms,opinions:records_('Opinions'),comments:records_('Comments').filter(c=>!c.deleted),plans:records_('Plans'),remindersEnabled:props_().getProperty('REMINDERS_ENABLED')==='true',invite:admin_(m)?props_().getProperty('TEAM_INVITE'):null};
}
function cleanRoom_(v){const r={};for(const k of ['name','provider','location','website','players','theme','date','notes'])r[k]=str_(v[k]||'',k==='notes'?3000:k==='website'?800:180);if(!r.name)fail_('Een kamer heeft een naam nodig.');if(r.website&&!/^https?:\/\//i.test(r.website))fail_('Gebruik een geldige website.');if(!['wishlist','booked','played'].includes(v.bucketId))fail_('Ongeldige status.');r.bucketId=v.bucketId;for(const k of ['price','duration']){r[k]=v[k]===''||v[k]==null?'':Number(v[k]);if(r[k]!==''&&(!Number.isFinite(r[k])||r[k]<0||r[k]>100000))fail_('Ongeldig bedrag of aantal.')}if(r.date&&!validDate_(r.date))fail_('Ongeldige datum.');return r}
function saveRoom_(x,m){
 const value=cleanRoom_(x.room||{}),id=x.room.id?str_(x.room.id,250):uid_(),s=db_().getSheetByName('Escaperooms'),all=legacyRooms_(),old=all.find(r=>r.id===id);let row=s.getLastRow()+1;
 if(old){const live=room_(id);if(x.version!==live.version)fail_('Deze kamer is intussen gewijzigd. Vernieuw de lijst en probeer opnieuw.','CONFLICT');row=s.getRange(2,1,s.getLastRow()-1,1).getValues().findIndex(r=>String(r[0])===id)+2}
 else if(all.length>=1000)fail_('De lijst is vol.');
 if(row>s.getMaxRows())s.insertRowsAfter(s.getMaxRows(),100);
 const r={id,...value};s.getRange(row,11).setNumberFormat('@');s.getRange(row,1,1,14).setValues([[id,r.name,r.provider,r.location,r.website,r.bucketId,r.players,r.price,r.duration,r.theme,r.date,r.notes,old?old.favorite:false,old?old.order:all.length]]);
 const meta=get_('RoomMeta',id)||{createdBy:old?null:m.id,createdAt:old?'':now_()};
 const image=(!old||old.website!==r.website||meta.imageQuality!==2)?findRoomImage_(r.website,r.name,r.theme):{imageUrl:meta.imageUrl||'',imageSource:meta.imageSource||'',imageCheckedAt:meta.imageCheckedAt,imageQuality:meta.imageQuality,imageWidth:meta.imageWidth,imageHeight:meta.imageHeight};
 put_('RoomMeta',id,{...meta,...image,updatedBy:m.id,updatedAt:now_()});
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
function gameReminderNeeded_(p,today){return p.status==='active'&&daysAway_(p.date,today)>=0&&daysAway_(p.date,today)<=2}
function planReminderMail_(p,a,u,room,members,includeGame,includePayment,today){
 const date=p.date.split('-').reverse().join('/'),time=p.time||'Nog af te spreken',place=[room.provider,room.location].filter(Boolean).join(' · ')||'Bekijk de tool voor de locatie.',names=p.participants.map(a=>{const m=members.find(m=>m.id===a.userId);return m?m.name:'Teamlid'}).join(', '),payer=members.find(m=>m.id===p.paidBy),days=daysAway_(p.date,today),when=days===2?'Over 2 dagen':days===1?'Morgen':'Vandaag';
 let body='<p>Hallo '+html_(u.name)+',</p>',plain='Hallo '+u.name+',\n\n';
 if(includeGame){body+='<p><strong>'+when+'</strong> spelen we <strong>'+html_(room.name)+'</strong>. Klaar voor ons volgende avontuur?</p><table role="presentation" style="width:100%;border-collapse:collapse;background:#00334f;border-radius:8px"><tr><td style="padding:18px"><strong>Datum</strong><br>'+date+'<br><strong>Starttijd</strong><br>'+html_(time)+'<br><strong>Locatie</strong><br>'+html_(place)+'<br><strong>Deelnemers</strong><br>'+html_(names)+'</td></tr></table>';
 plain+=when+' spelen we '+room.name+'.\nDatum: '+date+'\nStarttijd: '+time+'\nLocatie: '+place+'\nDeelnemers: '+names+'\n';
 if(p.notes){body+='<p><strong>Onze afspraken</strong><br>'+html_(p.notes).replace(/\n/g,'<br>')+'</p>';plain+='Afspraken: '+p.notes+'\n'}
 if(a.status==='paid'&&a.amountCents>0){body+='<p style="color:#68c9c4">Je bijdrage is al betaald. Helemaal klaar voor het avontuur.</p>';plain+='Je bijdrage is al betaald.\n'}
 else if(a.status==='reported'){body+='<p>Je hebt je betaling gemeld. De organisator kan die nog bevestigen.</p>';plain+='Je hebt je betaling gemeld; de organisator kan die nog bevestigen.\n'}
 }else{body+='<p>Ons avontuur <strong>'+html_(room.name)+'</strong> staat gepland op <strong>'+date+'</strong>.</p>';plain+='Ons avontuur '+room.name+' staat gepland op '+date+'.\n'}
 if(includePayment){body+='<p>Jouw openstaande bijdrage:</p><p style="font-size:34px;color:#f4eddc;font-weight:bold">'+eur_(a.amountCents)+'</p><p>'+html_(payer?payer.name:'De organisator')+' heeft voorgeschoten. Heb je al betaald? Meld dat in de tool; de organisator kan je betaling bevestigen.</p>';plain+='Je openstaande bijdrage: '+eur_(a.amountCents)+'. Voorschieter: '+(payer?payer.name:'de organisator')+'. Heb je al betaald? Meld dat in de tool.\n'}
 const url=EE_URL+'?room='+encodeURIComponent(p.roomId);body+='<p><a href="'+url+'" style="display:inline-block;padding:14px 22px;background:#db1f48;color:#ffffff;text-decoration:none;border-radius:6px;font-weight:bold">'+(includeGame?'Bekijk ons avontuur':'Bekijk mijn betaling')+' ↗</a></p>';plain+='\nBekijk ons avontuur: '+url;
 return {subject:(includeGame?when+' spelen we '+room.name:'Je bijdrage voor '+room.name)+' · De Expert Escapers',html:mailFrame_('Bijna tijd voor avontuur.',body),plain};
}
// Bestaande dagelijkse trigger blijft dezelfde functie gebruiken; geen tweede trigger nodig.
function sendPaymentReminders(){
 EE_DB_CACHE=null;EE_RECORD_CACHE={};const lock=LockService.getScriptLock();if(!lock.tryLock(15000))return;
 try{if(props_().getProperty('REMINDERS_ENABLED')!=='true')return;const today=today_(),plans=records_('Plans'),members=records_('Members'),log=records_('MailLog'),rooms=legacyRooms_();
 for(const p of plans)for(const a of p.participants){const u=members.find(m=>m.id===a.userId),room=rooms.find(r=>r.id===p.roomId);if(!u||!u.verifiedAt||u.disabled||!room)continue;
 const gameId=version_(['game',p.id,a.userId,p.date]),paymentId=version_([p.id,a.userId,p.date,a.amountCents]);
 const includeGame=u.gameNotifications!==false&&gameReminderNeeded_(p,today)&&!log.some(l=>l.id===gameId),includePayment=u.notifications!==false&&reminderNeeded_(p,a,today)&&!log.some(l=>l.id===paymentId);
 if(!includeGame&&!includePayment)continue;if(MailApp.getRemainingDailyQuota()<1)return;
 const ids=[];if(includeGame)ids.push({id:gameId,kind:'game'});if(includePayment)ids.push({id:paymentId,kind:'payment'});
 const record=(status,error)=>ids.forEach(item=>put_('MailLog',item.id,{planId:p.id,userId:u.id,kind:item.kind,status,at:now_(),...(error?{error:String(error).slice(0,300)}:{})}));
 // Eerst registreren voorkomt dubbele mails bij een onzekere providerrespons.
 record('attempted');const mail=planReminderMail_(p,a,u,room,members,includeGame,includePayment,today);
 try{sendMail_(u.email,mail.subject,mail.html,mail.plain);record('sent')}catch(err){record('needs_review',err.message)}
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
 switch(x.action){case 'saveRoom':saveRoom_(x,m);break;case 'refreshRoomImage':refreshRoomImage_(x,m);break;case 'archiveRoom':archiveRoom_(x,m);break;case 'saveOpinion':saveOpinion_(x,m);break;case 'addComment':addComment_(x,m);break;case 'deleteComment':deleteComment_(x,m);break;case 'savePlan':savePlan_(x,m);break;case 'payment':payment_(x,m);break;case 'cancelPlan':cancelPlan_(x,m);break;case 'profile':m.name=str_(x.name,80);if(m.name.length<2)fail_('Vul je naam in.');m.notifications=x.notifications!==false;if(x.gameNotifications!==undefined)m.gameNotifications=x.gameNotifications!==false;put_('Members',m.id,m);break;default:fail_('Onbekende actie.');}
 put_('Receipts',receiptId,{userId:m.id,action:x.action,at:now_()});return json_3_({ok:true,data:loadData_(m)});
 }catch(err){return json_3_({ok:false,error:String(err.message||err),code:err.code||'ERROR'})}finally{if(lock)lock.releaseLock()}}
function doGet(){return json_3_({ok:true,version:3,message:'De Expert Escapers · accounts vereist'})}
function json_3_(v){return ContentService.createTextOutput(JSON.stringify(v)).setMimeType(ContentService.MimeType.JSON)}

// Alleen publieke HTTP(S)-websites; maximaal drie gecontroleerde redirects.
function publicImageUrl_(raw,base){
 raw=String(raw||'').trim().replace(/&amp;/gi,'&').replace(/&#(\d+);/g,(_,n)=>String.fromCharCode(Number(n))).replace(/&quot;/gi,'"');
 if(!raw||raw.length>2000||/[\s\\<>"'\x00-\x1f]/.test(raw))return '';
 if(/^[a-z][a-z0-9+.-]*:/i.test(raw)&&!/^https?:\/\//i.test(raw))return '';
 if(raw.indexOf('//')===0)raw='https:'+raw;
 if(!/^https?:\/\//i.test(raw)){
  const b=/^(https?:\/\/[^/]+)(\/[^?#]*)?/i.exec(base||'');if(!b)return '';
  raw=raw[0]==='/'?b[1]+raw:b[1]+(b[2]||'/').replace(/[^/]*$/,'')+raw;
 }
 const m=/^(https?):\/\/([^/?#]+)([^#]*)/i.exec(raw);if(!m)return '';
 const host=m[2].toLowerCase();if(host.includes('@')||host.includes(':')||!host.includes('.')||/^[\d.]+$/.test(host)||!/^([a-z0-9-]+\.)+[a-z]{2,63}$/.test(host)||/(^|\.)(localhost|local|internal|test|invalid|example)$/.test(host)||host.endsWith('.google.internal'))return '';
 return m[1].toLowerCase()+'://'+host+(m[3]||'/');
}
function imageFetch_(url){
 for(let n=0;n<4;n++){
  url=publicImageUrl_(url);if(!url)throw Error('Geen publieke URL.');
  const r=UrlFetchApp.fetch(url,{muteHttpExceptions:true,followRedirects:false,validateHttpsCertificates:true,headers:{'User-Agent':'DeExpertEscapers/1.0 (room preview)'}}),code=r.getResponseCode();
  if([301,302,303,307,308].includes(code)){const h=r.getAllHeaders();url=publicImageUrl_(h.Location||h.location,url);continue}
  if(code<200||code>=300)throw Error('Website niet beschikbaar.');return {response:r,url};
 }
 throw Error('Te veel omleidingen.');
}
function imageAttrs_(tag){const attrs={};String(tag).replace(/([\w:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/g,(_,k,a,b,c)=>{attrs[k.toLowerCase()]=a??b??c;return ''});return attrs}
function imageCandidates_(html,base,name){
 const candidates=[],bad=/logo|favicon|icon|sprite|badge|award|teamfoto|team-photo|group-photo|groepsfoto|visitors|bezoekers|avatar|publieksfavoriet|stars?\.|flag|world\.png|united-kingdom|placeholder|avatar/i;
 const add=(url,score,text)=>{url=publicImageUrl_(url,base);if(!url||bad.test(url+' '+(text||'')))return;const old=candidates.find(c=>c.url===url);if(old)old.score=Math.max(old.score,score);else candidates.push({url,score})};
 const source=String(html).slice(0,500000),words=String(name||'').toLowerCase().split(/[^a-z0-9]+/).filter(w=>w.length>3);
 for(const tag of source.match(/<meta\b[^>]*>/gi)||[]){const a=imageAttrs_(tag),key=(a.property||a.name||'').toLowerCase();if(['og:image','og:image:secure_url','twitter:image','twitter:image:src'].includes(key))add(a.content,60,'')}
 for(const tag of source.match(/<img\b[^>]*>/gi)||[]){const a=imageAttrs_(tag),text=[a.alt,a.title,a.src,a['data-src']].join(' ');if(bad.test(text))continue;const match=words.filter(w=>text.toLowerCase().includes(w)).length,score=30+match*35+(/\.jpe?g|\.webp/i.test(text)?25:0)-(Number(a.height)>Number(a.width)*1.1?60:0);
 const set=(a['data-srcset']||a.srcset||'').split(',').map(v=>v.trim().split(/\s+/)).filter(v=>/^\d+w$/.test(v[1]||'')).sort((a,b)=>parseInt(b[1])-parseInt(a[1]));
 if(set[0])add(set[0][0],score+5,text);else add(a['data-src']||a.src,score,text)
 }
 for(const match of source.matchAll(/background-image\s*:\s*url\(\s*["']?([^"')\s]+)["']?\s*\)/gi)){const url=match[1],score=35+words.filter(w=>url.toLowerCase().includes(w)).length*35+(/\.jpe?g|\.webp/i.test(url)?25:0)+(/escaperoom|interieur|interior|room-|kamer-/i.test(url)?20:0);add(url,score,'')}
 return candidates.sort((a,b)=>b.score-a.score).slice(0,5);
}
function imageFromHtml_(html,base,name){const candidates=imageCandidates_(html,base,name);return candidates.length?candidates[0].url:''}
function imageDimensions_(bytes){
 const b=Array.from(bytes,v=>(v+256)%256),be=(i)=>b[i]*256+b[i+1],le32=(i)=>(b[i]+b[i+1]*256+b[i+2]*65536+b[i+3]*16777216)>>>0;
 if(b.length<24)return null;
 if(b[0]===137&&b[1]===80&&b[2]===78&&b[3]===71)return {width:be(16)*65536+be(18),height:be(20)*65536+be(22)};
 if(b[0]===255&&b[1]===216){let i=2;while(i+8<b.length){if(b[i]!==255){i++;continue}const marker=b[i+1];if(marker===218||marker===217)break;if(marker===216||marker===1||(marker>=208&&marker<=215)){i+=2;continue}const len=be(i+2);if(len<2)break;if([192,193,194,195,197,198,199,201,202,203,205,206,207].includes(marker))return {width:be(i+7),height:be(i+5)};i+=2+len}return null}
 if(String.fromCharCode(...b.slice(0,4))==='RIFF'&&String.fromCharCode(...b.slice(8,12))==='WEBP'){
 const type=String.fromCharCode(...b.slice(12,16));if(type==='VP8X'&&b.length>=30)return {width:1+b[24]+256*b[25]+65536*b[26],height:1+b[27]+256*b[28]+65536*b[29]};
 if(type==='VP8 '&&b.length>=30&&b[23]===157&&b[24]===1&&b[25]===42)return {width:(b[26]+256*b[27])&16383,height:(b[28]+256*b[29])&16383};
 if(type==='VP8L'&&b.length>=25&&b[20]===47){const bits=le32(21);return {width:(bits&16383)+1,height:((bits>>>14)&16383)+1}}
 }return null;
}
function goodRoomPhoto_(size){return !!size&&size.width>=800&&size.height>=450&&size.width/size.height<=3.2&&size.height/size.width<=1.1}
function findRoomImage_(website,name,theme){
 const result=moodRoomImage_(name,theme),url=publicImageUrl_(website);if(!url||/^https?:\/\/(www\.)?escapetalk\.nl\//i.test(url))return result;
 try{const page=imageFetch_(url),r=page.response,h=r.getAllHeaders(),type=String(h['Content-Type']||h['content-type']||'');if(type&&!/text\/html|application\/xhtml/i.test(type))return result;
 for(const candidate of imageCandidates_(r.getContentText(),page.url,name))try{const photo=imageFetch_(candidate.url),bytes=photo.response.getBlob().getBytes();if(bytes.length>10000000)continue;const size=imageDimensions_(bytes);if(goodRoomPhoto_(size)){return {...result,imageUrl:photo.url,imageSource:page.url,imageWidth:size.width,imageHeight:size.height,imageKind:'room',imageCredit:''}}}catch(e){}
 }catch(e){}return result;
}
function refreshRoomImage_(x,m){const r=room_(str_(x.roomId,250));if(r.meta.imagePinned){try{const photo=imageFetch_(r.meta.imageUrl),size=imageDimensions_(photo.response.getBlob().getBytes());if(goodRoomPhoto_(size)){put_('RoomMeta',r.id,{...r.meta,imageCheckedAt:now_()});return}}catch(e){}}
 put_('RoomMeta',r.id,{...r.meta,...findRoomImage_(r.website,r.name,r.theme),imagePinned:false})}
function backfillRoomImages(){
 // Handmatig door de eigenaar; kleine batches houden de uitvoering begrensd.
 EE_DB_CACHE=null;EE_RECORD_CACHE={};let count=0;
 for(const r of legacyRooms_()){const meta=get_('RoomMeta',r.id)||{};if(!meta.archived&&r.website&&meta.imageQuality!==2){put_('RoomMeta',r.id,{...meta,...findRoomImage_(r.website,r.name,r.theme)});count++;if(count>=10)break}}
 console.log('Afbeeldingen gecontroleerd voor '+count+' kamers.');
}

const EE_MOOD_PHOTOS={"jungle":{"imageUrl":"https://images.unsplash.com/photo-1753155438861-97e8d0025742?auto=format&fit=max&w=1600&q=85","imageSource":"https://unsplash.com/photos/ancient-temple-ruins-overgrown-with-green-moss-ssiwp2cmQys","imageWidth":1600,"imageHeight":1067,"imageKind":"mood","imageCredit":"Unsplash"},"prison":{"imageUrl":"https://images.unsplash.com/photo-1546450334-03038e2eaa39?auto=format&fit=max&w=1600&q=85","imageSource":"https://unsplash.com/s/photos/jail-cell","imageWidth":1600,"imageHeight":1063,"imageKind":"mood","imageCredit":"Unsplash"},"egypt":{"imageUrl":"https://images.unsplash.com/photo-1569307706396-9ac7325ff281?auto=format&fit=max&w=1600&q=85","imageSource":"https://unsplash.com/photos/a-room-that-has-some-writing-on-the-walls-r6Ic2Q1ffoE","imageWidth":1600,"imageHeight":1071,"imageKind":"mood","imageCredit":"Unsplash"},"forest":{"imageUrl":"https://images.unsplash.com/photo-1449158743715-0a90ebb6d2d8?auto=format&fit=max&w=1600&q=85","imageSource":"https://unsplash.com","imageWidth":1600,"imageHeight":1067,"imageKind":"mood","imageCredit":"Unsplash"},"castle":{"imageUrl":"https://www.castellobevilacqua.com/library/reservation/565/gallery_evento_CastelloBevilacqua-Ditatour1.jpg","imageSource":"https://www.castellobevilacqua.com/","imageWidth":800,"imageHeight":534,"imageKind":"mood","imageCredit":""},"mystery":{"imageUrl":"https://images.unsplash.com/photo-1542911077-46bd18a594f6?auto=format&fit=max&w=1600&q=85","imageSource":"https://unsplash.com/es/s/fotos/bibliotecas","imageWidth":1600,"imageHeight":1116,"imageKind":"mood","imageCredit":"Unsplash"},"science":{"imageUrl":"https://www.escapingrooms.be/filerep/contentblocks/67/1_Patient_x_escaping2.jpeg","imageSource":"https://www.escapingrooms.be/en/escape-rooms/patient-x","imageWidth":800,"imageHeight":533,"imageKind":"mood","imageCredit":""},"war":{"imageUrl":"https://framerusercontent.com/images/eDle9JhbNkpxk80HLH2VG3jPQ.webp?width=4290&height=2856","imageSource":"https://escapefortliezele.be/","imageWidth":4290,"imageHeight":2856,"imageKind":"mood","imageCredit":""},"space":{"imageUrl":"https://degoudenkooi.pluxit.be/web/assets/Heros/Back-Space.webp","imageSource":"https://www.degoudenkooi.be/nl/escape-rooms/lost-in-space","imageWidth":1280,"imageHeight":731,"imageKind":"mood","imageCredit":""}};
function moodRoomImage_(name,theme){const text=String(name||'')+' '+String(theme||'');let kind='mystery';
 if(/jumanji|jaguar|jungle|kalakmul/i.test(text))kind='jungle';else if(/prison|gevangenis|break.*bars/i.test(text))kind='prison';else if(/egypt|farao|phara|pyramid|tombe/i.test(text))kind='egypt';else if(/space|ruimte|alien|sci.?fi/i.test(text))kind='space';else if(/lab|patient|patiënt|room 22|science|wetenschap/i.test(text))kind='science';else if(/armor|armadura|kasteel|medieval|ridder/i.test(text))kind='castle';else if(/1943|war|oorlog|bunker/i.test(text))kind='war';else if(/cabin|woods|hunting|forest|bos/i.test(text))kind='forest';
 return {...EE_MOOD_PHOTOS[kind],imageKind:'mood',imageQuality:2,imageCheckedAt:now_()};
}

const EE_RESEARCHED_PHOTOS=[{"roomId":"et-de-gouden-kooi-escape-experience--arbreville","name":"Arbreville","imageUrl":"https://degoudenkooi.pluxit.be/web/assets/content-rooms/Arbreville/back-arbreville_2025-12-02-152252_jtfg.webp","imageSource":"https://www.degoudenkooi.be/nl/escape-rooms/arbreville","imageWidth":1280,"imageHeight":731,"imageKind":"mood","imageCredit":""},{"roomId":"et-la-fabrica--la-maldicion-del-jaguar","name":"La Maldición del Jaguar","imageUrl":"https://images.unsplash.com/photo-1753155438861-97e8d0025742?auto=format&fit=max&w=1600&q=85","imageSource":"https://unsplash.com/photos/ancient-temple-ruins-overgrown-with-green-moss-ssiwp2cmQys","imageWidth":1600,"imageHeight":1067,"imageKind":"mood","imageCredit":"Unsplash"},{"roomId":"et-the-x-door-alicante--cain","name":"CAIN","imageUrl":"https://www.thexdooralicante.com/wp-content/uploads/cain-01-thegem-gallery-sidebar.jpg","imageSource":"https://www.thexdooralicante.com/en/escape-room-cain/","imageWidth":867,"imageHeight":540,"imageKind":"room","imageCredit":""},{"roomId":"et-escape-hotel--waar-is-rita","name":"Where is Rita?","imageUrl":"https://7ir5cqnwkgzm.b-cdn.net/w_1920/s3-escape-hotel/ae0c396b-0af9-4f27-99dd-216899bf90c6.jpg","imageSource":"https://escapehotel.be/en-gb/escape-rooms/where-is-rita","imageWidth":1920,"imageHeight":1440,"imageKind":"room","imageCredit":""},{"roomId":"et-escaperoom-antwerp--monster-mashers-dark","name":"Monster Mashers DARK","imageUrl":"https://torchdale.be/public/storage/uploads/Bgh9MmcR/2500/wwMUv5lE7c.png","imageSource":"https://torchdale.be/nl/monster-mashers","imageWidth":2500,"imageHeight":1004,"imageKind":"room","imageCredit":""},{"roomId":"et-exitroom-leuven--prison-break-leuven","name":"Prison Break","imageUrl":"https://images.unsplash.com/photo-1546450334-03038e2eaa39?auto=format&fit=max&w=1600&q=85","imageSource":"https://unsplash.com/s/photos/jail-cell","imageWidth":1600,"imageHeight":1063,"imageKind":"mood","imageCredit":"Unsplash"},{"roomId":"et-best-escape-vlaanderen--the-factory","name":"The Factory","imageUrl":"https://bestescape.be/wp-content/uploads/2026/08/the-factory.png","imageSource":"https://bestescape.be/rooms/the-factory/","imageWidth":1535,"imageHeight":1024,"imageKind":"room","imageCredit":""},{"roomId":"et-escaping-belgium--le-prince","name":"Le Prince","imageUrl":"https://www.escaping.be/wp-content/uploads/2026/10/og-leprince.jpg","imageSource":"https://www.escaping.be/leprince/","imageWidth":1200,"imageHeight":630,"imageKind":"room","imageCredit":""},{"roomId":"et-de-gouden-kooi-escape-experience--de-wraak-van-han","name":"De wraak van Han","imageUrl":"https://degoudenkooi.pluxit.be/web/assets/Heros/Back-Han.webp","imageSource":"https://www.degoudenkooi.be/nl/escape-rooms/de-wraak-van-han","imageWidth":1280,"imageHeight":731,"imageKind":"mood","imageCredit":""},{"roomId":"et-de-kraakfabriek--de-illusionist-antwerpen","name":"De Illusionist","imageUrl":"https://kraakfabriek.be/img/Escape-Room/escaperoom39.jpg","imageSource":"https://kraakfabriek.be/index.html","imageWidth":1024,"imageHeight":683,"imageKind":"room","imageCredit":""},{"roomId":"et-escape-in-time--pyramid-of-doom","name":"Pyramid Of Doom","imageUrl":"https://images.unsplash.com/photo-1569307706396-9ac7325ff281?auto=format&fit=max&w=1600&q=85","imageSource":"https://unsplash.com/photos/a-room-that-has-some-writing-on-the-walls-r6Ic2Q1ffoE","imageWidth":1600,"imageHeight":1071,"imageKind":"mood","imageCredit":"Unsplash"},{"roomId":"et-escaperoom-antwerp--hunting-season","name":"Hunting Season","imageUrl":"https://images.unsplash.com/photo-1449158743715-0a90ebb6d2d8?auto=format&fit=max&w=1600&q=85","imageSource":"https://unsplash.com","imageWidth":1600,"imageHeight":1067,"imageKind":"mood","imageCredit":"Unsplash"},{"roomId":"et-locked-escape-room-leuven--flight-923","name":"Flight 923","imageUrl":"https://www.escapereview.be/wp-content/uploads/2021/08/11.png","imageSource":"https://www.escapereview.be/neergestort-in-de-jungle-van-peru-op-onze-huwelijksverjaardag/","imageWidth":900,"imageHeight":900,"imageKind":"mood","imageCredit":""},{"roomId":"et-escape-fort-liezele--the-mystery-of-anna-lockwood","name":"The Mystery Of Anna Lockwood","imageUrl":"https://framerusercontent.com/images/eDle9JhbNkpxk80HLH2VG3jPQ.webp?width=4290&height=2856","imageSource":"https://escapefortliezele.be/","imageWidth":4290,"imageHeight":2856,"imageKind":"room","imageCredit":""},{"roomId":"et-exitroom-leuven--box-of-mysteries","name":"Box Of Mysteries","imageUrl":"https://thexdooralicante.com/wp-content/uploads/home-slider.jpg","imageSource":"https://www.thexdooralicante.com/en/","imageWidth":1030,"imageHeight":684,"imageKind":"mood","imageCredit":""},{"roomId":"et-exitroom-leuven--jumanji","name":"Jumanji","imageUrl":"https://images.unsplash.com/photo-1753155438861-97e8d0025742?auto=format&fit=max&w=1600&q=85","imageSource":"https://unsplash.com/photos/ancient-temple-ruins-overgrown-with-green-moss-ssiwp2cmQys","imageWidth":1600,"imageHeight":1067,"imageKind":"mood","imageCredit":"Unsplash"},{"roomId":"et-escaping-antwerp--patient-x-antwerpen","name":"Patiënt X","imageUrl":"https://www.escapingrooms.be/filerep/escapeRoomTypeImages/5/58.jpeg","imageSource":"https://www.escapingrooms.be/en/escape-rooms/patient-x","imageWidth":1621,"imageHeight":1080,"imageKind":"room","imageCredit":""},{"roomId":"et-het-geheim-baron-beeckman--het-geheim-baron-beeckman","name":"Het Genootschap","imageUrl":"https://www.baronbeeckman.be/wp-content/uploads/2020/06/Escape-game-Diest-Het-geheim-Baron-Beeckman.jpg","imageSource":"https://www.baronbeeckman.be/","imageWidth":1280,"imageHeight":851,"imageKind":"mood","imageCredit":""},{"roomId":"et-het-geheim-baron-beeckman--het-testament","name":"Het Testament","imageUrl":"https://www.baronbeeckman.be/wp-content/uploads/2020/06/Baron-Beeckman-Escaperoom.jpg","imageSource":"https://www.baronbeeckman.be/","imageWidth":1280,"imageHeight":1011,"imageKind":"mood","imageCredit":""},{"roomId":"et-la-fabrica-altea--harrys-maze","name":"Harry's maze","imageUrl":"https://images.unsplash.com/photo-1542911077-46bd18a594f6?auto=format&fit=max&w=1600&q=85","imageSource":"https://unsplash.com/es/s/fotos/bibliotecas","imageWidth":1600,"imageHeight":1116,"imageKind":"mood","imageCredit":"Unsplash"},{"roomId":"et-paradox-escape-room-benidorm--armor-of-the-gods","name":"Armor of the Gods [Las Armaduras]","imageUrl":"https://www.castellobevilacqua.com/library/reservation/565/gallery_evento_CastelloBevilacqua-Ditatour1.jpg","imageSource":"https://www.castellobevilacqua.com/","imageWidth":800,"imageHeight":534,"imageKind":"mood","imageCredit":""},{"roomId":"et-xcape-room-finestrat--el-sombrerero-de-alicia-alices-hatter","name":"Alice’s Hatter [El sombrerero de Alicia]","imageUrl":"https://gpsescape.es/img/salas/xcape-finestrat-el-sombrerero-de-alicia.webp","imageSource":"https://gpsescape.es/alicante/sala/xcape-finestrat-el-sombrerero-de-alicia/","imageWidth":1280,"imageHeight":716,"imageKind":"mood","imageCredit":""},{"roomId":"et-de-gouden-kooi-escape-experience--de-nekker","name":"De Nekker","imageUrl":"https://degoudenkooi.pluxit.be/web/assets/content-rooms/De-Nekker/nekker.webp","imageSource":"https://www.degoudenkooi.be/nl/escape-rooms/de-nekker","imageWidth":1280,"imageHeight":731,"imageKind":"mood","imageCredit":""},{"roomId":"et-de-gouden-kooi--cabin-in-the-woods","name":"Cabin In The Woods","imageUrl":"https://images.unsplash.com/photo-1449158743715-0a90ebb6d2d8?auto=format&fit=max&w=1600&q=85","imageSource":"https://unsplash.com","imageWidth":1600,"imageHeight":1067,"imageKind":"mood","imageCredit":"Unsplash"},{"roomId":"et-de-gouden-kooi--de-meesterdief-van-mechelen","name":"The master thief of Mechelen","imageUrl":"https://degoudenkooi.pluxit.be/web/assets/Heros/Back-Dief.webp","imageSource":"https://www.degoudenkooi.be/nl/escape-rooms/de-meesterdief-van-mechelen","imageWidth":1280,"imageHeight":731,"imageKind":"mood","imageCredit":""},{"roomId":"et-de-gouden-kooi--de-schat-van-kalakmul","name":"The treasure of Kalakmul","imageUrl":"https://degoudenkooi.pluxit.be/web/assets/Heros/Back-Kalakmul.webp","imageSource":"https://www.degoudenkooi.be/nl/escape-rooms/de-schat-van-kalakmul","imageWidth":1280,"imageHeight":731,"imageKind":"mood","imageCredit":""},{"roomId":"et-escape-game-center-parcs-de-vossemeren--de-vloek-van-de-farao","name":"De vloek van de Farao","imageUrl":"https://images.unsplash.com/photo-1569307706396-9ac7325ff281?auto=format&fit=max&w=1600&q=85","imageSource":"https://unsplash.com/photos/a-room-that-has-some-writing-on-the-walls-r6Ic2Q1ffoE","imageWidth":1600,"imageHeight":1071,"imageKind":"mood","imageCredit":"Unsplash"},{"roomId":"et-escape-house-emerantia--emerantia","name":"Emerantia","imageUrl":"https://www.baronbeeckman.be/wp-content/uploads/2020/06/Baron-Beeckman-Escaperoom.jpg","imageSource":"https://www.baronbeeckman.be/","imageWidth":1280,"imageHeight":1011,"imageKind":"mood","imageCredit":""},{"roomId":"et-de-gouden-kooi-escape-experience--het-geheim-van-sint-rumoldus","name":"The secret of Saint-Rumoldus","imageUrl":"https://degoudenkooi.pluxit.be/web/assets/Heros/Back-Rumoldus.webp","imageSource":"https://www.degoudenkooi.be/nl/escape-rooms/het-geheim-van-sint-rumoldus","imageWidth":1280,"imageHeight":731,"imageKind":"mood","imageCredit":""},{"roomId":"et-escaping-antwerp--lets-rob-the-bank-antwerpen","name":"Let's Rob the Bank!","imageUrl":"https://www.escaping.nl/filerep/escapeRoomTypeImages/1/45.png","imageSource":"https://www.escaping.nl/en/escape-rooms/lets-rob-the-bank","imageWidth":1620,"imageHeight":1080,"imageKind":"room","imageCredit":""},{"roomId":"et-de-gouden-kooi--lost-in-space","name":"Lost In Space","imageUrl":"https://degoudenkooi.pluxit.be/web/assets/Heros/Back-Space.webp","imageSource":"https://www.degoudenkooi.be/nl/escape-rooms/lost-in-space","imageWidth":1280,"imageHeight":731,"imageKind":"mood","imageCredit":""},{"roomId":"et-the-lock-escape-mortsel--mortsel-1943","name":"Mortsel 1943","imageUrl":"https://framerusercontent.com/images/eDle9JhbNkpxk80HLH2VG3jPQ.webp?width=4290&height=2856","imageSource":"https://escapefortliezele.be/","imageWidth":4290,"imageHeight":2856,"imageKind":"mood","imageCredit":""},{"roomId":"et-keys-clues--newtons-theory","name":"Newton's Secret","imageUrl":"https://images.unsplash.com/photo-1542911077-46bd18a594f6?auto=format&fit=max&w=1600&q=85","imageSource":"https://unsplash.com/es/s/fotos/bibliotecas","imageWidth":1600,"imageHeight":1116,"imageKind":"mood","imageCredit":"Unsplash"},{"roomId":"et-escape-room-leuven--prison-orcus","name":"Prison Orcus","imageUrl":"https://www.exitgames.be/web/image/569686-3309644e/Exitgames%20Leuven%20-%20Prison%20Orcus%2000892.jpg","imageSource":"https://www.exitgames.be/escape-room-leuven","imageWidth":1920,"imageHeight":1280,"imageKind":"room","imageCredit":""},{"roomId":"et-escaping-antwerp--room-22-antwerp","name":"Room 22","imageUrl":"https://www.escapingrooms.be/filerep/contentblocks/67/1_Patient_x_escaping2.jpeg","imageSource":"https://www.escapingrooms.be/en/escape-rooms/patient-x","imageWidth":800,"imageHeight":533,"imageKind":"mood","imageCredit":""}];
function installPlayedRoomPhotos(){
 const lock=LockService.getScriptLock();lock.waitLock(20000);try{EE_DB_CACHE=null;EE_RECORD_CACHE={};const rooms=legacyRooms_();let count=0;
 for(const image of EE_RESEARCHED_PHOTOS){const r=rooms.find(r=>r.id===image.roomId);if(!r)continue;const meta=get_('RoomMeta',r.id)||{};if(meta.archived)continue;
 const {roomId,name,...fields}=image;if(!goodRoomPhoto_({width:fields.imageWidth,height:fields.imageHeight})||!publicImageUrl_(fields.imageUrl))throw Error('Ongeldige onderzoeksfoto.');
 put_('RoomMeta',r.id,{...meta,...fields,imageQuality:2,imagePinned:true,imageCheckedAt:now_()});count++}
 console.log('Gespeelde kamers bijgewerkt: '+count+'. 10 kamer-/locatiebeelden en 25 thematische sfeerbeelden. Geen teamfoto’s.');
 }finally{lock.releaseLock()}
}
