'use strict';
(() => {
 const $=id=>document.getElementById(id),view=$('view'),actions=$('actions'),status=$('status');
 const state={cfg:null,key:'',who:null,subject:'public',tab:'Calendar',sub:'Today',rows:[],shared:null,docs:{},people:[],my:null,day:today(),month:today().slice(0,7),rangeStart:today(),rangeEnd:today(),pending:null,busy:false};
 function today(){return new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Shanghai',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());}
 function node(tag,text,props={}){const e=document.createElement(tag);if(text!=null)e.textContent=text;Object.assign(e,props);return e;}
 function button(text,fn,primary=false){const b=node('button',text);b.type='button';if(primary)b.className='primary';b.addEventListener('click',fn);return b;}
 function say(text,error=false){status.textContent=text;status.className=error?'error':'';}
 function card(title){const c=node('div',null,{className:'card'});if(title)c.append(node('h2',title));view.append(c);return c;}
 function toolbar(items){const r=node('div',null,{className:'toolbar'});r.append(...items);return r;}
 function select(values,value,fn){const s=node('select');for(const[v,t]of values)s.append(node('option',t,{value:v}));s.value=value;s.onchange=()=>fn(s.value);return s;}
 function table(parent,headers,rows){const wrap=node('div',null,{className:'scroll'}),t=node('table'),head=node('tr');headers.forEach(h=>head.append(node('th',h)));t.append(head);for(const row of rows){const tr=node('tr');row.forEach(v=>{const td=node('td');if(v instanceof Node)td.append(v);else td.textContent=v??'';tr.append(td);});t.append(tr);}wrap.append(t);parent.append(wrap);if(!rows.length)parent.append(node('p','No records to show.'));}
 function modal(title,build){$('dialogBody').replaceChildren(node('h2',title));build($('dialogBody'));if(!$('dialog').open)$('dialog').showModal();}
 function input(parent,label,value='',type='text'){const id='f-'+crypto.randomUUID(),l=node('label',label,{htmlFor:id}),i=node(type==='textarea'?'textarea':'input',null,{id,value:value??''});if(type!=='textarea')i.type=type;parent.append(l,i);return i;}
 function download(name,value){const url=URL.createObjectURL(new Blob([JSON.stringify(value,null,2)],{type:'application/json'})),a=node('a',null,{href:url,download:name+'.json'});a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);}
 async function rpc(name,args={},withKey=true){const r=await fetch(state.cfg.url+'/rest/v1/rpc/'+name,{method:'POST',headers:{apikey:state.cfg.key,'Content-Type':'application/json'},body:JSON.stringify({...((withKey)?{p_key:state.key||null}:{}),...args}),cache:'no-store',signal:AbortSignal.timeout(30000)});const data=await r.json();if(!r.ok){const e=Error(data.message||'Request failed ('+r.status+')');e.code=data.code;throw e;}return data;}
 async function allDocuments(module){let rows=[];for(let offset=0;;offset+=500){const page=await rpc('ms_documents',{p_module:module,p_offset:offset,...(state.who?.role==='calendar-editor'?{p_key:null}:{})});if(!Array.isArray(page))throw Error('Unexpected database response');rows.push(...page);if(page.length<500)return rows;}}
 const owner=()=>state.who?.role==='owner';
 const admin=()=>['owner','administrator'].includes(state.who?.role);
 const canEdit=module=>(owner()&&(state.who.active_release||['assemblies','newsletter'].includes(module)))||(state.who?.active_release&&['editor','administrator'].includes(state.who.role)&&state.who.modules.includes(module))||(module==='calendar'&&state.who?.role==='calendar-editor');
 async function connect(key){
  let who;try{who=await rpc('ms_access',{p_key:key});}catch(e){if(e.code!=='42501')throw e;await rpc('calendar_shared_list',{p_key:key});who={role:'calendar-editor',modules:['calendar'],label:'Calendar editing key'};}
  state.key=key;state.who=who;
  const hash=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(key));state.subject=who.id||Array.from(new Uint8Array(hash),b=>b.toString(16).padStart(2,'0')).join('');
  try{sessionStorage.setItem('mastersheet553.access',key);}catch{}
  try{state.pending=JSON.parse(localStorage.getItem('mastersheet553.draft.'+state.subject));}catch{}
  state.shared=null;state.docs={};state.my=null;state.people=[];state.personId=null;drawAccess();navigation();await load();
 }
 function unlock(){modal('Unlock personal or editing access',body=>{
  body.append(node('p','Shared views and staff dropdowns are open. Your access key reveals authorised editing and management tools.'));
  const key=input(body,'Access key','','password'),feedback=node('p');key.autocomplete='off';body.append(feedback);
  body.append(button('Unlock',async()=>{try{if(!/^[a-f0-9]{64}$/.test(key.value.trim()))throw Error('Paste the full 64-character access key.');await connect(key.value.trim());key.value='';$('dialog').close();}catch(e){feedback.textContent=e.message;}},true));
 });}
 function ownerUnsaved(){const frame=view.querySelector('iframe');return !!frame?.contentWindow?.MS553Owner?.hasUnsaved();}
 function disconnect(){if(state.busy||state.pending||ownerUnsaved())return say('Finish the shared save or export and resolve the retained draft before locking.',true);try{sessionStorage.removeItem('mastersheet553.access');}catch{}state.key='';state.who=null;state.subject='public';state.my=null;state.rows=[];state.people=[];state.docs={};state.shared=null;state.tab='Calendar';state.sub='Today';drawAccess();navigation();void load();}
 function drawAccess(){const host=$('access');host.replaceChildren();if(state.who){host.append(node('span',state.who.label+' · '+state.who.role+' ',{className:'badge'}),button('Lock access',disconnect));if(owner())host.append(button('Manage access keys',manageKeys),button('Database manager',()=>databaseManager('calendar')),button('Install private setup / update',installPrivate));}else host.append(button('Unlock personal / editing access',unlock));if(state.pending)host.append(button('Retained save draft',recovery));}
 const tabs=['Calendar','Duties','Staff','Enrichment','School Information'];
 const subviews={Calendar:['Today','This Week','This Month','Calendar','Academic Year','Date range'],Duties:['Right Now','My Duties','Expectations'],Staff:['My Week'],Enrichment:["Today's ECA",'My ECA','Choices','Clubs & Teams','Sports','Saturday Enrichment'],'School Information':['Assembly','Newsletter','Handbooks']};
 function navigation(){const nav=$('navigation');nav.replaceChildren();for(const tab of tabs){const b=button(tab,()=>{state.tab=tab;state.sub=subviews[tab][0];void load();});if(tab===state.tab)b.classList.add('active');nav.append(b);}if(owner())nav.append(button('Original page',()=>{window.MS554EditingTools=false;void load();}));}
 async function load(){
 if(!window.MS554EditingTools){if(owner())return openOwner();return openOriginal();}

  if(state.busy||state.pending){recovery();return;}
  if(ownerUnsaved()){say('Finish the owner workspace save before leaving it.',true);return;}
  view.replaceChildren();actions.replaceChildren();navigation();
  for(const sub of subviews[state.tab]){const b=button(sub,()=>{state.sub=sub;void load();});if(sub===state.sub)b.classList.add('primary');actions.append(b);}
  say('Loading…');
  try{
   if(state.tab==='Calendar'){
    const key=canEdit('calendar')?state.key:null;state.rows=await rpc('calendar_shared_list',{p_key:key});calendar();
   }else if(['My Week','My Duties','My ECA'].includes(state.sub)){await personal();}
   else if(state.sub==='Assembly'||state.sub==='Newsletter'){const module=state.sub==='Assembly'?'assemblies':'newsletter';state.docs[module]=await allDocuments(module);if(module==='assemblies')assemblies();else newsletter();}
   else {if(!state.shared)state.shared=await rpc('ms_shared_staff',{},!!state.key&&state.who?.active_release===false);sharedViews();}
   say('Up to date · '+(state.who?state.who.label:'public read-only'));
  }catch(e){say(e.message+' No changes were made.',true);card('Connection needed').append(node('p','The new views require the matching v5.53 database preparation and activation. Calendar continues to use the existing database.'));}
 }
 function datePlus(date,days){const d=new Date(date+'T12:00:00Z');d.setUTCDate(d.getUTCDate()+days);return d.toISOString().slice(0,10);}
 function calendar(){
  const dayMode=['Today','This Week'].includes(state.sub),rangeMode=state.sub==='Date range',yearMode=state.sub==='Academic Year',schoolYear=+state.month.slice(0,4)-(+state.month.slice(5)<9?1:0),c=card(state.sub),row=toolbar([]),date=node('input',null,{type:yearMode?'number':dayMode||rangeMode?'date':'month',value:yearMode?schoolYear:rangeMode?state.rangeStart:dayMode?state.day:state.month});date.onchange=()=>{if(!date.value)return;if(yearMode)state.month=date.value+'-09';else if(rangeMode)state.rangeStart=date.value;else if(dayMode)state.day=date.value;else state.month=date.value;calendarRedraw();};
  const search=node('input',null,{type:'search',placeholder:'Search Calendar'}),section=select([['all','All sections'],['US','US'],['LS','LS'],['KG','KG'],['WS','Whole school']],'all',calendarRedraw);
  row.append(date,search,section,button('Today',()=>{state.day=today();state.month=today().slice(0,7);void load();}),button('Print',()=>window.print()));
  if(rangeMode){const until=node('input',null,{type:'date',value:state.rangeEnd});until.onchange=()=>{if(until.value){state.rangeEnd=until.value;calendarRedraw();}};row.insertBefore(node('span','to'),search);row.insertBefore(until,search);}
  if(canEdit('calendar'))row.append(button('Add event',()=>editEvent(null),true));c.append(row);const result=node('div');c.append(result);search.oninput=calendarRedraw;
  function calendarRedraw(){
   result.replaceChildren();let start,end;
   if(state.sub==='Today'){start=state.day;end=start;}
   else if(state.sub==='This Week'){const day=new Date(state.day+'T12:00:00Z').getUTCDay();start=datePlus(state.day,-((day+6)%7));end=datePlus(start,6);}
   else if(rangeMode){start=state.rangeStart;end=state.rangeEnd;}
   else if(yearMode){const year=+state.month.slice(0,4)-(+state.month.slice(5)<9?1:0);start=year+'-09-01';end=(year+1)+'-08-31';}
   else {start=state.month+'-01';end=datePlus(new Date(Date.UTC(+state.month.slice(0,4),+state.month.slice(5,7),1)).toISOString().slice(0,10),-1);}
   if(end<start){result.append(node('p','Choose an end date on or after the start date.'));return;}
   const rows=state.rows.filter(r=>!r.deleted_at&&r.data.date<=end&&(r.data.end||r.data.date)>=start&&(!search.value||r.data.text.toLowerCase().includes(search.value.toLowerCase()))&&(section.value==='all'||(r.data.labels||[]).includes(section.value)||(r.data.labels||[]).includes('WS'))).sort((a,b)=>a.data.date.localeCompare(b.data.date)||a.data.text.localeCompare(b.data.text));
   result.append(node('p',start+' to '+end+' · '+rows.length+' events',{className:'muted'}));
   if(state.sub==='Calendar'){
    const grid=node('div',null,{className:'grid'});for(const name of ['Mon','Tue','Wed','Thu','Fri','Sat','Sun'])grid.append(node('b',name,{className:'weekday'}));const offset=(new Date(start+'T12:00:00Z').getUTCDay()+6)%7;for(let i=0;i<offset;i++)grid.append(node('div',null,{className:'day blank'}));for(let day=start;day<=end;day=datePlus(day,1)){const cell=node('div',null,{className:'day'}),ds=day;cell.append(button(day.slice(8)+' · '+new Date(day+'T12:00:00Z').toLocaleDateString('en-GB',{weekday:'short'}),()=>{state.day=ds;state.sub='Today';void load();}));for(const r of rows.filter(r=>r.data.date<=day&&(r.data.end||r.data.date)>=day)){const a=node('a',r.data.text);a.onclick=()=>eventDetails(r);cell.append(a);}grid.append(cell);}result.append(grid);
   }else table(result,['Date','Event','Time','Location','Status'],rows.map(r=>[r.data.date+(r.data.end&&r.data.end!==r.data.date?' – '+r.data.end:''),button(r.data.text,()=>eventDetails(r)),r.data.timeText||({whole:'Whole day',first:'First half',second:'Second half'}[r.data.timeMode]||''),r.data.location,(r.data.cancelled?'Cancelled':r.data.confirmed?'Confirmed':'Draft')+(r.data.approved?' · Approved':'')]));
  }calendarRedraw();
 }
 function eventDetails(row){modal(row.data.text,body=>{for(const[key,label]of [['date','Date'],['end','Until'],['timeText','Time'],['location','Location'],['lead','Lead'],['notes','Notes'],['rundown','Rundown'],['vips','Key people']])if(row.data[key])body.append(node('h3',label),node('p',row.data[key]));if(canEdit('calendar'))body.append(button('Edit event',()=>editEvent(row),true));});}
 function editEvent(row){modal(row?'Edit event':'Add event',body=>{
  const data=row?.data||{id:'ev-'+crypto.randomUUID(),date:state.day,end:state.day,labels:['WS'],aud:[],type:'Events',approved:false,confirmed:true,rec:{freq:'none'}};
  const name=input(body,'Event name',data.text),date=input(body,'Start date',data.date,'date'),end=input(body,'End date',data.end||data.date,'date'),time=input(body,'Time',data.timeText||''),location=input(body,'Location',data.location||''),lead=input(body,'Lead',data.lead||''),notes=input(body,'Notes — private to Calendar editors',data.notes||'','textarea');
  const sections=input(body,'Sections (US, LS, KG, WS)',(data.labels||[]).join(','));
  const buttons=toolbar([button('Save to Calendar',async()=>{try{if(!name.value.trim())throw Error('Event name required');if(end.value<date.value)throw Error('End date precedes start date');const next={...data,text:name.value.trim(),date:date.value,end:end.value||date.value,timeText:time.value,timeMode:time.value?'custom':data.timeMode||'whole',location:location.value,lead:lead.value,notes:notes.value,labels:sections.value.split(',').map(x=>x.trim()).filter(Boolean)};await write('calendar_shared_apply',{p_request_id:crypto.randomUUID(),p_ops:[{id:data.id,action:'save',expected_revision:row?.revision||0,scope:row?.edit_scope||(next.labels.includes('WS')?'WS':next.labels.find(x=>['US','LS','KG'].includes(x))||'US'),data:next}]});}catch(e){say(e.message,true);}},true)]);
  if(row)buttons.append(button('Delete event',()=>modal('Delete this event?',m=>{m.append(node('p',row.data.text),button('Delete',()=>write('calendar_shared_apply',{p_request_id:crypto.randomUUID(),p_ops:[{id:row.id,action:'delete',expected_revision:row.revision}]})));})));body.append(buttons);
 });}
 async function personal(){
  const c=card(state.sub);if(!state.key||state.who?.role==='calendar-editor'){c.append(node('p','Use your personal staff key to see your own information. No staff search is available to other staff.'),button('Unlock my views',unlock,true));return;}
  if(admin()){
   if(!state.people.length)state.people=await rpc('ms_people');
   const selected=state.personId||state.who.staff_id||state.people[0]?.id;
   c.append(select(state.people.map(p=>[p.id,p.name+' ('+p.code+')']),selected,id=>{state.personId=id;state.my=null;void load();}));
   state.personId=selected;
  }
  state.my=await rpc('ms_my_week',{p_staff_id:admin()?state.personId:null});c.append(node('p',state.my.staff.name));
  if(state.my.unresolved_lessons)c.append(node('p','Some lesson rows need the owner to confirm their staff links. They are withheld rather than guessed.',{className:'muted'}));
  if(state.sub==='My Week')table(c,['Day','Time','Lesson','Room'],state.my.lessons.map(r=>[r.day,[r.start,r.end].filter(Boolean).join('–')||r.slot,r.subject||r.label,r.room]));
  if(state.sub==='My Week'||state.sub==='My Duties'){c.append(node('h3','Duties'));table(c,['Day','Time','Duty','Location','Rota week'],state.my.duties.map(r=>[r.day,r.time,r.scope,r.location,r.week]));}
  if(state.sub==='My Week'||state.sub==='My ECA'){c.append(node('h3','ECA assignments'));table(c,['Day','ECA','Role','Room','Grades'],state.my.ecas.map(r=>[r.day,r.name,r.role,r.room,r.grades]));}
 }
 function people(value){return Array.isArray(value)?value.filter(v=>v!=null&&String(v).trim()).map(String).join(', '):String(value??'');}
 function sharedViews(){const d=state.shared,c=card(state.sub);
  if(state.sub==='Right Now'){
   let day=new Intl.DateTimeFormat('en-GB',{weekday:'long',timeZone:'Asia/Shanghai'}).format(new Date()),now=new Intl.DateTimeFormat('en-GB',{hour:'2-digit',minute:'2-digit',hour12:false,timeZone:'Asia/Shanghai'}).format(new Date());const weekLabel=d.week_labels?.find(r=>r.date===today())?.week||'',weekMatch=/^Week\s+(\d+)$/.exec(weekLabel),dayType=d.school_days?.find(r=>r.date===today())?.type||'';
   let mode='now',week=weekMatch?(+weekMatch[1]%2?'A':'B'):'both';const days=['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],matches=(value,wanted)=>String(value||'').toLowerCase().startsWith(wanted.slice(0,3).toLowerCase());
   const result=node('div'),draw=()=>{result.replaceChildren();const yesterday=days[(days.indexOf(day)+6)%7],rows=d.duties.filter(r=>{const start=(r.start||'').slice(0,5),end=(r.end||'').slice(0,5);if(mode==='day')return matches(r.day,day);if(start&&end&&end<start)return(matches(r.day,day)&&now>=start)||(matches(r.day,yesterday)&&now<end);return matches(r.day,day)&&(!start||!end||(start<=now&&end>now));});table(result,['Duty','Time','Location','Assigned staff'],rows.map(r=>[r.scope+(r.section?' · '+r.section:''),r.time||[r.start,r.end].join('–'),r.location,r.biweekly?(week==='A'?people(r.people):week==='B'?people(r.peopleB):'A: '+people(r.people)+'\nB: '+people(r.peopleB)):people(r.people)]));const next=d.duties.filter(r=>matches(r.day,day)&&r.start&&r.start.slice(0,5)>now).sort((a,b)=>a.start.localeCompare(b.start))[0];if(mode==='now'&&next)result.append(node('p','Next scheduled slot: '+next.start.slice(0,5)+' · '+next.scope+' · '+(next.location||'')));};
   const clock=node('input',null,{type:'time',value:now});clock.onchange=()=>{now=clock.value;draw();};c.append(node('p','School time (Hangzhou) · '+today()+' · '+dayType+(weekLabel?' · '+weekLabel:'')),toolbar([select(days.map(v=>[v,v]),day,v=>{day=v;draw();}),clock,select([['now','Current slots'],['day','All selected day']],'now',v=>{mode=v;draw();}),select([['both','Show both rota weeks'],['A','Week A'],['B','Week B']],week,v=>{week=v;draw();}),button('Reset to now',()=>void load())]),result);draw();
  }else if(state.sub==='Expectations'){for(const[k,v]of Object.entries(d.expectations)){if(!v)continue;const label=k.replace(/([A-Z])/g,' $1');c.append(node('h3',label[0].toUpperCase()+label.slice(1)),node('div',readable(v),{className:'reader'}));}}
  else if(state.sub==='Handbooks'){
   const books=Object.entries(d.handbooks);const content=node('div'),draw=key=>{content.replaceChildren();const hb=d.handbooks[key];if(!hb)return;content.append(node('h3',hb.title),node('p',hb.blurb||''));for(const section of hb.sections||[]){content.append(node('h3',section.title));for(const sub of section.subs||[]){content.append(node('h4',sub.title));for(const p of sub.paras||[])content.append(node('p',plainText(typeof p==='string'?p:JSON.stringify(p)),{className:'reader'}));}}};c.append(select(books.map(([k,b])=>[k,b.title]),books[0]?.[0],draw),content);draw(books[0]?.[0]);
  }else if(state.sub==='Saturday Enrichment'){table(c,['Date','Theme','Lead','Support','Year groups','Venue'],d.saturday.map(r=>[r.date,r.theme,r.lead,r.support,Array.isArray(r.yearGroups)?r.yearGroups.join(', '):r.yearGroups,r.venue]));}
  else if(state.sub==='Sports'||state.sub==='Clubs & Teams'){
   table(c,['Sport','Team','Grade','Coach'],d.teams.map(t=>[d.sports.find(s=>s.id===t.sport_id)?.name,t.name,t.grade,t.coach]));
   if(state.sub==='Sports'){c.append(node('h3','Fixtures & results'));table(c,['Date','Team','Opponent','Location','Score / result'],d.fixtures.map(f=>[f.date,d.teams.find(t=>t.id===f.team_id)?.name,f.opponent,f.location,[f.score,f.result].filter(Boolean).join(' · ')]));}
   if(state.sub==='Clubs & Teams'){c.append(node('h3','Clubs'));ecaTable(c,d.ecas);if(d.clubTeams?.length){c.append(node('h3','Enrichment teams'));table(c,['Team','Day','Lead','Grades'],d.clubTeams.map(t=>[t.name,t.day,t.lead,t.grades]));}}
  }else{
   const result=node('div');let chosenDay=state.sub==="Today's ECA"?new Intl.DateTimeFormat('en-GB',{weekday:'long',timeZone:'Asia/Shanghai'}).format(new Date()):'all',school='all',grade='all';
   const grades=[...new Set(d.ecas.flatMap(r=>String(r.grades||'').split(/[\/,;]/).map(x=>x.trim()).filter(Boolean)))].sort();
   const draw=()=>{result.replaceChildren();ecaTable(result,d.ecas.filter(r=>(chosenDay==='all'||r.day===chosenDay)&&(school==='all'||r.school===school)&&(grade==='all'||String(r.grades||'').split(/[\/,;]/).map(x=>x.trim()).includes(grade))).sort((a,b)=>a.name.localeCompare(b.name)));if(state.sub==="Today's ECA"){result.append(node('h3','On call'));table(result,['Day','Section','Post','Assigned staff'],(d.oncall||[]).filter(r=>(chosenDay==='all'||r.day===chosenDay)&&(school==='all'||r.school===school)).map(r=>[r.day,r.school,r.post,people(r.who)]));}};
   c.append(toolbar([select([['all','All sections'],...['US','LS','KG'].map(x=>[x,x])],school,v=>{school=v;draw();}),select([['all','All days'],...['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'].map(x=>[x,x])],chosenDay,v=>{chosenDay=v;draw();}),select([['all','All grades'],...grades.map(x=>[x,x])],grade,v=>{grade=v;draw();})]),result);draw();
  }
 }
 function plainText(text){return new DOMParser().parseFromString(text,'text/html').body.textContent||'';}
 function readable(value){if(Array.isArray(value))return value.map(readable).join('\n');if(value&&typeof value==='object')return Object.entries(value).map(([k,v])=>k+': '+readable(v)).join('\n');return plainText(String(value??''));}
 function ecaTable(c,rows){table(c,['ECA','Day','Section / grades','Lead','Support','Room'],rows.map(r=>[r.name,r.day,[r.school,r.grades].filter(Boolean).join(' · '),people(r.lead),[people(r.support),people(r.support2)].filter(Boolean).join(', '),r.room]));}
 function docActive(module){return(state.docs[module]||[]).filter(r=>!r.deleted_at);}
 function docSave(module,row,data){return write('ms_documents_apply',{p_module:module,p_request_id:crypto.randomUUID(),p_ops:[{id:row?.id||'row:'+crypto.randomUUID(),action:'save',expected_revision:row?.revision||0,data}]});}
 function assemblies(){const c=card('Assembly schedule'),rows=docActive('assemblies'),slots=rows.find(r=>r.id==='settings:slots')?.data.slots||[];let school='US';const result=node('div');
  c.append(toolbar([select(['US','LS','KG','WS'].map(s=>[s,s]),school,s=>{school=s;draw();}),...(canEdit('assemblies')?[button('Add assembly date',()=>form(null),true)]:[])]),result);
  function draw(){result.replaceChildren();table(result,['Date',...slots.map(s=>s.time+' · '+s.label),'Whole assembly',...(canEdit('assemblies')?['Edit']:[])],rows.filter(r=>r.data.section===school).sort((a,b)=>a.data.date.localeCompare(b.data.date)).map(r=>[r.data.date,...slots.map(s=>r.data[s.id]||''),r.data.take||'',...(canEdit('assemblies')?[button('Edit',()=>form(r))]:[])]));}
  function form(row){modal(row?'Edit assembly':'Add assembly',body=>{const date=input(body,'Date',row?.data.date||today(),'date'),fields=slots.map(s=>[s.id,input(body,s.label,row?.data[s.id]||'')]),take=input(body,'Whole assembly takeover',row?.data.take||'');body.append(button('Save assembly',()=>{const data={...row?.data,section:school,date:date.value,take:take.value,...Object.fromEntries(fields.map(([id,i])=>[id,i.value]))};const id=school+':'+date.value;if(row&&row.id!==id)return say('Keep this record’s date unchanged. Add a new date separately.',true);docSave('assemblies',row||{id,revision:0},data);},true));});}draw();
 }
 function newsletter(){const c=card('Newsletter schedule'),rows=docActive('newsletter'),cats=rows.filter(r=>r.data.kind==='category').sort((a,b)=>a.data.order-b.data.order),dates=rows.filter(r=>r.data.kind==='row').sort((a,b)=>a.data.order-b.data.order);
  if(canEdit('newsletter'))c.append(toolbar([button('Add date',()=>form(null),true),button('Add category',()=>modal('Add newsletter category',body=>{const name=input(body,'Category name');body.append(button('Add',()=>docSave('newsletter',{id:'category:'+crypto.randomUUID(),revision:0},{kind:'category',label:name.value.trim(),order:cats.length}),true));}))]));
  table(c,['Date',...cats.map(r=>r.data.label),...(canEdit('newsletter')?['Edit']:[])],dates.map(r=>[r.data.date_label,...cats.map(cat=>r.data.cells?.[cat.id.replace(/^category:/,'')]||''),...(canEdit('newsletter')?[button('Edit',()=>form(r))]:[])]));
  function form(row){modal(row?'Edit newsletter date':'Add newsletter date',body=>{const date=input(body,'Date label (existing text dates are preserved)',row?.data.date_label||''),fields=cats.map(cat=>[cat.id.replace(/^category:/,''),input(body,cat.data.label,row?.data.cells?.[cat.id.replace(/^category:/,'')]||'','textarea')]);body.append(button('Save newsletter',()=>docSave('newsletter',row,{...row?.data,kind:'row',date_label:date.value,cells:Object.fromEntries(fields.map(([id,i])=>[id,i.value])),order:row?.data.order??dates.length}),true));});}
 }
 async function write(name,args){return writeMany([{name,args}]);}
 async function writeMany(jobs){if(state.busy||state.pending)return recovery();state.pending={jobs,subject:state.subject};try{localStorage.setItem('mastersheet553.draft.'+state.subject,JSON.stringify(state.pending));}catch{drawAccess();say('Draft is in memory because browser recovery storage is full. Export it before closing this page.',true);recovery();return false;}drawAccess();return sendPending();}
 async function sendPending(){if(!state.pending||state.busy)return;state.busy=true;say('Saving to shared database…');try{if(!state.pending.jobs)state.pending.jobs=[{name:state.pending.name,args:state.pending.args}];while(state.pending.jobs.length){const job=state.pending.jobs[0];await rpc(job.name,job.args);state.pending.jobs.shift();localStorage.setItem('mastersheet553.draft.'+state.subject,JSON.stringify(state.pending));}localStorage.removeItem('mastersheet553.draft.'+state.subject);state.pending=null;state.shared=null;$('dialog').close();drawAccess();state.busy=false;await load();say('Saved to shared database.');}catch(e){say(e.message+' Remaining requests are retained. Earlier completed batches stay saved.',true);state.busy=false;drawAccess();}}
 function recovery(){if(!state.pending)return;modal('Retained save draft',body=>{body.append(node('p','The original request is retained. Retry uses the same request ID. A conflict cannot overwrite another person’s newer revision.'),button('Export draft',()=>download('mastersheet-retained-draft-private',state.pending)),button('Retry exact request',sendPending),button('Compare with latest records',comparePending),button('Discard draft and reload',()=>modal('Discard this retained draft?',m=>{m.append(node('p','Export the draft first if you need it. The server may already have accepted an earlier request.'),button('Export draft',()=>download('mastersheet-retained-draft-private',state.pending)),button('Discard and reload',()=>{localStorage.removeItem('mastersheet553.draft.'+state.subject);state.pending=null;$('dialog').close();drawAccess();void load();}));})));});}
 async function comparePending(){
  const original=state.pending;if(!original||state.busy)return;
  try{const jobs=original.jobs||[{name:original.name,args:original.args}],nextJobs=[],comparisons=[];
   for(const job of jobs){let rows=[],kind;
    if(job.name==='ms_core_apply'){kind='core';for(let offset=0;;offset+=500){const page=await rpc('ms_core_list',{p_table:job.args.p_table,p_offset:offset,p_include_deleted:true});rows.push(...page);if(page.length<500)break;}}
    else if(job.name==='ms_documents_apply'){kind='document';rows=await allDocuments(job.args.p_module);}
    else{kind='calendar';rows=await rpc('calendar_shared_list');}
    if(job.name==='ms_calendar_govern'){
     const row=rows.find(r=>r.id===job.args.p_id);if(!row||row.deleted_at)throw Error('Event was deleted. Restore it separately.');comparisons.push({current:row,retained:job.args});nextJobs.push({name:job.name,args:{...job.args,p_expected_revision:row.revision,p_request_id:crypto.randomUUID()}});
    }else{
     comparisons.push(...job.args.p_ops.map(op=>({current:rows.find(r=>r.id===op.id)||{missing:true},retained:op})));
     const ops=rebaseOperations(job.args.p_ops,rows,kind);if(ops.length)nextJobs.push({name:job.name,args:{...job.args,p_ops:ops,p_request_id:crypto.randomUUID()}});
    }
   }
   modal('Compare current records and retained draft',body=>{body.append(node('p','Review the current server records against your proposed changes. Submitting creates new revision-checked requests. It may replace the fields shown; another change made after this comparison will still cause a conflict.'),node('pre',JSON.stringify(comparisons,null,2)),button('Export comparison',()=>download('mastersheet-draft-comparison-private',comparisons)),button(nextJobs.length?'Submit reviewed changes against these revisions':'No changes remain — clear retained draft',async()=>{if(state.pending!==original)return say('Draft changed; open a fresh comparison.',true);try{const replacement={jobs:nextJobs,subject:state.subject};localStorage.setItem('mastersheet553.draft.'+state.subject,JSON.stringify(replacement));state.pending=replacement;await sendPending();}catch(e){say(e.message,true);}},true));});
  }catch(e){say(e.message+' Original draft retained.',true);}
 }
 async function manageKeys(){try{
  const keys=await rpc('ms_access_keys');if(!state.people.length)state.people=await rpc('ms_people');
  modal('Access keys — owner only',body=>{
   body.append(node('p','Keys are bearer credentials, not verified Microsoft identities. Share each key privately with its intended staff member. Administrators can look up other staff; editors can write only selected modules.'));
   const label=input(body,'Label (staff name or purpose)'),role=select([['reader','Staff reader — own personal views'],['administrator','Administrator — staff lookup; editing only if granted below'],['editor','Editor — selected modules']],'reader',()=>{for(const[,check]of grants){check.disabled=role.value==='reader';if(check.disabled)check.checked=false;}}),person=select([['','No personal staff link'],...state.people.map(p=>[p.id,p.name+' ('+p.code+')'])],'',()=>{});
   body.append(node('label','Role'),role,node('label','Linked staff record'),person);const grants=[];for(const module of ['calendar','assemblies','newsletter']){const check=node('input',null,{type:'checkbox',disabled:true});body.append(node('label','May edit '+module),check);grants.push([module,check]);}
   body.append(button('Create key',async()=>{try{const created=await rpc('ms_issue_key',{p_label:label.value,p_role:role.value,p_modules:grants.filter(([,c])=>c.checked).map(([m])=>m),p_staff_id:person.value||null});modal('Save this key privately',m=>{m.append(node('p','This is shown only now. Do not put it into GitHub or chat.'),node('textarea',null,{value:created.key,readOnly:true}),button('Done',()=>{$('dialog').close();}));});}catch(e){say(e.message,true);}},true));
   table(body,['Label','Role','Linked staff','State','Action'],keys.map(k=>[k.label,k.role,state.people.find(p=>p.id===k.staff_id)?.name||'',k.active?'Active':'Revoked',k.role!=='owner'&&k.active?button('Revoke',()=>modal('Revoke this key?',m=>{m.append(node('p',k.label),button('Revoke',async()=>{await rpc('ms_revoke_key',{p_id:k.id});await manageKeys();}));})):'' ]));
   body.append(node('p','The older v5.52 Calendar key still works for Calendar only until you retire it.'),button('Retire the old Calendar key',()=>modal('Retire the old Calendar connection key?',m=>{m.append(node('p','Anyone using the old shared Calendar key will lose editing access. Public reading and the new owner/staff keys will continue to work.'),button('Retire old key',async()=>{try{await rpc('ms_retire_calendar_key');$('dialog').close();say('Old Calendar connection key retired.');}catch(e){say(e.message,true);}}));})));
  });
 }catch(e){say(e.message,true);}}
 async function openOwner(){if(!owner())return;if(ownerUnsaved())return say('Finish your owner save before reopening the workspace.',true);document.body.classList.add('original-mode');try{if(state.pending)return recovery();say('Opening owner workspace…');const base64=await rpc('ms_owner_bundle');if(!base64)throw Error('Private owner workspace not installed');const bytes=Uint8Array.from(atob(base64),c=>c.charCodeAt(0));const stream=new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));let html=await new Response(stream).text();html=html.replace("const VERSION = '5.53';","const VERSION = '5.56';").replace("['pastoral', 'Admin']","['pastoral', 'School Information']").replace("['duty', 'Pastoral']","['duty', 'Duties']").replace('autoSaveOn=true;','autoSaveOn=true;UI.mode="admin";render();').replace('</head>','<style>#ms550bar{display:none!important}</style></head>');html=html.replace("if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);","\nconst originalNewsletter556=newsView;\nnewsView=()=>{originalNewsletter556();const rows=window.MS556NewsletterArchive||DB.newsletter;if(!rows?.length)return;const host=$('#pastBody')||$('#view'),c=card('Original imported newsletter rota','Original dates retained as supplied; separate from the current 2026–27 schedule.');const table=el('table');table.appendChild(el('thead',{},[el('tr',{},['No.','Semester 1 date','SLT','GL / HOF','Semester 2 date','SLT','GL / HOF'].map(x=>el('th',{text:x})))]));table.appendChild(el('tbody',{},rows.map(r=>el('tr',{},[r.no,r.s1?.date,r.s1?.slt,r.s1?.gl,r.s2?.date,r.s2?.slt,r.s2?.gl].map(x=>el('td',{text:x||''}))))));c.body.appendChild(table);host.insertBefore(c,host.firstChild);};\n\nif (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);");view.replaceChildren();actions.replaceChildren();const frame=node('iframe',null,{title:'Owner workspace'});frame.srcdoc=html;view.append(frame);say('Owner workspace · full tools remain restricted to owner access');}catch(e){say(e.message,true);}}
 function installPrivate(){if(!owner())return;modal('Install private workspace',body=>{
  body.append(node('p','Choose PRIVATE-SETUP-v5.53.json from the private setup folder. The file goes to your existing Supabase project using your owner key. It installs the owner workspace and missing local-data records; existing records are preserved. Keep this file out of GitHub.'));
  const file=node('input',null,{type:'file',accept:'.json,application/json'}),feedback=node('p'),start=button('Install selected private setup',async()=>{
   if(state.busy||state.pending||ownerUnsaved())return say('Finish or resolve current saves before installing.',true);
   try{if(!file.files[0])throw Error('Choose the private setup JSON file first.');const pkg=JSON.parse(await file.files[0].text());if(pkg.version!=='5.53'||!pkg.payload_base64||!/^[a-f0-9]{64}$/.test(pkg.sha256))throw Error('Choose PRIVATE-SETUP-v5.53.json.');
    let upload;try{upload=JSON.parse(localStorage.getItem('mastersheet553.install'));}catch{}
    if(upload?.sha256!==pkg.sha256||upload?.subject!==state.subject)upload={id:crypto.randomUUID(),sha256:pkg.sha256,subject:state.subject};
    try{localStorage.setItem('mastersheet553.install',JSON.stringify(upload));}catch{}
    const total=Math.ceil(pkg.payload_base64.length/250000);if(total>200)throw Error('This setup file is larger than the supported installation limit.');
    state.busy=true;start.disabled=true;
    for(let part=0;part<total;part++){feedback.textContent='Uploading private setup: '+(part+1)+' of '+total+' parts…';await rpc('ms_install_part',{p_upload_id:upload.id,p_part:part,p_total:total,p_sha256:pkg.sha256,p_data:pkg.payload_base64.slice(part*250000,(part+1)*250000)});}
    feedback.textContent='Checking the file and finishing installation…';await rpc('ms_install_finish',{p_upload_id:upload.id});try{localStorage.removeItem('mastersheet553.install');}catch{}
    state.shared=null;feedback.textContent='Private setup installed. Existing records were preserved. Review the views, then run 04_ACTIVATE.sql when ready to switch over.';say('Private setup installed.');
   }catch(e){feedback.textContent=e.message+' Select the same file and retry; accepted parts can be safely retried.';}finally{state.busy=false;start.disabled=false;}
  },true);body.append(file,feedback,start);
 });}
 async function databaseManager(module){
  if(!owner())return;try{
   const isCalendar=module==='calendar',isDocument=['assemblies','newsletter','workspace'].includes(module);let rows=[];
   if(isCalendar)rows=await rpc('calendar_shared_list');
   else if(isDocument)rows=await allDocuments(module);
   else for(let offset=0;;offset+=500){const page=await rpc('ms_core_list',{p_table:module,p_offset:offset,p_include_deleted:true});rows.push(...page);if(page.length<500)break;}
   const revision=r=>isCalendar||isDocument?r.revision:r.ms_revision,deleted=r=>isCalendar||isDocument?r.deleted_at:r.ms_deleted_at;
   const clean=r=>isCalendar||isDocument?r.data:Object.fromEntries(Object.entries(r).filter(([k])=>!k.startsWith('ms_')));
   const endpoint=isCalendar?'calendar_shared_apply':isDocument?'ms_documents_apply':'ms_core_apply';
   const args=ops=>({...(!isCalendar?{[isDocument?'p_module':'p_table']:module}:{}),p_request_id:crypto.randomUUID(),p_ops:ops});
   const operation=(r,action,data)=>({id:r.id,action,expected_revision:revision(r)||0,...(data?{data}:{}),...(isCalendar?{scope:r.edit_scope||(data?.labels||[]).find(x=>['US','LS','KG','WS'].includes(x))||'US'}:{})});
   modal('Database manager',body=>{
    const modules=['calendar','staff','rooms','ecas','duties','timetables','sports','teams','players','fixtures','prizes','assemblies','newsletter','workspace'];
    body.append(select(modules.map(m=>[m,m]),module,m=>void databaseManager(m)),node('p','The same checked save, revision, history, export and restore controls apply here. “Workspace” contains the remaining shared collections, such as cover, room bookings, Saturday Enrichment, handbooks and settings.'));
    body.append(toolbar([button('Export records',()=>download('mastersheet-'+module+'-private',{module,exported:new Date().toISOString(),records:rows})),button('Back up all shared data',async()=>{try{say('Preparing private backup…');download('mastersheet-all-shared-data-private',await rpc('ms_backup'));say('Private data backup downloaded. Access keys are excluded.');}catch(e){say(e.message,true);}}),button('Preview import',()=>importForm()),...(module==='workspace'?[button('Import browser-only backup',browserImport)]:[]),button('Add record',()=>edit(null)),button('Change history',async()=>{try{const history=await rpc('ms_history',{p_module:module});modal('Recent history · '+module,m=>{m.append(node('p','Up to 500 recent changes. Keys do not establish a verified individual identity.'),node('pre',JSON.stringify(history,null,2)),button('Export history',()=>download(module+'-history-private',history)));});}catch(e){say(e.message,true);}})]));
    table(body,['Record','Revision','State','Actions'],rows.map(r=>[r.data?.text||r.full_name||r.name||r.data?.date_label||r.id,revision(r),deleted(r)?'Deleted':'Active',toolbar(deleted(r)?[button('Restore',()=>write(endpoint,args([operation(r,'restore')])) )]:[button('Edit',()=>edit(r)),...(isCalendar?[button(r.data.approved?'Revoke approval':'Approve',()=>govern(r,'approve',!r.data.approved)),button(r.published?'Unpublish':'Publish',()=>govern(r,'publish',!r.published))]:[]),button('Delete',()=>modal('Delete this record?',m=>{m.append(node('p',r.id+' — the record will remain recoverable.'),button('Delete',()=>write(endpoint,args([operation(r,'delete')]))));}))])]));
   });
   function govern(row,action,value){modal((value?'Confirm ':'Revoke ')+action,body=>{body.append(node('p',row.data.text),node('p','Public views already show event summaries, including drafts. Approval and publication are owner actions; editing content withdraws them.'),button('Confirm',()=>write('ms_calendar_govern',{p_request_id:crypto.randomUUID(),p_id:row.id,p_action:action,p_value:value,p_expected_revision:row.revision}),true));});}
   function edit(row){modal(row?'Edit record':'Add record',body=>{const id=input(body,'Stable record ID',row?.id||'record-'+crypto.randomUUID());id.readOnly=!!row;const data=input(body,'Record data (JSON)',JSON.stringify(row?clean(row):isCalendar?{id:id.value,text:'',date:today(),end:today(),labels:['WS'],approved:false}:isDocument?{}:{id:id.value},null,2),'textarea');data.style.minHeight='260px';const feedback=node('p');body.append(feedback,button('Save checked change',async()=>{try{const value=JSON.parse(data.value);if(!value||Array.isArray(value)||typeof value!=='object')throw Error('Use a JSON object.');if(!isDocument)value.id=id.value;await write(endpoint,args([operation(row||{id:id.value,revision:0,ms_revision:0},'save',value)]));}catch(e){feedback.textContent=e.message;}},true));});}
   function importForm(){modal('Preview import · '+module,body=>{body.append(node('p','Accepts this manager’s export format. Missing records are never deleted. Existing records use the revisions loaded when this manager was opened. Large imports run in batches; a failure retains the remaining requests.'));const file=node('input',null,{type:'file',accept:'.json,application/json'}),report=node('p'),submit=button('Apply previewed changes',()=>{},true);submit.disabled=true;body.append(file,report,submit);let ops=[];file.onchange=async()=>{submit.disabled=true;try{const raw=JSON.parse(await file.files[0].text()),incoming=Array.isArray(raw)?raw:raw.records;if(!Array.isArray(incoming))throw Error('Choose a manager export containing records.');if(raw.module&&raw.module!==module)throw Error('This export belongs to another table.');const ids=new Set();ops=[];for(const r of incoming){if(!r?.id||ids.has(r.id))throw Error('Missing or duplicate record ID.');ids.add(r.id);if(deleted(r))continue;const old=rows.find(x=>x.id===r.id);if(old&&deleted(old))throw Error('Restore deleted record '+r.id+' separately.');const data=clean(r);if(!old||canonical(clean(old))!==canonical(data))ops.push(operation(old||{id:r.id},'save',data));}report.textContent=ops.length+' additions/updates; no missing-record deletions. Back up first.';submit.disabled=!ops.length;submit.onclick=()=>writeMany(Array.from({length:Math.ceil(ops.length/500)},(_,i)=>({name:endpoint,args:args(ops.slice(i*500,i*500+500))})));}catch(e){report.textContent=e.message;}};});}
  }catch(e){say(e.message,true);}
 }
 async function browserImport(){
  if(!owner())return;modal('Import previously browser-only data',body=>{
   body.append(node('p','Choose your private Mastersheet JSON backup. This imports assemblies, newsletter and the remaining local collections. It excludes the shared staff, room, duty, ECA, timetable, Sports and Calendar tables. Missing records are never deleted.'));
   const file=node('input',null,{type:'file',accept:'.json,application/json'}),report=node('pre'),submit=button('Apply reviewed local-data changes',()=>{},true);submit.disabled=true;body.append(file,report,submit);
   file.onchange=async()=>{submit.disabled=true;try{
    const raw=JSON.parse(await file.files[0].text()),db=raw.db||raw;
    if(!db||!Array.isArray(db.staff)||!Array.isArray(db.duty))throw Error('Choose a Mastersheet data backup.');
    const modules=documentsFromDB(db),jobs=[],lines=[];
    for(const module of ['workspace','assemblies','newsletter']){
     const existing=await allDocuments(module),map=new Map(existing.map(r=>[r.id,r])),ops=[];
     for(const incoming of modules[module]){
      if(module==='newsletter'&&incoming.data.kind==='row'){
       const matches=existing.filter(r=>!r.deleted_at&&r.data.kind==='row'&&r.data.date_label===incoming.data.date_label);
       if(!map.has(incoming.id)){if(matches.length>1)throw Error('More than one newsletter row has date '+incoming.data.date_label+'. Export and reconcile those rows separately.');if(matches.length===1)incoming.id=matches[0].id;}
      }
      const old=map.get(incoming.id);if(old?.deleted_at)continue;
      if(!old||canonical(old.data)!==canonical(incoming.data))ops.push({id:incoming.id,action:'save',expected_revision:old?.revision||0,data:incoming.data});
     }
     lines.push(module+': '+ops.length+' additions/updates; 0 missing-record deletions');
     for(let i=0;i<ops.length;i+=100)jobs.push({name:'ms_documents_apply',args:{p_module:module,p_request_id:crypto.randomUUID(),p_ops:ops.slice(i,i+100)}});
    }
    report.textContent=lines.join('\n')+'\nShared core tables and Calendar: 0 changes.';submit.disabled=!jobs.length;submit.onclick=()=>writeMany(jobs);
   }catch(e){report.textContent=e.message;}};
  });
 }
 
 async function openOriginal(){document.body.classList.add('original-mode');if(view.querySelector('iframe[data-public-original]'))return;const frame=node('iframe',null,{title:'Mastersheet original interface'});frame.dataset.publicOriginal='true';frame.src='./original-v556.html';view.replaceChildren(frame);actions.replaceChildren();say('');}
 window.MS554Unlock=unlock;
 window.MS554Tools=()=>{window.MS554EditingTools=true;document.body.classList.remove('original-mode');void load();};
 const originalAccess=drawAccess;drawAccess=()=>{originalAccess();if(state.who&&state.who.role!=='reader'&&!owner())$('access').append(button('Editing tools',window.MS554Tools));for(const b of $('access').querySelectorAll('button'))if(b.textContent==='Unlock personal / editing access')b.textContent='Unlock';};

$('unlock').onclick=unlock;$('refresh').onclick=()=>{const frame=view.querySelector('iframe[data-public-original]');if(frame){frame.contentWindow.location.reload();return;}state.shared=null;void load();};$('print').onclick=()=>window.print();
 window.addEventListener('beforeunload',e=>{if(state.pending||state.busy){e.preventDefault();e.returnValue='';}});
 (async()=>{try{state.cfg=await window.MS550Config;state.cfg.url=state.cfg.url.replace(/\/$/,'');let key='';try{key=sessionStorage.getItem('mastersheet553.access')||'';}catch{}if(key){try{await connect(key);return;}catch{try{sessionStorage.removeItem('mastersheet553.access');}catch{}state.key='';state.who=null;}}drawAccess();navigation();await load();}catch(e){say(e.message,true);}})();
})();
