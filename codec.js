// Stable per-date assembly records; stable newsletter category/row IDs.
function canonical(value) {
  return JSON.stringify(value,(_key,item)=>item&&typeof item==='object'&&!Array.isArray(item)?Object.fromEntries(Object.keys(item).sort().map(k=>[k,item[k]])):item);
}
function rebaseOperations(ops,rows,kind){
  const current=new Map(rows.map(r=>[r.id,r])),result=[];
  for(const op of ops){const row=current.get(op.id),deleted=kind==='core'?row?.ms_deleted_at:row?.deleted_at;
    if(op.action==='delete'&&(!row||deleted))continue;
    if(op.action==='restore'&&row&&!deleted)continue;
    if(op.action==='save'&&deleted)throw Error('Restore deleted record '+op.id+' separately before reapplying a save.');
    if(op.action==='restore'&&!row)throw Error('The deleted record is missing.');
    const next={...op,expected_revision:row?(kind==='core'?row.ms_revision:row.revision):0};
    if(op.action==='save'){
      if(kind==='calendar'&&row){next.data={...op.data,approved:!!row.data.approved};next.scope=row.edit_scope;}
      const oldData=kind==='core'?row:row?.data;
      if(oldData&&Object.entries(next.data).every(([k,v])=>canonical(v)===canonical(oldData[k])))continue;
    }result.push(next);
  }return result;
}
function documentsFromDB(db, handbooks = db.hbooks) {
  const result = { assemblies: [], newsletter: [], workspace: [] };
  const record = (id,data) => ({id,data});
  for (const section of ['US','LS','KG','WS']) for (const [date,row] of Object.entries(db['assembly'+section] || {})) {
    result.assemblies.push(record(section+':'+date,{...row,section,date}));
  }
  if(db.asSlots)result.assemblies.push(record('settings:slots',{slots:db.asSlots}));
  let news=db.news27v2;
  if(!news){const seen=new Set(),rows=[];for(const sem of ['s1','s2'])for(const row of db.news27?.[sem]||[]){const id=row.date||JSON.stringify(row);if(seen.has(id))continue;seen.add(id);rows.push({date:row.date||'',cells:{slt:row.slt||'',gl:row.gl||''}});}news={cats:[{id:'slt',label:'SLT'},{id:'gl',label:'GL / HOF'}],rows};}
  news.cats.forEach((c,i)=>result.newsletter.push(record('category:'+c.id,{kind:'category',label:c.label,order:i,extra:Object.fromEntries(Object.entries(c).filter(([k])=>!['id','label'].includes(k)))})));
  news.rows.forEach((r,i)=>{if(!r._sharedId)r._sharedId='row:'+((globalThis.crypto?.randomUUID?.())||('seed-'+i));result.newsletter.push(record(r._sharedId,{kind:'row',date_label:r.date||'',cells:r.cells||{},order:i,extra:Object.fromEntries(Object.entries(r).filter(([k])=>!['date','cells','_sharedId'].includes(k)))}));});
  const excluded=new Set(['staff','rooms','duty','teaching','sports','icEvents','assemblyUS','assemblyLS','assemblyKG','assemblyWS','asSlots','news27','news27v2','meta','security','hbooks']);
  for(const [key,value]of Object.entries(db))if(!excluded.has(key)){
    const v=key==='enrich'?Object.fromEntries(Object.entries(value||{}).filter(([k])=>k!=='ecas')):value;
    result.workspace.push(record(key,{value:v}));
  }
  if(handbooks)result.workspace.push(record('hbooks',{value:handbooks}));
  return result;
}
function documentsToDB(db, modules) {
  const active=rows=>(rows||[]).filter(r=>!r.deleted_at);
  for(const r of modules.workspace||[])if(r.deleted_at)db[r.id]=Array.isArray(db[r.id])?[]:{};
  for(const r of active(modules.workspace)){
    if(r.id==='enrich')db.enrich={...(db.enrich||{}),...r.data.value};else db[r.id]=r.data.value;
  }
  for(const section of ['US','LS','KG','WS'])db['assembly'+section]={};
  for(const r of active(modules.assemblies)){
    if(r.id==='settings:slots')db.asSlots=r.data.slots;
    else {const{section,date,...row}=r.data;db['assembly'+section][date]=row;}
  }
  db.news27v2={cats:[],rows:[]};
  for(const r of active(modules.newsletter).sort((a,b)=>(a.data.order||0)-(b.data.order||0))){
    if(r.data.kind==='category')db.news27v2.cats.push({...r.data.extra,id:r.id.replace(/^category:/,''),label:r.data.label});
    else db.news27v2.rows.push({...r.data.extra,_sharedId:r.id,date:r.data.date_label,cells:r.data.cells||{}});
  }
}
