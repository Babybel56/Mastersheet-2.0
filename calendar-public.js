/* No administrator source, seed, notes or authentication tokens on this page. */
'use strict';
(() => {
  const $ = id => document.getElementById(id);
  let events = [], lastUpdated = '';
  const now = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Shanghai', year: 'numeric', month: '2-digit' }).formatToParts(new Date());
  $('month').value = `${now.find(p => p.type === 'year').value}-${now.find(p => p.type === 'month').value}`;
  const node = (tag, text, cls) => { const n = document.createElement(tag); n.textContent = text || ''; if (cls) n.className = cls; return n; };
  function paint() {
    const month = $('month').value, section = $('section').value, q = $('search').value.trim().toLowerCase();
    const first = month ? `${month}-01` : '', last = month ? `${month}-31` : '';
    const visible = events.filter(e => (!first || (e.date <= last && (e.end || e.date) >= first))
      && (section === 'all' || (e.labels || []).includes(section) || (section !== 'WS' && (e.labels || []).includes('WS')))
      && (!q || [e.text, e.type, e.location].join(' ').toLowerCase().includes(q)));
    $('events').replaceChildren();
    if (!visible.length) $('events').append(node('div', 'No Calendar events match this view.', 'empty'));
    visible.sort((a,b)=>String(a.date<first?first:a.date).localeCompare(String(b.date<first?first:b.date))||Number(b.sig||0)-Number(a.sig||0)||Number(b.date<first)-Number(a.date<first)||String(a.text).localeCompare(String(b.text)));for (const e of visible) {
      const row = node('article', '', 'event'+(!e.approved?' pending':'')), date = node('div', e.date + (e.end && e.end !== e.date ? ` – ${e.end}` : ''));
      const content = node('div'); content.append(node('h3', e.text));if(!e.approved)content.append(node('small','Awaiting approval'));
      content.append(node('p', ({ whole: 'Whole day', first: 'First half day', second: 'Second half day' })[e.timeMode] || e.timeText || 'Time to be confirmed'));
      if (e.location) content.append(node('p', e.location));
      const tags = node('div');if(e.ghp)tags.append(node('span','GHP','tag'));row.dataset.eventId=e.id;row.tabIndex=0;const open=()=>{const url=new URL('index.html',location.href);url.searchParams.set('event',e.id);location.assign(url.href);};row.onclick=open;row.onkeydown=ev=>{if(ev.key==='Enter')open();};Promise.resolve(window.MS550Config).then(async cfg=>{const r=await fetch(cfg.url+'/rest/v1/rpc/ms_event_attachments',{method:'POST',headers:{apikey:cfg.key,'Content-Type':'application/json'},body:JSON.stringify({p_event:e.id})});if(r.ok&&(await r.json()).some(f=>f.ready))tags.append(node('span','📎','tag'));}).catch(()=>{}); [e.type, ...(e.labels || [])].filter(Boolean).forEach(t => tags.append(node('span', t, 'tag')));
      content.append(tags); row.append(date, content); $('events').append(row);
    }
    if (lastUpdated) $('status').textContent = `${visible.length} Calendar event(s) · Updated ${lastUpdated} · School time: Asia/Shanghai`;
  }
  async function refresh() {
    $('refresh').disabled = true;
    try {
      const cfg = await window.MS550Config;
      if (!cfg.configured) throw new Error('Calendar preview: the shared backend has not been configured. No school events are published here.');
      const r = await fetch(`${cfg.url}/rest/v1/rpc/calendar_shared_list`, { method: 'POST', headers: { apikey: cfg.key, 'Content-Type': 'application/json' }, body: '{}', cache: 'no-store' });
      if (!r.ok) throw new Error('The published Calendar is temporarily unavailable. Please refresh.');
      const data = await r.json(); if (!Array.isArray(data)) throw new Error('Unexpected Calendar response.');
      events = data.filter(r=>!r.deleted_at).map(r=>r.data); lastUpdated = new Date().toLocaleTimeString(); $('status').classList.remove('error'); paint();
    } catch (e) {
      // Never retain withdrawn events on a failed refresh and imply freshness.
      events = []; lastUpdated = ''; $('events').replaceChildren(); $('status').textContent = e.message; $('status').classList.add('error');
    } finally { $('refresh').disabled = false; }
  }
  ['month','section','search'].forEach(id => $(id).addEventListener('input', paint));
  $('refresh').onclick = refresh; $('print').onclick = () => window.print();
  window.addEventListener('focus', refresh); setInterval(refresh, 60000); refresh();
})();
