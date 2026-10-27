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
    if (!visible.length) $('events').append(node('div', 'No published events match this view.', 'empty'));
    for (const e of visible) {
      const row = node('article', '', 'event'), date = node('div', e.date + (e.end && e.end !== e.date ? ` – ${e.end}` : ''));
      const content = node('div'); content.append(node('h3', e.text));
      content.append(node('p', ({ whole: 'Whole day', first: 'First half day', second: 'Second half day' })[e.timeMode] || e.timeText || 'Time to be confirmed'));
      if (e.location) content.append(node('p', e.location));
      const tags = node('div'); [e.type, ...(e.labels || [])].filter(Boolean).forEach(t => tags.append(node('span', t, 'tag')));
      content.append(tags); row.append(date, content); $('events').append(row);
    }
    if (lastUpdated) $('status').textContent = `${visible.length} published event(s) · Updated ${lastUpdated} · School time: Asia/Shanghai`;
  }
  async function refresh() {
    $('refresh').disabled = true;
    try {
      const cfg = await window.MS550Config;
      if (!cfg.configured) throw new Error('Calendar preview: the shared backend has not been configured. No school events are published here.');
      const r = await fetch(`${cfg.url}/rest/v1/rpc/calendar_public`, { method: 'POST', headers: { apikey: cfg.key, 'Content-Type': 'application/json' }, body: '{}', cache: 'no-store' });
      if (!r.ok) throw new Error('The published Calendar is temporarily unavailable. Please refresh.');
      const data = await r.json(); if (!Array.isArray(data)) throw new Error('Unexpected Calendar response.');
      events = data; lastUpdated = new Date().toLocaleTimeString(); $('status').classList.remove('error'); paint();
    } catch (e) {
      // Never retain withdrawn events on a failed refresh and imply freshness.
      events = []; lastUpdated = ''; $('events').replaceChildren(); $('status').textContent = e.message; $('status').classList.add('error');
    } finally { $('refresh').disabled = false; }
  }
  ['month','section','search'].forEach(id => $(id).addEventListener('input', paint));
  $('refresh').onclick = refresh; $('print').onclick = () => window.print();
  window.addEventListener('focus', refresh); setInterval(refresh, 60000); refresh();
})();
