import { createClient } from '@supabase/supabase-js';
import * as XLSX from 'xlsx';

const SUPABASE_URL = 'https://ddeoswpycvlfhztjlkdj.supabase.co';
const SUPABASE_KEY = 'sb_publishable_zEUf50vq4WcErjmNuCev3g_Lko70Obs';
const sb = createClient(SUPABASE_URL, SUPABASE_KEY, { auth: { persistSession: true, autoRefreshToken: true } });

/* ---------------- stato ---------------- */
const S = {
  session: null, member: null, loading: false,
  wines: [], locs: [], giac: [], movs: [], segn: [], membri: [],
  view: { name: 'cantina' }, history: [], sheet: null, q: '', filters: { tip: '', dw: '', loc: '', zero: false },
};
const $ = (s) => document.querySelector(s);
const esc = (v) => String(v ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
const fmtEur = (n) => n == null ? '' : n.toLocaleString('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });
const fmtDate = (d) => d ? new Date(d).toLocaleDateString('it-IT', { day: '2-digit', month: 'short', year: 'numeric' }) : '';
const YEAR = new Date().getFullYear();

/* ---------------- indici ---------------- */
let W = new Map(), L = new Map();
function reindex() {
  W = new Map(S.wines.map((w) => [w.id, w]));
  L = new Map(S.locs.map((l) => [l.id, l]));
}
const physLocs = () => S.locs.filter((l) => l.tipo === 'fisica' && l.attiva).sort((a, b) => (a.ordine ?? 50) - (b.ordine ?? 50) || a.nome.localeCompare(b.nome));
const unknownLoc = () => S.locs.find((l) => l.tipo === 'sconosciuta');
const stockOfWine = (wid) => S.giac.filter((g) => g.wine_id === wid && g.qty > 0);
const stockAt = (wid, lid) => S.giac.find((g) => g.wine_id === wid && g.location_id === lid)?.qty || 0;
const stockOfLoc = (lid) => S.giac.filter((g) => g.location_id === lid && g.qty > 0);
const totalWine = (wid) => stockOfWine(wid).reduce((a, g) => a + g.qty, 0);
const openMovs = () => S.movs.filter((m) => m.stato === 'in_corso' || m.stato === 'pianificato');
const openMovsOfWine = (wid) => openMovs().filter((m) => m.wine_id === wid);
const segnOpen = () => S.segn.filter((s) => s.aperta);
function avgPrice(wid) {
  const ms = S.movs.filter((m) => m.wine_id === wid && (m.tipo === 'acquisto' || m.tipo === 'carico_iniziale') && m.stato !== 'annullato' && m.prezzo_unitario != null);
  const q = ms.reduce((a, m) => a + m.qty, 0);
  return q ? ms.reduce((a, m) => a + m.qty * Number(m.prezzo_unitario), 0) / q : null;
}
function dwState(w) {
  const i = w.dw_inizio, f = w.dw_fine;
  if (!i && !f) return null;
  if (i && YEAR < i) return 'Da attendere';
  if (f && YEAR > f) return 'Oltre finestra';
  if (f && YEAR >= f - 1) return 'Bere presto';
  return 'In finestra';
}
const DWCLS = { 'Da attendere': 'wait', 'In finestra': 'ok', 'Bere presto': 'soon', 'Oltre finestra': 'late' };
const wineTitle = (w) => `${w.produttore} · ${w.vino}`;
const wineSub = (w) => [w.annata, w.denominazione, w.formato !== '750ml' ? w.formato : null].filter(Boolean).join(' · ');
const tipDot = (w) => `<span class="dot ${w.tipologia === 'Rosso' ? 'r' : w.tipologia === 'Bianco' ? 'b' : 'p'}"></span>`;

/* ---------------- dati ---------------- */
async function fetchAll(table, build) {
  const out = []; const step = 1000;
  for (let from = 0; ; from += step) {
    let q = sb.from(table).select('*').range(from, from + step - 1);
    if (build) q = build(q);
    const { data, error } = await q;
    if (error) throw error;
    out.push(...data);
    if (data.length < step) break;
  }
  return out;
}
async function loadAll() {
  S.loading = true; renderTop();
  try {
    const [wines, locs, giac, movs, segn, membri] = await Promise.all([
      fetchAll('wines', (q) => q.order('id')),
      fetchAll('locations', (q) => q.order('id')),
      fetchAll('giacenze'),
      fetchAll('movimenti', (q) => q.order('data_movimento', { ascending: false }).order('id', { ascending: false })),
      fetchAll('segnalazioni', (q) => q.order('id')),
      fetchAll('membri', (q) => q.order('email')),
    ]);
    Object.assign(S, { wines, locs, giac, movs, segn, membri });
    reindex();
  } catch (e) { toast('Errore nel caricamento: ' + (e.message || e)); }
  S.loading = false; render();
}
function errMsg(e) {
  const m = e?.message || String(e);
  if (/Giacenza insufficiente/.test(m)) return m.replace(/^.*?(Giacenza insufficiente[^"]*).*$/, '$1');
  if (/duplicate key|wines_unico/.test(m)) return 'Esiste già un vino con stesso produttore, nome, annata e formato.';
  return m;
}
async function run(promise, okMsg) {
  const { error } = await promise;
  if (error) { toast(errMsg(error)); return false; }
  if (okMsg) toast(okMsg);
  return true;
}

/* ---------------- navigazione ---------------- */
function go(view, push = true) {
  if (push) S.history.push(S.view);
  S.view = view; S.sheet = null; window.scrollTo(0, 0); render();
}
function back() { S.view = S.history.pop() || { name: 'cantina' }; render(); }
const TABS = [
  ['cantina', 'Cantina', 'M4 21V9l8-6 8 6v12h-5v-7H9v7z'],
  ['cerca', 'Cerca', 'M10 4a6 6 0 104.47 10l5.27 5.27 1.41-1.41L15.9 12.6A6 6 0 0010 4zm0 2a4 4 0 110 8 4 4 0 010-8z'],
  ['arrivi', 'In arrivo', 'M3 6h11v9H3zM14 9h4l3 3v3h-7zM6.5 18a1.5 1.5 0 100-3 1.5 1.5 0 000 3zm11 0a1.5 1.5 0 100-3 1.5 1.5 0 000 3z'],
  ['verifica', 'Verifica', 'M12 2l9 4v6c0 5-3.8 9.3-9 10-5.2-.7-9-5-9-10V6zm-1 13l6-6-1.4-1.4L11 12.2 8.4 9.6 7 11z'],
  ['altro', 'Altro', 'M5 10a2 2 0 110 4 2 2 0 010-4zm7 0a2 2 0 110 4 2 2 0 010-4zm7 0a2 2 0 110 4 2 2 0 010-4z'],
];

/* ---------------- render ---------------- */
function render() {
  const app = $('#app');
  if (!S.session) { app.innerHTML = loginView(); return; }
  if (S.member === false) { app.innerHTML = notMemberView(); return; }
  const v = S.view;
  const body = {
    cantina: cantinaView, loc: locView, cerca: cercaView, arrivi: arriviView, verifica: verificaView,
    altro: altroView, wine: wineView, membri: membriView, luoghi: luoghiView,
  }[v.name]?.(v) ?? '';
  const nArr = openMovs().length, nVer = segnOpen().length;
  const tabRoot = { loc: 'cantina', wine: S.history.length ? null : 'cerca', membri: 'altro', luoghi: 'altro' }[v.name] ?? v.name;
  app.innerHTML = `
    <header class="top" id="top"></header>
    <main>${body}</main>
    <nav class="tabs">${TABS.map(([k, label, d]) => `<button class="tab ${tabRoot === k ? 'on' : ''}" data-tab="${k}">
      <svg viewBox="0 0 24 24"><path d="${d}"/></svg><span>${label}</span>
      ${k === 'arrivi' && nArr ? `<i class="badge">${nArr}</i>` : ''}${k === 'verifica' && nVer ? `<i class="badge">${nVer}</i>` : ''}</button>`).join('')}</nav>
    <div id="sheet"></div><div id="toast"></div>`;
  renderTop(); renderSheet();
}
function renderTop() {
  const t = $('#top'); if (!t) return;
  const v = S.view;
  const title = { cantina: 'Cantina', cerca: 'Cerca', arrivi: 'In arrivo', verifica: 'Da verificare', altro: 'Altro', membri: 'Persone', luoghi: 'Location' }[v.name]
    ?? (v.name === 'loc' ? L.get(v.id)?.nome : v.name === 'wine' ? 'Scheda vino' : '');
  const showBack = ['loc', 'wine', 'membri', 'luoghi'].includes(v.name) && S.history.length;
  t.innerHTML = `${showBack ? '<button class="iconbtn" data-act="back" aria-label="Indietro">‹</button>' : '<span class="brand">◆</span>'}
    <h1>${esc(title)}</h1>
    <button class="iconbtn ${S.loading ? 'spin' : ''}" data-act="reload" aria-label="Aggiorna">⟳</button>`;
}

function loginView() {
  return `<div class="login">
    <div class="logo">◆</div><h1>Cantina</h1>
    <p class="note">Accedi con la tua email e password. La prima volta scegli "Crea account": entri se la tua email è tra quelle autorizzate.</p>
    <form id="loginform" class="card form">
      <label>Email<input name="email" type="email" autocomplete="email" required></label>
      <label>Password<input name="password" type="password" autocomplete="current-password" minlength="6" required></label>
      <button class="primary" type="submit" name="mode" value="in">Accedi</button>
      <button class="secondary" type="submit" name="mode" value="up">Crea account</button>
    </form></div>`;
}
function notMemberView() {
  return `<div class="login"><div class="logo">◆</div><h1>Accesso non autorizzato</h1>
    <p class="note">L'account ${esc(S.session?.user?.email)} esiste, ma non è nell'elenco delle persone autorizzate. Chiedi a Matteo di aggiungerti, poi riapri l'app.</p>
    <button class="secondary" data-act="logout">Esci</button></div>`;
}

function locCard(l) {
  const st = stockOfLoc(l.id);
  const n = st.reduce((a, g) => a + g.qty, 0);
  const inArr = openMovs().filter((m) => m.a_location_id === l.id).reduce((a, m) => a + m.qty, 0);
  const cap = l.capacita;
  const pct = cap ? Math.min(100, Math.round((n / cap) * 100)) : null;
  return `<button class="loccard ${l.tipo === 'sconosciuta' ? 'unk' : ''}" data-act="openloc" data-id="${l.id}">
    <div class="lc-h"><b>${esc(l.nome)}</b><span class="num">${n}</span></div>
    <div class="lc-s">${esc(l.ruolo || (l.tipo === 'sconosciuta' ? '' : ''))}${st.length ? ` · ${st.length} etichette` : ''}${inArr ? ` · <span class="accent">+${inArr} in arrivo</span>` : ''}</div>
    ${cap ? `<div class="bar ${n > cap ? 'over' : ''}"><i style="width:${pct}%"></i></div><div class="lc-s">${n}/${cap} posti</div>` : ''}
  </button>`;
}
function cantinaView() {
  const phys = physLocs();
  const unk = unknownLoc();
  const tot = S.giac.filter((g) => L.get(g.location_id)?.tipo === 'fisica').reduce((a, g) => a + g.qty, 0);
  const arr = openMovs().filter((m) => m.stato === 'in_corso').reduce((a, m) => a + m.qty, 0);
  const unkN = unk ? stockOfLoc(unk.id).reduce((a, g) => a + g.qty, 0) : 0;
  const shown = phys.filter((l) => stockOfLoc(l.id).length || openMovs().some((m) => m.a_location_id === l.id) || l.capacita);
  const empty = phys.filter((l) => !shown.includes(l));
  return `
    <section class="stats">
      <div><b class="num">${tot}</b><span>in posizione</span></div>
      <div><b class="num">${arr}</b><span>in arrivo</span></div>
      <div><b class="num">${unkN}</b><span>da ritrovare</span></div>
    </section>
    <div class="grid">${shown.map(locCard).join('')}${unk && unkN ? locCard(unk) : ''}</div>
    ${empty.length ? `<p class="note">Vuote: ${empty.map((l) => `<a href="#" data-act="openloc" data-id="${l.id}">${esc(l.nome)}</a>`).join(', ')}</p>` : ''}`;
}
function wineRow(w, qty, extra = '') {
  const st = dwState(w);
  return `<button class="row" data-act="openwine" data-id="${w.id}">
    <div class="r-main">${tipDot(w)}<div><b>${esc(wineTitle(w))}</b><span class="s">${esc(wineSub(w))}${extra}</span></div></div>
    <div class="r-side">${qty != null ? `<span class="qty num">${qty}</span>` : ''}${st ? `<span class="chip ${DWCLS[st]}">${st}</span>` : ''}${w.bere_presto ? '<span class="chip soon">★</span>' : ''}</div>
  </button>`;
}
function locView(v) {
  const l = L.get(v.id); if (!l) return '<p class="note">Location non trovata.</p>';
  const q = norm(S.q);
  let st = stockOfLoc(l.id).map((g) => ({ w: W.get(g.wine_id), qty: g.qty })).filter((x) => x.w);
  if (q) st = st.filter((x) => q.split(/\s+/).every((t) => norm(`${x.w.produttore} ${x.w.vino} ${x.w.annata} ${x.w.denominazione} ${x.w.regione}`).includes(t)));
  st.sort((a, b) => a.w.produttore.localeCompare(b.w.produttore) || String(a.w.annata).localeCompare(String(b.w.annata)));
  const incoming = openMovs().filter((m) => m.a_location_id === l.id);
  const n = stockOfLoc(l.id).reduce((a, g) => a + g.qty, 0);
  return `
    <p class="note">${esc(l.ruolo || '')}${l.capacita ? ` · ${n}/${l.capacita} posti` : ''}${l.note ? ` · ${esc(l.note)}` : ''}</p>
    <input class="search" data-bind="q" placeholder="Filtra in ${esc(l.nome)}…" value="${esc(S.q)}">
    <div class="list">${st.map((x) => wineRow(x.w, x.qty)).join('') || '<p class="note">Nessuna bottiglia.</p>'}</div>
    ${incoming.length ? `<div class="sec">In arrivo qui</div><div class="list">${incoming.map(movCard).join('')}</div>` : ''}`;
}
function cercaView() {
  const q = norm(S.q), f = S.filters;
  let ws = S.wines.map((w) => ({ w, qty: totalWine(w.id), open: openMovsOfWine(w.id).reduce((a, m) => a + (m.tipo === 'acquisto' ? m.qty : 0), 0) }));
  if (!f.zero) ws = ws.filter((x) => x.qty > 0 || x.open > 0);
  if (q) ws = ws.filter((x) => q.split(/\s+/).every((t) => norm(`${x.w.produttore} ${x.w.vino} ${x.w.annata} ${x.w.denominazione} ${x.w.regione} ${x.w.vitigno} ${x.w.paese}`).includes(t)));
  if (f.tip) ws = ws.filter((x) => x.w.tipologia === f.tip);
  if (f.dw) ws = ws.filter((x) => (dwState(x.w) || 'Senza finestra') === f.dw);
  if (f.loc) ws = ws.filter((x) => stockAt(x.w.id, Number(f.loc)) > 0);
  ws.sort((a, b) => a.w.produttore.localeCompare(b.w.produttore) || String(a.w.annata).localeCompare(String(b.w.annata)));
  const sel = (k, opts, label) => `<select data-filter="${k}"><option value="">${label}</option>${opts.map(([v, t]) => `<option value="${esc(v)}" ${String(f[k]) === String(v) ? 'selected' : ''}>${esc(t)}</option>`).join('')}</select>`;
  const tot = ws.reduce((a, x) => a + x.qty, 0);
  return `
    <input class="search" data-bind="q" placeholder="Produttore, vino, annata, zona…" value="${esc(S.q)}">
    <div class="filters">
      ${sel('tip', [['Rosso', 'Rosso'], ['Bianco', 'Bianco'], ['Rosé', 'Rosé']], 'Tipologia')}
      ${sel('dw', ['Da attendere', 'In finestra', 'Bere presto', 'Oltre finestra', 'Senza finestra'].map((x) => [x, x]), 'Finestra')}
      ${sel('loc', [...physLocs(), unknownLoc()].filter(Boolean).map((l) => [l.id, l.nome]), 'Location')}
      <label class="chk"><input type="checkbox" data-filter="zero" ${f.zero ? 'checked' : ''}> finiti</label>
    </div>
    <p class="note">${ws.length} vini · ${tot} bottiglie</p>
    <div class="list">${ws.slice(0, 300).map((x) => wineRow(x.w, x.qty, x.open ? ` · <span class="accent">+${x.open} in consegna</span>` : '')).join('')}</div>`;
}
const TIPO_LABEL = { carico_iniziale: 'Carico iniziale', acquisto: 'Acquisto', trasferimento: 'Trasferimento', consumo: 'Bevuta', rettifica: 'Rettifica' };
const STATO_LABEL = { pianificato: 'Pianificato', in_corso: 'In corso', confermato: 'Confermato', annullato: 'Annullato' };
function movLine(m) {
  const da = m.da_location_id ? L.get(m.da_location_id)?.nome : (m.fornitore || ({ rettifica: 'rettifica', carico_iniziale: 'Excel' }[m.tipo] || ''));
  const a = m.a_location_id ? L.get(m.a_location_id)?.nome : (m.tipo === 'consumo' ? 'bevuta' : 'uscita');
  return `${esc(da || '?')} → ${esc(a || '?')}`;
}
function movCard(m) {
  const w = W.get(m.wine_id); if (!w) return '';
  const label = m.tipo === 'acquisto' ? 'In consegna' : m.stato === 'pianificato' ? 'Da fare' : 'In transito';
  return `<div class="mcard">
    <button class="row flat" data-act="openwine" data-id="${w.id}"><div class="r-main">${tipDot(w)}<div><b>${m.qty}× ${esc(wineTitle(w))}</b>
      <span class="s">${esc(wineSub(w))}</span><span class="s">${movLine(m)} · <span class="chip ${m.stato === 'pianificato' ? 'wait' : 'soon'}">${label}</span> ${fmtDate(m.data_movimento)}</span>
      ${m.nota && !/^Import Cantina/.test(m.nota) ? `<span class="s">${esc(m.nota)}</span>` : ''}</div></div></button>
    <div class="acts">
      ${m.stato === 'pianificato' && m.tipo === 'trasferimento' ? `<button class="secondary sm" data-act="startmov" data-id="${m.id}">Partite</button>` : ''}
      <button class="primary sm" data-act="confirmmov" data-id="${m.id}">${m.stato === 'pianificato' ? 'Fatto' : 'Arrivate'}</button>
      <button class="secondary sm" data-act="destmov" data-id="${m.id}">Destinazione</button>
      <button class="ghost sm" data-act="cancelmov" data-id="${m.id}">Annulla</button>
    </div></div>`;
}
function arriviView() {
  const om = openMovs();
  const tr = om.filter((m) => m.tipo === 'trasferimento' && m.stato === 'in_corso');
  const pl = om.filter((m) => m.stato === 'pianificato');
  const acq = om.filter((m) => m.tipo === 'acquisto' && m.stato === 'in_corso');
  const byF = {};
  acq.forEach((m) => { const k = (m.fornitore || 'Senza fornitore') + (/Primeur/.test(m.nota || '') ? ' · Primeur' : ''); (byF[k] ||= []).push(m); });
  const sum = (a) => a.reduce((x, m) => x + m.qty, 0);
  return `
    <div class="sec">In transito <span class="muted">${sum(tr)}</span></div>
    <div class="list">${tr.map(movCard).join('') || '<p class="note">Nessun trasferimento in corso. Da una scheda vino usa "Sposta" e scegli "In transito".</p>'}</div>
    ${pl.length ? `<div class="sec">Da fare <span class="muted">${sum(pl)}</span></div><div class="list">${pl.map(movCard).join('')}</div>` : ''}
    <div class="sec">In consegna <span class="muted">${sum(acq)}</span></div>
    ${Object.entries(byF).sort().map(([k, ms]) => `<details class="group"><summary><b>${esc(k)}</b><span class="muted">${sum(ms)} bott. · ${ms.length} righe</span></summary>
      <div class="acts"><button class="primary sm" data-act="confirmgroup" data-key="${esc(k)}">Arrivata tutta la consegna</button></div>
      <div class="list">${ms.map(movCard).join('')}</div></details>`).join('') || '<p class="note">Nessuna consegna in attesa.</p>'}`;
}
function verificaView() {
  const ss = segnOpen();
  const groups = {};
  ss.forEach((s) => { const k = s.location_id; (groups[k] ||= []).push(s); });
  const ids = Object.keys(groups).map(Number).sort((a, b) => (L.get(a)?.tipo === 'sconosciuta' ? -1 : 0) - (L.get(b)?.tipo === 'sconosciuta' ? -1 : 0));
  return `<p class="note">Bottiglie di cui non c'è certezza. Risolvile quando le controlli: il resto dell'app resta coerente.</p>
    ${ids.map((lid) => `<div class="sec">${esc(L.get(lid)?.nome || 'Senza location')} <span class="muted">${groups[lid].length}</span></div>
      <div class="list">${groups[lid].map((s) => { const w = W.get(s.wine_id); if (!w) return ''; const n = stockAt(w.id, lid);
        return `<div class="mcard"><button class="row flat" data-act="openwine" data-id="${w.id}"><div class="r-main">${tipDot(w)}<div><b>${n}× ${esc(wineTitle(w))}</b><span class="s">${esc(wineSub(w))}</span><span class="s">${esc(s.motivo)}</span></div></div></button>
        <div class="acts"><button class="primary sm" data-act="resolve" data-id="${s.id}">Risolvi</button></div></div>`; }).join('')}</div>`).join('') || '<p class="note">Niente da verificare.</p>'}`;
}
function wineView(v) {
  const w = W.get(v.id); if (!w) return '<p class="note">Vino non trovato.</p>';
  const st = stockOfWine(w.id);
  const om = openMovsOfWine(w.id);
  const hist = S.movs.filter((m) => m.wine_id === w.id && (m.stato === 'confermato' || m.stato === 'annullato'));
  const dws = dwState(w); const pr = avgPrice(w.id);
  const segn = segnOpen().filter((s) => s.wine_id === w.id);
  return `
    <section class="card hero">${tipDot(w)}<div><h2>${esc(w.vino)}</h2><p>${esc(w.produttore)} · <b>${esc(w.annata)}</b></p>
      <p class="s">${esc([w.denominazione, w.regione, w.paese].filter(Boolean).join(' · '))}</p>
      <p class="s">${esc([w.tipologia, w.stile, w.vitigno, w.formato].filter(Boolean).join(' · '))}</p></div></section>
    <section class="card kv">
      <div><span>Finestra</span><b>${w.dw_inizio || '?'} – ${w.dw_fine || '?'}</b>${dws ? ` <span class="chip ${DWCLS[dws]}">${dws}</span>` : ''}${w.bere_presto ? ' <span class="chip soon">★ bere presto</span>' : ''}</div>
      ${w.fonte_dw ? `<div><span>Fonte</span><p class="s">${esc(w.fonte_dw)}</p></div>` : ''}
      ${pr != null ? `<div><span>Prezzo medio</span><b>${fmtEur(pr)}</b></div>` : ''}
      ${w.note ? `<div><span>Note</span><p class="s">${esc(w.note)}</p></div>` : ''}
    </section>
    <div class="sec">Dove si trova <span class="muted">${totalWine(w.id)}</span></div>
    <div class="list">${st.map((g) => `<div class="row flat"><div class="r-main"><b>${esc(L.get(g.location_id)?.nome)}</b></div><div class="r-side"><span class="qty num">${g.qty}</span></div></div>`).join('') || '<p class="note">Nessuna bottiglia in posizione.</p>'}</div>
    ${segn.length ? `<div class="list">${segn.map((s) => `<div class="mcard"><span class="s">⚠ ${esc(L.get(s.location_id)?.nome)}: ${esc(s.motivo)}</span><div class="acts"><button class="primary sm" data-act="resolve" data-id="${s.id}">Risolvi</button></div></div>`).join('')}</div>` : ''}
    <div class="actions">
      <button class="primary" data-act="move" data-id="${w.id}" ${st.length ? '' : 'disabled'}>Sposta</button>
      <button class="primary" data-act="drink" data-id="${w.id}" ${st.length ? '' : 'disabled'}>Bevuta</button>
      <button class="secondary" data-act="buy" data-id="${w.id}">Acquisto</button>
      <button class="secondary" data-act="adjust" data-id="${w.id}">Correggi</button>
      <button class="secondary" data-act="editwine" data-id="${w.id}">Modifica</button>
    </div>
    ${om.length ? `<div class="sec">In corso</div><div class="list">${om.map(movCard).join('')}</div>` : ''}
    <div class="sec">Storico</div>
    <div class="list">${hist.map((m) => `<div class="hist ${m.stato === 'annullato' ? 'canc' : ''}"><b>${TIPO_LABEL[m.tipo]} · ${m.qty}</b><span class="s">${movLine(m)} · ${fmtDate(m.data_conferma || m.data_movimento)}${m.creato_da && m.creato_da !== 'import' ? ` · ${esc(m.creato_da.split('@')[0])}` : ''}${m.stato === 'annullato' ? ' · annullato' : ''}</span>${m.nota && !/^Import Cantina/.test(m.nota) ? `<span class="s">${esc(m.nota)}</span>` : ''}${m.prezzo_unitario != null ? `<span class="s">${fmtEur(Number(m.prezzo_unitario))} cad.</span>` : ''}</div>`).join('') || '<p class="note">Nessun movimento.</p>'}</div>`;
}
function altroView() {
  return `<div class="list">
    <button class="row" data-act="buy"><div class="r-main"><div><b>Nuovo acquisto</b><span class="s">Vino nuovo o già in cantina, arrivato o in consegna</span></div></div></button>
    <button class="row" data-act="export"><div class="r-main"><div><b>Esporta in Excel</b><span class="s">Giacenze, movimenti e location in un file .xlsx</span></div></div></button>
    <button class="row" data-act="gomembri"><div class="r-main"><div><b>Persone autorizzate</b><span class="s">${S.membri.length} account</span></div></div></button>
    <button class="row" data-act="goluoghi"><div class="r-main"><div><b>Location</b><span class="s">Aggiungi o modifica posti e capacità</span></div></div></button>
    <button class="row" data-act="logout"><div class="r-main"><div><b>Esci</b><span class="s">${esc(S.session?.user?.email)}</span></div></div></button>
  </div>
  <p class="note">Il cruscotto con valori e grafici arriverà in una versione successiva.</p>`;
}
function membriView() {
  const me = norm(S.session?.user?.email);
  return `<p class="note">Solo queste email possono vedere e modificare la cantina. La persona deve poi aprire l'app e scegliere "Crea account" con la stessa email.</p>
    <div class="list">${S.membri.map((m) => `<div class="row flat"><div class="r-main"><div><b>${esc(m.nome || m.email)}</b><span class="s">${esc(m.email)}</span></div></div>
      ${norm(m.email) !== me ? `<div class="r-side"><button class="ghost sm" data-act="delmembro" data-email="${esc(m.email)}">Rimuovi</button></div>` : ''}</div>`).join('')}</div>
    <div class="actions"><button class="primary" data-act="addmembro">Aggiungi persona</button></div>`;
}
function luoghiView() {
  return `<div class="list">${S.locs.slice().sort((a, b) => (a.ordine ?? 50) - (b.ordine ?? 50)).map((l) => `<button class="row" data-act="editloc" data-id="${l.id}"><div class="r-main"><div><b>${esc(l.nome)}${l.attiva ? '' : ' (nascosta)'}</b><span class="s">${esc(l.ruolo || '')}${l.capacita != null ? ` · capacità ${l.capacita}` : ''}</span></div></div></button>`).join('')}</div>
    <div class="actions"><button class="primary" data-act="editloc">Nuova location</button></div>`;
}

/* ---------------- fogli (form) ---------------- */
function openSheet(kind, data = {}) { S.sheet = { kind, ...data }; renderSheet(); }
function closeSheet() { S.sheet = null; renderSheet(); }
const locOptions = (list, selected, placeholder = 'Scegli…') => `<option value="">${placeholder}</option>` + list.map((l) => `<option value="${l.id}" ${String(selected) === String(l.id) ? 'selected' : ''}>${esc(l.nome)}</option>`).join('');
const destLocs = () => physLocs();
function renderSheet() {
  const el = $('#sheet'); if (!el) return;
  const s = S.sheet;
  if (!s) { el.innerHTML = ''; return; }
  const w = s.wid ? W.get(s.wid) : null;
  const head = (t) => `<div class="sh-h"><b>${t}</b><button class="iconbtn" data-act="closesheet" aria-label="Chiudi">✕</button></div>${w ? `<p class="s">${esc(wineTitle(w))} · ${esc(w.annata)}</p>` : ''}`;
  let body = '';
  if (s.kind === 'move') {
    const st = stockOfWine(w.id);
    const from = s.from ?? (st.length === 1 ? st[0].location_id : '');
    const max = from ? stockAt(w.id, Number(from)) : 0;
    body = `${head('Sposta bottiglie')}<form id="f" class="form">
      <label>Da<select name="from" data-live>${locOptions(st.map((g) => L.get(g.location_id)), from)}</select></label>
      <label>Quante${max ? ` (max ${max})` : ''}<input name="qty" type="number" inputmode="numeric" min="1" max="${max || 1}" value="${Math.min(s.qty || 1, max || 1)}" required></label>
      <label>A<select name="to" required>${locOptions(destLocs().filter((l) => String(l.id) !== String(from)), s.to)}</select></label>
      <fieldset class="seg"><legend>Stato</legend>
        <label><input type="radio" name="stato" value="in_corso" checked><span>In transito<small>confermi all'arrivo</small></span></label>
        <label><input type="radio" name="stato" value="confermato"><span>Già arrivate</span></label>
        <label><input type="radio" name="stato" value="pianificato"><span>Da fare<small>non ancora partite</small></span></label>
      </fieldset>
      <label>Nota<input name="nota" placeholder="facoltativa"></label>
      <button class="primary" type="submit">Registra</button></form>`;
  } else if (s.kind === 'drink') {
    const st = stockOfWine(w.id);
    const from = s.from ?? (st.length === 1 ? st[0].location_id : '');
    const max = from ? stockAt(w.id, Number(from)) : 0;
    body = `${head('Bottiglia bevuta')}<form id="f" class="form">
      <label>Da dove<select name="from" data-live>${locOptions(st.map((g) => L.get(g.location_id)), from)}</select></label>
      <label>Quante${max ? ` (max ${max})` : ''}<input name="qty" type="number" inputmode="numeric" min="1" max="${max || 1}" value="1" required></label>
      <label>Nota<input name="nota" placeholder="occasione, impressioni…"></label>
      <button class="primary" type="submit">Registra</button></form>`;
  } else if (s.kind === 'adjust') {
    const all = [...physLocs(), unknownLoc()].filter(Boolean);
    const at = s.at ?? (stockOfWine(w.id)[0]?.location_id || '');
    const cur = at ? stockAt(w.id, Number(at)) : 0;
    body = `${head('Correggi quantità')}<p class="note">Per allineare il numero a quello che c'è davvero (bottiglie trovate in più, mancanti, rotte).</p><form id="f" class="form">
      <label>Location<select name="at" data-live>${locOptions(all, at)}</select></label>
      <label>Bottiglie presenti davvero (ora: ${cur})<input name="real" type="number" inputmode="numeric" min="0" value="${cur}" required></label>
      <label>Motivo<input name="nota" placeholder="es. rotta, trovata durante inventario"></label>
      <button class="primary" type="submit">Registra</button></form>`;
  } else if (s.kind === 'confirm') {
    const m = S.movs.find((x) => x.id === s.mid);
    body = `<div class="sh-h"><b>Conferma arrivo</b><button class="iconbtn" data-act="closesheet">✕</button></div><p class="s">${esc(wineTitle(W.get(m.wine_id)))} · ${movLine(m)}</p><form id="f" class="form">
      <label>Arrivate<input name="qty" type="number" inputmode="numeric" min="1" max="${m.qty}" value="${m.qty}" required></label>
      <p class="note">Se ne arrivano meno di ${m.qty}, le restanti restano in attesa.</p>
      <button class="primary" type="submit">Conferma</button></form>`;
  } else if (s.kind === 'dest') {
    const m = S.movs.find((x) => x.id === s.mid);
    body = `<div class="sh-h"><b>Cambia destinazione</b><button class="iconbtn" data-act="closesheet">✕</button></div><p class="s">${esc(wineTitle(W.get(m.wine_id)))} · ${m.qty} bott.</p><form id="f" class="form">
      <label>Nuova destinazione<select name="to" required>${locOptions(destLocs().filter((l) => l.id !== m.da_location_id), m.a_location_id)}</select></label>
      <button class="primary" type="submit">Salva</button></form>`;
  } else if (s.kind === 'resolve') {
    const sg = S.segn.find((x) => x.id === s.sid);
    const lw = W.get(sg.wine_id); const lid = sg.location_id; const n = stockAt(sg.wine_id, lid);
    const isUnk = L.get(lid)?.tipo === 'sconosciuta';
    body = `<div class="sh-h"><b>Risolvi</b><button class="iconbtn" data-act="closesheet">✕</button></div><p class="s">${esc(wineTitle(lw))} · ${esc(lw.annata)}</p><p class="note">${esc(sg.motivo)}. Registrate: ${n} in ${esc(L.get(lid)?.nome)}.</p>
      <form id="f" class="form">
      ${isUnk ? `<fieldset class="seg"><legend>Cosa ne sai?</legend>
          <label><input type="radio" name="esito" value="trovate" checked><span>Le ho trovate</span></label>
          <label><input type="radio" name="esito" value="transito"><span>Sono in transito</span></label>
          <label><input type="radio" name="esito" value="perse"><span>Non ci sono</span></label></fieldset>
        <label>Quante<input name="qty" type="number" inputmode="numeric" min="1" max="${n}" value="${n}" required></label>
        <label>Dove sono / verso dove<select name="to">${locOptions(destLocs(), '')}</select></label>`
      : `<label>Quante ce ne sono davvero in ${esc(L.get(lid)?.nome)}?<input name="real" type="number" inputmode="numeric" min="0" value="${n}" required></label>
         <p class="note">Se sono meno, la differenza esce dalla giacenza come rettifica.</p>`}
      <label>Nota<input name="nota" placeholder="facoltativa"></label>
      <button class="primary" type="submit">Conferma</button></form>`;
  } else if (s.kind === 'buy') {
    const mode = s.mode || (w ? 'existing' : 'existing');
    const q = norm(s.q || '');
    const res = !w && mode === 'existing' && q ? S.wines.filter((x) => q.split(/\s+/).every((t) => norm(`${x.produttore} ${x.vino} ${x.annata}`).includes(t))).slice(0, 25) : [];
    const pick = s.pick ? W.get(s.pick) : w;
    body = `<div class="sh-h"><b>Acquisto</b><button class="iconbtn" data-act="closesheet">✕</button></div>
      ${!w ? `<div class="segbtn"><button class="${mode === 'existing' ? 'on' : ''}" data-act="buymode" data-mode="existing">Vino già presente</button><button class="${mode === 'new' ? 'on' : ''}" data-act="buymode" data-mode="new">Vino nuovo</button></div>` : ''}
      ${!w && mode === 'existing' && !pick ? `<input class="search" id="buyq" placeholder="Cerca produttore, vino, annata…" value="${esc(s.q || '')}">
        <div class="list">${res.map((x) => `<button class="row" data-act="buypick" data-id="${x.id}"><div class="r-main">${tipDot(x)}<div><b>${esc(wineTitle(x))}</b><span class="s">${esc(wineSub(x))}</span></div></div></button>`).join('')}</div>` : ''}
      ${(pick || mode === 'new') ? `<form id="f" class="form">
        ${pick ? `<p class="s"><b>${esc(wineTitle(pick))}</b> · ${esc(wineSub(pick))}</p>` : wineFields({})}
        <div class="grid2"><label>Bottiglie<input name="qty" type="number" inputmode="numeric" min="1" value="1" required></label>
        <label>Prezzo cad. €<input name="prezzo" type="number" inputmode="decimal" step="0.01" min="0"></label></div>
        <label>Fornitore<input name="fornitore" list="fornitori"><datalist id="fornitori">${[...new Set(S.movs.map((m) => m.fornitore).filter(Boolean))].map((f) => `<option value="${esc(f)}">`).join('')}</datalist></label>
        <label>Destinazione<select name="to" required>${locOptions(destLocs(), S.locs.find((l) => l.nome === 'Bixio')?.id)}</select></label>
        <fieldset class="seg"><legend>Stato</legend>
          <label><input type="radio" name="stato" value="in_corso" checked><span>In consegna</span></label>
          <label><input type="radio" name="stato" value="confermato"><span>Già arrivato</span></label></fieldset>
        <label>Nota<input name="nota" placeholder="ordine, fattura…"></label>
        <button class="primary" type="submit">Registra acquisto</button></form>` : ''}`;
  } else if (s.kind === 'editwine') {
    body = `${head('Modifica vino')}<form id="f" class="form">${wineFields(w)}
      <label class="chk"><input type="checkbox" name="bere_presto" ${w.bere_presto ? 'checked' : ''}> Bere presto</label>
      <label>Note<textarea name="note" rows="2">${esc(w.note || '')}</textarea></label>
      <button class="primary" type="submit">Salva</button></form>`;
  } else if (s.kind === 'membro') {
    body = `<div class="sh-h"><b>Aggiungi persona</b><button class="iconbtn" data-act="closesheet">✕</button></div><form id="f" class="form">
      <label>Email<input name="email" type="email" required></label><label>Nome<input name="nome"></label>
      <button class="primary" type="submit">Aggiungi</button></form>`;
  } else if (s.kind === 'loc') {
    const l = s.lid ? L.get(s.lid) : {};
    body = `<div class="sh-h"><b>${s.lid ? 'Modifica location' : 'Nuova location'}</b><button class="iconbtn" data-act="closesheet">✕</button></div><form id="f" class="form">
      <label>Nome<input name="nome" value="${esc(l.nome || '')}" required ${l.tipo === 'sconosciuta' ? 'readonly' : ''}></label>
      <label>Ruolo<input name="ruolo" value="${esc(l.ruolo || '')}" placeholder="es. Casa – pronta beva"></label>
      <div class="grid2"><label>Capacità<input name="capacita" type="number" min="0" value="${l.capacita ?? ''}"></label>
      <label>Ordine<input name="ordine" type="number" value="${l.ordine ?? ''}"></label></div>
      <label>Note<input name="note" value="${esc(l.note || '')}"></label>
      ${l.tipo !== 'sconosciuta' ? `<label class="chk"><input type="checkbox" name="attiva" ${l.attiva !== false ? 'checked' : ''}> Visibile</label>` : ''}
      <button class="primary" type="submit">Salva</button></form>`;
  }
  el.innerHTML = `<div class="backdrop" data-act="closesheet"></div><div class="sheet">${body}</div>`;
  const bq = $('#buyq'); if (bq) { bq.focus(); bq.setSelectionRange(bq.value.length, bq.value.length); }
}
function wineFields(w) {
  const f = (name, label, attrs = '') => `<label>${label}<input name="${name}" value="${esc(w[name] ?? '')}" ${attrs}></label>`;
  return `${f('produttore', 'Produttore', 'required')}${f('vino', 'Vino', 'required')}
    <div class="grid2">${f('annata', 'Annata', 'required placeholder="2019 o NV" pattern="[0-9]{4}|NV"')}${f('formato', 'Formato', 'placeholder="750ml"')}</div>
    <div class="grid2"><label>Tipologia<select name="tipologia">${['', 'Rosso', 'Bianco', 'Rosé'].map((t) => `<option ${w.tipologia === t ? 'selected' : ''}>${t}</option>`).join('')}</select></label>${f('stile', 'Stile', 'placeholder="Fermo secco"')}</div>
    ${f('denominazione', 'Denominazione')}
    <div class="grid2">${f('regione', 'Regione')}${f('paese', 'Paese')}</div>
    ${f('vitigno', 'Vitigno')}
    <div class="grid2">${f('dw_inizio', 'Inizio finestra', 'type="number" inputmode="numeric"')}${f('dw_fine', 'Fine finestra', 'type="number" inputmode="numeric"')}</div>
    ${f('fonte_dw', 'Fonte finestra')}`;
}
function readWine(fd) {
  const g = (k) => { const v = (fd.get(k) ?? '').toString().trim(); return v === '' ? null : v; };
  const int = (k) => (g(k) == null ? null : parseInt(g(k), 10));
  return {
    produttore: g('produttore'), vino: g('vino'), annata: (g('annata') || '').toUpperCase(), formato: g('formato') || '750ml',
    tipologia: g('tipologia'), stile: g('stile'), denominazione: g('denominazione'), regione: g('regione'), paese: g('paese'),
    vitigno: g('vitigno'), dw_inizio: int('dw_inizio'), dw_fine: int('dw_fine'), fonte_dw: g('fonte_dw'),
  };
}

/* ---------------- azioni ---------------- */
async function submitSheet(form) {
  const fd = new FormData(form); const s = S.sheet;
  const num = (k) => parseInt(fd.get(k), 10);
  const nota = (fd.get('nota') || '').toString().trim() || null;
  let ok = false;
  if (s.kind === 'move') {
    const from = num('from'), to = num('to'), qty = num('qty');
    if (!from || !to) return toast('Scegli origine e destinazione');
    ok = await run(sb.from('movimenti').insert({ wine_id: s.wid, tipo: 'trasferimento', qty, da_location_id: from, a_location_id: to, stato: fd.get('stato'), nota }),
      fd.get('stato') === 'in_corso' ? 'In transito: lo trovi in "In arrivo"' : 'Registrato');
  } else if (s.kind === 'drink') {
    const from = num('from'); if (!from) return toast('Scegli da dove');
    ok = await run(sb.from('movimenti').insert({ wine_id: s.wid, tipo: 'consumo', qty: num('qty'), da_location_id: from, nota }), 'Alla salute');
  } else if (s.kind === 'adjust') {
    const at = num('at'), real = num('real'); const cur = stockAt(s.wid, at); const d = real - cur;
    if (!at) return toast('Scegli la location');
    if (d === 0) { closeSheet(); return toast('Nessuna differenza'); }
    ok = await run(sb.from('movimenti').insert(d > 0 ? { wine_id: s.wid, tipo: 'rettifica', qty: d, a_location_id: at, nota } : { wine_id: s.wid, tipo: 'rettifica', qty: -d, da_location_id: at, nota }), 'Quantità corretta');
  } else if (s.kind === 'confirm') {
    const m = S.movs.find((x) => x.id === s.mid); const qty = num('qty');
    if (qty < m.qty) {
      const { id, created_at, updated_at, data_conferma, ...rest } = m;
      // prima confermo la parte arrivata, poi creo il resto ancora in attesa (così il controllo giacenze resta coerente)
      ok = await run(sb.from('movimenti').update({ qty, stato: 'confermato' }).eq('id', m.id));
      if (ok) ok = await run(sb.from('movimenti').insert({ ...rest, qty: m.qty - qty, creato_da: undefined, nota: [m.nota, 'resto di consegna parziale'].filter(Boolean).join('; ') }), 'Arrivo parziale registrato');
    } else ok = await run(sb.from('movimenti').update({ stato: 'confermato' }).eq('id', m.id), 'Arrivo confermato');
  } else if (s.kind === 'dest') {
    ok = await run(sb.from('movimenti').update({ a_location_id: num('to') }).eq('id', s.mid), 'Destinazione aggiornata');
  } else if (s.kind === 'resolve') {
    const sg = S.segn.find((x) => x.id === s.sid); const lid = sg.location_id; const n = stockAt(sg.wine_id, lid);
    const me = S.session.user.email; const close = (esito) => sb.from('segnalazioni').update({ aperta: false, esito, chiusa_da: me, chiusa_at: new Date().toISOString() }).eq('id', sg.id);
    if (L.get(lid)?.tipo === 'sconosciuta') {
      const esito = fd.get('esito'), qty = num('qty'), to = num('to');
      if (esito !== 'perse' && !to) return toast('Scegli la location');
      const mv = esito === 'perse' ? { wine_id: sg.wine_id, tipo: 'rettifica', qty, da_location_id: lid, nota: nota || 'Non trovate' }
        : { wine_id: sg.wine_id, tipo: 'trasferimento', qty, da_location_id: lid, a_location_id: to, stato: esito === 'trovate' ? 'confermato' : 'in_corso', nota };
      ok = await run(sb.from('movimenti').insert(mv));
      if (ok) ok = await run(n - qty <= 0 ? close(esito + (nota ? ': ' + nota : '')) : sb.from('segnalazioni').update({ qty: n - qty }).eq('id', sg.id), 'Fatto');
    } else {
      const real = num('real'); const d = real - n;
      if (d !== 0) ok = await run(sb.from('movimenti').insert(d > 0 ? { wine_id: sg.wine_id, tipo: 'rettifica', qty: d, a_location_id: lid, nota: nota || 'Verifica' } : { wine_id: sg.wine_id, tipo: 'rettifica', qty: -d, da_location_id: lid, nota: nota || 'Verifica: mancanti' }));
      else ok = true;
      if (ok) ok = await run(close(`verificate ${real}/${n}` + (nota ? ': ' + nota : '')), 'Verifica registrata');
    }
  } else if (s.kind === 'buy') {
    let wid = s.pick || s.wid;
    if (!wid) {
      const nw = readWine(fd);
      const { data, error } = await sb.from('wines').insert(nw).select('id').single();
      if (error) return toast(errMsg(error));
      wid = data.id;
    }
    const prezzo = fd.get('prezzo') ? Number(fd.get('prezzo')) : null;
    ok = await run(sb.from('movimenti').insert({ wine_id: wid, tipo: 'acquisto', qty: num('qty'), a_location_id: num('to'), stato: fd.get('stato'), fornitore: (fd.get('fornitore') || '').toString().trim() || null, prezzo_unitario: prezzo, nota }), 'Acquisto registrato');
    if (ok && !s.wid) { closeSheet(); await loadAll(); return go({ name: 'wine', id: wid }); }
  } else if (s.kind === 'editwine') {
    const upd = { ...readWine(fd), bere_presto: fd.get('bere_presto') === 'on', note: (fd.get('note') || '').toString().trim() || null };
    ok = await run(sb.from('wines').update(upd).eq('id', s.wid), 'Salvato');
  } else if (s.kind === 'membro') {
    ok = await run(sb.from('membri').insert({ email: fd.get('email').toString().trim().toLowerCase(), nome: (fd.get('nome') || '').toString().trim() || null }), 'Persona aggiunta');
  } else if (s.kind === 'loc') {
    const g = (k) => { const v = (fd.get(k) ?? '').toString().trim(); return v === '' ? null : v; };
    const row = { nome: g('nome'), ruolo: g('ruolo'), capacita: g('capacita') == null ? null : Number(g('capacita')), ordine: g('ordine') == null ? null : Number(g('ordine')), note: g('note') };
    if (L.get(s.lid)?.tipo !== 'sconosciuta') row.attiva = fd.get('attiva') === 'on';
    ok = await run(s.lid ? sb.from('locations').update(row).eq('id', s.lid) : sb.from('locations').insert(row), 'Salvato');
  }
  if (ok) { closeSheet(); await loadAll(); }
}

async function exportXlsx() {
  const wb = XLSX.utils.book_new();
  const head = ['Quantità', 'Vino', 'Produttore', 'Annata', 'Denominazione', 'Regione', 'Paese', 'Tipologia', 'Stile', 'Vitigno', 'Formato', 'Prezzo medio (€)', 'Location', 'Stato', 'Destinazione', 'Fornitore', 'Inizio DW', 'Fine DW', 'Stato DW', 'Fonte DW', 'Bere presto'];
  const rows = [];
  const segnKey = new Set(segnOpen().map((s) => `${s.wine_id}|${s.location_id}`));
  const base = (w) => [w.vino, w.produttore, w.annata, w.denominazione, w.regione, w.paese, w.tipologia, w.stile, w.vitigno, w.formato, avgPrice(w.id) == null ? null : Math.round(avgPrice(w.id) * 100) / 100];
  const tail = (w) => [w.dw_inizio, w.dw_fine, dwState(w), w.fonte_dw, w.bere_presto ? 'Sì' : ''];
  S.giac.filter((g) => g.qty > 0).forEach((g) => { const w = W.get(g.wine_id); const l = L.get(g.location_id);
    rows.push([g.qty, ...base(w), l?.nome, l?.tipo === 'sconosciuta' || segnKey.has(`${g.wine_id}|${g.location_id}`) ? 'Da verificare' : 'In posizione', null, null, ...tail(w)]); });
  // i pianificati sono ancora fisicamente nell'origine: già contati nelle giacenze, li trovi nel foglio Movimenti
  openMovs().filter((m) => m.stato === 'in_corso').forEach((m) => { const w = W.get(m.wine_id);
    rows.push([m.qty, ...base(w), m.da_location_id ? L.get(m.da_location_id)?.nome : null, m.tipo === 'acquisto' ? 'In consegna' : m.stato === 'pianificato' ? 'Pianificato' : 'In transito', L.get(m.a_location_id)?.nome, m.fornitore, ...tail(w)]); });
  rows.sort((a, b) => String(a[2]).localeCompare(String(b[2])) || String(a[3]).localeCompare(String(b[3])));
  const ws1 = XLSX.utils.aoa_to_sheet([head, ...rows]); ws1['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: head.length - 1 } }) };
  XLSX.utils.book_append_sheet(wb, ws1, 'Cantina');
  const mh = ['Data', 'Tipo', 'Stato', 'Quantità', 'Produttore', 'Vino', 'Annata', 'Da', 'A', 'Fornitore', 'Prezzo cad. (€)', 'Nota', 'Registrato da', 'Confermato il'];
  const mr = S.movs.map((m) => { const w = W.get(m.wine_id); return [fmtDate(m.data_movimento), TIPO_LABEL[m.tipo], STATO_LABEL[m.stato], m.qty, w?.produttore, w?.vino, w?.annata, L.get(m.da_location_id)?.nome, L.get(m.a_location_id)?.nome, m.fornitore, m.prezzo_unitario == null ? null : Number(m.prezzo_unitario), m.nota, m.creato_da, fmtDate(m.data_conferma)]; });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([mh, ...mr]), 'Movimenti');
  const lh = ['Location', 'Ruolo', 'Capacità', 'Occupazione', 'In arrivo', 'Note'];
  const lr = S.locs.map((l) => [l.nome, l.ruolo, l.capacita, stockOfLoc(l.id).reduce((a, g) => a + g.qty, 0), openMovs().filter((m) => m.a_location_id === l.id).reduce((a, m) => a + m.qty, 0), l.note]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([lh, ...lr]), 'Location');
  XLSX.writeFile(wb, `Cantina_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

/* ---------------- eventi ---------------- */
let toastT;
function toast(msg) { const t = $('#toast'); if (!t) return alertFallback(msg); t.textContent = msg; t.className = 'show'; clearTimeout(toastT); toastT = setTimeout(() => (t.className = ''), 3200); }
function alertFallback(msg) { console.log(msg); }

document.addEventListener('click', async (e) => {
  const t = e.target.closest('[data-act],[data-tab]'); if (!t) return;
  if (t.tagName === 'A') e.preventDefault();
  if (t.dataset.tab) { S.history = []; S.q = ''; return go({ name: t.dataset.tab }, false); }
  const a = t.dataset.act, id = t.dataset.id ? Number(t.dataset.id) : null;
  if (a === 'back') back();
  else if (a === 'reload') loadAll();
  else if (a === 'openloc') { S.q = ''; go({ name: 'loc', id }); }
  else if (a === 'openwine') go({ name: 'wine', id });
  else if (a === 'closesheet') closeSheet();
  else if (a === 'move') openSheet('move', { wid: id });
  else if (a === 'drink') openSheet('drink', { wid: id });
  else if (a === 'adjust') openSheet('adjust', { wid: id });
  else if (a === 'editwine') openSheet('editwine', { wid: id });
  else if (a === 'buy') openSheet('buy', { wid: id || null, mode: 'existing' });
  else if (a === 'buymode') { S.sheet.mode = t.dataset.mode; S.sheet.pick = null; renderSheet(); }
  else if (a === 'buypick') { S.sheet.pick = id; renderSheet(); }
  else if (a === 'confirmmov') openSheet('confirm', { mid: id });
  else if (a === 'destmov') openSheet('dest', { mid: id });
  else if (a === 'startmov') { if (await run(sb.from('movimenti').update({ stato: 'in_corso' }).eq('id', id), 'In transito')) loadAll(); }
  else if (a === 'cancelmov') { if (confirm('Annullare questo movimento?') && await run(sb.from('movimenti').update({ stato: 'annullato' }).eq('id', id), 'Annullato')) loadAll(); }
  else if (a === 'confirmgroup') {
    const key = t.dataset.key;
    const ms = openMovs().filter((m) => m.tipo === 'acquisto' && m.stato === 'in_corso' && ((m.fornitore || 'Senza fornitore') + (/Primeur/.test(m.nota || '') ? ' · Primeur' : '')) === key);
    if (!confirm(`Confermare l'arrivo di ${ms.reduce((x, m) => x + m.qty, 0)} bottiglie (${key})?`)) return;
    if (await run(sb.from('movimenti').update({ stato: 'confermato' }).in('id', ms.map((m) => m.id)), 'Consegna arrivata')) loadAll();
  }
  else if (a === 'resolve') openSheet('resolve', { sid: id });
  else if (a === 'export') { try { await exportXlsx(); } catch (err) { toast('Export non riuscito: ' + err.message); } }
  else if (a === 'gomembri') go({ name: 'membri' });
  else if (a === 'goluoghi') go({ name: 'luoghi' });
  else if (a === 'addmembro') openSheet('membro');
  else if (a === 'delmembro') { if (confirm(`Rimuovere ${t.dataset.email}?`) && await run(sb.from('membri').delete().eq('email', t.dataset.email), 'Rimossa')) loadAll(); }
  else if (a === 'editloc') openSheet('loc', { lid: id });
  else if (a === 'logout') { await sb.auth.signOut(); }
});
document.addEventListener('submit', async (e) => {
  e.preventDefault();
  const form = e.target;
  const btns = form.querySelectorAll('button'); btns.forEach((b) => (b.disabled = true));
  try {
    if (form.id === 'loginform') {
      const fd = new FormData(form); const mode = e.submitter?.value;
      const creds = { email: fd.get('email').toString().trim(), password: fd.get('password').toString() };
      const { error } = mode === 'up' ? await sb.auth.signUp(creds) : await sb.auth.signInWithPassword(creds);
      if (error) {
        const m = error.message || '';
        toast(/Invalid login/.test(m) ? 'Email o password errate' : /already registered/.test(m) ? 'Account già esistente: usa "Accedi"' : /not authorized|confirm/i.test(m) ? 'Registrazione non completata: serve disattivare la conferma email in Supabase' : m);
      }
    } else await submitSheet(form);
  } finally { btns.forEach((b) => (b.disabled = false)); }
});
document.addEventListener('input', (e) => {
  const t = e.target;
  if (t.dataset.bind === 'q') { S.q = t.value; const pos = t.selectionStart; render(); const n = document.querySelector('[data-bind="q"]'); if (n) { n.focus(); n.setSelectionRange(pos, pos); } }
  if (t.id === 'buyq') { S.sheet.q = t.value; renderSheet(); }
});
document.addEventListener('change', (e) => {
  const t = e.target;
  if (t.dataset.filter) { S.filters[t.dataset.filter] = t.type === 'checkbox' ? t.checked : t.value; render(); }
  if (t.hasAttribute('data-live') && S.sheet) { S.sheet[t.name] = t.value; renderSheet(); }
});
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && S.session && S.member) loadAll(); });

/* ---------------- avvio ---------------- */
async function onSession(session) {
  S.session = session;
  if (!session) { S.member = null; render(); return; }
  const { data, error } = await sb.rpc('is_member');
  S.member = error ? null : !!data;
  render();
  if (S.member) loadAll();
}
sb.auth.onAuthStateChange((_ev, session) => {
  if ((session?.access_token || null) !== (S.session?.access_token || null)) setTimeout(() => onSession(session), 0);
});
sb.auth.getSession().then(({ data }) => onSession(data.session));
if ('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js').catch(() => {});
