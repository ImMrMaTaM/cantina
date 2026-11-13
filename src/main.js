import { createClient } from '@supabase/supabase-js';
import * as XLSX from 'xlsx';

const SUPABASE_URL = 'https://ddeoswpycvlfhztjlkdj.supabase.co';
const SUPABASE_KEY = 'sb_publishable_zEUf50vq4WcErjmNuCev3g_Lko70Obs';
const sb = createClient(SUPABASE_URL, SUPABASE_KEY, { auth: { persistSession: true, autoRefreshToken: true } });

/* ================= stato ================= */
const S = {
  session: null, member: null, loading: false,
  wines: [], locs: [], giac: [], movs: [], segn: [], membri: [],
  view: { name: 'home' }, history: [], sheet: null,
  q: '', filters: { tip: '', dw: '', loc: '', zero: false },
  comp: { loc: 'all', dim: 'tip' }, outQ: '',
};
const $ = (s) => document.querySelector(s);
const esc = (v) => String(v ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = (s) => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
const fmtEur = (n) => n == null ? '' : n.toLocaleString('it-IT', { style: 'currency', currency: 'EUR', maximumFractionDigits: 0 });
const fmtDate = (d) => d ? new Date(d).toLocaleDateString('it-IT', { day: '2-digit', month: 'short', year: 'numeric' }) : '';
const YEAR = new Date().getFullYear();
const sum = (a, f = (x) => x.qty) => a.reduce((t, x) => t + f(x), 0);

/* ================= indici e calcoli ================= */
let W = new Map(), L = new Map();
function reindex() { W = new Map(S.wines.map((w) => [w.id, w])); L = new Map(S.locs.map((l) => [l.id, l])); }
const byOrder = (a, b) => (a.ordine ?? 50) - (b.ordine ?? 50) || a.nome.localeCompare(b.nome);
const children = (gid) => S.locs.filter((l) => l.parent_id === gid && l.attiva).sort(byOrder);
const isGroup = (l) => l?.tipo === 'gruppo';
const physLocs = () => S.locs.filter((l) => l.tipo === 'fisica' && l.attiva).sort(byOrder);
const topLocs = () => S.locs.filter((l) => (l.tipo === 'gruppo' || l.tipo === 'fisica') && !l.parent_id && l.attiva).sort(byOrder);
const unknownLoc = () => S.locs.find((l) => l.tipo === 'sconosciuta');
const homeGroup = () => S.locs.find((l) => isGroup(l) && norm(l.ruolo) === 'casa') || S.locs.find(isGroup);
const locIdsOf = (l) => (isGroup(l) ? children(l.id).map((c) => c.id) : [l.id]);
const locLabel = (l) => (l?.parent_id ? `${L.get(l.parent_id)?.nome} · ${l.nome}` : l?.nome || '');
const stockOfWine = (wid) => S.giac.filter((g) => g.wine_id === wid && g.qty > 0);
const stockAt = (wid, lid) => S.giac.find((g) => g.wine_id === wid && g.location_id === lid)?.qty || 0;
const stockIn = (wid, ids) => sum(S.giac.filter((g) => g.wine_id === wid && ids.includes(g.location_id) && g.qty > 0));
const stockOfLoc = (lid) => S.giac.filter((g) => g.location_id === lid && g.qty > 0);
const stockOfLocs = (ids) => S.giac.filter((g) => ids.includes(g.location_id) && g.qty > 0);
const physIds = () => physLocs().map((l) => l.id);
const totalWine = (wid) => sum(stockOfWine(wid));
const openMovs = () => S.movs.filter((m) => m.stato === 'in_corso' || m.stato === 'pianificato');
const openMovsOfWine = (wid) => openMovs().filter((m) => m.wine_id === wid);
const segnOpen = () => S.segn.filter((s) => s.aperta);
function avgPrice(wid) {
  const ms = S.movs.filter((m) => m.wine_id === wid && (m.tipo === 'acquisto' || m.tipo === 'carico_iniziale') && m.stato !== 'annullato' && m.prezzo_unitario != null);
  const q = sum(ms);
  return q ? sum(ms, (m) => m.qty * Number(m.prezzo_unitario)) / q : null;
}
/* Finestra di beva a 3 stati (deciso il 13/11/2026) */
const DW = { asap: 'Da bere asap', bere: 'Si può bere', aspettare: 'Meglio aspettare', senza: 'Senza finestra' };
const DWORDER = ['asap', 'bere', 'aspettare', 'senza'];
function dwKey(w) {
  const i = w.dw_inizio, f = w.dw_fine;
  if (w.bere_presto || (f && f <= YEAR + 1)) return 'asap';
  if (!i && !f) return 'senza';
  if (i && i > YEAR) return 'aspettare';
  return 'bere';
}
/* Tipologie come nella Panoramica */
const TIPS = ['Rossi', 'Bianchi', 'Bollicine', 'Rosati', 'Dolci'];
function tipKey(w) {
  if (w.stile === 'Bollicine') return 'Bollicine';
  if (w.stile === 'Dolce' || w.stile === 'Fermo abboccato') return 'Dolci';
  return w.tipologia === 'Rosso' ? 'Rossi' : w.tipologia === 'Rosé' ? 'Rosati' : 'Bianchi';
}
const TIPCLS = { Rossi: 'r', Bianchi: 'b', Bollicine: 'f', Rosati: 'p', Dolci: 'd' };
const tipDot = (w) => `<span class="dot ${TIPCLS[tipKey(w)]}"></span>`;
const wineTitle = (w) => `${w.produttore} · ${w.vino}`;
const wineSub = (w) => [w.annata, w.denominazione, w.formato !== '750ml' ? w.formato : null].filter(Boolean).join(' · ');
const wineShort = (w) => `${w.vino.startsWith(w.produttore) ? w.vino : w.produttore + ' ' + w.vino} ${w.annata}`;
function whereLine(wid) {
  const st = stockOfWine(wid).slice().sort((a, b) => byOrder(L.get(a.location_id), L.get(b.location_id)));
  return st.map((g) => `${L.get(g.location_id)?.nome} ${g.qty}`).join(' · ');
}
const MOTIVI = { bevuta: 'Bevuta', regalata: 'Regalata', rotta: 'Rotta', venduta: 'Venduta', altro: 'Altro' };

/* ================= dati ================= */
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
  if (/sotto-location/.test(m)) return 'Scegli una sotto-location, non il gruppo';
  if (/duplicate key|wines_unico/.test(m)) return 'Esiste già un vino con stesso produttore, nome, annata e formato.';
  return m;
}
async function run(promise, okMsg) {
  const { error } = await promise;
  if (error) { toast(errMsg(error)); return false; }
  if (okMsg) toast(okMsg);
  return true;
}

/* ================= navigazione ================= */
function go(view, push = true) {
  if (push) S.history.push(S.view);
  S.view = view; S.sheet = null; window.scrollTo(0, 0); render();
}
function back() { S.view = S.history.pop() || { name: 'home' }; S.sheet = null; render(); }
function home() { S.history = []; S.view = { name: 'home' }; S.sheet = null; window.scrollTo(0, 0); render(); }
const ICON = {
  add: '<rect x="4" y="9" width="16" height="11" rx="1.5"/><path d="M3 9h18M12 9v11M12 9c-1.5-3-5-3.5-5-1.2C7 9 9.5 9 12 9zM12 9c1.5-3 5-3.5 5-1.2C17 9 14.5 9 12 9z"/>',
  search: '<circle cx="10.5" cy="10.5" r="6.5"/><path d="M15.5 15.5L21 21"/>',
  out: '<path d="M8 10h8c0 3.3-1.8 5.5-4 5.5S8 13.3 8 10zM12 15.5V20M9 20.5h6M14.5 2.5l4 3-3.2 4.2-2.2-.4-.6-2.4zM13.6 9.6L12.4 11"/>',
  truck: '<path d="M3 7h11v9H3zM14 10h4l3 3v3h-7"/><circle cx="7" cy="18" r="1.6"/><circle cx="17" cy="18" r="1.6"/>',
  check: '<circle cx="11" cy="11" r="6"/><path d="M20 20l-4.5-4.5M11 8.5v3M11 14h.01"/>',
  found: '<path d="M12 3v12M6 9l6 6 6-6M4 20h16"/>',
};
const svg = (k, size = 28) => `<svg width="${size}" height="${size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON[k]}</svg>`;

/* ================= render ================= */
function render() {
  const app = $('#app');
  if (!S.session) { app.innerHTML = loginView(); return; }
  if (S.member === false) { app.innerHTML = notMemberView(); return; }
  const v = S.view;
  const body = {
    home: homeView, cerca: cercaView, wine: wineView, loc: locView, arrivi: arriviView, verifica: verificaView,
    altro: altroView, membri: membriView, luoghi: luoghiView, aggiungi: aggiungiView, togli: togliView,
  }[v.name]?.(v) ?? '';
  const act = { aggiungi: 'aggiungi', cerca: 'cerca', togli: 'togli' }[v.name] || '';
  app.innerHTML = `
    <header class="top" id="top"></header>
    <main>${body}</main>
    <nav class="bar" aria-label="Azioni principali">
      <button class="${act === 'aggiungi' ? 'on' : ''}" data-act="goaggiungi">${svg('add')}<span>Aggiungi</span></button>
      <button class="${act === 'cerca' ? 'on' : ''}" data-act="gocerca">${svg('search')}<span>Cerca</span></button>
      <button class="${act === 'togli' ? 'on' : ''}" data-act="gotogli">${svg('out')}<span>Togli</span></button>
    </nav>
    <div id="sheet"></div><div id="toast"></div>`;
  renderTop(); renderSheet();
}
function renderTop() {
  const t = $('#top'); if (!t) return;
  const v = S.view;
  const titles = { home: 'Panoramica', cerca: 'Cerca', arrivi: 'In arrivo', verifica: 'Da verificare', altro: 'Altro', membri: 'Persone', luoghi: 'Location', aggiungi: 'Aggiungi', togli: 'Togli' };
  const title = titles[v.name] ?? (v.name === 'loc' ? L.get(v.id)?.nome : v.name === 'wine' ? 'Scheda vino' : '');
  const badge = openMovs().length + segnOpen().length;
  t.innerHTML = v.name === 'home'
    ? `<div class="ttl"><span class="kicker" data-act="home">Cantina</span><h1>Panoramica</h1></div>`
    : `<button class="iconbtn" data-act="back" aria-label="Indietro">‹</button><div class="ttl"><span class="kicker" data-act="home">Cantina</span><h1>${esc(title)}</h1></div>`;
  t.innerHTML += `<button class="iconbtn ${S.loading ? 'spin' : ''}" data-act="reload" aria-label="Aggiorna">⟳</button>
    <button class="morebtn" data-act="goaltro" aria-label="Altro"><i></i><i></i><i></i>${v.name !== 'home' && badge ? '<b class="pip"></b>' : ''}</button>`;
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

/* ---------- Panoramica ---------- */
function homeView() {
  const pids = physIds();
  const inCellar = stockOfLocs(pids);
  const nBott = sum(inCellar);
  const nLab = new Set(inCellar.map((g) => g.wine_id)).size;
  const arr = sum(openMovs().filter((m) => m.stato === 'in_corso'));
  const unk = unknownLoc();
  const nUnk = unk ? sum(stockOfLoc(unk.id)) : 0;
  const nCheck = sum(segnOpen().filter((s) => L.get(s.location_id)?.tipo !== 'sconosciuta'), (s) => stockAt(s.wine_id, s.location_id));
  // quando berle
  const byDw = Object.fromEntries(DWORDER.map((k) => [k, 0]));
  inCellar.forEach((g) => { const w = W.get(g.wine_id); if (w) byDw[dwKey(w)] += g.qty; });
  // asap
  const hg = homeGroup(); const homeIds = hg ? locIdsOf(hg) : [];
  const asapW = S.wines.filter((w) => dwKey(w) === 'asap' && stockIn(w.id, pids) > 0);
  asapW.sort((a, b) => (stockIn(b.id, homeIds) > 0) - (stockIn(a.id, homeIds) > 0) || (a.dw_fine ?? 9999) - (b.dw_fine ?? 9999));
  // stasera
  const homeStock = stockOfLocs(homeIds);
  const byTipHome = Object.fromEntries(TIPS.map((k) => [k, 0]));
  homeStock.forEach((g) => { const w = W.get(g.wine_id); if (w) byTipHome[tipKey(w)] += g.qty; });

  return `
    <section class="kpis">
      <div class="kpi"><b>${nBott}</b><span>bottiglie</span></div>
      <div class="kpi"><b>${nLab}</b><span>etichette</span></div>
      <div class="kpi"><b>${arr}</b><span>in arrivo</span></div>
    </section>
    <section class="stack">
      <button class="alert" data-act="goarrivi"><span class="ic wine">${svg('truck', 18)}</span><span class="grow"><b>${arr} bottiglie in arrivo</b><small>consegne e trasferimenti da confermare</small></span><span class="chev">›</span></button>
      <button class="alert" data-act="goverifica"><span class="ic warn">${svg('check', 18)}</span><span class="grow"><b>${nUnk + nCheck} bottiglie da verificare</b><small>${nUnk} da ritrovare, ${nCheck} da controllare</small></span><span class="chev">›</span></button>
    </section>
    <section class="card">
      <h2>Quando berle</h2>
      <div class="stackbar">${DWORDER.map((k) => byDw[k] ? `<span class="c-${k}" style="flex-grow:${byDw[k]}"></span>` : '').join('')}</div>
      <div class="legend">${DWORDER.map((k) => `<button data-act="cercadw" data-dw="${k}"><span class="dot c-${k}"></span><span><b>${byDw[k]}</b><small>${DW[k]}</small></span></button>`).join('')}</div>
    </section>
    ${compCard()}
    <section class="card">
      <div class="h2row"><h2>Da bere asap</h2><button class="link" data-act="cercadw" data-dw="asap">Tutti i ${asapW.length} vini</button></div>
      ${hg ? `<p class="note">Prima quelli a ${esc(hg.nome)}</p>` : ''}
      <div class="rows">${asapW.slice(0, 5).map((w) => {
        const f = w.dw_fine; const tag = f ? (f < YEAR ? `chiusa ${f}` : `chiude ${f}`) : 'bere presto';
        return `<button class="lrow" data-act="openwine" data-id="${w.id}">${tipDot(w)}<span class="grow"><b>${esc(wineShort(w))}</b><small>${esc(whereLine(w.id))}</small></span><span class="tag">${tag}</span></button>`;
      }).join('') || '<p class="note">Niente di urgente.</p>'}</div>
    </section>
    ${hg ? `<section class="card">
      <div class="h2row"><h2>Da aprire stasera</h2><span class="muted">${sum(homeStock)} a ${esc(hg.nome)}</span></div>
      <div class="chips">${TIPS.filter((k) => byTipHome[k]).map((k) => `<button class="pill" data-act="cercatip" data-tip="${k}" data-loc="g${hg.id}"><span class="dot ${TIPCLS[k]}"></span><b>${byTipHome[k]}</b> ${k.toLowerCase()}</button>`).join('')}</div>
      <p class="note">Un tocco apre la lista filtrata.</p>
    </section>` : ''}
    ${locSection()}`;
}
function compCard() {
  const tops = topLocs().filter((l) => sum(stockOfLocs(locIdsOf(l))) > 0);
  const sel = S.comp.loc === 'all' ? null : L.get(Number(S.comp.loc));
  const ids = sel ? locIdsOf(sel) : physIds();
  const st = stockOfLocs(ids);
  const tot = sum(st);
  const acc = {};
  st.forEach((g) => {
    const w = W.get(g.wine_id); if (!w) return;
    const k = S.comp.dim === 'tip' ? tipKey(w) : S.comp.dim === 'fin' ? dwKey(w) : (w.regione || 'Senza regione');
    acc[k] = (acc[k] || 0) + g.qty;
  });
  let rows = Object.entries(acc);
  if (S.comp.dim === 'tip') rows.sort((a, b) => b[1] - a[1]);
  else if (S.comp.dim === 'fin') rows.sort((a, b) => DWORDER.indexOf(a[0]) - DWORDER.indexOf(b[0]));
  else rows = rows.sort((a, b) => b[1] - a[1]).slice(0, 6);
  const max = Math.max(1, ...rows.map((r) => r[1]));
  const cls = (k) => (S.comp.dim === 'tip' ? `t-${TIPCLS[k]}` : S.comp.dim === 'fin' ? `c-${k}` : 't-r');
  const label = (k) => (S.comp.dim === 'fin' ? DW[k] : k);
  return `<section class="card">
    <div class="h2row"><h2>Com'è fatta la cantina</h2><span class="muted">${tot} bottiglie</span></div>
    <div class="chips scroll" role="group" aria-label="Location">
      <button class="fchip ${!sel ? 'on' : ''}" data-act="comploc" data-loc="all" aria-pressed="${!sel}">Tutta</button>
      ${tops.map((l) => `<button class="fchip ${sel?.id === l.id ? 'on' : ''}" data-act="comploc" data-loc="${l.id}" aria-pressed="${sel?.id === l.id}">${esc(l.nome)}</button>`).join('')}
    </div>
    <div class="seg" role="group" aria-label="Ripartisci per">
      ${[['tip', 'Tipologia'], ['reg', 'Regione'], ['fin', 'Finestra']].map(([k, t]) => `<button class="${S.comp.dim === k ? 'on' : ''}" data-act="compdim" data-dim="${k}" aria-pressed="${S.comp.dim === k}">${t}</button>`).join('')}
    </div>
    <div class="bars">${rows.map(([k, n]) => `<button class="brow" data-act="compgo" data-key="${esc(k)}"><span>${esc(label(k))}</span><span class="track"><i class="${cls(k)}" style="width:${(n / max * 100).toFixed(1)}%"></i></span><b>${n}</b></button>`).join('') || '<p class="note">Nessuna bottiglia.</p>'}</div>
  </section>`;
}
function locSection() {
  const tops = topLocs();
  const incoming = (ids) => sum(openMovs().filter((m) => ids.includes(m.a_location_id)));
  const shown = tops.filter((l) => { const ids = locIdsOf(l); return sum(stockOfLocs(ids)) || incoming(ids) || l.capacita; });
  const empty = tops.filter((l) => !shown.includes(l));
  const capRow = (l, n) => (l.capacita ? `<span class="track thin ${n > l.capacita ? 'over' : ''}"><i style="width:${Math.min(100, n / l.capacita * 100).toFixed(1)}%"></i></span><small>${n} di ${l.capacita} posti</small>` : l.ruolo ? `<small>${esc(l.ruolo)}</small>` : '');
  return `<section class="card">
    <h2>Location</h2>
    <div class="rows">${shown.map((l) => {
      const ids = locIdsOf(l); const n = sum(stockOfLocs(ids)); const inc = incoming(ids);
      if (isGroup(l)) {
        return `<button class="lrow top" data-act="openloc" data-id="${l.id}"><span class="grow"><b>${esc(l.nome)}</b>${inc ? `<small class="accent">+${inc} in arrivo</small>` : ''}</span><b class="num">${n}</b></button>
          <div class="sub">${children(l.id).map((c) => { const cn = sum(stockOfLoc(c.id)); return `<button class="lrow" data-act="openloc" data-id="${c.id}"><span class="grow"><span class="between"><span>${esc(c.nome)}</span><b>${cn}</b></span>${capRow(c, cn)}</span></button>`; }).join('')}</div>`;
      }
      return `<button class="lrow top" data-act="openloc" data-id="${l.id}"><span class="grow"><b>${esc(l.nome)}</b>${l.capacita ? capRow(l, n) : `<small>${esc(l.ruolo || '')}${inc ? ` <span class="accent">+${inc} in arrivo</span>` : ''}</small>`}</span><b class="num">${n}</b></button>`;
    }).join('')}</div>
    ${empty.length ? `<p class="note">Vuote: ${empty.map((l) => `<button class="link" data-act="openloc" data-id="${l.id}">${esc(l.nome)}</button>`).join(', ')}</p>` : ''}
  </section>`;
}

/* ---------- liste ---------- */
function wineRow(w, qty, extra = '') {
  const k = dwKey(w);
  return `<button class="lrow" data-act="openwine" data-id="${w.id}">${tipDot(w)}<span class="grow"><b>${esc(wineTitle(w))}</b><small>${esc(wineSub(w))}${extra}</small></span>
    <span class="side">${qty != null ? `<b class="num">${qty}</b>` : ''}<span class="chip c-${k}">${DW[k]}</span></span></button>`;
}
function locFilterOptions(selected, { withGroups = true, withUnknown = true } = {}) {
  const out = [];
  topLocs().forEach((l) => {
    if (isGroup(l)) {
      if (withGroups) out.push(`<option value="g${l.id}" ${selected === `g${l.id}` ? 'selected' : ''}>${esc(l.nome)} (tutto)</option>`);
      children(l.id).forEach((c) => out.push(`<option value="${c.id}" ${String(selected) === String(c.id) ? 'selected' : ''}>${withGroups ? '  ' : ''}${esc(locLabel(c))}</option>`));
    } else out.push(`<option value="${l.id}" ${String(selected) === String(l.id) ? 'selected' : ''}>${esc(l.nome)}</option>`);
  });
  const u = unknownLoc();
  if (withUnknown && u) out.push(`<option value="${u.id}" ${String(selected) === String(u.id) ? 'selected' : ''}>${esc(u.nome)}</option>`);
  return out.join('');
}
const filterIds = (v) => (!v ? null : String(v).startsWith('g') ? locIdsOf(L.get(Number(String(v).slice(1)))) : [Number(v)]);
function cercaView() {
  const q = norm(S.q), f = S.filters;
  const ids = filterIds(f.loc);
  let ws = S.wines.map((w) => ({ w, qty: ids ? stockIn(w.id, ids) : totalWine(w.id), open: sum(openMovsOfWine(w.id).filter((m) => m.tipo === 'acquisto')) }));
  if (!f.zero) ws = ws.filter((x) => x.qty > 0 || (!ids && x.open > 0));
  if (q) ws = ws.filter((x) => q.split(/\s+/).every((t) => norm(`${x.w.produttore} ${x.w.vino} ${x.w.annata} ${x.w.denominazione} ${x.w.regione} ${x.w.vitigno} ${x.w.paese}`).includes(t)));
  if (f.tip) ws = ws.filter((x) => tipKey(x.w) === f.tip);
  if (f.dw) ws = ws.filter((x) => dwKey(x.w) === f.dw);
  if (f.reg) ws = ws.filter((x) => (x.w.regione || 'Senza regione') === f.reg);
  ws.sort((a, b) => (f.dw === 'asap' ? (a.w.dw_fine ?? 9999) - (b.w.dw_fine ?? 9999) : 0) || a.w.produttore.localeCompare(b.w.produttore) || String(a.w.annata).localeCompare(String(b.w.annata)));
  const sel = (k, opts, label) => `<select data-filter="${k}" aria-label="${label}"><option value="">${label}</option>${opts.map(([v, t]) => `<option value="${esc(v)}" ${String(f[k]) === String(v) ? 'selected' : ''}>${esc(t)}</option>`).join('')}</select>`;
  return `
    <input class="search" data-bind="q" placeholder="Produttore, vino, annata, zona…" value="${esc(S.q)}">
    <div class="filters">
      ${sel('tip', TIPS.map((x) => [x, x]), 'Tipologia')}
      ${sel('dw', DWORDER.map((k) => [k, DW[k]]), 'Finestra')}
      <select data-filter="loc" aria-label="Location"><option value="">Location</option>${locFilterOptions(f.loc)}</select>
      <label class="chk"><input type="checkbox" data-filter="zero" ${f.zero ? 'checked' : ''}> finiti</label>
    </div>
    ${f.reg ? `<p class="note">Regione: ${esc(f.reg)} · <button class="link" data-act="clearreg">togli</button></p>` : ''}
    <p class="note">${ws.length} vini · ${sum(ws)} bottiglie</p>
    <div class="rows">${ws.slice(0, 300).map((x) => wineRow(x.w, x.qty, x.open && !ids ? ` · <span class="accent">+${x.open} in consegna</span>` : '')).join('')}</div>`;
}
function locView(v) {
  const l = L.get(v.id); if (!l) return '<p class="note">Location non trovata.</p>';
  const q = norm(S.q);
  const block = (loc) => {
    let st = stockOfLoc(loc.id).map((g) => ({ w: W.get(g.wine_id), qty: g.qty })).filter((x) => x.w);
    if (q) st = st.filter((x) => q.split(/\s+/).every((t) => norm(`${x.w.produttore} ${x.w.vino} ${x.w.annata} ${x.w.denominazione} ${x.w.regione}`).includes(t)));
    st.sort((a, b) => a.w.produttore.localeCompare(b.w.produttore) || String(a.w.annata).localeCompare(String(b.w.annata)));
    return st;
  };
  const head = (loc, n) => `<p class="note">${esc(loc.ruolo || '')}${loc.capacita ? ` · ${n}/${loc.capacita} posti` : ''}${loc.note ? ` · ${esc(loc.note)}` : ''}</p>`;
  const incoming = openMovs().filter((m) => locIdsOf(l).includes(m.a_location_id));
  let body;
  if (isGroup(l)) {
    body = children(l.id).map((c) => { const st = block(c); return `<div class="sec">${esc(c.nome)} <span class="muted">${sum(stockOfLoc(c.id))}${c.capacita ? `/${c.capacita}` : ''}</span></div><div class="rows">${st.map((x) => wineRow(x.w, x.qty)).join('') || '<p class="note">Nessuna bottiglia.</p>'}</div>`; }).join('');
  } else {
    const st = block(l);
    body = `${head(l, sum(stockOfLoc(l.id)))}<div class="rows">${st.map((x) => wineRow(x.w, x.qty)).join('') || '<p class="note">Nessuna bottiglia.</p>'}</div>`;
  }
  return `<input class="search" data-bind="q" placeholder="Filtra in ${esc(l.nome)}…" value="${esc(S.q)}">${body}
    ${incoming.length ? `<div class="sec">In arrivo qui</div><div class="rows">${incoming.map(movCard).join('')}</div>` : ''}`;
}

/* ---------- movimenti aperti e verifiche ---------- */
const TIPO_LABEL = { carico_iniziale: 'Carico iniziale', acquisto: 'Acquisto', trasferimento: 'Trasferimento', consumo: 'Uscita', rettifica: 'Rettifica' };
const STATO_LABEL = { pianificato: 'Pianificato', in_corso: 'In corso', confermato: 'Confermato', annullato: 'Annullato' };
function movLine(m) {
  const da = m.da_location_id ? locLabel(L.get(m.da_location_id)) : (m.fornitore || ({ rettifica: 'rettifica', carico_iniziale: 'Excel' }[m.tipo] || ''));
  const a = m.a_location_id ? locLabel(L.get(m.a_location_id)) : (m.tipo === 'consumo' ? (MOTIVI[m.motivo] || 'uscita').toLowerCase() : 'uscita');
  return `${esc(da || '?')} → ${esc(a || '?')}`;
}
function movCard(m) {
  const w = W.get(m.wine_id); if (!w) return '';
  const label = m.tipo === 'acquisto' ? 'In consegna' : m.stato === 'pianificato' ? 'Da fare' : 'In transito';
  return `<div class="mcard">
    <button class="lrow flat" data-act="openwine" data-id="${w.id}">${tipDot(w)}<span class="grow"><b>${m.qty}× ${esc(wineShort(w))}</b>
      <small>${movLine(m)} · <span class="chip ${m.stato === 'pianificato' ? 'c-aspettare' : 'c-asap'}">${label}</span> ${fmtDate(m.data_movimento)}</small>
      ${m.nota && !/^Import Cantina/.test(m.nota) ? `<small>${esc(m.nota)}</small>` : ''}</span></button>
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
  return `
    <div class="sec">In transito <span class="muted">${sum(tr)}</span></div>
    <div class="rows">${tr.map(movCard).join('') || '<p class="note">Nessun trasferimento in corso. Da una scheda vino usa "Sposta" e scegli "In transito".</p>'}</div>
    ${pl.length ? `<div class="sec">Da fare <span class="muted">${sum(pl)}</span></div><div class="rows">${pl.map(movCard).join('')}</div>` : ''}
    <div class="sec">In consegna <span class="muted">${sum(acq)}</span></div>
    ${Object.entries(byF).sort().map(([k, ms]) => `<details class="group"><summary><b>${esc(k)}</b><span class="muted">${sum(ms)} bott. · ${ms.length} righe</span></summary>
      <div class="acts"><button class="primary sm" data-act="confirmgroup" data-key="${esc(k)}">Arrivata tutta la consegna</button></div>
      <div class="rows">${ms.map(movCard).join('')}</div></details>`).join('') || '<p class="note">Nessuna consegna in attesa.</p>'}`;
}
function verificaView() {
  const ss = segnOpen();
  const groups = {};
  ss.forEach((s) => { (groups[s.location_id] ||= []).push(s); });
  const ids = Object.keys(groups).map(Number).sort((a, b) => (L.get(a)?.tipo === 'sconosciuta' ? -1 : 0) - (L.get(b)?.tipo === 'sconosciuta' ? -1 : 0));
  return `<p class="note">Bottiglie di cui non c'è certezza. Risolvile quando le controlli.</p>
    ${ids.map((lid) => `<div class="sec">${esc(locLabel(L.get(lid)) || 'Senza location')} <span class="muted">${groups[lid].length}</span></div>
      <div class="rows">${groups[lid].map((s) => { const w = W.get(s.wine_id); if (!w) return ''; const n = stockAt(w.id, lid);
        return `<div class="mcard"><button class="lrow flat" data-act="openwine" data-id="${w.id}">${tipDot(w)}<span class="grow"><b>${n}× ${esc(wineShort(w))}</b><small>${esc(s.motivo)}</small></span></button>
        <div class="acts"><button class="primary sm" data-act="resolve" data-id="${s.id}">Risolvi</button></div></div>`; }).join('')}</div>`).join('') || '<p class="note">Niente da verificare.</p>'}`;
}

/* ---------- scheda vino ---------- */
function wineView(v) {
  const w = W.get(v.id); if (!w) return '<p class="note">Vino non trovato.</p>';
  const st = stockOfWine(w.id).slice().sort((a, b) => byOrder(L.get(a.location_id), L.get(b.location_id)));
  const om = openMovsOfWine(w.id);
  const hist = S.movs.filter((m) => m.wine_id === w.id && (m.stato === 'confermato' || m.stato === 'annullato'));
  const k = dwKey(w); const pr = avgPrice(w.id);
  const segn = segnOpen().filter((s) => s.wine_id === w.id);
  return `
    <section class="card hero">${tipDot(w)}<div><h2>${esc(w.vino)}</h2><p>${esc(w.produttore)} · <b>${esc(w.annata)}</b></p>
      <small>${esc([w.denominazione, w.regione, w.paese].filter(Boolean).join(' · '))}</small>
      <small>${esc([w.tipologia, w.stile, w.vitigno, w.formato].filter(Boolean).join(' · '))}</small></div></section>
    <section class="card kv">
      <div><span>Finestra</span><b>${w.dw_inizio || '?'} – ${w.dw_fine || '?'}</b> <span class="chip c-${k}">${DW[k]}</span></div>
      ${w.fonte_dw ? `<div><span>Fonte</span><p>${esc(w.fonte_dw)}</p></div>` : ''}
      <div><span>Finestra controllata</span><p>${w.dw_controllata_il ? fmtDate(w.dw_controllata_il) : 'non ancora'}</p></div>
      ${pr != null ? `<div><span>Prezzo medio</span><b>${fmtEur(pr)}</b></div>` : ''}
      ${w.note ? `<div><span>Note</span><p>${esc(w.note)}</p></div>` : ''}
    </section>
    <div class="sec">Dove si trova <span class="muted">${totalWine(w.id)}</span></div>
    <div class="rows">${st.map((g) => `<div class="lrow flat boxed"><span class="grow">${esc(locLabel(L.get(g.location_id)))}</span><b class="num">${g.qty}</b></div>`).join('') || '<p class="note">Nessuna bottiglia in posizione.</p>'}</div>
    ${segn.length ? `<div class="rows">${segn.map((s) => `<div class="mcard"><small>⚠ ${esc(locLabel(L.get(s.location_id)))}: ${esc(s.motivo)}</small><div class="acts"><button class="primary sm" data-act="resolve" data-id="${s.id}">Risolvi</button></div></div>`).join('')}</div>` : ''}
    <div class="actions">
      <button class="primary" data-act="move" data-id="${w.id}" ${st.length ? '' : 'disabled'}>Sposta</button>
      <button class="primary" data-act="out" data-id="${w.id}" ${st.length ? '' : 'disabled'}>Togli</button>
      <button class="secondary" data-act="buy" data-id="${w.id}">Acquisto</button>
      <button class="secondary" data-act="adjust" data-id="${w.id}">Correggi</button>
      <button class="secondary wide" data-act="editwine" data-id="${w.id}">Modifica dati</button>
    </div>
    ${om.length ? `<div class="sec">In corso</div><div class="rows">${om.map(movCard).join('')}</div>` : ''}
    <div class="sec">Storico</div>
    <div class="rows">${hist.map((m) => `<div class="hist ${m.stato === 'annullato' ? 'canc' : ''}"><b>${m.tipo === 'consumo' ? (MOTIVI[m.motivo] || 'Uscita') : TIPO_LABEL[m.tipo]} · ${m.qty}</b><small>${movLine(m)} · ${fmtDate(m.data_conferma || m.data_movimento)}${m.creato_da && m.creato_da !== 'import' ? ` · ${esc(m.creato_da.split('@')[0])}` : ''}${m.stato === 'annullato' ? ' · annullato' : ''}</small>${m.nota && !/^Import Cantina/.test(m.nota) ? `<small>${esc(m.nota)}</small>` : ''}${m.prezzo_unitario != null ? `<small>${fmtEur(Number(m.prezzo_unitario))} cad.</small>` : ''}</div>`).join('') || '<p class="note">Nessun movimento.</p>'}</div>`;
}

/* ---------- Aggiungi, Togli, Altro ---------- */
function aggiungiView() {
  return `<p class="note">Cosa entra in cantina?</p>
    <div class="bigchoices">
      <button class="bigchoice" data-act="buy"><span class="ic wine">${svg('add', 22)}</span><span class="grow"><b>Nuovo acquisto</b><small>Vino nuovo o già in cantina, arrivato o in consegna</small></span><span class="chev">›</span></button>
      <button class="bigchoice" data-act="goarrivi"><span class="ic wine">${svg('truck', 22)}</span><span class="grow"><b>Arrivo di una consegna</b><small>${sum(openMovs().filter((m) => m.tipo === 'acquisto' && m.stato === 'in_corso'))} bottiglie in consegna da confermare</small></span><span class="chev">›</span></button>
      <button class="bigchoice" data-act="found"><span class="ic wine">${svg('found', 22)}</span><span class="grow"><b>Bottiglie trovate</b><small>Erano in cantina ma non registrate</small></span><span class="chev">›</span></button>
    </div>`;
}
function togliView() {
  const q = norm(S.outQ);
  const hg = homeGroup(); const homeIds = hg ? locIdsOf(hg) : [];
  let ws = S.wines.filter((w) => totalWine(w.id) > 0);
  if (q) ws = ws.filter((w) => q.split(/\s+/).every((t) => norm(`${w.produttore} ${w.vino} ${w.annata} ${w.denominazione} ${w.regione}`).includes(t)));
  ws.sort((a, b) => (stockIn(b.id, homeIds) > 0) - (stockIn(a.id, homeIds) > 0) || a.produttore.localeCompare(b.produttore) || String(a.annata).localeCompare(String(b.annata)));
  return `<p class="note">Cosa esce dalla cantina? ${hg ? `Prima i vini a ${esc(hg.nome)}.` : ''}</p>
    <input class="search" data-bind="outQ" placeholder="Cerca il vino…" value="${esc(S.outQ)}">
    <div class="rows">${ws.slice(0, 60).map((w) => `<button class="lrow" data-act="out" data-id="${w.id}">${tipDot(w)}<span class="grow"><b>${esc(wineShort(w))}</b><small>${esc(whereLine(w.id))}</small></span><span class="side"><b class="num">${totalWine(w.id)}</b></span></button>`).join('') || '<p class="note">Nessun vino trovato.</p>'}</div>
    ${ws.length > 60 ? '<p class="note">Scrivi qualcosa per restringere la lista.</p>' : ''}`;
}
function altroView() {
  const nArr = openMovs().length, nVer = segnOpen().length;
  return `<div class="rows">
    <button class="lrow" data-act="goarrivi"><span class="grow"><b>In arrivo</b><small>Consegne, trasferimenti, cose da fare</small></span>${nArr ? `<span class="count">${nArr}</span>` : ''}<span class="chev">›</span></button>
    <button class="lrow" data-act="goverifica"><span class="grow"><b>Da verificare</b><small>Bottiglie da ritrovare o controllare</small></span>${nVer ? `<span class="count">${nVer}</span>` : ''}<span class="chev">›</span></button>
    <div class="lrow flat boxed disabled"><span class="grow"><b>Modalità inventario</b><small>Arriva con lo Sprint 3</small></span></div>
    <button class="lrow" data-act="goluoghi"><span class="grow"><b>Location</b><small>Aggiungi o modifica posti, gruppi e capienze</small></span><span class="chev">›</span></button>
    <button class="lrow" data-act="gomembri"><span class="grow"><b>Persone autorizzate</b><small>${S.membri.length} account</small></span><span class="chev">›</span></button>
    <button class="lrow" data-act="export"><span class="grow"><b>Esporta in Excel</b><small>Giacenze, movimenti e location</small></span></button>
    <button class="lrow" data-act="logout"><span class="grow"><b>Esci</b><small>${esc(S.session?.user?.email)}</small></span></button>
  </div>`;
}
function membriView() {
  const me = norm(S.session?.user?.email);
  return `<p class="note">Solo queste email possono vedere e modificare la cantina. La persona deve poi aprire l'app e scegliere "Crea account" con la stessa email.</p>
    <div class="rows">${S.membri.map((m) => `<div class="lrow flat boxed"><span class="grow"><b>${esc(m.nome || m.email)}</b><small>${esc(m.email)}</small></span>
      ${norm(m.email) !== me ? `<button class="ghost sm" data-act="delmembro" data-email="${esc(m.email)}">Rimuovi</button>` : ''}</div>`).join('')}</div>
    <div class="actions"><button class="primary wide" data-act="addmembro">Aggiungi persona</button></div>`;
}
function luoghiView() {
  const row = (l, sub) => `<button class="lrow ${sub ? 'subrow' : ''}" data-act="editloc" data-id="${l.id}"><span class="grow"><b>${esc(l.nome)}${l.attiva ? '' : ' (nascosta)'}${isGroup(l) ? ' · gruppo' : ''}</b><small>${esc(l.ruolo || '')}${l.capacita != null ? ` · capienza ${l.capacita}` : ''}</small></span><span class="chev">›</span></button>`;
  const tops = S.locs.filter((l) => !l.parent_id).sort(byOrder);
  return `<div class="rows">${tops.map((l) => row(l) + (isGroup(l) ? S.locs.filter((c) => c.parent_id === l.id).sort(byOrder).map((c) => row(c, true)).join('') : '')).join('')}</div>
    <div class="actions"><button class="primary wide" data-act="editloc">Nuova location</button></div>`;
}

/* ================= fogli (form) ================= */
function openSheet(kind, data = {}) { S.sheet = { kind, ...data }; renderSheet(); }
function closeSheet() { S.sheet = null; renderSheet(); }
const optLocs = (list, selected, placeholder = 'Scegli…') => `<option value="">${placeholder}</option>` + list.map((l) => `<option value="${l.id}" ${String(selected) === String(l.id) ? 'selected' : ''}>${esc(locLabel(l))}</option>`).join('');
const destOptions = (selected, excludeId) => `<option value="">Scegli…</option>` + topLocs().map((l) => isGroup(l)
  ? `<optgroup label="${esc(l.nome)}">${children(l.id).filter((c) => c.id !== excludeId).map((c) => `<option value="${c.id}" ${String(selected) === String(c.id) ? 'selected' : ''}>${esc(c.nome)}</option>`).join('')}</optgroup>`
  : (l.id !== excludeId ? `<option value="${l.id}" ${String(selected) === String(l.id) ? 'selected' : ''}>${esc(l.nome)}</option>` : '')).join('');
function defaultFrom(wid) {
  const st = stockOfWine(wid);
  const hg = homeGroup(); const homeIds = hg ? locIdsOf(hg) : [];
  return (st.find((g) => homeIds.includes(g.location_id)) || st[0])?.location_id ?? '';
}
function renderSheet() {
  const el = $('#sheet'); if (!el) return;
  const s = S.sheet;
  if (!s) { el.innerHTML = ''; return; }
  const w = s.wid ? W.get(s.wid) : null;
  const head = (t) => `<div class="sh-h"><b>${t}</b><button class="iconbtn" data-act="closesheet" aria-label="Chiudi">✕</button></div>${w ? `<p class="note">${esc(wineTitle(w))} · ${esc(w.annata)}</p>` : ''}`;
  const fromSelect = (from) => `<label>Da dove<select name="from" data-live>${optLocs(stockOfWine(w.id).map((g) => L.get(g.location_id)).sort(byOrder), from)}</select></label>`;
  let body = '';
  if (s.kind === 'move') {
    const from = s.from ?? defaultFrom(w.id);
    const max = from ? stockAt(w.id, Number(from)) : 0;
    body = `${head('Sposta bottiglie')}<form id="f" class="form">
      ${fromSelect(from)}
      <label>Quante${max ? ` (max ${max})` : ''}<input name="qty" type="number" inputmode="numeric" min="1" max="${max || 1}" value="1" required></label>
      <label>A<select name="to" required>${destOptions(s.to, Number(from))}</select></label>
      <fieldset class="seg-r"><legend>Stato</legend>
        <label><input type="radio" name="stato" value="in_corso" checked><span>In transito<small>confermi all'arrivo</small></span></label>
        <label><input type="radio" name="stato" value="confermato"><span>Già arrivate</span></label>
        <label><input type="radio" name="stato" value="pianificato"><span>Da fare<small>non ancora partite</small></span></label>
      </fieldset>
      <label>Nota<input name="nota" placeholder="facoltativa"></label>
      <button class="primary" type="submit">Registra</button></form>`;
  } else if (s.kind === 'out') {
    const from = s.from ?? defaultFrom(w.id);
    const max = from ? stockAt(w.id, Number(from)) : 0;
    body = `${head('Togli dalla cantina')}<form id="f" class="form">
      ${fromSelect(from)}
      <label>Quante${max ? ` (max ${max})` : ''}<input name="qty" type="number" inputmode="numeric" min="1" max="${max || 1}" value="1" required></label>
      <fieldset class="seg-r inline"><legend>Motivo</legend>
        ${Object.entries(MOTIVI).map(([k, t], i) => `<label><input type="radio" name="motivo" value="${k}" ${i === 0 ? 'checked' : ''}><span>${t}</span></label>`).join('')}
      </fieldset>
      <label>Nota<input name="nota" placeholder="occasione, a chi, impressioni…"></label>
      <button class="primary" type="submit">Registra uscita</button></form>`;
  } else if (s.kind === 'adjust') {
    const at = s.at ?? (stockOfWine(w.id)[0]?.location_id || '');
    const cur = at ? stockAt(w.id, Number(at)) : 0;
    body = `${head('Correggi quantità')}<p class="note">Per allineare il numero a quello che c'è davvero (bottiglie in più o mancanti).</p><form id="f" class="form">
      <label>Location<select name="at" data-live>${optLocs([...physLocs(), unknownLoc()].filter(Boolean), at)}</select></label>
      <label>Bottiglie presenti davvero (ora: ${cur})<input name="real" type="number" inputmode="numeric" min="0" value="${cur}" required></label>
      <label>Motivo<input name="nota" placeholder="es. trovata durante inventario"></label>
      <button class="primary" type="submit">Registra</button></form>`;
  } else if (s.kind === 'confirm') {
    const m = S.movs.find((x) => x.id === s.mid);
    body = `<div class="sh-h"><b>Conferma arrivo</b><button class="iconbtn" data-act="closesheet" aria-label="Chiudi">✕</button></div><p class="note">${esc(wineShort(W.get(m.wine_id)))} · ${movLine(m)}</p><form id="f" class="form">
      <label>Arrivate<input name="qty" type="number" inputmode="numeric" min="1" max="${m.qty}" value="${m.qty}" required></label>
      <p class="note">Se ne arrivano meno di ${m.qty}, le restanti restano in attesa.</p>
      <button class="primary" type="submit">Conferma</button></form>`;
  } else if (s.kind === 'dest') {
    const m = S.movs.find((x) => x.id === s.mid);
    body = `<div class="sh-h"><b>Cambia destinazione</b><button class="iconbtn" data-act="closesheet" aria-label="Chiudi">✕</button></div><p class="note">${esc(wineShort(W.get(m.wine_id)))} · ${m.qty} bott.</p><form id="f" class="form">
      <label>Nuova destinazione<select name="to" required>${destOptions(m.a_location_id, m.da_location_id)}</select></label>
      <button class="primary" type="submit">Salva</button></form>`;
  } else if (s.kind === 'resolve') {
    const sg = S.segn.find((x) => x.id === s.sid);
    const lw = W.get(sg.wine_id); const lid = sg.location_id; const n = stockAt(sg.wine_id, lid);
    const isUnk = L.get(lid)?.tipo === 'sconosciuta';
    body = `<div class="sh-h"><b>Risolvi</b><button class="iconbtn" data-act="closesheet" aria-label="Chiudi">✕</button></div><p class="note">${esc(wineShort(lw))} · ${esc(sg.motivo)}. Registrate: ${n} in ${esc(locLabel(L.get(lid)))}.</p>
      <form id="f" class="form">
      ${isUnk ? `<fieldset class="seg-r"><legend>Cosa ne sai?</legend>
          <label><input type="radio" name="esito" value="trovate" checked><span>Le ho trovate</span></label>
          <label><input type="radio" name="esito" value="transito"><span>Sono in transito</span></label>
          <label><input type="radio" name="esito" value="perse"><span>Non ci sono</span></label></fieldset>
        <label>Quante<input name="qty" type="number" inputmode="numeric" min="1" max="${n}" value="${n}" required></label>
        <label>Dove sono / verso dove<select name="to">${destOptions('')}</select></label>`
      : `<label>Quante ce ne sono davvero in ${esc(L.get(lid)?.nome)}?<input name="real" type="number" inputmode="numeric" min="0" value="${n}" required></label>
         <p class="note">Se sono meno, la differenza esce dalla giacenza come rettifica.</p>`}
      <label>Nota<input name="nota" placeholder="facoltativa"></label>
      <button class="primary" type="submit">Conferma</button></form>`;
  } else if (s.kind === 'buy' || s.kind === 'found') {
    const isFound = s.kind === 'found';
    const mode = isFound ? 'existing' : (s.mode || 'existing');
    const q = norm(s.q || '');
    const res = !w && mode === 'existing' && q ? S.wines.filter((x) => q.split(/\s+/).every((t) => norm(`${x.produttore} ${x.vino} ${x.annata}`).includes(t))).slice(0, 25) : [];
    const pick = s.pick ? W.get(s.pick) : w;
    const bixio = S.locs.find((l) => l.nome === 'Bixio')?.id;
    body = `<div class="sh-h"><b>${isFound ? 'Bottiglie trovate' : 'Acquisto'}</b><button class="iconbtn" data-act="closesheet" aria-label="Chiudi">✕</button></div>
      ${!w && !isFound ? `<div class="segbtn"><button class="${mode === 'existing' ? 'on' : ''}" data-act="buymode" data-mode="existing">Vino già presente</button><button class="${mode === 'new' ? 'on' : ''}" data-act="buymode" data-mode="new">Vino nuovo</button></div>` : ''}
      ${!w && mode === 'existing' && !pick ? `<input class="search" id="buyq" placeholder="Cerca produttore, vino, annata…" value="${esc(s.q || '')}">
        <div class="rows">${res.map((x) => `<button class="lrow" data-act="buypick" data-id="${x.id}">${tipDot(x)}<span class="grow"><b>${esc(wineTitle(x))}</b><small>${esc(wineSub(x))}</small></span></button>`).join('')}</div>
        ${isFound ? '<p class="note">Se il vino non c\'è, usa Nuovo acquisto e scegli "Già arrivato".</p>' : ''}` : ''}
      ${(pick || mode === 'new') ? `<form id="f" class="form">
        ${pick ? `<p><b>${esc(wineTitle(pick))}</b><br><small class="muted">${esc(wineSub(pick))}</small></p>` : wineFields({})}
        ${isFound ? `<label>Bottiglie<input name="qty" type="number" inputmode="numeric" min="1" value="1" required></label>
          <label>Dove le hai trovate<select name="to" required>${destOptions('')}</select></label>
          <label>Nota<input name="nota" placeholder="facoltativa"></label>
          <button class="primary" type="submit">Registra</button>`
        : `<div class="grid2"><label>Bottiglie<input name="qty" type="number" inputmode="numeric" min="1" value="1" required></label>
          <label>Prezzo cad. €<input name="prezzo" type="number" inputmode="decimal" step="0.01" min="0"></label></div>
          <label>Fornitore<input name="fornitore" list="fornitori"><datalist id="fornitori">${[...new Set(S.movs.map((m) => m.fornitore).filter(Boolean))].map((f) => `<option value="${esc(f)}">`).join('')}</datalist></label>
          <label>Destinazione<select name="to" required>${destOptions(bixio)}</select></label>
          <fieldset class="seg-r"><legend>Stato</legend>
            <label><input type="radio" name="stato" value="in_corso" checked><span>In consegna</span></label>
            <label><input type="radio" name="stato" value="confermato"><span>Già arrivato</span></label></fieldset>
          <label>Nota<input name="nota" placeholder="ordine, fattura…"></label>
          <button class="primary" type="submit">Registra acquisto</button>`}
        </form>` : ''}`;
  } else if (s.kind === 'editwine') {
    body = `${head('Modifica vino')}<form id="f" class="form">${wineFields(w)}
      <label class="chk"><input type="checkbox" name="bere_presto" ${w.bere_presto ? 'checked' : ''}> Da bere asap (anche se la finestra dice altro)</label>
      <label>Note<textarea name="note" rows="2">${esc(w.note || '')}</textarea></label>
      <button class="primary" type="submit">Salva</button></form>`;
  } else if (s.kind === 'membro') {
    body = `<div class="sh-h"><b>Aggiungi persona</b><button class="iconbtn" data-act="closesheet" aria-label="Chiudi">✕</button></div><form id="f" class="form">
      <label>Email<input name="email" type="email" required></label><label>Nome<input name="nome"></label>
      <button class="primary" type="submit">Aggiungi</button></form>`;
  } else if (s.kind === 'loc') {
    const l = s.lid ? L.get(s.lid) : {};
    const groups = S.locs.filter((x) => isGroup(x) && x.id !== l.id);
    const canParent = !isGroup(l) && l.tipo !== 'sconosciuta';
    body = `<div class="sh-h"><b>${s.lid ? 'Modifica location' : 'Nuova location'}</b><button class="iconbtn" data-act="closesheet" aria-label="Chiudi">✕</button></div><form id="f" class="form">
      <label>Nome<input name="nome" value="${esc(l.nome || '')}" required ${l.tipo === 'sconosciuta' ? 'readonly' : ''}></label>
      ${!s.lid ? `<label class="chk"><input type="checkbox" name="gruppo"> È un gruppo (contiene altre location)</label>` : ''}
      ${canParent ? `<label>Dentro il gruppo<select name="parent_id"><option value="">Nessuno</option>${groups.map((g) => `<option value="${g.id}" ${l.parent_id === g.id ? 'selected' : ''}>${esc(g.nome)}</option>`).join('')}</select></label>` : ''}
      <label>Ruolo<input name="ruolo" value="${esc(l.ruolo || '')}" placeholder="es. Casa, Stoccaggio"></label>
      <div class="grid2"><label>Capienza<input name="capacita" type="number" min="0" value="${l.capacita ?? ''}"></label>
      <label>Ordine<input name="ordine" type="number" value="${l.ordine ?? ''}"></label></div>
      <label>Note<input name="note" value="${esc(l.note || '')}"></label>
      ${l.tipo !== 'sconosciuta' ? `<label class="chk"><input type="checkbox" name="attiva" ${l.attiva !== false ? 'checked' : ''}> Visibile</label>` : ''}
      <button class="primary" type="submit">Salva</button></form>`;
  }
  el.innerHTML = `<div class="backdrop" data-act="closesheet"></div><div class="sheet" role="dialog" aria-modal="true">${body}</div>`;
  const bq = $('#buyq'); if (bq) { bq.focus(); bq.setSelectionRange(bq.value.length, bq.value.length); }
}
function wineFields(w) {
  const f = (name, label, attrs = '') => `<label>${label}<input name="${name}" value="${esc(w[name] ?? '')}" ${attrs}></label>`;
  return `${f('produttore', 'Produttore', 'required')}${f('vino', 'Vino', 'required')}
    <div class="grid2">${f('annata', 'Annata', 'required placeholder="2019 o NV" pattern="[0-9]{4}|NV|nv"')}${f('formato', 'Formato', 'placeholder="750ml"')}</div>
    <div class="grid2"><label>Tipologia<select name="tipologia">${['', 'Rosso', 'Bianco', 'Rosé'].map((t) => `<option ${w.tipologia === t ? 'selected' : ''}>${t}</option>`).join('')}</select></label>
    <label>Stile<select name="stile">${['', 'Fermo secco', 'Bollicine', 'Fermo abboccato', 'Dolce'].map((t) => `<option ${w.stile === t ? 'selected' : ''}>${t}</option>`).join('')}</select></label></div>
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

/* ================= azioni ================= */
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
  } else if (s.kind === 'out') {
    const from = num('from'); if (!from) return toast('Scegli da dove');
    const motivo = fd.get('motivo') || 'bevuta';
    ok = await run(sb.from('movimenti').insert({ wine_id: s.wid, tipo: 'consumo', motivo, qty: num('qty'), da_location_id: from, nota }), motivo === 'bevuta' ? 'Alla salute' : 'Uscita registrata');
  } else if (s.kind === 'adjust') {
    const at = num('at'), real = num('real'); const cur = stockAt(s.wid, at); const d = real - cur;
    if (!at) return toast('Scegli la location');
    if (d === 0) { closeSheet(); return toast('Nessuna differenza'); }
    ok = await run(sb.from('movimenti').insert(d > 0 ? { wine_id: s.wid, tipo: 'rettifica', qty: d, a_location_id: at, nota } : { wine_id: s.wid, tipo: 'rettifica', qty: -d, da_location_id: at, nota }), 'Quantità corretta');
  } else if (s.kind === 'confirm') {
    const m = S.movs.find((x) => x.id === s.mid); const qty = num('qty');
    if (qty < m.qty) {
      const { id, created_at, updated_at, data_conferma, ...rest } = m;
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
  } else if (s.kind === 'found') {
    const to = num('to'); if (!to) return toast('Scegli dove le hai trovate');
    ok = await run(sb.from('movimenti').insert({ wine_id: s.pick, tipo: 'rettifica', qty: num('qty'), a_location_id: to, nota: nota || 'Bottiglie trovate' }), 'Registrate');
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
    const w0 = W.get(s.wid);
    if (upd.dw_inizio !== w0.dw_inizio || upd.dw_fine !== w0.dw_fine || upd.fonte_dw !== w0.fonte_dw) upd.dw_controllata_il = new Date().toISOString().slice(0, 10);
    ok = await run(sb.from('wines').update(upd).eq('id', s.wid), 'Salvato');
  } else if (s.kind === 'membro') {
    ok = await run(sb.from('membri').insert({ email: fd.get('email').toString().trim().toLowerCase(), nome: (fd.get('nome') || '').toString().trim() || null }), 'Persona aggiunta');
  } else if (s.kind === 'loc') {
    const g = (k) => { const v = (fd.get(k) ?? '').toString().trim(); return v === '' ? null : v; };
    const cur = s.lid ? L.get(s.lid) : null;
    const row = { nome: g('nome'), ruolo: g('ruolo'), capacita: g('capacita') == null ? null : Number(g('capacita')), ordine: g('ordine') == null ? null : Number(g('ordine')), note: g('note') };
    if (!cur) row.tipo = fd.get('gruppo') === 'on' ? 'gruppo' : 'fisica';
    if (form.querySelector('[name=parent_id]')) row.parent_id = g('parent_id') == null ? null : Number(g('parent_id'));
    if (cur?.tipo !== 'sconosciuta') row.attiva = fd.get('attiva') === 'on';
    ok = await run(cur ? sb.from('locations').update(row).eq('id', s.lid) : sb.from('locations').insert(row), 'Salvato');
  }
  if (ok) { closeSheet(); await loadAll(); }
}

async function exportXlsx() {
  const wb = XLSX.utils.book_new();
  const head = ['Quantità', 'Vino', 'Produttore', 'Annata', 'Denominazione', 'Regione', 'Paese', 'Tipologia', 'Stile', 'Vitigno', 'Formato', 'Prezzo medio (€)', 'Gruppo', 'Location', 'Stato', 'Destinazione', 'Fornitore', 'Inizio DW', 'Fine DW', 'Stato DW', 'Fonte DW', 'DW controllata il'];
  const rows = [];
  const segnKey = new Set(segnOpen().map((s) => `${s.wine_id}|${s.location_id}`));
  const base = (w) => { const p = avgPrice(w.id); return [w.vino, w.produttore, w.annata, w.denominazione, w.regione, w.paese, w.tipologia, w.stile, w.vitigno, w.formato, p == null ? null : Math.round(p * 100) / 100]; };
  const tail = (w) => [w.dw_inizio, w.dw_fine, DW[dwKey(w)], w.fonte_dw, w.dw_controllata_il];
  const grp = (l) => (l?.parent_id ? L.get(l.parent_id)?.nome : null);
  S.giac.filter((g) => g.qty > 0).forEach((g) => { const w = W.get(g.wine_id); const l = L.get(g.location_id);
    rows.push([g.qty, ...base(w), grp(l), l?.nome, l?.tipo === 'sconosciuta' || segnKey.has(`${g.wine_id}|${g.location_id}`) ? 'Da verificare' : 'In posizione', null, null, ...tail(w)]); });
  // i pianificati sono ancora nell'origine: già contati sopra, li trovi nel foglio Movimenti
  openMovs().filter((m) => m.stato === 'in_corso').forEach((m) => { const w = W.get(m.wine_id); const l = L.get(m.da_location_id);
    rows.push([m.qty, ...base(w), grp(l), l?.nome ?? null, m.tipo === 'acquisto' ? 'In consegna' : 'In transito', locLabel(L.get(m.a_location_id)), m.fornitore, ...tail(w)]); });
  rows.sort((a, b) => String(a[2]).localeCompare(String(b[2])) || String(a[3]).localeCompare(String(b[3])));
  const ws1 = XLSX.utils.aoa_to_sheet([head, ...rows]); ws1['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: head.length - 1 } }) };
  XLSX.utils.book_append_sheet(wb, ws1, 'Cantina');
  const mh = ['Data', 'Tipo', 'Motivo', 'Stato', 'Quantità', 'Produttore', 'Vino', 'Annata', 'Da', 'A', 'Fornitore', 'Prezzo cad. (€)', 'Nota', 'Registrato da', 'Confermato il'];
  const mr = S.movs.map((m) => { const w = W.get(m.wine_id); return [fmtDate(m.data_movimento), TIPO_LABEL[m.tipo], MOTIVI[m.motivo] || null, STATO_LABEL[m.stato], m.qty, w?.produttore, w?.vino, w?.annata, locLabel(L.get(m.da_location_id)) || null, locLabel(L.get(m.a_location_id)) || null, m.fornitore, m.prezzo_unitario == null ? null : Number(m.prezzo_unitario), m.nota, m.creato_da, fmtDate(m.data_conferma)]; });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([mh, ...mr]), 'Movimenti');
  const lh = ['Location', 'Gruppo', 'Ruolo', 'Capienza', 'Occupazione', 'In arrivo', 'Note'];
  const lr = S.locs.filter((l) => !isGroup(l)).map((l) => [l.nome, grp(l), l.ruolo, l.capacita, sum(stockOfLoc(l.id)), sum(openMovs().filter((m) => m.a_location_id === l.id)), l.note]);
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([lh, ...lr]), 'Location');
  XLSX.writeFile(wb, `Cantina_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

/* ================= eventi ================= */
let toastT;
function toast(msg) { const t = $('#toast'); if (!t) return; t.textContent = msg; t.className = 'show'; clearTimeout(toastT); toastT = setTimeout(() => (t.className = ''), 3200); }
function cercaWith(f) { S.q = ''; S.filters = { tip: '', dw: '', loc: '', reg: '', zero: false, ...f }; go({ name: 'cerca' }); }

document.addEventListener('click', async (e) => {
  const t = e.target.closest('[data-act]'); if (!t) return;
  const a = t.dataset.act, id = t.dataset.id ? Number(t.dataset.id) : null;
  if (a === 'back') back();
  else if (a === 'home') home();
  else if (a === 'reload') loadAll();
  else if (a === 'goaltro') go({ name: 'altro' });
  else if (a === 'goaggiungi') go({ name: 'aggiungi' });
  else if (a === 'gocerca') { S.q = ''; S.filters = { tip: '', dw: '', loc: '', reg: '', zero: false }; go({ name: 'cerca' }); }
  else if (a === 'gotogli') { S.outQ = ''; go({ name: 'togli' }); }
  else if (a === 'goarrivi') go({ name: 'arrivi' });
  else if (a === 'goverifica') go({ name: 'verifica' });
  else if (a === 'gomembri') go({ name: 'membri' });
  else if (a === 'goluoghi') go({ name: 'luoghi' });
  else if (a === 'openloc') { S.q = ''; go({ name: 'loc', id }); }
  else if (a === 'openwine') go({ name: 'wine', id });
  else if (a === 'cercadw') cercaWith({ dw: t.dataset.dw });
  else if (a === 'cercatip') cercaWith({ tip: t.dataset.tip, loc: t.dataset.loc || '' });
  else if (a === 'comploc') { S.comp.loc = t.dataset.loc; render(); }
  else if (a === 'compdim') { S.comp.dim = t.dataset.dim; render(); }
  else if (a === 'compgo') {
    const k = t.dataset.key; const loc = S.comp.loc === 'all' ? '' : (isGroup(L.get(Number(S.comp.loc))) ? `g${S.comp.loc}` : S.comp.loc);
    cercaWith({ loc, ...(S.comp.dim === 'tip' ? { tip: k } : S.comp.dim === 'fin' ? { dw: k } : { reg: k }) });
  }
  else if (a === 'clearreg') { S.filters.reg = ''; render(); }
  else if (a === 'closesheet') closeSheet();
  else if (a === 'move') openSheet('move', { wid: id });
  else if (a === 'out') openSheet('out', { wid: id });
  else if (a === 'adjust') openSheet('adjust', { wid: id });
  else if (a === 'editwine') openSheet('editwine', { wid: id });
  else if (a === 'buy') openSheet('buy', { wid: id || null, mode: 'existing' });
  else if (a === 'found') openSheet('found', { wid: null });
  else if (a === 'buymode') { S.sheet.mode = t.dataset.mode; S.sheet.pick = null; renderSheet(); }
  else if (a === 'buypick') { S.sheet.pick = id; renderSheet(); }
  else if (a === 'confirmmov') openSheet('confirm', { mid: id });
  else if (a === 'destmov') openSheet('dest', { mid: id });
  else if (a === 'startmov') { if (await run(sb.from('movimenti').update({ stato: 'in_corso' }).eq('id', id), 'In transito')) loadAll(); }
  else if (a === 'cancelmov') { if (confirm('Annullare questo movimento?') && await run(sb.from('movimenti').update({ stato: 'annullato' }).eq('id', id), 'Annullato')) loadAll(); }
  else if (a === 'confirmgroup') {
    const key = t.dataset.key;
    const ms = openMovs().filter((m) => m.tipo === 'acquisto' && m.stato === 'in_corso' && ((m.fornitore || 'Senza fornitore') + (/Primeur/.test(m.nota || '') ? ' · Primeur' : '')) === key);
    if (!confirm(`Confermare l'arrivo di ${sum(ms)} bottiglie (${key})?`)) return;
    if (await run(sb.from('movimenti').update({ stato: 'confermato' }).in('id', ms.map((m) => m.id)), 'Consegna arrivata')) loadAll();
  }
  else if (a === 'resolve') openSheet('resolve', { sid: id });
  else if (a === 'export') { try { await exportXlsx(); } catch (err) { toast('Export non riuscito: ' + err.message); } }
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
        toast(/Invalid login/.test(m) ? 'Email o password errate' : /already registered/.test(m) ? 'Account già esistente: usa "Accedi"' : m);
      }
    } else await submitSheet(form);
  } finally { btns.forEach((b) => (b.disabled = false)); }
});
document.addEventListener('input', (e) => {
  const t = e.target;
  const key = t.dataset.bind;
  if (key === 'q' || key === 'outQ') { S[key] = t.value; const pos = t.selectionStart; render(); const n = document.querySelector(`[data-bind="${key}"]`); if (n) { n.focus(); n.setSelectionRange(pos, pos); } }
  if (t.id === 'buyq') { S.sheet.q = t.value; renderSheet(); }
});
document.addEventListener('change', (e) => {
  const t = e.target;
  if (t.dataset.filter) { S.filters[t.dataset.filter] = t.type === 'checkbox' ? t.checked : t.value; render(); }
  if (t.hasAttribute('data-live') && S.sheet) { S.sheet[t.name] = t.value; renderSheet(); }
});
document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && S.session && S.member) loadAll(); });

/* ================= avvio ================= */
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
