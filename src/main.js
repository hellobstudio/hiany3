import './style.css';
import * as XLSX from 'xlsx';
import { Capacitor } from '@capacitor/core';
import { Filesystem, Directory } from '@capacitor/filesystem';
import { Share } from '@capacitor/share';

/* ---------- Adattár (localStorage, offline) ---------- */
const KEY = 'hianyzas_v1';
const def = { students: [], reasons: ['Orvosi vizsgálat', 'Családi ok', 'Verseny', 'Zeneiskola', 'Betegség'], reporters: ['Tanár', 'Szülő', 'Diák'], absences: [], theme: '' };
let S;
try { S = { ...def, ...JSON.parse(localStorage.getItem(KEY) || '{}') }; } catch { S = { ...def }; }
const save = () => localStorage.setItem(KEY, JSON.stringify(S));

/* ---------- Segédek ---------- */
const $ = (s, r = document) => r.querySelector(s);
const esc = t => String(t ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const today = () => new Date(Date.now() - new Date().getTimezoneOffset() * 60000).toISOString().slice(0, 10);
const fmt = d => (d ? String(d).replace(/-/g, '.') + '.' : '');
const dur = a => (a.full ? 'Egész nap' : `${a.l1}–${a.l2}. óra`);
const rng = a => (a.from === a.to ? fmt(a.from) : `${fmt(a.from)} – ${fmt(a.to)}`);
let tmr;
function toast(m) { const t = $('#toast'); t.textContent = m; t.classList.add('show'); clearTimeout(tmr); tmr = setTimeout(() => t.classList.remove('show'), 2400); }
function applyTheme() {
  const t = S.theme || (matchMedia('(prefers-color-scheme:dark)').matches ? 'dark' : 'light');
  document.documentElement.dataset.theme = t;
  $('#th').textContent = t === 'dark' ? '☀️' : '🌙';
}

/* ---------- Közös űrlaprészlet (rögzítés + szerkesztés) ---------- */
function detHtml(o) {
  const opt = (arr, sel, ph) => `<option value="">${ph}</option>` + arr.map(x => `<option ${x === sel ? 'selected' : ''}>${esc(x)}</option>`).join('');
  const hrs = sel => [1, 2, 3, 4, 5, 6, 7, 8].map(n => `<option ${n == sel ? 'selected' : ''}>${n}</option>`).join('');
  return `<div class="row"><label>Kezdő dátum<input type="date" name="from" value="${o.from}"></label><label>Záró dátum<input type="date" name="to" value="${o.to}"></label></div>
  <div class="seg"><label><input type="radio" name="full" value="1" ${o.full ? 'checked' : ''}><span>Egész tanítási nap</span></label><label><input type="radio" name="full" value="0" ${o.full ? '' : 'checked'}><span>Óra intervallum</span></label></div>
  <div class="row hrs" ${o.full ? 'hidden' : ''}><label>Első óra<select name="l1">${hrs(o.l1)}</select></label><label>Utolsó óra<select name="l2">${hrs(o.l2)}</select></label></div>
  <label>Bejelentő<select name="reporter">${opt(S.reporters, o.reporter, 'Válassz…')}</select></label>
  <label style="margin-top:8px">Ok<select name="reason">${opt(S.reasons, o.reason, 'Válassz…')}</select></label>`;
}
function bindDet(form) { form.addEventListener('change', e => { if (e.target.name === 'full') $('.hrs', form).hidden = e.target.value === '1'; }); }
function readDet(form) {
  const f = new FormData(form);
  return { from: f.get('from'), to: f.get('to'), full: f.get('full') === '1', l1: +f.get('l1'), l2: +f.get('l2'), reporter: f.get('reporter'), reason: f.get('reason') };
}
const check = o => !o.from || !o.to ? 'Add meg a dátumokat.' : o.to < o.from ? 'A záró dátum nem lehet korábbi a kezdőnél.' : (!o.full && o.l1 > o.l2) ? 'Az első óra nem lehet később az utolsónál.' : !o.reporter ? 'Válaszd ki a bejelentőt.' : !o.reason ? 'Válaszd ki az okot.' : '';

/* ---------- 1. Rögzítés ---------- */
const rec = { checked: new Set(), q: '', from: today(), to: today(), full: true, l1: 1, l2: 8, reporter: '', reason: '' };
function vRec() {
  $('#view').innerHTML = `<section class="card"><h2>Hiányzó diákok</h2>
    <input type="search" id="q" placeholder="Keresés név szerint…" value="${esc(rec.q)}">
    <div id="list"></div><p class="muted" id="cnt"></p></section>
    <form class="card" id="rf"><h2>Hiányzás részletei</h2>${detHtml(rec)}<button class="btn">Hiányzás rögzítése</button></form>`;
  const form = $('#rf'); bindDet(form); drawStuds();
  form.onsubmit = e => {
    e.preventDefault();
    if (!rec.checked.size) return toast('Jelölj ki legalább egy diákot.');
    const o = readDet(form), er = check(o); if (er) return toast(er);
    Object.assign(rec, o);
    rec.checked.forEach(id => { const s = S.students.find(x => x.id === id); if (s) S.absences.push({ id: uid(), sid: id, name: s.name, ...o }); });
    const n = rec.checked.size; rec.checked.clear(); save(); toast(`${n} hiányzás rögzítve.`); vRec();
  };
}
function drawStuds() {
  const q = rec.q.toLowerCase();
  const l = S.students.filter(s => s.name.toLowerCase().includes(q)).sort((a, b) => a.name.localeCompare(b.name, 'hu'));
  $('#list').innerHTML = l.map(s => `<label class="chk"><input type="checkbox" data-id="${s.id}" ${rec.checked.has(s.id) ? 'checked' : ''}><span>${esc(s.name)}</span><small>${esc(fmt(s.birth))}</small></label>`).join('')
    || '<p class="muted">Még nincs diák. Importáld őket a Beállítások fülön.</p>';
  $('#cnt').textContent = `Kijelölt hiányzók: ${rec.checked.size}`;
}

/* ---------- 3. Statisztika ---------- */
const st = { mode: 'date', from: '', to: '', q: '' };
function rows() {
  const q = st.q.toLowerCase();
  return S.absences.filter(a => (!st.from || a.to >= st.from) && (!st.to || a.from <= st.to) && (!q || a.name.toLowerCase().includes(q)))
    .sort((a, b) => st.mode === 'date' ? (a.from.localeCompare(b.from) || a.name.localeCompare(b.name, 'hu')) : (a.name.localeCompare(b.name, 'hu') || a.from.localeCompare(b.from)));
}
function vStats() {
  $('#view').innerHTML = `<div class="card"><div class="seg">
    <label><input type="radio" name="m" value="date" ${st.mode === 'date' ? 'checked' : ''}><span>Dátum szerint</span></label>
    <label><input type="radio" name="m" value="student" ${st.mode === 'student' ? 'checked' : ''}><span>Diákok szerint</span></label></div>
    <input type="search" id="sq" placeholder="Diák keresése…" value="${esc(st.q)}">
    <div class="row"><label>Ettől<input type="date" id="sf" value="${st.from}"></label><label>Eddig<input type="date" id="sto" value="${st.to}"></label></div>
    <button class="btn" data-act="export">📤 Megosztás Excelben (e-mail, Drive…)</button></div><div id="slist"></div>`;
  drawList();
}
function ent(a) {
  const d = st.mode === 'date';
  return `<div class="ent"><div><b>${d ? esc(a.name) : rng(a)}</b><small>${d ? rng(a) + ' · ' : ''}${dur(a)} · ${esc(a.reporter)} · ${esc(a.reason)}</small></div>
  <div><button class="ib" data-act="edit" data-id="${a.id}" aria-label="Szerkesztés">✏️</button><button class="ib" data-act="del" data-id="${a.id}" aria-label="Törlés">🗑️</button></div></div>`;
}
function drawList() {
  const r = rows(), g = new Map();
  r.forEach(a => { const k = st.mode === 'date' ? a.from : a.name; if (!g.has(k)) g.set(k, []); g.get(k).push(a); });
  $('#slist').innerHTML = `<p class="muted">Összesen: ${r.length} bejegyzés</p>` + ([...g].map(([k, arr]) => {
    const h = st.mode === 'date' ? `${fmt(k)} <small class="muted">${new Date(k + 'T12:00').toLocaleDateString('hu-HU', { weekday: 'long' })}</small>` : esc(k);
    return `<h3>${h}<span class="badge">${arr.length}</span></h3>` + arr.map(ent).join('');
  }).join('') || '<p class="muted">Nincs megjeleníthető hiányzás.</p>');
}
function edit(id) {
  const a = S.absences.find(x => x.id === id); if (!a) return;
  const d = $('#dlg');
  d.innerHTML = `<form id="ef"><h2>${esc(a.name)}</h2>${detHtml(a)}<div class="row"><button type="button" class="btn2" data-act="close">Mégse</button><button class="btn">Mentés</button></div></form>`;
  const f = $('#ef'); bindDet(f);
  f.onsubmit = e => {
    e.preventDefault(); const o = readDet(f), er = check(o); if (er) return toast(er);
    Object.assign(a, o); save(); d.close(); drawList(); toast('Módosítva.');
  };
  d.showModal();
}
async function exportXlsx(r) {
  const head = ['Diák', 'Születési dátum', 'Anyja neve', 'Kezdő dátum', 'Záró dátum', 'Időtartam', 'Bejelentő', 'Ok'];
  const data = [head, ...r.map(a => { const s = S.students.find(x => x.id === a.sid) || {}; return [a.name, s.birth || '', s.mother || '', a.from, a.to, dur(a), a.reporter, a.reason]; })];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), 'Hiányzások');
  const fn = `hianyzasok_${today()}.xlsx`;
  if (Capacitor.isNativePlatform()) {
    const res = await Filesystem.writeFile({ path: fn, data: XLSX.write(wb, { type: 'base64', bookType: 'xlsx' }), directory: Directory.Cache });
    await Share.share({ title: 'Hiányzások', text: 'Hiányzási statisztika', files: [res.uri], dialogTitle: 'Megosztás' });
  } else XLSX.writeFile(wb, fn);
}

/* ---------- 2. Beállítások ---------- */
function chips(k, title, ph) {
  return `<section class="card"><h2>${title}</h2><div class="chips">${S[k].map((x, i) => `<span class="chip">${esc(x)}<button data-act="rmlist" data-k="${k}" data-i="${i}" aria-label="Törlés">✕</button></span>`).join('')}</div>
  <div class="add"><input id="n_${k}" placeholder="${ph}"><button class="btn" data-act="addlist" data-k="${k}">Hozzáad</button></div></section>`;
}
function vSet() {
  $('#view').innerHTML = chips('reasons', 'Hiányzás okai', 'Új ok…') + chips('reporters', 'Bejelentők', 'Új bejelentő…') + `
  <section class="card"><h2>Diákok importálása Excelből</h2>
    <p class="muted">Az első munkalap első sora a fejléc. Oszlopok: <b>Név</b>, <b>Születési dátum</b>, <b>Anyja neve</b>. A már meglévő diákokat (név + születési dátum) kihagyja.</p>
    <input type="file" id="xf" accept=".xlsx,.xls"></section>
  <section class="card"><h2>Diák hozzáadása kézzel</h2>
    <input id="sn" placeholder="Név"><input id="sb" type="date" style="margin-top:8px"><input id="sm" placeholder="Anyja neve" style="margin-top:8px">
    <button class="btn" data-act="addstud">Diák mentése</button></section>
  <details class="card"><summary>Diákok (${S.students.length})</summary>
    ${[...S.students].sort((a, b) => a.name.localeCompare(b.name, 'hu')).map(s => `<div class="ent" style="margin-top:6px"><div><b>${esc(s.name)}</b><small>${esc(fmt(s.birth))} · ${esc(s.mother)}</small></div><button class="ib" data-act="rmstud" data-id="${s.id}" aria-label="Törlés">🗑️</button></div>`).join('')}
    ${S.students.length ? '<button class="btn2" data-act="clrstud" style="margin-top:10px">Összes diák törlése</button>' : ''}</details>`;
}
const isoDate = v => {
  if (v instanceof Date) { const t = new Date(v.getTime() + 43200000); return `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, '0')}-${String(t.getDate()).padStart(2, '0')}`; }
  const m = String(v).match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  return m ? `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}` : String(v).trim();
};
async function importXlsx(file) {
  try {
    const wb = XLSX.read(await file.arrayBuffer(), { cellDates: true });
    const r = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: '' });
    const h = (r[0] || []).map(c => String(c).toLowerCase());
    let im = h.findIndex(x => /any/.test(x)), ib = h.findIndex(x => /sz[üu]let|d[áa]tum/.test(x)), iN = h.findIndex((x, i) => /n[ée]v/.test(x) && i !== im);
    if (im < 0 || ib < 0 || iN < 0) { iN = 0; ib = 1; im = 2; }
    let add = 0, dup = 0;
    r.slice(1).forEach(x => {
      const name = String(x[iN]).trim(); if (!name) return;
      const birth = isoDate(x[ib]), mother = String(x[im]).trim();
      if (S.students.some(s => s.name.toLowerCase() === name.toLowerCase() && s.birth === birth)) { dup++; return; }
      S.students.push({ id: uid(), name, birth, mother }); add++;
    });
    save(); toast(`${add} diák importálva, ${dup} már létezett.`); vSet();
  } catch (err) { toast('Import hiba: ' + err.message); }
}

/* ---------- Események ---------- */
let tab = 'rec';
function render() {
  document.querySelectorAll('nav button').forEach(b => b.classList.toggle('on', b.dataset.id === tab));
  ({ rec: vRec, st: vStats, set: vSet })[tab](); scrollTo(0, 0);
}
document.addEventListener('click', async e => {
  const b = e.target.closest('[data-act]'); if (!b) return;
  const { act, id, k, i } = b.dataset;
  if (act === 'tab') { tab = id; render(); }
  else if (act === 'theme') { S.theme = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark'; save(); applyTheme(); }
  else if (act === 'edit') edit(id);
  else if (act === 'close') $('#dlg').close();
  else if (act === 'del') { if (confirm('Biztosan törlöd ezt a bejegyzést?')) { S.absences = S.absences.filter(a => a.id !== id); save(); drawList(); toast('Törölve.'); } }
  else if (act === 'export') { const r = rows(); if (!r.length) return toast('Nincs exportálható adat.'); try { await exportXlsx(r); } catch (err) { toast('Megosztási hiba: ' + err.message); } }
  else if (act === 'addlist') { const v = $('#n_' + k).value.trim(); if (!v) return; if (!S[k].includes(v)) S[k].push(v); save(); vSet(); }
  else if (act === 'rmlist') { S[k].splice(+i, 1); save(); vSet(); }
  else if (act === 'addstud') { const name = $('#sn').value.trim(); if (!name) return toast('Add meg a nevet.'); S.students.push({ id: uid(), name, birth: $('#sb').value, mother: $('#sm').value.trim() }); save(); toast('Diák mentve.'); vSet(); }
  else if (act === 'rmstud') { if (confirm('Törlöd a diákot? A korábbi hiányzásai megmaradnak.')) { S.students = S.students.filter(s => s.id !== id); rec.checked.delete(id); save(); vSet(); } }
  else if (act === 'clrstud') { if (confirm('Biztosan törlöd az összes diákot?')) { S.students = []; rec.checked.clear(); save(); vSet(); } }
});
document.addEventListener('input', e => {
  const t = e.target;
  if (t.id === 'q') { rec.q = t.value; drawStuds(); }
  else if (t.matches('#list input')) { t.checked ? rec.checked.add(t.dataset.id) : rec.checked.delete(t.dataset.id); $('#cnt').textContent = `Kijelölt hiányzók: ${rec.checked.size}`; }
  else if (t.name === 'm') { st.mode = t.value; drawList(); }
  else if (t.id === 'sq') { st.q = t.value; drawList(); }
  else if (t.id === 'sf') { st.from = t.value; drawList(); }
  else if (t.id === 'sto') { st.to = t.value; drawList(); }
  else if (t.id === 'xf' && t.files[0]) { importXlsx(t.files[0]); t.value = ''; }
});
applyTheme(); render();
