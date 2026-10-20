/* Interfaz de la herramienta SIRCE. Todo ocurre en el navegador. */
(function () {
  'use strict';
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const lsGet = (k, d) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } };
  const lsSet = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* sin almacenamiento */ } };

  const S = {
    reportes: [], reporteRows: new Map(), catalogo: null, clavesBase: null, claves: null, cursos: null, previas: [], info: null,
    ov: Object.assign(Core.emptyOv(), lsGet('sirce.ov', {})), ex: Object.assign({ cursos: [], empresas: [] }, lsGet('sirce.ex', {})),
    res: null, limit: {},
  };
  const saveOv = () => lsSet('sirce.ov', S.ov);
  const saveEx = () => lsSet('sirce.ex', S.ex);

  /* ---------- lectura de archivos ---------- */
  async function loadWith(file, parser, prefer) {
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const names = [...wb.SheetNames];
    if (prefer) names.sort((a, b) => (Core.norm(b) === prefer) - (Core.norm(a) === prefer));
    let lastErr = '';
    for (const n of names) {
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' });
      const r = parser(aoa);
      if (r && r.error) { lastErr = r.error; continue; }
      if (r) return { r, sheet: n };
    }
    return { error: lastErr || 'No se reconoció el formato del archivo (revisa que tenga los encabezados esperados).' };
  }
  function cloneClaves(c) {
    const cp = (x) => ({ map: new Map(x.map), names: new Map(x.names) });
    return { estados: cp(c.estados), municipios: cp(c.municipios), puestos: cp(c.puestos), sucursales: cp(c.sucursales), niveles: new Map(c.niveles) };
  }
  function rebuildClaves() {
    if (!S.clavesBase) { S.claves = null; return; }
    S.claves = cloneClaves(S.clavesBase);
    if (S.info && S.info.data) Core.mergeInfoDC3(S.claves, S.info.data);
    for (const p of S.previas) Core.mergePrevias(S.claves, p.data);
  }

  const handlers = {
    async reportes(files) {
      for (const f of files) {
        const { r, sheet, error } = await loadWith(f, Core.parseReporte);
        const entry = { name: f.name, error };
        if (r) { r.rows.forEach((x) => { x.src = f.name; }); S.reporteRows.set(f.name, r.rows); Object.assign(entry, { sheet, count: r.rows.length }); }
        S.reportes = S.reportes.filter((x) => x.name !== f.name).concat(entry);
        if (!r) S.reporteRows.delete(f.name);
      }
    },
    async catalogo(files) {
      const { r, sheet, error } = await loadWith(files[0], Core.parseCatalogo);
      S.catalogo = r ? { name: files[0].name, sheet, map: r.map, dups: r.dups } : { name: files[0].name, error };
    },
    async claves(files) {
      const { r, sheet, error } = await loadWith(files[0], Core.parseClaves, 'CLAVES');
      if (r) { S.clavesBase = r; S.clavesInfo = { name: files[0].name, sheet }; } else { S.clavesBase = null; S.clavesInfo = { name: files[0].name, error }; }
      rebuildClaves();
    },
    async cursos(files) {
      const { r, sheet, error } = await loadWith(files[0], Core.parseClavesCursos);
      S.cursos = r ? { name: files[0].name, sheet, map: r.map, conflictos: r.conflictos } : { name: files[0].name, error };
    },
    async info(files) {
      const { r, sheet, error } = await loadWith(files[0], Core.parseInfoDC3);
      S.info = { name: files[0].name, sheet, data: r, error };
      rebuildClaves();
    },
    async previas(files) {
      for (const f of files) {
        const { r, sheet, error } = await loadWith(f, Core.parseCargaPrevia);
        S.previas = S.previas.filter((x) => x.name !== f.name).concat({ name: f.name, sheet, data: r, error });
      }
      rebuildClaves();
    },
  };

  function wireDrops() {
    $$('.drop').forEach((box) => {
      const kind = box.dataset.kind; const input = $('input', box);
      const go = async (files) => {
        if (!files.length) return;
        const st = $('#st-' + kind); st.innerHTML = '<span class="muted">Leyendo archivos…</span>';
        try { await handlers[kind](files); } catch (e) { st.innerHTML = `<div class="msg bad">${esc(e.message || e)}</div>`; return; }
        recalc();
      };
      input.addEventListener('change', () => { go([...input.files]); input.value = ''; });
      ['dragover', 'dragenter'].forEach((ev) => box.addEventListener(ev, (e) => { e.preventDefault(); box.classList.add('over'); }));
      ['dragleave', 'drop'].forEach((ev) => box.addEventListener(ev, (e) => { e.preventDefault(); box.classList.remove('over'); }));
      box.addEventListener('drop', (e) => go([...e.dataTransfer.files]));
    });
  }

  /* ---------- cálculo ---------- */
  const allRows = () => [...S.reporteRows.values()].flat();
  function recalc() {
    const ready = S.reportes.some((x) => x.count) && S.catalogo && S.catalogo.map && S.claves;
    S.res = ready ? Core.build({
      reportRows: allRows(), catalogo: S.catalogo.map, claves: S.claves,
      clavesCursos: S.cursos && S.cursos.map ? S.cursos : { map: new Map() },
      ov: S.ov, excluirCursos: S.ex.cursos, excluirEmpresas: S.ex.empresas,
    }) : null;
    renderAll();
  }

  /* ---------- pestañas ---------- */
  function showTab(t) {
    $$('#nav button').forEach((b) => b.classList.toggle('on', b.dataset.tab === t));
    $$('section.tab').forEach((s) => s.classList.toggle('on', s.id === 'tab-' + t));
  }
  $('#nav').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) showTab(b.dataset.tab); });

  function renderAll() { renderStatus(); renderRevision(); renderEquiv(); renderGenerar(); renderDC3(); }

  /* ---------- estado de archivos ---------- */
  const tag = (t, c) => `<span class="tag ${c || ''}">${esc(t)}</span>`;
  const fileLine = (name, right) => `<div class="f"><span>${esc(name)}</span><span>${right}</span></div>`;
  function renderStatus() {
    $('#st-reportes').innerHTML = S.reportes.map((x) => fileLine(x.name, x.error ? tag(x.error, 'bad') : tag(`${x.count} filas · hoja ${x.sheet}`, 'ok'))).join('');
    const c = S.catalogo;
    $('#st-catalogo').innerHTML = c ? fileLine(c.name, c.error ? tag(c.error, 'bad') : tag(`${c.map.size} colaboradores${c.dups ? ' · ' + c.dups + ' IDs repetidos' : ''}`, 'ok')) : '';
    const k = S.clavesInfo;
    $('#st-claves').innerHTML = k ? fileLine(k.name, k.error ? tag(k.error, 'bad') : tag(`hoja ${k.sheet} · ${S.claves.puestos.map.size} puestos · ${S.claves.sucursales.map.size} sucursales`, 'ok')) : '';
    const cu = S.cursos;
    $('#st-cursos').innerHTML = cu ? fileLine(cu.name, cu.error ? tag(cu.error, 'bad') : tag(`${cu.map.size} claves de curso`, 'ok')) +
      (cu.conflictos && cu.conflictos.length ? `<div class="msg warn">${cu.conflictos.length} curso(s) con dos claves distintas para el mismo nombre, fecha y empresa; se usa la primera: ${esc(cu.conflictos.map((x) => x.nombre + ' ' + x.fecha + ' (' + x.claves.join(' / ') + ')').join('; '))}</div>` : '') : '';
    const inf = S.info;
    $('#st-info').innerHTML = inf ? fileLine(inf.name, inf.data ? tag(`${inf.data.map.size} sucursales · ${inf.data.razones.size} empresas`, 'ok') : tag(inf.error || 'No se reconoció el formato', 'bad')) : '';
    $('#st-previas').innerHTML = S.previas.map((x) => fileLine(x.name, x.data ? tag(`${x.data.sucursales.size} sucursales · ${x.data.puestos.size} puestos`, 'ok') : tag(x.error || 'Sin formato de layout', 'bad'))).join('');
    const m = $('#readyMsg');
    if (S.res) { m.hidden = false; m.className = 'msg ok'; m.textContent = 'Archivos listos. Continúa en “2 · Revisión”.' + (S.cursos && S.cursos.map ? '' : ' Falta el archivo de claves de cursos: todas las filas aparecerán sin clave de curso.'); }
    else { m.hidden = false; m.className = 'msg warn'; m.textContent = 'Para continuar carga al menos: reportes, catálogo de colaboradores y catálogo de claves.'; }
  }

  /* ---------- revisión ---------- */
  const TITLES = {
    PUESTO: 'Puestos sin clave de ocupación', NIVEL: 'Escolaridades sin equivalencia', SUCURSAL: 'Sucursales sin clave de establecimiento',
    CURSO: 'Cursos sin clave de curso', ESTADO: 'Estados sin clave', MUN: 'Estados sin municipio principal',
    FECHA: 'Fechas dudosas (revisar)', HORAS: 'Cursos con horas inválidas', CURP: 'CURP inválidas o vacías',
  };
  const ORDER = ['FECHA', 'CURSO', 'SUCURSAL', 'PUESTO', 'NIVEL', 'ESTADO', 'MUN', 'HORAS', 'CURP'];

  function renderRevision() {
    const R = S.res;
    const st = $('#stats'), fl = $('#filtros'), er = $('#errores');
    if (!R) { st.innerHTML = ''; fl.innerHTML = '<p class="muted">Carga los archivos en el paso 1 para ver la revisión.</p>'; er.innerHTML = ''; $('#badgeErr').hidden = true; return; }
    const ok = R.records.filter((x) => !x.errors.length).length;
    const bad = R.records.length - ok;
    const corr = R.records.filter((x) => x.fechaCorregida).length;
    const warns = R.records.filter((x) => x.warnings.length).length;
    st.innerHTML = [
      ['ok', ok, 'Filas listas para cargar'], ['bad', bad, 'Filas con errores'], ['', R.bajas.length, 'Bajas ignoradas (no están en el catálogo)'],
      ['', corr, 'Fechas corregidas (día/mes invertidos)'], ['', R.duplicados, 'Duplicados omitidos'], ['', R.excluidos, 'Excluidas por filtro'], ['', warns, 'Filas con advertencias'],
    ].map(([c, n, l]) => `<div class="stat ${c}"><b>${n}</b><span>${l}</span></div>`).join('');
    const badge = $('#badgeErr'); badge.hidden = !bad; badge.textContent = bad;

    renderFiltros(fl);

    const groups = Core.groupErrors(R.records);
    let html = '';
    if (!groups.length) html += '<div class="msg ok">No hay errores pendientes. Puedes generar el archivo en el paso 4.</div>';
    for (const code of ORDER) {
      const gs = groups.filter((g) => g.code === code);
      if (!gs.length) continue;
      html += renderGrupo(code, gs);
    }
    html += `<div class="row-actions"><button class="btn sec" id="btnErrXlsx">Descargar errores y advertencias (Excel)</button><button class="btn sec" id="btnBajas">Descargar bajas ignoradas (Excel)</button></div>`;
    er.innerHTML = html;
    wireRevision();
  }

  function renderFiltros(el) {
    const rows = allRows();
    const count = (key) => rows.reduce((m, r) => (m.set(r[key], (m.get(r[key]) || 0) + 1), m), new Map());
    const cursos = [...count('curso')].sort((a, b) => a[0].localeCompare(b[0]));
    const emps = [...count('compania')].sort((a, b) => b[1] - a[1]);
    const exC = new Set(S.ex.cursos.map(Core.norm)), exE = new Set(S.ex.empresas.map(Core.norm));
    const list = (arr, ex, kind) => arr.map(([n, c]) => `<label><input type="checkbox" data-kind="${kind}" data-name="${esc(n)}" ${ex.has(Core.norm(n)) ? '' : 'checked'}> ${esc(n)} <span class="c">${c}</span></label>`).join('');
    el.innerHTML = `<details ${S.filtrosOpen ? 'open' : ''}><summary>Filtros: cursos y compañías que no aplican (${S.ex.cursos.length + S.ex.empresas.length} excluidos)</summary>
      <p class="hint">Desmarca lo que no deba cargarse en SIRCE. Se recuerda en este navegador.</p>
      <div class="grid"><div><b>Cursos</b><div class="checks">${list(cursos, exC, 'cursos')}</div></div><div><b>Compañías</b><div class="checks">${list(emps, exE, 'empresas')}</div></div></div></details>`;
    $('details', el).addEventListener('toggle', (e) => { S.filtrosOpen = e.target.open; });
    $$('input[type=checkbox]', el).forEach((cb) => cb.addEventListener('change', () => {
      const k = cb.dataset.kind, n = cb.dataset.name;
      const arr = S.ex[k].filter((x) => Core.norm(x) !== Core.norm(n));
      if (!cb.checked) arr.push(n);
      S.ex[k] = arr; saveEx(); recalc();
    }));
  }

  const inp = (id, ph, w, type) => `<input type="${type || 'text'}" data-f="${id}" placeholder="${esc(ph || '')}" style="width:${w || 110}px">`;
  function renderGrupo(code, gs) {
    const lim = S.limit[code] || 40;
    const total = gs.reduce((s, g) => s + g.count, 0);
    const shown = gs.slice(0, lim);
    let head, body;
    const names = S.claves ? [...S.claves.sucursales.names.values()] : [];
    if (code === 'SUCURSAL') {
      const sug = gs.map((g) => Core.sugerirSucursal(g.label, names));
      const n = sug.filter(Boolean).length;
      body = shown.map((g, i) => {
        const s = sug[i];
        const k = s ? S.claves.sucursales.map.get(Core.norm(s.nombre)) : '';
        return `<tr data-code="${code}" data-key="${esc(g.key)}"><td>${esc(g.label)}</td><td class="n">${g.count}</td>
          <td>${s ? `<button class="btn sec sm" data-act="usar" data-val="${esc(k)}">Usar “${esc(s.nombre)}” (${esc(k)})</button>` : '<span class="muted small">sin sugerencia</span>'}</td>
          <td><input list="dlSuc" data-f="v" placeholder="Clave o nombre de Claves" style="width:200px"> <button class="btn sm" data-act="save">Guardar</button></td></tr>`;
      }).join('');
      head = `<th>Sucursal (del catálogo)</th><th>Filas</th><th>Sugerencia</th><th>Clave de establecimiento</th>`;
      const dl = `<datalist id="dlSuc">${names.map((x) => `<option value="${esc(x)}">`).join('')}</datalist>`;
      return wrapGrupo(code, gs.length, total, `${n ? `<div class="row-actions"><button class="btn sm" id="btnAllSug">Aceptar las ${n} sugerencias</button><span class="muted small">Revisa que sean correctas antes de aceptarlas.</span></div>` : ''}<table><tr>${head}</tr>${body}</table>${dl}`, gs, lim);
    }
    if (code === 'PUESTO') {
      head = '<th>Puesto</th><th>Filas</th><th>Clave de ocupación</th>';
      body = shown.map((g) => `<tr data-code="${code}" data-key="${esc(g.key)}"><td>${esc(g.label)}</td><td class="n">${g.count}</td><td>${inp('v', 'Ej. 3375', 100, 'number')} <button class="btn sm" data-act="save">Guardar</button></td></tr>`).join('');
    } else if (code === 'NIVEL') {
      const opts = S.claves ? [...S.claves.niveles.values()].map((n) => n.nombre) : [];
      head = '<th>Escolaridad</th><th>Filas</th><th>Nivel de estudios</th><th>Claves (nivel · doc. · inst.)</th>';
      body = shown.map((g) => {
        const guess = Core.inferNivel(g.label);
        return `<tr data-code="${code}" data-key="${esc(g.key)}"><td>${esc(g.label)}</td><td class="n">${g.count}</td>
        <td><select data-f="lvl"><option value="">Elegir…</option>${[...new Set(opts)].map((o) => `<option ${Core.norm(o) === guess ? 'selected' : ''}>${esc(o)}</option>`).join('')}</select></td>
        <td>${inp('n', 'nivel', 60, 'number')} ${inp('d', 'doc.', 55, 'number')} ${inp('i', 'inst.', 55, 'number')} <button class="btn sm" data-act="save">Guardar</button></td></tr>`;
      }).join('');
    } else if (code === 'CURSO') {
      head = '<th>Curso · fecha de inicio · empresa</th><th>Filas</th><th>Clave de curso</th>';
      body = shown.map((g) => `<tr data-code="${code}" data-key="${esc(g.key)}"><td>${esc(g.label)}</td><td class="n">${g.count}</td><td>${inp('v', 'Clave', 110, 'number')} <button class="btn sm" data-act="save">Guardar</button></td></tr>`).join('');
    } else if (code === 'ESTADO') {
      head = '<th>Estado</th><th>Filas</th><th>Clave de estado</th>';
      body = shown.map((g) => `<tr data-code="${code}" data-key="${esc(g.key)}"><td>${esc(g.label)}</td><td class="n">${g.count}</td><td>${inp('v', 'Clave', 90, 'number')} <button class="btn sm" data-act="save">Guardar</button></td></tr>`).join('');
    } else if (code === 'MUN') {
      head = '<th>Estado</th><th>Filas</th><th>Municipio principal y clave</th>';
      body = shown.map((g) => `<tr data-code="${code}" data-key="${esc(g.key)}"><td>${esc(g.label)}</td><td class="n">${g.count}</td><td>${inp('m', 'Municipio', 160)} ${inp('v', 'Clave', 70, 'number')} <button class="btn sm" data-act="save">Guardar</button></td></tr>`).join('');
    } else if (code === 'FECHA') {
      head = '<th>Colaborador · curso</th><th>Motivo</th><th>Fecha inicial</th><th>Fecha final</th>';
      body = shown.map((g) => { const r = g.rows[0]; const iso = (x) => (x ? Core.ymdStr(x) : '');
        return `<tr data-code="${code}" data-key="${esc(g.key)}"><td>${esc(r.nombreCompleto)}<br><span class="muted small">${esc(r.curso)}</span></td><td class="small">${esc(g.label)}</td>
        <td>${inp('fi', '', 140, 'date').replace('placeholder', `value="${iso(r.fi)}" placeholder`)}</td><td>${inp('ff', '', 140, 'date').replace('placeholder', `value="${iso(r.ff)}" placeholder`)} <button class="btn sm" data-act="save">Confirmar</button></td></tr>`; }).join('');
    } else if (code === 'HORAS') {
      head = '<th>Colaborador · curso</th><th>Problema</th><th>Horas</th>';
      body = shown.map((g) => { const r = g.rows[0];
        return `<tr data-code="${code}" data-key="${esc(g.key)}"><td>${esc(r.nombreCompleto)}<br><span class="muted small">${esc(r.curso)}</span></td><td class="small">${esc(g.label)}</td><td>${inp('v', 'Horas', 80, 'number')} <button class="btn sm" data-act="save">Guardar</button></td></tr>`; }).join('');
    } else if (code === 'CURP') {
      head = '<th>CURP actual</th><th>Filas</th><th>CURP correcta</th>';
      body = shown.map((g) => `<tr data-code="${code}" data-key="${esc(g.key)}"><td>${esc(g.label)}<br><span class="muted small">ID ${esc(g.key)} · ${esc(g.rows[0].nombreCompleto)}</span></td><td class="n">${g.count}</td><td>${inp('v', '18 caracteres', 190)} <button class="btn sm" data-act="save">Guardar</button></td></tr>`).join('');
    }
    return wrapGrupo(code, gs.length, total, `<table><tr>${head}</tr>${body}</table>`, gs, lim);
  }
  function wrapGrupo(code, nGroups, total, inner, gs, lim) {
    const more = gs.length > lim ? `<div class="row-actions"><button class="btn sec sm" data-more="${code}">Mostrar más (${gs.length - lim} restantes)</button></div>` : '';
    return `<details open class="card"><summary>${TITLES[code]} <span class="tag bad">${nGroups} distintos · ${total} filas</span></summary>${inner}${more}</details>`;
  }

  function saveFix(tr) {
    const code = tr.dataset.code, key = tr.dataset.key; const f = (n) => { const e = $(`[data-f="${n}"]`, tr); return e ? e.value.trim() : ''; };
    const num = (v) => (v !== '' && !isNaN(Number(v)) ? Number(v) : null);
    const ov = S.ov;
    if (code === 'PUESTO') { const v = num(f('v')); if (v === null) return alert('Escribe la clave de ocupación (número).'); ov.puesto[key] = v; }
    else if (code === 'CURSO') { const v = num(f('v')); if (v === null) return alert('Escribe la clave del curso (número).'); ov.cursoClave[key] = v; }
    else if (code === 'ESTADO') { const v = num(f('v')); if (v === null) return alert('Escribe la clave del estado.'); ov.estadoMun[key] = Object.assign({}, ov.estadoMun[key], { claveEstado: v }); }
    else if (code === 'MUN') { const v = num(f('v')); if (!f('m') || v === null) return alert('Escribe el municipio y su clave.'); ov.estadoMun[key] = Object.assign({}, ov.estadoMun[key], { municipio: f('m'), clave: v }); }
    else if (code === 'HORAS') { const v = num(f('v')); if (!(v > 0)) return alert('Escribe las horas (mayor a 0).'); ov.horas[key] = v; }
    else if (code === 'CURP') { const v = f('v').toUpperCase(); if (v.length !== 18) return alert('La CURP debe tener 18 caracteres.'); ov.curp[key] = v; }
    else if (code === 'FECHA') { if (!f('fi')) return alert('Elige la fecha inicial.'); ov.fechas[key] = { fi: f('fi'), ff: f('ff') || f('fi') }; }
    else if (code === 'NIVEL') {
      const lvl = f('lvl'); if (!lvl) return alert('Elige el nivel de estudios.');
      ov.escolaridad[key] = Core.norm(lvl);
      const n = num(f('n'));
      if (n !== null) { const base = S.claves.niveles.get(Core.norm(lvl)) || {}; ov.nivel[Core.norm(lvl)] = { nombre: lvl, nivel: n, doc: num(f('d')) ?? base.doc ?? null, inst: num(f('i')) ?? base.inst ?? null }; }
    }
    else if (code === 'SUCURSAL') {
      const v = f('v'); let k = num(v);
      if (k === null) { k = S.claves.sucursales.map.get(Core.norm(v)); }
      if (k === undefined || k === null) return alert('Escribe una clave numérica o elige un nombre de la lista.');
      ov.sucursal[key] = k;
    }
    saveOv(); recalc();
  }

  function wireRevision() {
    const er = $('#errores');
    er.onclick = onRevClick;
    er.onkeydown = (e) => { if (e.key === 'Enter' && e.target.matches('input')) { const tr = e.target.closest('tr'); if (tr) saveFix(tr); } };
    const all = $('#btnAllSug');
    if (all) all.addEventListener('click', () => {
      const names = [...S.claves.sucursales.names.values()];
      Core.groupErrors(S.res.records).filter((g) => g.code === 'SUCURSAL').forEach((g) => {
        const s = Core.sugerirSucursal(g.label, names);
        if (s) S.ov.sucursal[g.key] = S.claves.sucursales.map.get(Core.norm(s.nombre));
      });
      saveOv(); recalc();
    });
    const be = $('#btnErrXlsx'); if (be) be.addEventListener('click', downloadErrores);
    const bb = $('#btnBajas'); if (bb) bb.addEventListener('click', downloadBajas);
  }
  function onRevClick(e) {
    const more = e.target.closest('[data-more]');
    if (more) { S.limit[more.dataset.more] = (S.limit[more.dataset.more] || 40) + 80; return renderRevision(); }
    const b = e.target.closest('button[data-act]'); if (!b) return;
    const tr = b.closest('tr');
    if (b.dataset.act === 'usar') { S.ov.sucursal[tr.dataset.key] = Number(b.dataset.val); saveOv(); recalc(); }
    else if (b.dataset.act === 'save') saveFix(tr);
  }

  /* ---------- equivalencias ---------- */
  function renderEquiv() {
    const el = $('#equiv'); const R = S.res; const ov = S.ov;
    if (!R) { el.innerHTML = '<div class="card"><p class="muted">Carga los archivos en el paso 1.</p></div>'; return; }
    const empresas = [...new Set(allRows().map((r) => r.compania))].sort();
    const estados = [...new Set(R.records.map((r) => r.layout[5]))].filter(Boolean).sort();
    const aliasRows = empresas.map((n) => `<tr data-alias="${esc(Core.norm(n))}"><td>${esc(n)}</td><td><input type="text" data-f="a" value="${esc(ov.empresaAlias[Core.norm(n)] || '')}" placeholder="${esc(n)}" style="width:300px"></td></tr>`).join('');
    const munRows = estados.map((n) => { const k = Core.estadoNorm(n); const d = Core.MUN_PRINCIPAL[k]; const o = ov.estadoMun[k] || {};
      return `<tr data-est="${esc(k)}"><td>${esc(n)}</td><td><input type="text" data-f="m" value="${esc(o.municipio || '')}" placeholder="${esc(d ? d[0] : '')}" style="width:200px"></td><td><input type="number" data-f="c" value="${esc(o.clave ?? '')}" placeholder="${esc(S.claves.municipios.map.get(Core.norm(d ? d[0] : '')) ?? (d ? d[1] : ''))}" style="width:90px"></td></tr>`; }).join('');
    const saved = (obj) => Object.entries(obj);
    const savedTable = (title, obj, kind) => { const e = saved(obj); return e.length ? `<details><summary>${title} (${e.length})</summary><table>${e.map(([k, v]) => `<tr><td>${esc(k)}</td><td>${esc(typeof v === 'object' ? JSON.stringify(v) : v)}</td><td class="n"><button class="btn sec sm" data-del="${kind}" data-key="${esc(k)}">Quitar</button></td></tr>`).join('')}</table></details>` : ''; };
    el.innerHTML = `
      <div class="card"><h2>Alias de empresa para claves de curso</h2><p class="hint">La clave de curso se busca por nombre del curso + fecha de inicio + empresa. Si en el archivo de claves de cursos la empresa se llama distinto que en el reporte, escribe aquí el nombre que usa el archivo de claves. Vacío = mismo nombre.</p>
        <table><tr><th>Empresa en el reporte</th><th>Nombre en claves de cursos</th></tr>${aliasRows}</table><div class="row-actions"><button class="btn" id="saveAlias">Guardar alias</button></div></div>
      <div class="card"><h2>Municipio principal por estado</h2><p class="hint">Se asigna el municipio principal del estado (por defecto la capital). Si la clave queda vacía se toma de la hoja Claves o de la tabla interna.</p>
        <table><tr><th>Estado</th><th>Municipio</th><th>Clave</th></tr>${munRows}</table><div class="row-actions"><button class="btn" id="saveMun">Guardar municipios</button></div></div>
      <div class="card"><h2>Equivalencias guardadas en este navegador</h2><p class="hint">Se guardan para la próxima carga. Para compartirlas con tu equipo, exporta el archivo y que lo importen en su navegador (no incluye CURP ni datos de personas).</p>
        ${savedTable('Puestos → clave de ocupación', ov.puesto, 'puesto')}${savedTable('Escolaridades → nivel', ov.escolaridad, 'escolaridad')}${savedTable('Niveles → claves', ov.nivel, 'nivel')}
        ${savedTable('Sucursales → clave de establecimiento', ov.sucursal, 'sucursal')}${savedTable('Cursos → clave', ov.cursoClave, 'cursoClave')}
        ${Object.keys(ov.fechas).length + Object.keys(ov.horas).length + Object.keys(ov.curp).length ? `<p class="muted small">Correcciones de filas puntuales: ${Object.keys(ov.fechas).length} fechas, ${Object.keys(ov.horas).length} horas, ${Object.keys(ov.curp).length} CURP.</p>` : ''}
        <div class="row-actions"><button class="btn sec" id="expEq">Exportar equivalencias</button><label class="btn sec" style="cursor:pointer">Importar equivalencias<input type="file" id="impEq" accept=".json" hidden></label><button class="btn sec" id="clrEq">Borrar todas</button></div></div>`;
    $('#saveAlias').onclick = () => { $$('[data-alias]', el).forEach((tr) => { const v = $('[data-f=a]', tr).value.trim(); if (v) ov.empresaAlias[tr.dataset.alias] = v; else delete ov.empresaAlias[tr.dataset.alias]; }); saveOv(); recalc(); };
    $('#saveMun').onclick = () => { $$('[data-est]', el).forEach((tr) => { const m = $('[data-f=m]', tr).value.trim(), c = $('[data-f=c]', tr).value.trim();
      const prev = ov.estadoMun[tr.dataset.est] || {};
      if (m) ov.estadoMun[tr.dataset.est] = Object.assign({}, prev, { municipio: m, clave: c !== '' ? Number(c) : undefined }); else { delete prev.municipio; delete prev.clave; if (!Object.keys(prev).length) delete ov.estadoMun[tr.dataset.est]; } }); saveOv(); recalc(); };
    el.onclick = (e) => { const b = e.target.closest('[data-del]'); if (b) { delete ov[b.dataset.del][b.dataset.key]; saveOv(); recalc(); } };
    $('#expEq').onclick = () => { const { puesto, escolaridad, nivel, sucursal, estadoMun, empresaAlias, cursoClave } = ov; download(new Blob([JSON.stringify({ version: 1, puesto, escolaridad, nivel, sucursal, estadoMun, empresaAlias, cursoClave }, null, 1)], { type: 'application/json' }), 'equivalencias_sirce.json'); };
    $('#impEq').onchange = async (e) => { try { const j = JSON.parse(await e.target.files[0].text()); for (const k of ['puesto', 'escolaridad', 'nivel', 'sucursal', 'estadoMun', 'empresaAlias', 'cursoClave']) Object.assign(ov[k], j[k] || {}); saveOv(); recalc(); } catch (err) { alert('Archivo no válido: ' + err.message); } };
    $('#clrEq').onclick = () => { if (confirm('¿Borrar todas las equivalencias guardadas en este navegador?')) { S.ov = Core.emptyOv(); saveOv(); recalc(); } };
  }

  /* ---------- generar ---------- */
  function renderGenerar() {
    const el = $('#generar'); const R = S.res;
    if (!R) { el.innerHTML = '<div class="card"><p class="muted">Carga los archivos en el paso 1.</p></div>'; return; }
    const sheets = Core.sheetsByEmpresa(R.records);
    const bad = R.records.filter((x) => x.errors.length).length;
    const total = sheets.reduce((s, x) => s + x.rows.length, 0);
    el.innerHTML = `<div class="card"><h2>Archivo para cargar en SIRCE</h2>
      <p class="hint">Una pestaña por empresa, con las 27 columnas del layout. Solo se incluyen las filas sin errores.</p>
      ${bad ? `<div class="msg warn">${bad} filas con errores no se incluirán. Corrígelas en “2 · Revisión” o genera el archivo con lo que ya está listo.</div>` : '<div class="msg ok">Todas las filas están listas.</div>'}
      <table><tr><th>Pestaña (empresa)</th><th class="n">Filas</th></tr>${sheets.map((s) => `<tr><td>${esc(s.empresa)}</td><td class="n">${s.rows.length}</td></tr>`).join('')}<tr><td><b>Total</b></td><td class="n"><b>${total}</b></td></tr></table>
      <div class="row-actions"><button class="btn" id="btnGen" ${total ? '' : 'disabled'}>Descargar Excel para SIRCE</button></div></div>`;
    const b = $('#btnGen'); if (b) b.onclick = () => downloadLayout(sheets);
  }

  /* ---------- descargas ---------- */
  function download(blob, name) {
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  }
  const stamp = () => new Date().toISOString().slice(0, 10);
  function writeBook(wb, name) { const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true }); download(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }), name); }
  function downloadLayout(sheets) {
    const wb = XLSX.utils.book_new(); const used = new Set();
    for (const s of sheets) {
      const ws = XLSX.utils.aoa_to_sheet([Core.LAYOUT, ...s.rows]);
      const range = XLSX.utils.decode_range(ws['!ref']);
      for (let r = 1; r <= range.e.r; r++) for (const c of [18, 19]) { const cell = ws[XLSX.utils.encode_cell({ r, c })]; if (cell && cell.t === 'n') cell.z = 'dd/mm/yyyy'; }
      ws['!cols'] = Core.LAYOUT.map((h) => ({ wch: Math.max(12, Math.min(34, h.length + 2)) }));
      XLSX.utils.book_append_sheet(wb, ws, Core.sheetName(s.empresa, used));
    }
    writeBook(wb, `Layout_SIRCE_${stamp()}.xlsx`);
  }
  function downloadErrores() {
    const err = [['Tipo', 'Detalle', 'ID', 'Colaborador', 'Curso', 'Empresa', 'Fecha inicio']]; const adv = [['Tipo', 'Detalle', 'ID', 'Colaborador', 'Curso', 'Empresa']];
    for (const r of S.res.records) {
      for (const e of r.errors) err.push([TITLES[e.code] || e.code, e.label, r.id, r.nombreCompleto, r.curso, r.empresa, r.fi ? Core.ymdStr(r.fi) : '']);
      for (const w of r.warnings) adv.push([w.code, w.label, r.id, r.nombreCompleto, r.curso, r.empresa]);
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(err), 'Errores');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(adv), 'Advertencias');
    writeBook(wb, `Errores_SIRCE_${stamp()}.xlsx`);
  }
  function downloadBajas() {
    const rows = [['ID', 'Colaborador', 'Curso', 'Empresa', 'Sucursal']];
    S.res.bajas.forEach((r) => rows.push([r.id, `${r.apPat} ${r.apMat} ${r.nombre}`.replace(/\s+/g, ' ').trim(), r.curso, r.compania, r.sucursal]));
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Bajas');
    writeBook(wb, `Bajas_ignoradas_${stamp()}.xlsx`);
  }


  /* ---------- DC-3 ---------- */
  const dc3Cfg = Object.assign({ patron: 'CARLOS MANUEL VILLARREAL TUYUB', orden: 'nombre', excluir: [] }, lsGet('sirce.dc3', {}));
  const dc3Save = () => lsSet('sirce.dc3', dc3Cfg);
  const dc3State = { busy: false, msg: '' };

  function dc3Items() {
    const R = S.res; if (!R || !S.info || !S.info.data) return null;
    const items = [], sinDatos = [], sinRep = new Map();
    for (const r of R.records.filter((x) => !x.errors.length)) {
      const e = Core.datosEmpresa(S.info.data, r.sucursal, r.empresa);
      if (!e) { sinDatos.push(r); continue; }
      if (!e.rep) sinRep.set(r.sucursal, (sinRep.get(r.sucursal) || 0) + 1);
      const d = r.det;
      const nombre = dc3Cfg.orden === 'apellidos' ? `${d.apPat} ${d.apMat} ${d.nombre}` : `${d.nombre} ${d.apPat} ${d.apMat}`;
      items.push({ rec: r, suc: r.sucursal, data: {
        nombre: nombre.replace(/\s+/g, ' ').trim(), curp: r.curp, ocupacion: d.claveOcup, puesto: r.puesto, razon: e.razon, rfc: e.rfc,
        curso: r.curso, horas: r.horas, fi: r.fi, ff: r.ff, agente: '', instructor: '', patron: dc3Cfg.patron, representante: e.rep,
      } });
    }
    return { items, sinDatos, sinRep };
  }

  function renderDC3() {
    const el = $('#dc3'); const R = S.res;
    if (!R) { el.innerHTML = '<div class="card"><p class="muted">Carga los archivos en el paso 1.</p></div>'; return; }
    if (!S.info || !S.info.data) { el.innerHTML = '<div class="card"><h2>Constancias DC-3</h2><div class="msg warn">Carga el archivo “Información de sucursales (DC-3)” en el paso 1 para tener razón social, RFC y representante de los trabajadores.</div></div>'; return; }
    const D = dc3Items();
    const bySuc = new Map(); D.items.forEach((i) => bySuc.set(i.suc, (bySuc.get(i.suc) || 0) + 1));
    const sucs = [...bySuc].sort((a, b) => a[0].localeCompare(b[0]));
    const ex = new Set(dc3Cfg.excluir);
    const sel = D.items.filter((i) => !ex.has(i.suc)).length;
    const pend = new Map(); D.sinDatos.forEach((r) => pend.set(r.sucursal + ' · ' + r.empresa, (pend.get(r.sucursal + ' · ' + r.empresa) || 0) + 1));
    el.innerHTML = `<div class="card"><h2>Constancias DC-3 (PDF por colaborador)</h2>
      <p class="hint">Se generan con las filas listas para SIRCE (sin errores), un PDF por colaborador y curso, en un ZIP con una carpeta por sucursal. Nombre del archivo: SUCURSAL - COLABORADOR - CURSO - DC3. El instructor y el agente capacitador van vacíos.</p>
      <div class="row-actions" style="margin:0 0 12px"><label>Patrón o representante legal <input type="text" id="dc3Patron" value="${esc(dc3Cfg.patron)}" style="width:300px"></label>
        <label>Orden del nombre <select id="dc3Orden"><option value="nombre" ${dc3Cfg.orden === 'nombre' ? 'selected' : ''}>Nombre y apellidos (como el ejemplo)</option><option value="apellidos" ${dc3Cfg.orden === 'apellidos' ? 'selected' : ''}>Apellidos y nombre (como indica el formato)</option></select></label></div>
      ${D.sinDatos.length ? `<div class="msg bad">${D.sinDatos.length} constancias no se pueden generar porque falta la razón social o el RFC de su sucursal y empresa: ${esc([...pend].slice(0, 6).map(([k, v]) => k + ' (' + v + ')').join('; '))}${pend.size > 6 ? '…' : ''}</div>` : ''}
      ${D.sinRep.size ? `<div class="msg warn">${[...D.sinRep.values()].reduce((a, b) => a + b, 0)} constancias saldrán sin representante de los trabajadores (${D.sinRep.size} sucursales sin dato en el archivo).</div>` : ''}
      <details open><summary>Sucursales (${sucs.length}) · ${sel} constancias seleccionadas</summary>
        <div class="row-actions" style="margin:0 0 8px"><button class="btn sec sm" id="dc3All">Todas</button><button class="btn sec sm" id="dc3None">Ninguna</button></div>
        <div class="checks">${sucs.map(([n, c]) => `<label><input type="checkbox" data-suc="${esc(n)}" ${ex.has(n) ? '' : 'checked'}> ${esc(n)} <span class="c">${c}</span></label>`).join('')}</div></details>
      <div class="row-actions"><button class="btn" id="dc3Go" ${sel && !dc3State.busy ? '' : 'disabled'}>Generar PDFs (ZIP)</button><button class="btn sec" id="dc3Pend">Descargar pendientes (Excel)</button></div>
      <div id="dc3Prog" class="muted small" style="margin-top:10px">${esc(dc3State.msg)}</div></div>`;
    $('#dc3Patron').onchange = (e) => { dc3Cfg.patron = e.target.value.trim(); dc3Save(); };
    $('#dc3Orden').onchange = (e) => { dc3Cfg.orden = e.target.value; dc3Save(); };
    $$('[data-suc]', el).forEach((cb) => cb.addEventListener('change', () => { const n = cb.dataset.suc; dc3Cfg.excluir = dc3Cfg.excluir.filter((x) => x !== n); if (!cb.checked) dc3Cfg.excluir.push(n); dc3Save(); renderDC3(); }));
    $('#dc3All').onclick = () => { dc3Cfg.excluir = []; dc3Save(); renderDC3(); };
    $('#dc3None').onclick = () => { dc3Cfg.excluir = sucs.map((x) => x[0]); dc3Save(); renderDC3(); };
    $('#dc3Go').onclick = () => dc3Generate(D);
    $('#dc3Pend').onclick = () => {
      const rows = [['Problema', 'Sucursal', 'Empresa', 'Constancias']];
      pend.forEach((v, k) => { const [s, e] = k.split(' · '); rows.push(['Sin razón social o RFC', s, e, v]); });
      D.sinRep.forEach((v, k) => rows.push(['Sin representante de los trabajadores', k, '', v]));
      const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Pendientes DC-3'); writeBook(wb, `Pendientes_DC3_${stamp()}.xlsx`);
    };
  }

  const tick = () => new Promise((r) => setTimeout(r, 0));
  const carpeta = (s) => String(s || 'Sin sucursal').replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim() || 'Sin sucursal';
  async function dc3Generate(D) {
    const ex = new Set(dc3Cfg.excluir);
    const items = D.items.filter((i) => !ex.has(i.suc)).sort((a, b) => a.suc.localeCompare(b.suc) || a.data.nombre.localeCompare(b.data.nombre));
    const PART = 1000, parts = Math.ceil(items.length / PART);
    dc3State.busy = true; $('#dc3Go').disabled = true; const prog = $('#dc3Prog');
    try {
      for (let p = 0; p < parts; p++) {
        const zip = new JSZip(); const used = new Set();
        const slice = items.slice(p * PART, (p + 1) * PART);
        for (let i = 0; i < slice.length; i++) {
          const it = slice[i];
          const folder = carpeta(it.suc);
          let base = DC3.nombreArchivo(it.suc, it.data.nombre, it.data.curso);
          let path = `${folder}/${base}.pdf`;
          if (used.has(path)) { base = base.replace(/ - DC3$/, ` ${Core.ymdStr(it.data.fi)} - DC3`); path = `${folder}/${base}.pdf`; }
          for (let n = 2; used.has(path); n++) path = `${folder}/${base} (${n}).pdf`;
          used.add(path);
          const doc = DC3.crear(window.jspdf.jsPDF, it.data);
          zip.file(path, doc.output('arraybuffer'));
          if (i % 10 === 0) { prog.textContent = `Parte ${p + 1} de ${parts}: ${i + 1} de ${slice.length} constancias…`; await tick(); }
        }
        prog.textContent = `Comprimiendo parte ${p + 1} de ${parts}…`; await tick();
        const blob = await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 3 } });
        download(blob, `DC3_${stamp()}${parts > 1 ? '_parte' + (p + 1) : ''}.zip`);
        await tick();
      }
      dc3State.msg = `Listo: ${items.length} constancias en ${parts} archivo(s) ZIP.`;
    } catch (e) { dc3State.msg = 'Error al generar: ' + (e.message || e); }
    dc3State.busy = false; renderDC3();
  }

  wireDrops();
  recalc();
  window.__S = S; // útil para depuración
})();
