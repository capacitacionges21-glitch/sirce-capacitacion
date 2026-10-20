/* Lógica de cruce para el layout de carga SIRCE.
 * No depende del DOM ni de SheetJS: recibe matrices (arreglo de filas) y devuelve objetos.
 * Funciona en el navegador (window.Core) y en Node (module.exports) para pruebas. */
(function (root) {
  'use strict';
  const Core = {};

  /* ---------- utilidades ---------- */
  function norm(s) {
    if (s === null || s === undefined) return '';
    return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toUpperCase().replace(/\s+/g, ' ').trim();
  }
  const isMissing = (v) => v === null || v === undefined || v === '' || v === '#N/A' || v === 0 || v === '0';
  const toKey = (v) => (isMissing(v) ? null : (typeof v === 'number' ? v : (isNaN(Number(v)) ? null : Number(v))));
  const pad7 = (v) => {
    if (v === null || v === undefined || v === '') return '';
    let s = String(v).trim();
    if (/^\d+(\.0+)?$/.test(s)) s = s.replace(/\.0+$/, '').padStart(7, '0');
    return s;
  };
  const str = (v) => (v === null || v === undefined ? '' : String(v).trim());
  Core.norm = norm;

  function findHeaderRow(aoa, required) {
    for (let i = 0; i < Math.min(aoa.length, 15); i++) {
      const cells = (aoa[i] || []).map(norm);
      if (required.every((r) => cells.includes(r))) return i;
    }
    return -1;
  }
  function colFinder(headerRow) {
    const cells = headerRow.map(norm);
    return (name, nth = 0) => {
      let seen = -1;
      for (let i = 0; i < cells.length; i++) {
        if (cells[i] === name) { seen++; if (seen === nth) return i; }
      }
      return -1;
    };
  }

  /* ---------- fechas ---------- */
  function serialToYMD(n) {
    const d = new Date(Math.floor(n - 25569) * 86400000);
    return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
  }
  function parseDate(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate(), kind: 'serial' };
    if (typeof v === 'number') return Object.assign(serialToYMD(v), { kind: 'serial' });
    const s = String(v).trim();
    let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (m) return { y: +m[3], m: +m[2], d: +m[1], kind: 'text' };
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return { y: +m[1], m: +m[2], d: +m[3], kind: 'text' };
    return null;
  }
  const ymdStr = (x) => (x ? `${x.y}-${String(x.m).padStart(2, '0')}-${String(x.d).padStart(2, '0')}` : '');
  const cmpYMD = (a, b) => (a.y - b.y) || (a.m - b.m) || (a.d - b.d);
  const swapDM = (x) => ({ y: x.y, m: x.d, d: x.m, kind: x.kind });
  const validDM = (x) => x.m >= 1 && x.m <= 12 && x.d >= 1 && x.d <= 31;
  Core.ymdStr = ymdStr;

  /* El reporte NOM exporta fechas con día y mes invertidos; se corrige comparando contra la columna Mes. */
  function resolveDates(fiV, ffV, mes, anio) {
    const a = parseDate(fiV);
    if (!a) return { fi: null, ff: null, fixed: false, dudosa: true, motivo: 'Sin fecha inicial' };
    let fi = { y: a.y, m: a.m, d: a.d };
    let fixed = false, dudosa = false, motivo = '';
    const mesN = Number(mes) || 0;
    if (mesN && a.m !== mesN) {
      const s = swapDM(a);
      if (a.d === mesN && validDM(s)) { fi = { y: s.y, m: s.m, d: s.d }; fixed = true; }
      else { dudosa = true; motivo = `El mes de la fecha (${a.m}) no coincide con el mes del curso (${mesN})`; }
    }
    const anioN = Number(anio) || 0;
    if (anioN && fi.y !== anioN) { dudosa = true; motivo = motivo || `El año de la fecha (${fi.y}) no coincide con ${anioN}`; }
    let ff = fi;
    const b = parseDate(ffV);
    if (b) {
      ff = { y: b.y, m: b.m, d: b.d };
      if (fixed && validDM(swapDM(b))) ff = { y: b.y, m: b.d, d: b.m };
      if (cmpYMD(ff, fi) < 0) {
        const alt = { y: b.y, m: b.d, d: b.m };
        if (validDM(alt) && cmpYMD(alt, fi) >= 0) ff = alt;
        else { dudosa = true; motivo = motivo || 'La fecha final es anterior a la inicial'; }
      }
    }
    return { fi, ff, fixed, dudosa, motivo };
  }

  /* ---------- lectura de archivos (matrices) ---------- */
  Core.parseReporte = function (aoa) {
    const hr = findHeaderRow(aoa, ['TRABAJADOR', 'CURSO', 'FECHA INICIAL']);
    if (hr < 0) return null;
    const c = colFinder(aoa[hr]);
    const I = {
      id: c('TRABAJADOR'), apPat: c('APELLIDO PATERNO'), apMat: c('APELLIDO MATERNO'), nombre: c('NOMBRE'),
      curso: c('CURSO'), mes: c('MES'), anio: c('ANO'), fi: c('FECHA INICIAL'), ff: c('FECHA FINAL'),
      comp: c('COMPANIA TRABAJADOR'), cadena: c('CADENA'), estado: c('ESTADO'), suc: c('SUCURSAL'),
      puesto: c('PUESTO'), horas: c('HORAS'), inst: c('INSTRUCTOR'), status: c('STATUS PROGRAMA DE CURSOS'),
    };
    const faltan = ['id', 'curso', 'mes', 'fi', 'comp', 'horas'].filter((k) => I[k] < 0);
    if (faltan.length) return { error: 'Faltan columnas en el reporte: ' + faltan.join(', ') };
    const g = (r, k) => (I[k] >= 0 ? r[I[k]] : '');
    const rows = [];
    for (let i = hr + 1; i < aoa.length; i++) {
      const r = aoa[i] || [];
      const id = pad7(g(r, 'id'));
      if (!id || !str(g(r, 'curso'))) continue;
      rows.push({
        rowNum: i + 1, id,
        apPat: str(g(r, 'apPat')), apMat: str(g(r, 'apMat')), nombre: str(g(r, 'nombre')),
        curso: str(g(r, 'curso')), mes: g(r, 'mes'), anio: g(r, 'anio'),
        fiRaw: g(r, 'fi'), ffRaw: g(r, 'ff'),
        compania: str(g(r, 'comp')), cadena: str(g(r, 'cadena')), estado: str(g(r, 'estado')),
        sucursal: str(g(r, 'suc')), puesto: str(g(r, 'puesto')), horas: g(r, 'horas'),
        instructor: str(g(r, 'inst')), status: str(g(r, 'status')),
      });
    }
    return { rows };
  };

  Core.parseCatalogo = function (aoa) {
    const hr = findHeaderRow(aoa, ['ID', 'CURP']);
    if (hr < 0) return null;
    const c = colFinder(aoa[hr]);
    const I = { id: c('ID'), curp: c('CURP'), nombre: c('NOMBRE'), puesto: c('PUESTO'), esc: c('ESCOLARIDAD'),
      estado: c('ESTADO'), suc: c('SUCURSAL'), emp: c('EMPRESA'), cadena: c('CADENA') };
    const g = (r, k) => (I[k] >= 0 ? r[I[k]] : '');
    const map = new Map();
    let dups = 0;
    for (let i = hr + 1; i < aoa.length; i++) {
      const r = aoa[i] || [];
      const id = pad7(g(r, 'id'));
      if (!id) continue;
      if (map.has(id)) { dups++; continue; }
      map.set(id, {
        id, curp: norm(g(r, 'curp')), nombre: str(g(r, 'nombre')), puesto: str(g(r, 'puesto')),
        escolaridad: str(g(r, 'esc')), estado: str(g(r, 'estado')).replace(/^\d+\s*/, ''),
        sucursal: str(g(r, 'suc')), empresa: str(g(r, 'emp')), cadena: str(g(r, 'cadena')),
      });
    }
    return { map, dups };
  };

  /* Hoja "Claves". Las columnas de usuario y contraseña de SIRCE NO se leen. */
  Core.parseClaves = function (aoa) {
    const hr = findHeaderRow(aoa, ['CLAVE ESTADO', 'CLAVE MUNICIPIO']);
    if (hr < 0) return null;
    const c = colFinder(aoa[hr]);
    const body = aoa.slice(hr + 1);
    const pairs = (kName, vName, kNth = 0, vNth = 0) => {
      const ki = c(kName, kNth), vi = c(vName, vNth);
      const m = new Map(); const names = new Map();
      if (ki < 0 || vi < 0) return { map: m, names };
      for (const r of body) {
        const name = str(r[ki]); if (!name) continue;
        const k = toKey(r[vi]);
        names.set(norm(name), name);
        if (k !== null && !m.has(norm(name))) m.set(norm(name), k);
      }
      return { map: m, names };
    };
    const estados = pairs('ESTADO', 'CLAVE ESTADO');
    const municipios = pairs('MUNICIPIO', 'CLAVE MUNICIPIO');
    const p1 = pairs('PUESTO', 'CLAVE OCUPACION', 0, 0);
    const p2 = pairs('PUESTO', 'CLAVE', 1, 0);
    const puestos = { map: new Map(p1.map), names: new Map([...p1.names, ...p2.names]) };
    for (const [k, v] of p2.map) if (!puestos.map.has(k)) puestos.map.set(k, v);
    const sucursales = pairs('SUCURSAL', '# DE SUCURSAL');
    const ni = c('NIVEL DE ESTUDIO'), nk = c('CLAVE NIV ESTUDIOS'), nd = c('CLAVE DOC PROBATORIO'), nin = c('CLAVE INSTITUCION');
    const niveles = new Map();
    if (ni >= 0) {
      for (const r of body) {
        const name = str(r[ni]); if (!name) continue;
        const rec = { nombre: name, nivel: toKey(r[nk]), doc: toKey(r[nd]), inst: toKey(r[nin]) };
        const prev = niveles.get(norm(name));
        if (!prev || (prev.nivel === null && rec.nivel !== null)) niveles.set(norm(name), rec);
      }
    }
    return { estados, municipios, puestos, sucursales, niveles };
  };

  Core.parseClavesCursos = function (aoa) {
    const hr = findHeaderRow(aoa, ['NOMBRE CURSO', 'CLAVE', 'EMPRESA', 'FEC INICIO']);
    if (hr < 0) return null;
    const c = colFinder(aoa[hr]);
    const I = { n: c('NOMBRE CURSO'), k: c('CLAVE'), e: c('EMPRESA'), f: c('FEC INICIO'), d: c('DURACION'),
      t: c('CLAVE TIP AGENTE'), m: c('CLAVE MODALIDAD'), o: c('CLAVE CAPACITACION (OBJETIVO)') };
    const map = new Map(); const conflictos = [];
    for (let i = hr + 1; i < aoa.length; i++) {
      const r = aoa[i] || [];
      const nombre = str(r[I.n]); const clave = toKey(r[I.k]); const f = parseDate(r[I.f]);
      if (!nombre || clave === null || !f) continue;
      const key = Core.cursoKey(nombre, f, str(r[I.e]));
      const rec = { clave, duracion: I.d >= 0 ? r[I.d] : '', tip: str(r[I.t]), modalidad: str(r[I.m]), objetivo: str(r[I.o]), nombre, empresa: str(r[I.e]) };
      if (map.has(key) && map.get(key).clave !== clave) conflictos.push({ key, nombre, empresa: rec.empresa, fecha: ymdStr(f), claves: [map.get(key).clave, clave] });
      else if (!map.has(key)) map.set(key, rec);
    }
    return { map, conflictos };
  };
  Core.cursoKey = (nombre, fecha, empresa) => `${norm(nombre)}|${ymdStr(fecha)}|${norm(empresa)}`;

  /* ---------- reglas por defecto ---------- */
  const ESTADO_ALIAS = {
    'MEXICO': 'ESTADO DE MEXICO', 'EDO MEXICO': 'ESTADO DE MEXICO', 'EDO DE MEXICO': 'ESTADO DE MEXICO',
    'CIUDAD DE MEXICO': 'CDMX', 'DISTRITO FEDERAL': 'CDMX', 'DF': 'CDMX', 'MICHOACAN': 'MICHOACAN DE OCAMPO',
    'VERACRUZ DE IGNACIO DE LA LLAVE': 'VERACRUZ', 'COAHUILA DE ZARAGOZA': 'COAHUILA',
  };
  /* Municipio principal (capital) por estado; la clave sale de la hoja Claves por nombre.
   * `inegi` solo se usa si el nombre no existe en Claves (se marca como advertencia). */
  const MUN_PRINCIPAL = {
    'AGUASCALIENTES': ['Aguascalientes', 1], 'BAJA CALIFORNIA': ['Mexicali', 2], 'BAJA CALIFORNIA SUR': ['La Paz', 3],
    'CAMPECHE': ['Campeche', 2], 'COAHUILA': ['Saltillo', 30], 'COLIMA': ['Colima', 2], 'CHIAPAS': ['Tuxtla Gutiérrez', 101],
    'CHIHUAHUA': ['Chihuahua', 19], 'CDMX': ['Cuauhtémoc', 15], 'DURANGO': ['Durango', 5], 'GUANAJUATO': ['Guanajuato', 15],
    'GUERRERO': ['Chilpancingo de los Bravo', 29], 'HIDALGO': ['Pachuca de Soto', 48], 'JALISCO': ['Guadalajara', 39],
    'ESTADO DE MEXICO': ['Toluca', 106], 'MICHOACAN DE OCAMPO': ['Morelia', 53], 'MORELOS': ['Cuernavaca', 7],
    'NAYARIT': ['Tepic', 17], 'NUEVO LEON': ['Monterrey', 39], 'OAXACA': ['Oaxaca de Juárez', 67], 'PUEBLA': ['Puebla', 114],
    'QUERETARO': ['Querétaro', 14], 'QUINTANA ROO': ['Othón P. Blanco', 4], 'SAN LUIS POTOSI': ['San Luis Potosí', 28],
    'SINALOA': ['Culiacán', 6], 'SONORA': ['Hermosillo', 30], 'TABASCO': ['Centro', 4], 'TAMAULIPAS': ['Victoria', 41],
    'TLAXCALA': ['Tlaxcala', 33], 'VERACRUZ': ['Xalapa', 87], 'YUCATAN': ['Mérida', 50], 'ZACATECAS': ['Zacatecas', 56],
  };
  const ESTADO_CLAVE = {
    'AGUASCALIENTES': 1, 'BAJA CALIFORNIA': 2, 'BAJA CALIFORNIA SUR': 3, 'CAMPECHE': 4, 'COAHUILA': 5, 'COLIMA': 6, 'CHIAPAS': 7,
    'CHIHUAHUA': 8, 'CDMX': 9, 'DURANGO': 10, 'GUANAJUATO': 11, 'GUERRERO': 12, 'HIDALGO': 13, 'JALISCO': 14, 'ESTADO DE MEXICO': 15,
    'MICHOACAN DE OCAMPO': 16, 'MORELOS': 17, 'NAYARIT': 18, 'NUEVO LEON': 19, 'OAXACA': 20, 'PUEBLA': 21, 'QUERETARO': 22,
    'QUINTANA ROO': 23, 'SAN LUIS POTOSI': 24, 'SINALOA': 25, 'SONORA': 26, 'TABASCO': 27, 'TAMAULIPAS': 28, 'TLAXCALA': 29,
    'VERACRUZ': 30, 'YUCATAN': 31, 'ZACATECAS': 32,
  };
  Core.MUN_PRINCIPAL = MUN_PRINCIPAL;

  /* Cargas anteriores a SIRCE (layout ya subido): sirven para aprender sucursal->establecimiento y puesto->ocupación */
  Core.parseCargaPrevia = function (aoa) {
    const hr = findHeaderRow(aoa, ['CURP', 'SUCURSAL', 'CLAVE ESTABLECIMIENTO']);
    if (hr < 0) return null;
    const c = colFinder(aoa[hr]);
    const si = c('SUCURSAL'), ki = c('CLAVE ESTABLECIMIENTO'), pi = c('PUESTO'), oi = c('CLAVE OCUPACION');
    const sucursales = new Map(), puestos = new Map(), nombres = new Map();
    for (let i = hr + 1; i < aoa.length; i++) {
      const r = aoa[i] || [];
      const k = toKey(r[ki]);
      if (str(r[si]) && k !== null) { sucursales.set(norm(r[si]), k); nombres.set(norm(r[si]), str(r[si])); }
      const o = oi >= 0 ? toKey(r[oi]) : null;
      if (pi >= 0 && str(r[pi]) && o !== null) puestos.set(norm(r[pi]), o);
    }
    return { sucursales, puestos, nombres };
  };
  Core.mergePrevias = function (claves, previas) {
    for (const [k, v] of previas.sucursales) if (!claves.sucursales.map.has(k)) claves.sucursales.map.set(k, v);
    for (const [k, v] of previas.nombres) if (!claves.sucursales.names.has(k)) claves.sucursales.names.set(k, v);
    for (const [k, v] of previas.puestos) if (!claves.puestos.map.has(k)) claves.puestos.map.set(k, v);
  };
  Core.estadoNorm = (s) => { const n = norm(String(s || '').replace(/^\d+\s*/, '')); return ESTADO_ALIAS[n] || n; };

  /* Escolaridad en texto libre -> nivel de la hoja Claves */
  Core.inferNivel = function (esc) {
    const n = norm(esc);
    if (!n) return null;
    if (/SIN ESTUD|NINGUN|NO ESTUD|SIN INSTRUC/.test(n)) return 'SIN ESTUDIOS';
    if (/PRIMARIA/.test(n)) return 'PRIMARIA';
    if (/SECUNDARIA/.test(n)) return 'SECUNDARIA';
    if (/MAESTR|POSGRADO|MBA|DOCTORAD|ESPECIALIDAD/.test(n)) return 'MAESTRIA';
    if (/BACHILLER|PREPARATORIA|\bPREPA\b|CARRERA TECNICA|TECNICO|COMERCIO|\bCBTIS\b|\bCONALEP\b|\bCETIS\b/.test(n)) return 'BACHILLERATO';
    if (/\bLIC\b|LICENCIATURA|\bING\b|INGENIER|CONTADOR|\bC\.?P\.?\b|ARQUITECT|\bARQ\b|ABOGADO|DERECHO|MEDIC|CIRUJANO|PSICOLOG|ENFERMER|QUIMIC|BIOLOG|NUTRI|ADMINISTRACION|ADMON|GASTRONOM|PEDAGOG|TRABAJO SOCIAL|MERCADOTECNIA|COMUNICACION|CONTADURIA|INFORMATICA|SISTEMAS|TURISMO/.test(n)) return 'LICENCIATURA';
    return null;
  };

  /* ---------- cruce principal ---------- */
  const CURP_RE = /^[A-Z]{4}\d{6}[HM][A-Z]{2}[B-DF-HJ-NP-TV-Z]{3}[A-Z0-9]\d$/;
  const emptyOv = () => ({ puesto: {}, escolaridad: {}, nivel: {}, sucursal: {}, estadoMun: {}, empresaAlias: {}, cursoClave: {}, fechas: {}, curp: {}, horas: {} });
  Core.emptyOv = emptyOv;
  Core.LAYOUT = ['CURP', 'NOMBRE', 'PRIMER APELLIDO', 'SEGUNDO APELLIDO', 'CLAVE ESTADO', 'ESTADO', 'CLAVE MUNICIPIO', 'MUNICIPIO',
    'CLAVE OCUPACION', 'PUESTO', 'CLAVE NIV ESTUDIOS', 'NIVEL DE ESTUDIO', 'CLAVE DOC PROBATORIO', 'CLAVE INSTITUCION',
    'CLAVE CURSO', 'NOMBRE CURSO', 'CLAVE AREA TEMATICA (3134)', 'DURACION', 'FEC INICIO', 'FEC TERMINO', 'CLAVE TIP AGENTE',
    'RFC AGENTE STPS', 'CLAVE MODALIDAD', 'CLAVE CAPACITACION (OBJETIVO)', 'CLAVE ESTABLECIMIENTO', 'SUCURSAL', 'Empresa'];

  Core.build = function (ctx) {
    const { reportRows, catalogo, claves, clavesCursos } = ctx;
    const ov = Object.assign(emptyOv(), ctx.ov || {});
    const exCursos = new Set((ctx.excluirCursos || []).map(norm));
    const exEmpresas = new Set((ctx.excluirEmpresas || []).map(norm));
    const records = [], bajas = [], seen = new Set();
    let excluidos = 0, duplicados = 0;

    for (const r of reportRows) {
      if (exCursos.has(norm(r.curso)) || exEmpresas.has(norm(r.compania))) { excluidos++; continue; }
      const cat = catalogo.get(r.id);
      if (!cat) { bajas.push(r); continue; }

      const errors = [], warnings = [];
      const rowId = `${r.src || ''}:${r.rowNum}`;
      const ovF = ov.fechas[rowId];
      let dates = resolveDates(r.fiRaw, r.ffRaw, r.mes, r.anio);
      if (ovF) dates = { fi: parseDate(ovF.fi), ff: parseDate(ovF.ff || ovF.fi), fixed: false, dudosa: false, motivo: '', manual: true };
      if (dates.dudosa) errors.push({ code: 'FECHA', key: rowId, label: dates.motivo });

      const curp = ov.curp[r.id] || cat.curp;
      if (!CURP_RE.test(curp)) errors.push({ code: 'CURP', key: r.id, label: curp || '(vacía)' });

      const estadoN = Core.estadoNorm(cat.estado || r.estado);
      const estadoNombre = claves.estados.names.get(estadoN) || estadoN;
      let claveEstado = ov.estadoMun[estadoN] && ov.estadoMun[estadoN].claveEstado ? ov.estadoMun[estadoN].claveEstado : claves.estados.map.get(estadoN);
      if (claveEstado === undefined && ESTADO_CLAVE[estadoN]) { claveEstado = ESTADO_CLAVE[estadoN]; warnings.push({ code: 'ESTADO_INTERNO', label: `Clave de estado ${estadoN} tomada de la tabla interna` }); }
      if (claveEstado === undefined) errors.push({ code: 'ESTADO', key: estadoN, label: estadoN || '(vacío)' });
      let munNombre = '', claveMun = null;
      const mo = ov.estadoMun[estadoN] && ov.estadoMun[estadoN].municipio ? ov.estadoMun[estadoN] : null;
      const mp = mo ? [mo.municipio, mo.clave] : MUN_PRINCIPAL[estadoN];
      if (mp) {
        munNombre = mp[0];
        const k = claves.municipios.map.get(norm(mp[0]));
        if (mo && mo.clave) claveMun = mo.clave;
        else if (k !== undefined) claveMun = k;
        else { claveMun = mp[1]; warnings.push({ code: 'MUN_INTERNO', label: `Clave de municipio ${mp[0]} tomada de la tabla interna` }); }
      } else if (estadoN) errors.push({ code: 'MUN', key: estadoN, label: `Sin municipio principal para ${estadoN}` });

      const puestoTxt = cat.puesto || r.puesto; const puestoN = norm(puestoTxt);
      const claveOcup = ov.puesto[puestoN] ?? claves.puestos.map.get(puestoN);
      if (claveOcup === undefined || claveOcup === null) errors.push({ code: 'PUESTO', key: puestoN, label: puestoTxt || '(vacío)' });

      const escN = norm(cat.escolaridad);
      let nivelNombre = ov.escolaridad[escN] || (claves.niveles.has(escN) ? escN : Core.inferNivel(cat.escolaridad));
      const nk = nivelNombre ? (ov.nivel[nivelNombre] || claves.niveles.get(nivelNombre)) : null;
      if (!nk || nk.nivel === null || nk.nivel === undefined) errors.push({ code: 'NIVEL', key: escN, label: cat.escolaridad || '(vacía)' });

      const sucTxt = cat.sucursal || r.sucursal; const sucN = norm(sucTxt);
      const claveSuc = ov.sucursal[sucN] ?? claves.sucursales.map.get(sucN);
      if (claveSuc === undefined || claveSuc === null) errors.push({ code: 'SUCURSAL', key: sucN, label: sucTxt || '(vacía)' });

      const empresaBase = ov.empresaAlias[norm(r.compania)] || r.compania;
      let claveCurso = null, cc = null;
      if (dates.fi) {
        const ckey = Core.cursoKey(r.curso, dates.fi, empresaBase);
        cc = clavesCursos.map.get(ckey);
        const man = ov.cursoClave[ckey];
        if (man) { claveCurso = man; if (!cc) cc = { tip: 'INTERNO', modalidad: 'MIXTA', objetivo: 'ACTUALIZAR Y PERFECCIONAR' }; }
        else if (cc) claveCurso = cc.clave;
        else errors.push({ code: 'CURSO', key: ckey, label: `${r.curso} · ${ymdStr(dates.fi)} · ${empresaBase}` });
      }
      const horas = Number(ov.horas[rowId] !== undefined ? ov.horas[rowId] : r.horas);
      if (!(horas > 0)) errors.push({ code: 'HORAS', key: rowId, label: `Horas inválidas (${r.horas})` });
      else if (cc && cc.duracion && Number(cc.duracion) !== horas) warnings.push({ code: 'DURACION', label: `Horas del reporte (${horas}) distintas a las de claves de curso (${cc.duracion})` });

      const dk = `${curp}|${norm(r.curso)}|${ymdStr(dates.fi)}`;
      if (seen.has(dk)) { duplicados++; continue; }
      seen.add(dk);

      /* Fecha como número de serie de Excel (entero): evita desfases de zona horaria al escribir el archivo */
      const dt = (x) => (x ? Math.round(Date.UTC(x.y, x.m - 1, x.d) / 86400000) + 25569 : null);
      records.push({
        rowId, id: r.id, empresa: r.compania, curso: r.curso, sucursal: sucTxt,
        nombreCompleto: `${r.apPat} ${r.apMat} ${r.nombre}`.replace(/\s+/g, ' ').trim(),
        fi: dates.fi, ff: dates.ff, fechaCorregida: dates.fixed, errors, warnings,
        puesto: puestoTxt, instructor: r.instructor, curp, horas,
        layout: [
          curp, r.nombre, r.apPat, r.apMat, claveEstado ?? '', estadoNombre, claveMun ?? '', munNombre,
          claveOcup ?? '', puestoTxt, nk && nk.nivel !== null ? nk.nivel : '', nivelNombre ? (claves.niveles.get(nivelNombre) || { nombre: nivelNombre }).nombre : '',
          nk && nk.doc !== null ? nk.doc : '', nk && nk.inst !== null ? nk.inst : '',
          claveCurso ?? '', r.curso, 3134, horas > 0 ? horas : '', dt(dates.fi), dt(dates.ff),
          cc ? cc.tip : '', '', cc ? cc.modalidad : '', cc ? cc.objetivo : '', claveSuc ?? '', sucTxt, r.compania,
        ],
      });
    }
    return { records, bajas, excluidos, duplicados };
  };

  /* Agrupa los errores por tipo y valor para corregirlos de una vez */
  Core.groupErrors = function (records) {
    const groups = new Map();
    for (const rec of records) for (const e of rec.errors) {
      const g = `${e.code}|${e.key}`;
      if (!groups.has(g)) groups.set(g, { code: e.code, key: e.key, label: e.label, count: 0, rows: [] });
      const o = groups.get(g); o.count++; o.rows.push(rec);
    }
    return [...groups.values()].sort((a, b) => b.count - a.count);
  };

  /* Solo filas sin errores, agrupadas por empresa (una pestaña por empresa) */
  Core.sheetsByEmpresa = function (records) {
    const by = new Map();
    for (const r of records.filter((x) => x.errors.length === 0)) {
      if (!by.has(r.empresa)) by.set(r.empresa, []);
      by.get(r.empresa).push(r);
    }
    const out = [];
    for (const [empresa, list] of by) {
      list.sort((a, b) => a.curso.localeCompare(b.curso) || ymdStr(a.fi).localeCompare(ymdStr(b.fi)) || a.nombreCompleto.localeCompare(b.nombreCompleto));
      out.push({ empresa, rows: list.map((x) => x.layout) });
    }
    return out.sort((a, b) => a.empresa.localeCompare(b.empresa));
  };
  Core.sheetName = function (name, used) {
    let s = String(name).replace(/[\\\/?*\[\]:]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 31) || 'Hoja';
    let out = s, i = 2;
    while (used.has(out.toUpperCase())) { const suf = ' ' + i++; out = s.slice(0, 31 - suf.length) + suf; }
    used.add(out.toUpperCase());
    return out;
  };

  /* Sugerencia de sucursal por similitud de palabras */
  Core.sugerirSucursal = function (texto, nombres) {
    const tok = (s) => new Set(norm(s).replace(/[^A-Z0-9 ]/g, ' ').split(' ').filter((w) => w.length > 1 && !['DE', 'LA', 'EL', 'LOS', 'LAS', 'EST', 'SERV'].includes(w)));
    const a = tok(texto); if (!a.size) return null;
    let best = null, bs = 0;
    for (const n of nombres) {
      const b = tok(n); if (!b.size) continue;
      let inter = 0; for (const w of a) if (b.has(w)) inter++;
      const score = inter / Math.max(a.size, b.size);
      if (score > bs) { bs = score; best = n; }
    }
    return bs >= 0.6 ? { nombre: best, score: bs } : null;
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = Core; else root.Core = Core;
})(typeof window !== 'undefined' ? window : globalThis);
