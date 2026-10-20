/* Generador del formato DC-3 (Constancia de competencias o de habilidades laborales) en PDF.
 * Recibe el constructor de jsPDF para funcionar en el navegador y en Node. Unidades: puntos, tamaño carta. */
(function (root) {
  'use strict';
  const DC3 = {};
  const X = 30, W = 552;

  const AREAS_L = [['01', 'Cultivo, crianza y aprovechamiento'], ['01.1', 'Agricultura y silvicultura'], ['01.2', 'Ganadería'], ['01.3', 'Pesca y acuacultura'], null,
    ['02', 'Extracción y suministro'], ['02.1', 'Exploración'], ['02.2', 'Extracción'], ['02.3', 'Refinación y beneficio'], ['02.4', 'Provisión de energía'], ['02.5', 'Provisión de agua'], null,
    ['03', 'Construcción'], ['03.1', 'Planeación y dirección de obras'], ['03.2', 'Edificación y urbanización'], ['03.3', 'Acabado'], ['03.4', 'Instalación y mantenimiento'], null,
    ['04', 'Tecnología'], ['04.1', 'Mecánica'], ['04.2', 'Electricidad'], ['04.3', 'Electrónica'], ['04.4', 'Informática'], ['04.5', 'Telecomunicaciones'], ['04.6', 'Procesos industriales'], null,
    ['05', 'Procesamiento y fabricación'], ['05.1', 'Minerales no metálicos'], ['05.2', 'Metales'], ['05.3', 'Alimentos y bebidas'], ['05.4', 'Textiles y prendas de vestir'], ['05.5', 'Materia orgánica'],
    ['05.6', 'Productos químicos'], ['05.7', 'Productos metálicos y de hule y plástico'], ['05.8', 'Productos eléctricos y electrónicos'], ['05.9', 'Productos impresos']];
  const AREAS_R = [['06', 'Transporte'], ['06.1', 'Ferroviario'], ['06.2', 'Autotransporte'], ['06.3', 'Aéreo'], ['06.4', 'Marítimo y fluvial'], ['06.5', 'Servicios de apoyo'], null,
    ['07', 'Provisión de bienes y servicios'], ['07.1', 'Comercio'], ['07.2', 'Alimentación y hospedaje'], ['07.3', 'Turismo'], ['07.4', 'Deporte y esparcimiento'], ['07.5', 'Servicios personales'],
    ['07.6', 'Reparación de artículos de uso doméstico y personal'], ['07.7', 'Limpieza'], ['07.8', 'Servicio postal y mensajería'], null,
    ['08', 'Gestión y soporte administrativo'], ['08.1', 'Bolsa, banca y seguros'], ['08.2', 'Administración'], ['08.3', 'Servicios legales'], null,
    ['09', 'Salud y protección social'], ['09.1', 'Servicios médicos'], ['09.2', 'Inspección sanitaria y del medio ambiente'], ['09.3', 'Seguridad social'], ['09.4', 'Protección de bienes y/o personas'], null,
    ['10', 'Comunicación'], ['10.1', 'Publicación'], ['10.2', 'Radio, cine, televisión y teatro'], ['10.3', 'Interpretación artística'], ['10.4', 'Traducción e interpretación lingüística'],
    ['10.5', 'Publicidad, propaganda y relaciones públicas'], null, ['11', 'Desarrollo y extensión del conocimiento'], ['11.1', 'Investigación'], ['11.2', 'Enseñanza'], ['11.3', 'Difusión cultural']];
  const TEM_L = [['1000', 'Producción general'], ['2000', 'Servicios'], ['3000', 'Administración, contabilidad y economía'], ['4000', 'Comercialización'], ['5000', 'Mantenimiento y reparación']];
  const TEM_R = [['6000', 'Seguridad'], ['7000', 'Desarrollo personal y familiar'], ['8000', 'Uso de tecnologías de la información y comunicación'], ['9000', 'Participación social']];

  const pad2 = (n) => String(n).padStart(2, '0');

  function shrink(doc, text, maxW, size, min) {
    let s = size; doc.setFontSize(s);
    while (s > (min || 6) && doc.getTextWidth(text) > maxW) { s -= 0.5; doc.setFontSize(s); }
    return s;
  }
  function cell(doc, x, y, w, h, label, lsize) {
    doc.rect(x, y, w, h);
    if (label) { doc.setFont('helvetica', 'normal'); doc.setFontSize(lsize || 7.5); doc.text(label, x + 3, y + 9.5); }
  }
  function bar(doc, y, text) {
    doc.setFillColor(0, 0, 0); doc.rect(X, y, W, 15, 'F');
    doc.setTextColor(255, 255, 255); doc.setFont('helvetica', 'bold'); doc.setFontSize(9);
    doc.text(text, X + W / 2, y + 10.5, { align: 'center' });
    doc.setTextColor(0, 0, 0); doc.setFont('helvetica', 'normal');
  }
  /* Casillas: una por carácter, dentro de la franja [y, y+h] */
  function boxes(doc, x, y, w, h, chars, size) {
    doc.setFont('helvetica', 'normal'); doc.setFontSize(size || 9);
    const n = chars.length;
    for (let i = 0; i < n; i++) {
      const bx = x + i * w;
      doc.rect(bx, y, w, h);
      if (chars[i] !== ' ') doc.text(String(chars[i]), bx + w / 2, y + h / 2 + 3.2, { align: 'center' });
    }
  }
  const digits = (s, n) => String(s).padStart(n, '0').split('');

  function dateGroup(doc, x, y, d) {
    const bw = 22.5;
    const heads = [['Año', 4], ['Mes', 2], ['Día', 2]];
    let hx = x;
    doc.setFontSize(7.5);
    for (const [t, n] of heads) { doc.rect(hx, y, bw * n, 14); doc.text(t, hx + (bw * n) / 2, y + 10, { align: 'center' }); hx += bw * n; }
    boxes(doc, x, y + 14, bw, 32, [...digits(d.y, 4), ...digits(d.m, 2), ...digits(d.d, 2)], 10);
  }

  DC3.nombreArchivo = function (suc, nombre, curso) {
    const clean = (s) => String(s).replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim();
    return `${clean(suc)} - ${clean(nombre)} - ${clean(curso)} - DC3`.slice(0, 180).trim();
  };

  /* d: { nombre, curp, ocupacion, puesto, razon, rfc, curso, horas, fi:{y,m,d}, ff:{y,m,d}, agente, instructor, patron, representante } */
  DC3.render = function (doc, d) {
    doc.setLineWidth(0.6); doc.setDrawColor(0, 0, 0); doc.setTextColor(0, 0, 0);
    doc.setFont('helvetica', 'bold'); doc.setFontSize(12);
    doc.text('FORMATO DC-3', X + W / 2, 36, { align: 'center' });
    doc.text('CONSTANCIA DE COMPETENCIAS O DE HABILIDADES LABORALES', X + W / 2, 50, { align: 'center' });
    doc.setFont('helvetica', 'normal');

    bar(doc, 58, 'DATOS DEL TRABAJADOR');
    cell(doc, X, 73, W, 36, 'Nombre (Anotar apellido paterno, apellido materno y nombre (s))');
    shrink(doc, d.nombre, W - 14, 10, 7); doc.text(d.nombre, X + 6, 73 + 28);
    // CURP + ocupación
    cell(doc, X, 109, 372, 46, 'Clave Única de Registro de Población');
    boxes(doc, X, 109 + 20, 372 / 18, 26, String(d.curp || '').padEnd(18, ' ').slice(0, 18).split(''), 9.5);
    cell(doc, X + 372, 109, 180, 46, 'Ocupación específica (Catálogo Nacional de');
    doc.setFontSize(7.5); doc.text('Ocupaciones)1/', X + 375, 109 + 18.5);
    doc.setFontSize(10); doc.text(String(d.ocupacion ?? ''), X + 378, 109 + 38);
    cell(doc, X, 155, W, 34, 'Puesto*');
    shrink(doc, d.puesto || '', W - 14, 10, 7); doc.text(d.puesto || '', X + 6, 155 + 27);

    bar(doc, 201, 'DATOS DE LA EMPRESA');
    cell(doc, X, 216, W, 36, 'Nombre o razón social (En caso de persona física, anotar apellido paterno, apellido materno y nombre(s))');
    shrink(doc, d.razon || '', W - 14, 10, 7); doc.text(d.razon || '', X + 6, 216 + 28);
    cell(doc, X, 252, W, 44, 'Registro Federal de Contribuyentes con homoclave (SHCP)');
    const rfc = String(d.rfc || '').toUpperCase().replace(/[^A-Z0-9&Ñ]/g, '');
    const letras = rfc.length === 13 ? 4 : 3;
    const rfcFmt = rfc.length >= 12 ? `${rfc.slice(0, letras)}-${rfc.slice(letras, letras + 6)}-${rfc.slice(letras + 6)}` : rfc;
    boxes(doc, X + 34, 252 + 17, 22, 22, rfcFmt.split(''), 10);

    bar(doc, 308, 'DATOS DEL PROGRAMA DE CAPACITACIÓN, ADIESTRAMIENTO Y PRODUCTIVIDAD');
    cell(doc, X, 323, W, 34, 'Nombre del curso');
    shrink(doc, d.curso || '', W - 14, 10, 7); doc.text(d.curso || '', X + 6, 323 + 27);
    // duración y periodo
    cell(doc, X, 357, 90, 46, 'Duración en horas');
    doc.setFontSize(10); doc.text(String(d.horas ?? ''), X + 6, 357 + 37);
    cell(doc, X + 90, 357, 78, 46, '');
    doc.setFontSize(7.5); doc.text('Periodo de', X + 93, 357 + 26); doc.text('ejecución:', X + 93, 357 + 35);
    doc.text('De', X + 90 + 78 - 5, 357 + 43, { align: 'right' });
    dateGroup(doc, X + 168, 357, d.fi);
    cell(doc, X + 348, 357, 24, 46, '');
    doc.setFontSize(7.5); doc.text('a', X + 348 + 12, 357 + 43, { align: 'center' });
    dateGroup(doc, X + 372, 357, d.ff);
    cell(doc, X, 403, W, 30, 'Área temática del curso 2/');
    doc.setFontSize(10); doc.text('3134', X + 6, 403 + 25);
    cell(doc, X, 433, W, 34, 'Nombre del agente capacitador o STPS 3/');
    if (d.agente) { doc.setFontSize(10); doc.text(d.agente, X + 6, 433 + 28); }

    // declaración y firmas
    const sy = 481, sh = 178;
    doc.rect(X, sy, W, sh);
    doc.setFont('helvetica', 'bold'); doc.setFontSize(8.5);
    doc.text('Los datos se asientan en esta constancia bajo protesta de decir verdad, apercibidos de la responsabilidad en que incurre todo', X + W / 2, sy + 22, { align: 'center' });
    doc.text('aquel que no se conduce con verdad.', X + W / 2, sy + 34, { align: 'center' });
    doc.setFont('helvetica', 'normal');
    const cols = [[X + 92, 'Instructor o tutor', d.instructor], [X + 276, 'Patrón o representante legal 4/', d.patron], [X + 460, 'Representante de los trabajadores 5/', d.representante]];
    for (const [cx, label, name] of cols) {
      doc.setFontSize(8.5); doc.text(label, cx, sy + 110, { align: 'center' });
      if (name) {
        doc.setFontSize(8); const lines = doc.splitTextToSize(String(name), 168);
        doc.text(lines.slice(0, 2), cx, sy + 128, { align: 'center' });
      }
      doc.line(cx - 86, sy + 144, cx + 86, sy + 144);
      doc.setFontSize(8.5); doc.text('Nombre y firma', cx, sy + 157, { align: 'center' });
    }

    // instrucciones
    let y = 674; doc.setFont('helvetica', 'bold'); doc.setFontSize(7); doc.text('INSTRUCCIONES', X, y); doc.setFont('helvetica', 'normal'); doc.setFontSize(6.5);
    const ins = ['- Llenar a máquina o con letra de molde.',
      '- Deberá entregarse al trabajador dentro de los veinte días hábiles siguientes al término del curso de capacitación aprobado.',
      '1/ Las áreas y subáreas ocupacionales del Catálogo Nacional de Ocupaciones se encuentran disponibles en el reverso de este formato y en la página www.stps.gob.mx',
      '2/ Las áreas temáticas de los cursos se encuentran disponibles en el reverso de este formato y en la página www.stps.gob.mx',
      '3/ Cursos impartidos por el área competente de la Secretaría del Trabajo y Previsión Social.',
      '4/ Para empresas con menos de 51 trabajadores. Para empresas con más de 50 trabajadores firmaría el representante del patrón ante la Comisión mixta de capacitación, adiestramiento y productividad.',
      '5/ Solo para empresas con más de 50 trabajadores.', '* Dato no obligatorio.'];
    for (const t of ins) for (const ln of doc.splitTextToSize(t, 490)) { y += 8.2; doc.text(ln, X, y); }
    doc.setFontSize(7); doc.text('DC-3', X + W, y - 8.2, { align: 'right' }); doc.text('ANVERSO', X + W, y, { align: 'right' });

    // reverso
    doc.addPage('letter', 'portrait');
    doc.setFont('helvetica', 'bold'); doc.setFontSize(8);
    doc.text('CLAVES Y DENOMINACIONES DE ÁREAS Y SUBÁREAS DEL CATÁLOGO NACIONAL DE OCUPACIONES', X + W / 2, 52, { align: 'center' });
    const colHead = (x, y2) => { doc.setFontSize(6.5); doc.text('CLAVE DEL', x, y2); doc.text('ÁREA/SUBÁREA', x, y2 + 7); doc.text('DENOMINACIÓN', x + 75, y2 + 7); };
    colHead(X + 20, 74); colHead(X + 290, 74);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(7);
    const list = (arr, x, y0) => { let yy = y0; for (const it of arr) { if (it) { doc.text(it[0], x + 18, yy); doc.text(it[1], x + 93, yy); } yy += 10.4; } };
    list(AREAS_L, X + 2, 100); list(AREAS_R, X + 272, 100);
    const ty = 100 + Math.max(AREAS_L.length, AREAS_R.length) * 10.4 + 14;
    doc.setFont('helvetica', 'bold'); doc.setFontSize(8);
    doc.text('CLAVES Y DENOMINACIONES DEL CATÁLOGO DE ÁREAS TEMÁTICAS DE LOS CURSOS', X + W / 2, ty, { align: 'center' });
    doc.setFontSize(6.5); const hy = ty + 16;
    doc.text('CLAVE DEL', X + 20, hy); doc.text('ÁREA', X + 20, hy + 7); doc.text('DENOMINACIÓN', X + 95, hy + 7);
    doc.text('CLAVE DEL', X + 290, hy); doc.text('ÁREA', X + 290, hy + 7); doc.text('DENOMINACIÓN', X + 365, hy + 7);
    doc.setFont('helvetica', 'normal'); doc.setFontSize(7);
    list(TEM_L, X + 2, hy + 24); list(TEM_R, X + 272, hy + 24);
    doc.setFontSize(7); doc.text('DC-3', X + W, hy + 24 + 5 * 10.4 + 14, { align: 'right' }); doc.text('REVERSO', X + W, hy + 24 + 5 * 10.4 + 23, { align: 'right' });
    return doc;
  };

  DC3.crear = function (jsPDF, d) {
    const doc = new jsPDF({ unit: 'pt', format: 'letter', orientation: 'portrait', compress: true });
    return DC3.render(doc, d);
  };

  if (typeof module !== 'undefined' && module.exports) module.exports = DC3; else root.DC3 = DC3;
})(typeof window !== 'undefined' ? window : globalThis);
