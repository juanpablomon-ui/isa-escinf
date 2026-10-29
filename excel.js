// El Excel de las ISA: armarlo desde el PowerPoint (con listas desplegables) y leerlo.
// Para leer se usa SheetJS (lib/xlsx.full.min.js). Para escribir, este archivo arma el .xlsx a mano,
// porque SheetJS gratuito no guarda las listas desplegables.
(function (raiz) {
  'use strict';
  const enNode = typeof module !== 'undefined' && module.exports;
  const L = enNode ? require('./leer_ppt.js') : raiz.LeerPpt;
  const P = enNode ? require('./plazos.js') : raiz.Plazos;
  const { normal } = L;

  const HOJA_ISA = 'ISA', HOJA_DIL = 'Diligencias', HOJA_JUD = 'Procesos judiciales', HOJA_AYUDA = 'Instrucciones';
  const LISTAS = {
    tipo: ['MÉDICA', 'LICENCIAS MÉDICAS', 'ACCIDENTE', 'MÉDICA/ACCIDENTE', 'DISCIPLINARIA', 'PÉRDIDA DE MATERIAL', 'PÉRDIDA Y/O ROBO'],
    estado: ['ABIERTA', 'EN INVESTIGACIÓN', 'EN TRÁMITE', 'CERRADA'],
    siNo: ['SÍ', 'NO', 'X'],
    hecha: ['Hecha', 'Pendiente']
  };
  // Columnas (el título es lo que se busca al leer; el orden puede cambiar en el Excel)
  const COL_ISA = [
    { k: 'n', t: 'N.°', ancho: 6, tipo: 'numero' },
    { k: 'titulo', t: 'Afectado / ISA', ancho: 34 },
    { k: 'unidad', t: 'Unidad', ancho: 20 },
    { k: 'tipo', t: 'Tipo de ISA', ancho: 18, lista: LISTAS.tipo },
    { k: 'inicio', t: 'Fecha de inicio', ancho: 13, tipo: 'fecha' },
    { k: 'estado', t: 'Estado', ancho: 17, lista: LISTAS.estado },
    { k: 'csDiveduc', t: 'CS DIVEDUC', ancho: 11, lista: LISTAS.siNo },
    { k: 'cse', t: 'CSE', ancho: 8, lista: LISTAS.siNo },
    { k: 'dictamen', t: 'Dictamen', ancho: 10, lista: LISTAS.siNo },
    { k: 'fiscal', t: 'Fiscal', ancho: 22 },
    { k: 'secretario', t: 'Secretario', ancho: 20 },
    { k: 'prorrogas', t: 'N.° prórrogas', ancho: 11, tipo: 'numero' },
    { k: 'vence', t: 'Vence la prórroga', ancho: 13, tipo: 'fecha' },
    { k: 'resumen', t: 'Resumen del hecho', ancho: 70 },
    { k: 'notas', t: 'Notas', ancho: 40 }
  ];
  const COL_DIL = [
    { k: 'n', t: 'N.° ISA', ancho: 8, tipo: 'numero' },
    { k: 'afectado', t: 'Afectado (se llena solo)', ancho: 30, auto: true },
    { k: 'fecha', t: 'Fecha', ancho: 13, tipo: 'fecha' },
    { k: 'tipo', t: 'Tipo', ancho: 20, lista: P.TIPOS },
    { k: 'hecha', t: '¿Hecha o pendiente?', ancho: 12, lista: LISTAS.hecha },
    { k: 'detalle', t: 'Detalle', ancho: 80 },
    { k: 'revisar', t: 'Revisar', ancho: 32 }
  ];
  const FILAS_DIL = 600; // filas con la fórmula del afectado ya puesta

  // ---------- Fechas de Excel ----------
  const BASE = Date.UTC(1899, 11, 30);
  const aSerie = (iso) => { const [a, m, d] = iso.split('-').map(Number); return (Date.UTC(a, m - 1, d) - BASE) / 86400000; };
  const deSerie = (n) => new Date(BASE + Math.round(n) * 86400000).toISOString().slice(0, 10);
  function fechaDeCelda(v, anioRef) {
    if (v == null || v === '') return '';
    if (typeof v === 'number') return v > 1000 ? deSerie(v) : '';
    if (v instanceof Date) return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, '0')}-${String(v.getDate()).padStart(2, '0')}`;
    const t = String(v).trim();
    const dmy = t.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
    if (dmy) return `${dmy[3]}-${dmy[2].padStart(2, '0')}-${dmy[1].padStart(2, '0')}`;
    const f = L.leerFecha(t, anioRef);
    return f ? f.iso : '';
  }

  // ---------- Del PowerPoint al Excel ----------
  const deLista = (valor, lista) => lista.find((x) => normal(x) === normal(valor).replace(/\s+/g, ' ').trim());
  const casilla = (v) => { const t = normal(v).trim(); return t === 'SI' ? 'SÍ' : t === 'NO' ? 'NO' : t === 'X' ? 'X' : ''; };
  function tipoIsa(v) {
    const t = normal(v).replace(/\s+/g, ' ').trim();
    if (t === 'PERDIDA MATERIAL') return 'PÉRDIDA DE MATERIAL';
    return deLista(v, LISTAS.tipo) || String(v || '').toUpperCase();
  }

  function armarLibro(datos) {
    const anioRef = +(datos.fechaExpo || new Date().toISOString()).slice(0, 4);
    // El mismo fiscal escrito de varias formas → la forma más completa
    const nombreFiscal = {};
    for (const f of datos.fichas) if ((f.fiscal || '').length > (nombreFiscal[f.fiscalClave] || '').length) nombreFiscal[f.fiscalClave] = f.fiscal;

    const filasIsa = [], filasDil = [];
    datos.fichas.forEach((f0, i) => {
      const f = P.completar({ ...f0 }, anioRef);
      const n = i + 1, notas = [];
      if (f.inicioSinAnio) notas.push(`Fecha de inicio sin año en el PowerPoint («${f.inicioTexto}»): se supuso ${f.inicio.slice(0, 4)}.`);
      if (f.inicioTexto && !f.inicio) notas.push(`Fecha de inicio no reconocida: «${f.inicioTexto}».`);
      filasIsa.push({
        n, titulo: f.titulo, unidad: f.unidad, tipo: tipoIsa(f.tipo), inicio: f.inicio,
        estado: deLista(f.estado, LISTAS.estado) || String(f.estado || '').toUpperCase(),
        csDiveduc: casilla(f.csDiveduc), cse: casilla(f.cse), dictamen: casilla(f.dictamen),
        fiscal: nombreFiscal[f.fiscalClave] || f.fiscal, secretario: f.secretario,
        prorrogas: f.prorrogas, vence: f.vence, resumen: f.resumen, notas: notas.join(' ')
      });
      for (const d of f.diligencias) {
        const revisar = [];
        if (!d.fecha && d.hecha) revisar.push('Sin fecha: póngala si la sabe.');
        if (d.sinAnio) revisar.push(`Fecha sin año: se supuso ${d.fecha.slice(0, 4)}.`);
        if ((normal(d.texto).match(/\d{1,2}(ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)/g) || []).length > 1) revisar.push('El texto trae dos fechas: revise cuál es la del trámite.');
        filasDil.push({ n, afectado: f.titulo, fecha: d.fecha, tipo: d.tipo, hecha: d.hecha ? 'Hecha' : 'Pendiente', detalle: d.resto || d.texto, revisar: revisar.join(' ') });
      }
    });

    const hojas = [
      { nombre: HOJA_ISA, columnas: COL_ISA, filas: filasIsa, filtro: true, extra: 300 },
      { nombre: HOJA_DIL, columnas: COL_DIL, filas: filasDil, filtro: true, extra: FILAS_DIL - filasDil.length },
    ];
    for (const t of datos.tablas || []) {
      hojas.push({
        nombre: HOJA_JUD,
        columnas: t.titulos.map((c, i) => ({ k: i, t: c, ancho: [6, 8, 30, 14, 12, 40, 60][i] || 20 })),
        filas: t.filas.map((fila) => Object.assign({}, fila)),
        extra: 50
      });
      break;
    }
    hojas.push({ nombre: HOJA_AYUDA, ayuda: true });
    return hojas;
  }

  const AYUDA = [
    ['ISA ESCINF — cómo llenar este Excel', 1],
    [''],
    ['👉 Hoja «ISA»: una fila por investigación sumaria. El «N.°» no se repite: con él se unen las diligencias.'],
    ['👉 Hoja «Diligencias»: una fila por diligencia. Escriba el N.° de la ISA y el nombre aparece solo.'],
    ['   · Fecha: el día del trámite (escriba 24-10-2026). En «Citación CSE», el día que debe presentarse.'],
    ['   · Tipo: elíjalo de la lista. De él dependen los plazos.'],
    ['   · «Hecha» = ya se hizo (últimas diligencias). «Pendiente» = por hacer (próximas diligencias).'],
    ['👉 Hoja «Procesos judiciales»: igual que la tabla de la última lámina.'],
    [''],
    ['Tipos de diligencia', 1],
    ['   Solicitud CSE · Reiteración CSE: pedir hora a la CSE. Se repite cada 30 días hasta que lo citen.'],
    ['   Citación CSE: la CSE dio hora. Con fecha futura, no corre el plazo de 30 días.'],
    ['   Asistencia CSE: el afectado fue a la CSE. Desde ahí se pide el informe.'],
    ['   Pide informe CSE: pedir o reiterar el informe a la CSE (cuenta como gestión de 15 días).'],
    ['   Informe CSE recibido: llegó el informe.'],
    ['   Prórroga: pedir prórroga (cuenta como gestión de 15 días). La prórroga dura hasta 20 días.'],
    ['   Otra gestión: todo lo demás (oficios, declaraciones, CSD, antecedentes…). Cada 15 días.'],
    [''],
    ['Plazos que revisa la app', 1],
    ['   · La prórroga dura hasta 20 días: «Vence la prórroga» no puede estar a más de 20 días de cuando se pidió.'],
    ['   · Mientras no lo citen: solicitud o reiteración a la CSE cada 30 días.'],
    ['   · Todas las demás gestiones (incluidas prórroga y pedir el informe a la CSE): cada 15 días.'],
    [''],
    ['Este Excel se armó solo desde el PowerPoint: los tipos de diligencia se asignaron por las palabras.', 1],
    ['Revise la columna «Revisar» de la hoja «Diligencias» y la columna «Notas» de la hoja «ISA».'],
    ['No cambie los títulos de las columnas (la app los busca por su nombre). Puede agregar filas al final.']
  ];

  // ---------- Escribir el .xlsx ----------
  const xml = (t) => String(t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]))
    // Excel no acepta caracteres de control en el XML
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
  const letra = (i) => { let s = ''; i++; while (i) { const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); } return s; };
  const ESTILO = { titulo: 1, fecha: 2, texto: 3, auto: 4, numero: 5, grande: 6 };

  function hojaXml(h) {
    const filas = [];
    if (h.ayuda) {
      AYUDA.forEach(([t, negrita], i) => {
        if (t) filas.push(`<row r="${i + 1}"><c r="A${i + 1}" t="inlineStr" s="${negrita ? ESTILO.grande : 0}"><is><t xml:space="preserve">${xml(t)}</t></is></c></row>`);
      });
      return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="1" width="120" customWidth="1"/></cols><sheetData>${filas.join('')}</sheetData></worksheet>`;
    }
    const cols = h.columnas;
    filas.push(`<row r="1" ht="32" customHeight="1">${cols.map((c, j) => `<c r="${letra(j)}1" t="inlineStr" s="${ESTILO.titulo}"><is><t>${xml(c.t)}</t></is></c>`).join('')}</row>`);
    const total = h.filas.length + (h.extra || 0);
    for (let i = 0; i < total; i++) {
      const r = i + 2, fila = h.filas[i] || {}, celdas = [];
      cols.forEach((c, j) => {
        const ref = letra(j) + r, v = fila[c.k];
        if (c.auto) {
          const cache = v == null ? '' : v;
          celdas.push(`<c r="${ref}" t="str" s="${ESTILO.auto}"><f>IFERROR(VLOOKUP(A${r},${HOJA_ISA}!$A:$B,2,FALSE),"")</f><v>${xml(cache)}</v></c>`);
          return;
        }
        if (v == null || v === '') return;
        if (c.tipo === 'fecha') celdas.push(`<c r="${ref}" s="${ESTILO.fecha}"><v>${aSerie(v)}</v></c>`);
        else if (c.tipo === 'numero' && typeof v === 'number') celdas.push(`<c r="${ref}" s="${ESTILO.numero}"><v>${v}</v></c>`);
        else celdas.push(`<c r="${ref}" t="inlineStr" s="${ESTILO.texto}"><is><t xml:space="preserve">${xml(v)}</t></is></c>`);
      });
      if (celdas.length) filas.push(`<row r="${r}">${celdas.join('')}</row>`);
    }
    const ultima = total + 1;
    const val = [];
    cols.forEach((c, j) => {
      const rango = `${letra(j)}2:${letra(j)}${Math.max(ultima, 2)}`;
      if (c.lista) val.push(`<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorStyle="warning" errorTitle="Valor fuera de la lista" error="Elija un valor de la lista (puede dejarlo igual si es necesario)." sqref="${rango}"><formula1>"${xml(c.lista.join(','))}"</formula1></dataValidation>`);
      if (c.tipo === 'fecha') val.push(`<dataValidation type="date" operator="greaterThan" allowBlank="1" showErrorMessage="1" errorStyle="warning" errorTitle="Fecha" error="Escriba una fecha, por ejemplo 24-10-2026." sqref="${rango}"><formula1>36526</formula1></dataValidation>`);
    });
    if (h.nombre === HOJA_DIL) val.push(`<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorStyle="warning" error="Ese N.° no está en la hoja ISA." sqref="A2:A${ultima}"><formula1>${HOJA_ISA}!$A$2:$A$500</formula1></dataValidation>`);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>` +
      `<sheetFormatPr defaultRowHeight="15"/>` +
      `<cols>${cols.map((c, j) => `<col min="${j + 1}" max="${j + 1}" width="${c.ancho || 14}" customWidth="1"/>`).join('')}</cols>` +
      `<sheetData>${filas.join('')}</sheetData>` +
      (h.filtro ? `<autoFilter ref="A1:${letra(cols.length - 1)}${Math.max(h.filas.length + 1, 2)}"/>` : '') +
      (val.length ? `<dataValidations count="${val.length}">${val.join('')}</dataValidations>` : '') +
      `</worksheet>`;
  }

  const ESTILOS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="dd\\-mm\\-yyyy"/></numFmts>
<fonts count="3"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="13"/><name val="Calibri"/><family val="2"/></font></fonts>
<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF24364F"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFE4EAF3"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="left" vertical="top"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="top"/></xf>
<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  // CRC-32 (lo exige el .zip)
  const TABLA_CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = (b) => { let c = 0xffffffff; for (let i = 0; i < b.length; i++) c = TABLA_CRC[(c ^ b[i]) & 0xff] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; };

  function zip(archivos) {
    const enc = new TextEncoder(), partes = [], central = [];
    let pos = 0;
    const u16 = (v) => [v & 255, (v >>> 8) & 255], u32 = (v) => [v & 255, (v >>> 8) & 255, (v >>> 16) & 255, (v >>> 24) & 255];
    for (const [nombre, texto] of archivos) {
      const datos = enc.encode(texto), n = enc.encode(nombre), crc = crc32(datos);
      const comun = [...u16(20), ...u16(0x0800), ...u16(0), ...u16(0), ...u16(0x21), ...u32(crc), ...u32(datos.length), ...u32(datos.length), ...u16(n.length), ...u16(0)];
      partes.push(new Uint8Array([...u32(0x04034b50), ...comun]), n, datos);
      central.push(new Uint8Array([...u32(0x02014b50), ...u16(20), ...comun, ...u16(0), ...u16(0), ...u16(0), ...u32(0), ...u32(pos)]), n);
      pos += 30 + n.length + datos.length;
    }
    const largoCentral = central.reduce((s, b) => s + b.length, 0);
    const fin = new Uint8Array([...u32(0x06054b50), ...u16(0), ...u16(0), ...u16(archivos.length), ...u16(archivos.length), ...u32(largoCentral), ...u32(pos), ...u16(0)]);
    const todo = [...partes, ...central, fin];
    const salida = new Uint8Array(todo.reduce((s, b) => s + b.length, 0));
    let k = 0;
    for (const b of todo) { salida.set(b, k); k += b.length; }
    return salida;
  }

  function escribir(hojas) {
    const nombresDef = [];
    hojas.forEach((h, i) => {
      if (h.filtro) nombresDef.push(`<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">${h.nombre.includes(' ') ? `'${h.nombre}'` : h.nombre}!$A$1:$${letra(h.columnas.length - 1)}$${Math.max(h.filas.length + 1, 2)}</definedName>`);
    });
    const archivos = [
      ['[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${hojas.map((h, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`],
      ['_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`],
      ['xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${hojas.map((h, i) => `<sheet name="${xml(h.nombre)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${nombresDef.length ? `<definedNames>${nombresDef.join('')}</definedNames>` : ''}<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>`],
      ['xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${hojas.map((h, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${hojas.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],
      ['xl/styles.xml', ESTILOS],
      ...hojas.map((h, i) => [`xl/worksheets/sheet${i + 1}.xml`, hojaXml(h)])
    ];
    return zip(archivos);
  }

  // ---------- Leer el Excel ----------
  const claveTitulo = (t) => normal(t).replace(/[.:¿?()]/g, '').replace(/\s+/g, ' ').trim();
  function filasDe(XLSX, libro, nombre) {
    const hoja = libro.SheetNames.find((n) => claveTitulo(n) === claveTitulo(nombre));
    return hoja ? XLSX.utils.sheet_to_json(libro.Sheets[hoja], { header: 1, raw: true, defval: '' }) : null;
  }
  function conTitulos(filas, columnas) {
    if (!filas || !filas.length) return [];
    const titulos = filas[0].map(claveTitulo);
    const pos = {};
    for (const c of columnas) {
      const i = titulos.indexOf(claveTitulo(c.t));
      if (i >= 0) pos[c.k] = i;
    }
    return filas.slice(1).map((f) => { const o = {}; for (const k in pos) o[k] = f[pos[k]]; return o; });
  }
  const texto = (v) => (v == null ? '' : String(v).trim());

  function leer(XLSX, libro, nombreArchivo) {
    const filasIsa = filasDe(XLSX, libro, HOJA_ISA);
    if (!filasIsa) throw new Error('El Excel no tiene la hoja «ISA».');
    const hoy = new Date().toISOString().slice(0, 10);
    const anioRef = +hoy.slice(0, 4);
    const fichas = [], porN = {};
    for (const r of conTitulos(filasIsa, COL_ISA)) {
      const n = Number(r.n);
      if (!n || !texto(r.titulo)) continue;
      const f = {
        id: 'isa' + n, lamina: n, titulo: texto(r.titulo), unidad: texto(r.unidad),
        tipo: texto(r.tipo), tipoGrupo: L.tipoGrupo(r.tipo), estado: texto(r.estado), estadoGrupo: L.estadoGrupo(r.estado),
        inicioTexto: '', inicio: fechaDeCelda(r.inicio, anioRef), inicioSinAnio: false,
        csDiveduc: texto(r.csDiveduc), cse: texto(r.cse), dictamen: texto(r.dictamen),
        fiscal: texto(r.fiscal), fiscalClave: L.claveFiscal(r.fiscal), secretario: texto(r.secretario),
        prorrogas: r.prorrogas === '' || r.prorrogas == null || isNaN(Number(r.prorrogas)) ? null : Number(r.prorrogas),
        vence: fechaDeCelda(r.vence, anioRef), venceTexto: '', resumen: texto(r.resumen), notas: texto(r.notas), diligencias: []
      };
      if (!f.inicio && texto(r.inicio)) f.inicioTexto = texto(r.inicio);
      fichas.push(f);
      porN[n] = f;
    }
    const sinIsa = [];
    for (const r of conTitulos(filasDe(XLSX, libro, HOJA_DIL), COL_DIL)) {
      const detalle = texto(r.detalle);
      if (!detalle && !r.fecha) continue;
      const f = porN[Number(r.n)];
      if (!f) { sinIsa.push(r.n); continue; }
      const fecha = fechaDeCelda(r.fecha, anioRef);
      const tipo = P.TIPOS.find((t) => claveTitulo(t) === claveTitulo(r.tipo)) || P.clasificar(detalle);
      const h = claveTitulo(r.hecha);
      const hecha = h ? /^(HECHA|SI|LISTA|OK)/.test(h) : !!(fecha && fecha <= hoy);
      f.diligencias.push({ texto: detalle, resto: detalle, fecha, tipo, hecha, sinAnio: false });
    }
    for (const f of fichas) {
      f.diligencias.sort((a, b) => (a.fecha || '9999').localeCompare(b.fecha || '9999'));
      P.completar(f, anioRef);
    }
    const tablas = [];
    const jud = filasDe(XLSX, libro, HOJA_JUD);
    if (jud && jud.length) {
      const filas = jud.map((f) => f.map(texto)).filter((f) => f.some(Boolean));
      if (filas.length) tablas.push({ titulo: 'Procesos judiciales', lamina: 0, titulos: filas[0], filas: filas.slice(1) });
    }
    return { fuente: 'excel', archivo: nombreArchivo || '', fechaExpo: '', cargado: new Date().toISOString(), fichas, tablas, sinIsa };
  }

  const Excel = { armarLibro, escribir, leer, fechaDeCelda, aSerie, deSerie, HOJA_ISA, HOJA_DIL, COL_ISA, COL_DIL };
  if (enNode) module.exports = Excel;
  else raiz.ExcelIsa = Excel;
})(this);
