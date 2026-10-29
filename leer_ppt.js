// Lee el PowerPoint de la exposición de ISA (una lámina por ISA, con tablas) y lo deja en datos simples.
// Sirve en el navegador y en Node (pruebas). No usa bibliotecas: el .pptx es un .zip con XML adentro.
(function (raiz) {
  'use strict';

  const MESES = { ENE: 1, FEB: 2, MAR: 3, ABR: 4, MAY: 5, JUN: 6, JUL: 7, AGO: 8, SEP: 9, SET: 9, OCT: 10, NOV: 11, DIC: 12 };
  const NOMBRE_MES = ['', 'ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC'];

  // Título de la celda en la lámina → nombre del dato. null = título de grupo, sin dato.
  const ETIQUETAS = {
    'UNIDAD': 'unidad', 'TIPO DE ISA': 'tipo', 'FECHA DE INICIO': 'inicioTexto', 'ESTADO DE LA ISA': 'estado',
    'CS DIVEDUC': 'csDiveduc', 'CSE': 'cse', 'DICTAMEN': 'dictamen', 'FISCAL': 'fiscal', 'SECRETARIO': 'secretario',
    'PRORROGAS': 'prorrogasTexto', 'RESUMEN DEL HECHO': 'resumen', 'ULTIMAS DILIGENCIAS': 'ultimasTexto',
    'PROXIMAS DILIGENCIAS': 'proximasTexto', 'ISA': null, 'ESTADO': null
  };

  const normal = (t) => String(t || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
  const clave = (t) => normal(t).replace(/[:.]/g, '').replace(/\s+/g, ' ').trim();

  // ---------- El .zip ----------
  async function abrirZip(bytes, inflar) {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let e = bytes.length - 22;
    while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
    if (e < 0) throw new Error('El archivo no es un PowerPoint (.pptx).');
    const cuantos = dv.getUint16(e + 10, true);
    let p = dv.getUint32(e + 16, true);
    const archivos = {};
    const deco = new TextDecoder();
    for (let i = 0; i < cuantos && dv.getUint32(p, true) === 0x02014b50; i++) {
      const largoNombre = dv.getUint16(p + 28, true);
      archivos[deco.decode(bytes.subarray(p + 46, p + 46 + largoNombre))] = {
        metodo: dv.getUint16(p + 10, true), comprimido: dv.getUint32(p + 20, true), local: dv.getUint32(p + 42, true)
      };
      p += 46 + largoNombre + dv.getUint16(p + 30, true) + dv.getUint16(p + 32, true);
    }
    return {
      nombres: Object.keys(archivos),
      async texto(nombre) {
        const a = archivos[nombre];
        if (!a) return null;
        const ini = a.local + 30 + dv.getUint16(a.local + 26, true) + dv.getUint16(a.local + 28, true);
        const datos = bytes.subarray(ini, ini + a.comprimido);
        return deco.decode(a.metodo === 0 ? datos : await inflar(datos));
      }
    };
  }

  // ---------- El XML de una lámina ----------
  const entidades = (t) => t.replace(/&(#x?[0-9a-f]+|amp|lt|gt|quot|apos);/gi, (m, c) => {
    if (c[0] === '#') return String.fromCodePoint(c[1] === 'x' || c[1] === 'X' ? parseInt(c.slice(2), 16) : parseInt(c.slice(1), 10));
    return { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[c.toLowerCase()];
  });
  function textoDe(xml) {
    const lineas = [];
    for (const p of xml.match(/<a:p>[\s\S]*?<\/a:p>/g) || []) {
      const conSaltos = p.replace(/<a:br\b[^>]*\/>/g, '<a:t>\n</a:t>');
      let t = '';
      for (const m of conSaltos.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)) t += m[1];
      t = entidades(t).replace(/[ \t ]+/g, ' ');
      for (const l of t.split('\n')) if (l.trim()) lineas.push(l.trim());
    }
    return lineas.join('\n');
  }
  function tablasDe(xml) {
    return (xml.match(/<a:tbl>[\s\S]*?<\/a:tbl>/g) || []).map((tbl) =>
      (tbl.match(/<a:tr\b[\s\S]*?<\/a:tr>/g) || []).map((tr) => {
        const celdas = [];
        for (const m of tr.matchAll(/<a:tc\b([^>]*)>([\s\S]*?)<\/a:tc>/g)) {
          // Las celdas combinadas (continuación de otra) no traen dato propio
          if (/\b[hv]Merge="(1|true)"/.test(m[1])) continue;
          celdas.push(textoDe(m[2]));
        }
        return celdas;
      }));
  }
  const formasDe = (xml) => (xml.match(/<p:sp\b[\s\S]*?<\/p:sp>/g) || []).map(textoDe).filter(Boolean);

  // ---------- Fechas ----------
  // 24SEP2026 · 280900JUL2026 (grupo fecha-hora) · 30 MAYO 2025 · 30.MAY.2025 · 15 SEPTIEMBRE 2025 · 12SEP (sin año)
  const RE_FECHA = /(?<!\d)(\d{1,2})(\d{4})?[\s.\-/]*(?:DE\s+)?(ENE(?:RO)?|FEB(?:RERO)?|MAR(?:ZO)?|ABR(?:IL)?|MAY(?:O)?|JUN(?:IO)?|JUL(?:IO)?|AGO(?:STO)?|SEPT?(?:IEMBRE)?|SET|OCT(?:UBRE)?|NOV(?:IEMBRE)?|DIC(?:IEMBRE)?)(?![A-Z])(?:[\s.\-/]*(?:DE(?:L)?\s+)?(\d{4}))?/;
  function leerFecha(texto, anioRef) {
    const m = normal(texto).match(RE_FECHA);
    if (!m) return null;
    const dia = +m[1], mes = MESES[m[3].slice(0, 3)];
    if (!mes || dia < 1 || dia > 31) return null;
    const anio = m[4] ? +m[4] : anioRef;
    if (!anio) return null;
    return { iso: `${anio}-${String(mes).padStart(2, '0')}-${String(dia).padStart(2, '0')}`, sinAnio: !m[4], escrita: m[0].trim() };
  }
  const fechaMilitar = (iso) => {
    if (!iso) return '';
    const [a, m, d] = iso.split('-');
    return d + NOMBRE_MES[+m] + a;
  };
  function diasEntre(isoDesde, isoHasta) {
    const f = (iso) => { const [a, m, d] = iso.split('-').map(Number); return Date.UTC(a, m - 1, d); };
    return Math.round((f(isoHasta) - f(isoDesde)) / 86400000);
  }

  // Diligencias: una por línea; si una línea trae varias («05OCT2026 Solicita… .06OCT2026 Toma…»), se separan.
  function diligencias(texto, anioRef) {
    const lineas = [];
    for (const l of String(texto || '').split('\n')) {
      for (const parte of l.split(/(?<=\.)\s*(?=\d{1,2}[A-ZÁÉÍÓÚ]{3}\d{4}\s+\S)/)) if (parte.trim()) lineas.push(parte.trim());
    }
    return lineas.map((t) => {
      const f = leerFecha(t, anioRef);
      // «24SEP2026. Contacto…» → la fecha va aparte y el resto se muestra sin repetirla
      const alInicio = f && normal(t).indexOf(f.escrita) === 0;
      const resto = alInicio ? t.slice(f.escrita.length).replace(/^[\s.,:;\-–]+/, '') : t;
      return { texto: t, resto: resto || t, fecha: f ? f.iso : '', sinAnio: !!(f && f.sinAnio) };
    });
  }

  function tipoGrupo(tipo) {
    const t = normal(tipo);
    if (/MEDIC|LICENCIA/.test(t)) return 'Médica';
    if (/ACCIDENT/.test(t)) return 'Accidente';
    if (/DISCIPLIN/.test(t)) return 'Disciplinaria';
    if (/PERDID|ROBO|EXTRAV/.test(t)) return 'Pérdida o robo';
    return t ? 'Otra' : 'Sin tipo';
  }
  function estadoGrupo(estado) {
    const t = normal(estado);
    if (/CERRAD|TERMINAD/.test(t)) return 'Cerrada';
    if (/INVESTIGA/.test(t)) return 'En investigación';
    if (/TRAMITE/.test(t)) return 'En trámite';
    if (/ABIERT/.test(t)) return 'Abierta';
    return t ? 'Otro' : 'Sin estado';
  }
  // «CAP S. HENRÍQUEZ» y «CAP. S. HENRÍQUEZ B.» son el mismo fiscal: grado, inicial y apellido.
  const claveFiscal = (f) => clave(f).replace(/°/g, '').split(' ').slice(0, 3).join(' ');

  // ---------- Una lámina ----------
  function leerFicha(formas, tablas, anioRef) {
    const crudo = {};
    let actual = null;
    const agregar = (k, v) => { if (k && v) crudo[k] = crudo[k] ? crudo[k] + '\n' + v : v; };
    for (const tabla of tablas) {
      actual = null;
      for (const celdas of tabla) {
        if (!celdas.length) continue;
        const k = clave(celdas[0]);
        if (k in ETIQUETAS) { actual = ETIQUETAS[k]; agregar(actual, celdas.slice(1).join('\n').trim()); } else agregar(actual, celdas.join('\n').trim());
      }
    }
    const titulo = formas.find((f) => !(clave(f) in ETIQUETAS)) || '';
    const pr = crudo.prorrogasTexto || '';
    const nPr = normal(pr).match(/PRORROGAS?\s*:?\s*(\d+)/);
    const lineaVen = pr.split('\n').find((l) => /VEN/i.test(l) && leerFecha(l, anioRef)) || pr.split('\n').filter((l) => leerFecha(l, anioRef)).pop();
    const vence = lineaVen ? leerFecha(lineaVen, anioRef) : null;
    const inicio = leerFecha(crudo.inicioTexto || '', anioRef);
    return {
      titulo: titulo.replace(/\n/g, ' '),
      unidad: crudo.unidad || '', tipo: crudo.tipo || '', tipoGrupo: tipoGrupo(crudo.tipo),
      estado: crudo.estado || '', estadoGrupo: estadoGrupo(crudo.estado),
      inicioTexto: crudo.inicioTexto || '', inicio: inicio ? inicio.iso : '', inicioSinAnio: !!(inicio && inicio.sinAnio),
      csDiveduc: crudo.csDiveduc || '', cse: crudo.cse || '', dictamen: crudo.dictamen || '',
      fiscal: (crudo.fiscal || '').replace(/\n/g, ' '), fiscalClave: claveFiscal(crudo.fiscal), secretario: (crudo.secretario || '').replace(/\n/g, ' '),
      prorrogas: nPr ? +nPr[1] : null, vence: vence ? vence.iso : '', venceTexto: vence ? vence.escrita : '',
      resumen: crudo.resumen || '',
      ultimas: diligencias(crudo.ultimasTexto, anioRef), proximas: diligencias(crudo.proximasTexto, anioRef)
    };
  }

  // ---------- La presentación completa ----------
  async function leer(bytes, inflar, nombreArchivo) {
    const zip = await abrirZip(bytes, inflar);
    // Orden real de las láminas (el que se ve al presentar), desde presentation.xml
    let rutas = [];
    const pres = await zip.texto('ppt/presentation.xml');
    const rels = await zip.texto('ppt/_rels/presentation.xml.rels');
    if (pres && rels) {
      const destino = {};
      for (const m of rels.matchAll(/<Relationship\b[^>]*>/g)) {
        const id = (m[0].match(/\bId="([^"]+)"/) || [])[1], t = (m[0].match(/\bTarget="([^"]+)"/) || [])[1];
        if (id && t) destino[id] = 'ppt/' + t.replace(/^\/?ppt\//, '').replace(/^\.\//, '');
      }
      for (const m of pres.matchAll(/<p:sldId\b[^>]*\br:id="([^"]+)"/g)) if (destino[m[1]]) rutas.push(destino[m[1]]);
    }
    if (!rutas.length) {
      rutas = zip.nombres.filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
        .sort((a, b) => +a.match(/(\d+)\.xml/)[1] - +b.match(/(\d+)\.xml/)[1]);
    }
    if (!rutas.length) throw new Error('El PowerPoint no tiene láminas.');

    const laminas = [];
    for (const r of rutas) {
      const xml = await zip.texto(r);
      if (!xml || /<p:sld\b[^>]*\bshow="(0|false)"/.test(xml)) continue; // lámina oculta
      laminas.push({ formas: formasDe(xml), tablas: tablasDe(xml) });
    }

    // Fecha de la exposición: la primera fecha de la portada (si no hay, la del nombre del archivo)
    let fechaExpo = null;
    for (const l of laminas.slice(0, 2)) { fechaExpo = fechaExpo || leerFecha(l.formas.join(' '), null); }
    fechaExpo = fechaExpo || leerFecha(nombreArchivo || '', null);
    const anioRef = fechaExpo ? +fechaExpo.iso.slice(0, 4) : new Date().getFullYear();

    const fichas = [], tablas = [];
    laminas.forEach((l, i) => {
      const plano = l.tablas.flat(2).map(clave);
      if (plano.includes('TIPO DE ISA') || plano.includes('FISCAL')) {
        const f = leerFicha(l.formas, l.tablas, anioRef);
        f.lamina = i + 1;
        f.id = 'isa' + (i + 1);
        fichas.push(f);
      } else if (l.tablas.length) {
        // Otras tablas (por ejemplo «Procesos Judiciales»): la primera fila son los títulos
        for (const t of l.tablas) {
          const filas = t.filter((c) => c.some(Boolean));
          if (filas.length) tablas.push({ titulo: (l.formas[0] || 'Tabla').replace(/\n/g, ' '), lamina: i + 1, titulos: filas[0], filas: filas.slice(1) });
        }
      }
    });
    return { archivo: nombreArchivo || '', fechaExpo: fechaExpo ? fechaExpo.iso : '', cargado: new Date().toISOString(), laminas: laminas.length, fichas, tablas };
  }

  // ---------- Qué revisar en cada ISA (al día de «hoy», en formato AAAA-MM-DD) ----------
  function semaforo(f, hoy) {
    if (f.estadoGrupo === 'Cerrada') return 'cerrada';
    if (!f.vence) return 'sin_fecha';
    const d = diasEntre(hoy, f.vence);
    return d < 0 ? 'vencida' : d <= 7 ? 'por_vencer' : 'al_dia';
  }
  function revisar(f, hoy) {
    const avisos = [];
    const abierta = f.estadoGrupo !== 'Cerrada';
    if (abierta && f.vence) {
      const d = diasEntre(hoy, f.vence);
      if (d < 0) avisos.push({ nivel: 'rojo', texto: `Vencida el ${fechaMilitar(f.vence)} (hace ${-d} día${d === -1 ? '' : 's'}). Revisar si se pidió la prórroga.` });
      else if (d <= 7) avisos.push({ nivel: 'amarillo', texto: d === 0 ? 'Vence hoy.' : `Vence en ${d} día${d === 1 ? '' : 's'} (${fechaMilitar(f.vence)}).` });
    }
    if (abierta && !f.vence) avisos.push({ nivel: 'amarillo', texto: 'No tiene fecha de vencimiento en la lámina.' });
    if (abierta && !f.proximas.length) avisos.push({ nivel: 'amarillo', texto: 'No tiene próximas diligencias.' });
    for (const p of f.proximas) {
      if (p.fecha && p.fecha < hoy) avisos.push({ nivel: 'amarillo', texto: `Próxima diligencia con fecha ya pasada: «${p.texto}». ¿Se hizo? Pasarla a últimas diligencias.` });
    }
    const ultima = f.ultimas.map((u) => u.fecha).filter(Boolean).sort().pop();
    for (const p of f.proximas) {
      if (p.fecha && ultima && p.fecha < ultima && p.fecha >= hoy) avisos.push({ nivel: 'info', texto: `La próxima diligencia «${p.texto}» es anterior a la última (${fechaMilitar(ultima)}).` });
    }
    if (f.inicioTexto && !/^\d{2}[A-Z]{3}\d{4}$/.test(normal(f.inicioTexto).replace(/\s/g, ''))) {
      avisos.push({ nivel: 'info', texto: `Fecha de inicio escrita de otra forma: «${f.inicioTexto}»${f.inicioSinAnio ? ' (sin año)' : ''}. Lo normal es DDMMMAAAA, por ejemplo 14DIC2024.` });
    }
    const sinAnio = [...f.ultimas, ...f.proximas].filter((x) => x.sinAnio).length;
    if (sinAnio) avisos.push({ nivel: 'info', texto: `${sinAnio} diligencia${sinAnio === 1 ? '' : 's'} con fecha sin año (por ejemplo «15OCT»).` });
    return avisos;
  }

  const LeerPpt = { leer, leerFecha, diligencias, fechaMilitar, diasEntre, semaforo, revisar, normal, tipoGrupo, estadoGrupo, claveFiscal };
  if (typeof module !== 'undefined' && module.exports) module.exports = LeerPpt;
  else raiz.LeerPpt = LeerPpt;
})(this);
