// ISA ESCINF: consulta rápida de las investigaciones sumarias desde el Excel del S-1 (o el PowerPoint de la exposición).
(function () {
  'use strict';

  const CLAVE = 'isa_escinf_datos';
  const SEMAFORO = {
    vencida: { texto: 'Vencida', punto: '🔴' },
    por_vencer: { texto: 'Vence en 7 días', punto: '🟡' },
    al_dia: { texto: 'Al día', punto: '🟢' },
    sin_fecha: { texto: 'Sin fecha de vencimiento', punto: '⚪' },
    cerrada: { texto: 'Cerrada', punto: '⚪' }
  };
  const PLAZO = { atrasado: '🔴', semana: '🟡', al_dia: '🟢' };

  const $ = (id) => document.getElementById(id);
  const estado = { datos: null, vista: 'tablero', expo: 0 };
  const { normal, fechaMilitar, diasEntre } = LeerPpt;
  const { semaforo, diasAbierta } = Plazos;
  const esc = (t) => String(t == null ? '' : t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const hoy = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
  const vacio = (t) => !String(t || '').trim() || /^-\.?-$/.test(String(t).trim());
  const deExcel = () => estado.datos && estado.datos.fuente === 'excel';
  const numero = (f) => (deExcel() ? 'N.° ' : 'Lámina ') + f.lamina;
  const avisos = (f, h) => [...LeerPpt.revisar(Plazos.conVenceEfectivo(f), h), ...Plazos.revisar(f, h)];
  const INICIO = () => fechaMilitar(Plazos.REGLAS.inicioControl);
  // «1 año y 3 meses», «4 meses», «12 días»
  function duracion(dias) {
    if (dias == null) return '—';
    if (dias < 60) return plural(dias, 'día');
    const meses = Math.floor(dias / 30.44), anios = Math.floor(meses / 12), resto = meses % 12;
    if (!anios) return plural(meses, 'mes', 'meses');
    return plural(anios, 'año') + (resto ? ' y ' + plural(resto, 'mes', 'meses') : '');
  }
  // Tareas que exigen algo (las «programadas» son solo recordatorio)
  const atrasos = (f, h) => Plazos.tareas(f, h).filter((t) => t.estado === 'atrasado' && t.regla !== 'programada');
  const plural = (n, s, p) => n + ' ' + (n === 1 ? s : p || s + 's');
  const cuando = (dias) => dias < 0 ? `atrasado ${plural(-dias, 'día')}` : dias === 0 ? 'hoy' : `en ${plural(dias, 'día')}`;

  // ---------- Guardar / recuperar (solo si el usuario lo pide) ----------
  function guardar(datos) {
    try { localStorage.setItem(CLAVE, JSON.stringify(datos)); } catch (e) { /* sin memoria: igual se muestra */ }
  }
  function recuperar() {
    try { return JSON.parse(localStorage.getItem(CLAVE) || 'null'); } catch (e) { return null; }
  }
  function olvidar() {
    try { localStorage.removeItem(CLAVE); } catch (e) { /* nada */ }
  }

  // ---------- Cargar el Excel o el PowerPoint ----------
  async function inflar(datos) {
    const flujo = new Blob([datos]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(flujo).arrayBuffer());
  }
  async function cargarArchivo(archivo) {
    if (!archivo) return;
    const esExcel = /\.xlsx$/i.test(archivo.name), esPpt = /\.pptx$/i.test(archivo.name);
    if (!esExcel && !esPpt) { alert('Cargue el Excel (.xlsx) de las ISA, o el PowerPoint (.pptx) de la exposición.\n\nSi está en formato antiguo (.xls o .ppt), ábralo y guárdelo como .xlsx o .pptx.'); return; }
    try {
      const bytes = new Uint8Array(await archivo.arrayBuffer());
      let datos;
      if (esExcel) {
        datos = ExcelIsa.leer(XLSX, XLSX.read(bytes, { type: 'array' }), archivo.name);
        if (!datos.fichas.length) throw new Error('La hoja «ISA» no tiene filas con N.° y nombre.');
      } else {
        if (typeof DecompressionStream === 'undefined') throw new Error('Este navegador es muy antiguo para leer el PowerPoint. Use Chrome, Edge o Safari actualizados.');
        datos = await LeerPpt.leer(bytes, inflar, archivo.name);
        if (!datos.fichas.length) throw new Error('No encontré láminas de ISA (con casillas «TIPO DE ISA» o «FISCAL»).');
        datos.fuente = 'ppt';
        const anio = +(datos.fechaExpo || hoy()).slice(0, 4);
        datos.fichas.forEach((f) => Plazos.completar(f, anio));
      }
      if ($('recordar').checked) guardar(datos); else olvidar();
      mostrarDatos(datos);
      if (datos.sinIsa && datos.sinIsa.length) alert(`Hay ${plural(datos.sinIsa.length, 'diligencia')} con un N.° de ISA que no está en la hoja «ISA» (${[...new Set(datos.sinIsa)].join(', ')}). No se muestran.`);
    } catch (e) {
      alert('No pude leer el archivo.\n\n' + e.message);
    }
  }

  // ---------- Mostrar ----------
  function mostrarDatos(datos) {
    estado.datos = datos;
    $('bienvenida').hidden = true;
    $('pestanas').hidden = false;
    $('imprimir').hidden = false;
    $('exponer').hidden = false;
    $('pasarExcel').hidden = deExcel();
    $('olvidar').hidden = !recuperar();
    $('fuente').textContent = (datos.archivo || 'Archivo cargado') + (datos.fechaExpo ? ' · exposición del ' + fechaMilitar(datos.fechaExpo) : '') +
      (deExcel() ? '' : ' · ⚠️ desde PowerPoint: los tipos de diligencia son adivinados');
    $('pestanaJudiciales').hidden = !datos.tablas.length;
    $('ahoraRevisar').innerHTML = `👉 <b>Ahora:</b> corrija esto en el ${deExcel() ? 'Excel' : 'PowerPoint'} y vuelva a cargarlo.`;

    const opciones = (sel, titulo, valores) => {
      const antes = sel.value;
      sel.innerHTML = `<option value="">${titulo}</option>` + valores.map(([v, t]) => `<option value="${esc(v)}">${esc(t)}</option>`).join('');
      if ([...sel.options].some((o) => o.value === antes)) sel.value = antes;
    };
    const unicos = (fn) => [...new Set(datos.fichas.map(fn).filter(Boolean))].sort();
    opciones($('filtroEstado'), 'Estado: todos', unicos((f) => f.estadoGrupo).map((v) => [v, v]));
    opciones($('filtroTipo'), 'Tipo: todos', unicos((f) => f.tipoGrupo).map((v) => [v, v]));
    const fis = fiscales().map((g) => [g.clave, g.nombre]);
    opciones($('filtroFiscal'), 'Fiscal: todos', fis);
    opciones($('agendaFiscal'), 'Fiscal: todos', fis);

    const h = hoy();
    $('nRevisar').textContent = datos.fichas.filter((f) => avisos(f, h).some((a) => a.nivel !== 'info')).length || '';
    $('nAtrasos').textContent = datos.fichas.filter((f) => atrasos(f, h).length).length || '';
    cambiarVista(estado.vista);
  }

  function cambiarVista(v) {
    estado.vista = v;
    document.querySelectorAll('.pestanas button').forEach((b) => b.classList.toggle('activa', b.dataset.vista === v));
    document.querySelectorAll('.vista').forEach((s) => { s.hidden = s.id !== 'vista-' + v; });
    ({ tablero: verTablero, agenda: verAgenda, estadistica: verEstadistica, buscar: verLista, fiscales: verFiscales, judiciales: verJudiciales, revisar: verRevisar }[v] || (() => {}))();
    window.scrollTo(0, 0);
  }

  // Una ISA en una línea de la lista
  function tarjeta(f, resaltar) {
    const h = hoy(), s = semaforo(f, h);
    const marcar = (t) => {
      let html = esc(t);
      for (const p of resaltar || []) {
        const re = new RegExp('(' + p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').split('').map((c) => c + '[̀-ͯ]?').join('') + ')', 'gi');
        html = html.replace(re, '<mark>$1</mark>');
      }
      return html;
    };
    let derecha;
    if (s === 'cerrada') derecha = '<b>Cerrada</b>';
    else if (!f.vence) derecha = '<b>Sin fecha</b><small>de vencimiento</small>';
    else {
      const d = diasEntre(h, Plazos.venceEfectivo(f));
      const txt = (Plazos.venceEfectivo(f) !== f.vence ? `<span class="ayuda">venció antes del ${INICIO()}</span><br>` : '') + (d < 0 ? `<span class="txt-rojo">venció hace ${d === -1 ? '1 día' : -d + ' d'}</span>` : d === 0 ? '<span class="txt-amarillo">vence hoy</span>' : d <= 7 ? `<span class="txt-amarillo">vence en ${d} d</span>` : `vence en ${d} d`);
      derecha = `<b>${SEMAFORO[s].punto} ${fechaMilitar(f.vence)}</b><small>${txt}</small>`;
    }
    const atr = atrasos(f, h).filter((t) => t.regla !== 'prórroga');
    const prox = f.proximas[0];
    return `<button class="item ${s}" data-id="${f.id}">
      <span class="nombre">${marcar(f.titulo)}</span>
      <span class="derecha">${derecha}${f.prorrogas != null ? `<small>${plural(f.prorrogas, 'prórroga')}</small>` : ''}</span>
      <span class="linea">${marcar(f.unidad)} · <span class="etiqueta">${esc(f.tipoGrupo)}</span> <span class="etiqueta">${esc(f.estadoGrupo)}</span> · Fiscal: ${marcar(f.fiscal || '—')}${diasAbierta(f, h) != null ? ` · ⏳ abierta hace ${duracion(diasAbierta(f, h))}` : ''}</span>
      ${atr.length ? `<span class="proxima txt-rojo">⏱ ${esc(atr[0].que)} (${cuando(atr[0].dias)})</span>` : prox ? `<span class="proxima">➡️ ${marcar(prox.resto || prox.texto)}${f.proximas.length > 1 ? ` <small class="ayuda">(+${f.proximas.length - 1})</small>` : ''}</span>` : ''}
    </button>`;
  }

  function barras(destino, pares) {
    const max = Math.max(1, ...pares.map((p) => p[1]));
    $(destino).innerHTML = pares.map(([t, n]) => `<div class="barra-dato"><span>${esc(t)}</span><div class="fondo-barra"><div class="relleno" style="width:${(100 * n) / max}%"></div></div><span class="num">${n}</span></div>`).join('') || '<p class="ayuda">—</p>';
  }
  const contar = (lista, fn) => {
    const c = {};
    for (const f of lista) { const k = fn(f); if (k) c[k] = (c[k] || 0) + 1; }
    return Object.entries(c).sort((a, b) => b[1] - a[1]);
  };
  const ordenVence = (a, b) => (a.estadoGrupo === 'Cerrada') - (b.estadoGrupo === 'Cerrada') || (a.vence || '9999').localeCompare(b.vence || '9999');

  // ---------- Tablero ----------
  function verTablero() {
    const fs = estado.datos.fichas, h = hoy();
    const por = (s) => fs.filter((f) => semaforo(f, h) === s).length;
    const abiertas = fs.filter((f) => f.estadoGrupo !== 'Cerrada');
    const cifras = [
      ['', fs.length, 'ISA en total', ''],
      ['abiertas', abiertas.length, 'abiertas o en trámite', ''],
      ['vencida', por('vencida'), 'vencidas', 'rojo'],
      ['por_vencer', por('por_vencer'), 'vencen en 7 días', 'amarillo'],
      ['atrasos', fs.filter((f) => atrasos(f, h).length).length, 'con plazos atrasados', 'rojo'],
      ['cerrada', por('cerrada'), 'cerradas', '']
    ];
    $('cifras').innerHTML = cifras.map(([filtro, n, t, c]) => `<button class="cifra ${c}" data-filtro="${filtro}"><b>${n}</b><span>${t}</span></button>`).join('');
    $('hoyTexto').textContent = `Al día de hoy, ${fechaMilitar(h)}. Las ISA abiertas, de la que vence primero a la última.`;
    $('listaVencimientos').innerHTML = abiertas.slice().sort(ordenVence).map((f) => tarjeta(f)).join('') || '<p class="ayuda">No hay ISA abiertas.</p>';
    barras('barrasTipo', contar(fs, (f) => f.tipoGrupo));
    barras('barrasEstado', contar(fs, (f) => f.estadoGrupo));
    barras('barrasCse', contar(abiertas, (f) => Plazos.etapaCse(f, h).etapa || 'Sin trámite en la CSE'));
    $('listaProrrogas').innerHTML = abiertas.filter((f) => f.prorrogas != null)
      .sort((a, b) => b.prorrogas - a.prorrogas).slice(0, 5)
      .map((f) => `<button class="item ${semaforo(f, h)}" data-id="${f.id}"><span class="nombre">${esc(f.titulo)}</span><span class="derecha"><b>${f.prorrogas}</b><small>prórrogas</small></span><span class="linea">Desde ${f.inicio ? fechaMilitar(f.inicio) : esc(f.inicioTexto)} · ${esc(f.fiscal)}</span></button>`).join('');
  }

  // ---------- Agenda ----------
  function agendaDe(filtroFiscal, filtroEstado) {
    const h = hoy();
    return fiscales().filter((g) => !filtroFiscal || g.clave === filtroFiscal).map((g) => {
      const tareas = [];
      for (const f of g.fichas) {
        for (const t of Plazos.tareas(f, h)) {
          if (filtroEstado === 'atrasado' && t.estado !== 'atrasado') continue;
          if (filtroEstado === 'urgente' && t.estado === 'al_dia') continue;
          tareas.push({ ...t, f });
        }
      }
      tareas.sort((a, b) => a.limite.localeCompare(b.limite));
      return { ...g, tareas };
    }).filter((g) => g.tareas.length);
  }
  function verAgenda() {
    const grupos = agendaDe($('agendaFiscal').value, $('agendaEstado').value);
    $('agenda').innerHTML = grupos.map((g) => {
      const atr = g.tareas.filter((t) => t.estado === 'atrasado').length;
      return `<div class="panel">
        <h3>👤 ${esc(g.nombre)} <small class="ayuda">${plural(g.tareas.length, 'tarea')}${atr ? ` · <span class="txt-rojo">${atr} atrasada${atr === 1 ? '' : 's'}</span>` : ''}</small></h3>
        <div class="tareas">${g.tareas.map((t) => `
          <button class="tarea ${t.estado}" data-id="${t.f.id}">
            <span class="cuando">${PLAZO[t.estado]} <b>${fechaMilitar(t.limite)}</b><small>${cuando(t.dias)}</small></span>
            <span class="que"><b>${esc(t.f.titulo)}</b><span>${esc(t.que)}</span><small class="ayuda">${esc(t.regla)}</small></span>
          </button>`).join('')}</div>
      </div>`;
    }).join('') || '<p class="panel">✅ No hay plazos con este filtro.</p>';
  }

  // ---------- Buscar ----------
  function verLista() {
    const fs = estado.datos.fichas, h = hoy();
    const palabras = normal($('busqueda').value).split(/\s+/).filter(Boolean);
    const fe = $('filtroEstado').value, ft = $('filtroTipo').value, ff = $('filtroFiscal').value, fsem = $('filtroSemaforo').value;
    let lista = fs.filter((f) => {
      if (fe && f.estadoGrupo !== fe) return false;
      if (ft && f.tipoGrupo !== ft) return false;
      if (ff && f.fiscalClave !== ff) return false;
      if (fsem === 'abiertas') { if (f.estadoGrupo === 'Cerrada') return false; }
      else if (fsem === 'atrasos') { if (!atrasos(f, h).length) return false; }
      else if (fsem && semaforo(f, h) !== fsem) return false;
      if (!palabras.length) return true;
      const texto = normal([f.titulo, f.unidad, f.tipo, f.estado, f.fiscal, f.secretario, f.resumen, f.notas, ...f.diligencias.map((d) => d.texto)].join(' '));
      return palabras.every((p) => texto.includes(p));
    });
    const cmp = {
      vence: ordenVence,
      lamina: (a, b) => a.lamina - b.lamina,
      prorrogas: (a, b) => (b.prorrogas || 0) - (a.prorrogas || 0),
      inicio: (a, b) => (a.inicio || '9999').localeCompare(b.inicio || '9999'),
      nombre: (a, b) => a.titulo.localeCompare(b.titulo, 'es')
    }[$('orden').value];
    lista = lista.sort(cmp);
    $('contador').textContent = lista.length === fs.length ? `${fs.length} ISA` : `${lista.length} de ${fs.length} ISA`;
    $('lista').innerHTML = lista.map((f) => tarjeta(f, palabras)).join('') || '<p class="ayuda">No hay ISA con esa búsqueda. Pruebe con menos palabras o quite los filtros.</p>';
  }

  // ---------- Por fiscal ----------
  function fiscales() {
    const g = {};
    for (const f of estado.datos.fichas) {
      const k = f.fiscalClave || '(sin fiscal)';
      g[k] = g[k] || { clave: f.fiscalClave, nombre: '', fichas: [] };
      if ((f.fiscal || '').length > g[k].nombre.length) g[k].nombre = f.fiscal;
      g[k].fichas.push(f);
    }
    return Object.values(g).map((x) => ({ ...x, nombre: x.nombre || '(sin fiscal)' })).sort((a, b) => b.fichas.length - a.fichas.length || a.nombre.localeCompare(b.nombre));
  }
  function verFiscales() {
    const h = hoy();
    $('tarjetasFiscales').innerHTML = fiscales().map((g) => {
      const abiertas = g.fichas.filter((f) => f.estadoGrupo !== 'Cerrada');
      const venc = abiertas.filter((f) => semaforo(f, h) === 'vencida').length;
      const conAtraso = abiertas.filter((f) => atrasos(f, h).length).length;
      const filas = g.fichas.slice().sort(ordenVence)
        .map((f) => { const s = semaforo(f, h); return `<div class="fila"><span>${esc(f.titulo)}</span><span>${SEMAFORO[s].punto} ${f.vence ? fechaMilitar(f.vence) : esc(SEMAFORO[s].texto)}</span></div>`; }).join('');
      return `<button class="tarjeta" data-fiscal="${esc(g.clave)}">
        <h3>${esc(g.nombre)}</h3>
        <div class="cuenta">${g.fichas.length} ISA (${abiertas.length} abierta${abiertas.length === 1 ? '' : 's'})${venc ? ` · <span class="txt-rojo">${venc} vencida${venc === 1 ? '' : 's'}</span>` : ''}${conAtraso ? ` · <span class="txt-rojo">${conAtraso} con plazos atrasados</span>` : ''}</div>
        ${filas}
      </button>`;
    }).join('');
  }

  // ---------- Procesos judiciales ----------
  function verJudiciales() {
    const palabras = normal($('busquedaJudicial').value).split(/\s+/).filter(Boolean);
    $('tablasJudiciales').innerHTML = estado.datos.tablas.map((t) => {
      const filas = t.filas.filter((fila) => { const tx = normal(fila.join(' ')); return palabras.every((p) => tx.includes(p)); });
      return `<h3>${esc(t.titulo)}${t.lamina ? ` <small class="ayuda">(lámina ${t.lamina})</small>` : ''}</h3>
        <div class="tabla-scroll"><table><thead><tr>${t.titulos.map((c) => `<th>${esc(c)}</th>`).join('')}</tr></thead>
        <tbody>${filas.map((fila) => `<tr>${t.titulos.map((c, i) => `<td>${esc(fila[i])}</td>`).join('')}</tr>`).join('') || `<tr><td colspan="${t.titulos.length}">Nada con esa búsqueda.</td></tr>`}</tbody></table></div>`;
    }).join('');
  }

  // ---------- Revisar planilla ----------
  function verRevisar() {
    const h = hoy();
    const conAvisos = estado.datos.fichas.map((f) => ({ f, av: avisos(f, h) })).filter((x) => x.av.length);
    const peso = (av) => (av.some((a) => a.nivel === 'rojo') ? 0 : av.some((a) => a.nivel === 'amarillo') ? 1 : 2);
    conAvisos.sort((a, b) => peso(a.av) - peso(b.av) || a.f.lamina - b.f.lamina);
    const icono = { rojo: '🔴', amarillo: '🟡', info: 'ℹ️' };
    $('listaRevisar').innerHTML = conAvisos.length ? conAvisos.map(({ f, av }) => `
      <div class="bloque ${['rojo', 'amarillo', ''][peso(av)]}">
        <h4 data-id="${f.id}">${numero(f)} · ${esc(f.titulo)}</h4>
        <ul>${av.map((a) => `<li>${icono[a.nivel]} ${esc(a.texto)}</li>`).join('')}</ul>
      </div>`).join('') : '<p class="panel">✅ No hay nada que revisar. ¡Bien!</p>';
  }

  // ---------- Estadística (solo las ISA abiertas) ----------
  function estadisticaHtml() {
    const h = hoy(), abiertas = estado.datos.fichas.filter((f) => f.estadoGrupo !== 'Cerrada');
    const conDias = abiertas.filter((f) => diasAbierta(f, h) != null);
    const prom = (lista, fn) => { const v = lista.map(fn).filter((x) => x != null); return v.length ? Math.round(v.reduce((a, b) => a + b, 0) / v.length) : null; };
    const promDias = prom(conDias, (f) => diasAbierta(f, h));
    const masAntigua = conDias.slice().sort((a, b) => diasAbierta(b, h) - diasAbierta(a, h));
    const tramos = [['Menos de 3 meses', 0, 91], ['3 a 6 meses', 91, 183], ['6 meses a 1 año', 183, 366], ['Más de 1 año', 366, 1e9]]
      .map(([t, a, b]) => [t, conDias.filter((f) => diasAbierta(f, h) >= a && diasAbierta(f, h) < b).length]);
    const max = Math.max(1, ...tramos.map((x) => x[1]));
    const tabla = (titulo, clave) => {
      const g = {};
      for (const f of abiertas) { const k = clave(f) || '—'; (g[k] = g[k] || []).push(f); }
      const filas = Object.entries(g).sort((a, b) => b[1].length - a[1].length || (prom(b[1], (f) => diasAbierta(f, h)) || 0) - (prom(a[1], (f) => diasAbierta(f, h)) || 0));
      return `<div class="panel"><h3>${titulo}</h3><div class="tabla-scroll"><table><thead><tr><th></th><th>ISA abiertas</th><th>Abierta (promedio)</th><th>Prórrogas (promedio)</th><th>Vencidas</th><th>Con plazos atrasados</th></tr></thead><tbody>${filas.map(([k, l]) => {
        const v = l.filter((f) => semaforo(f, h) === 'vencida').length, a = l.filter((f) => atrasos(f, h).length).length;
        return `<tr><td><b>${esc(k)}</b></td><td>${l.length}</td><td>${duracion(prom(l, (f) => diasAbierta(f, h)))}</td><td>${prom(l, (f) => f.prorrogas) ?? '—'}</td><td>${v ? `<span class="txt-rojo">${v}</span>` : 0}</td><td>${a ? `<span class="txt-rojo">${a}</span>` : 0}</td></tr>`;
      }).join('')}</tbody></table></div></div>`;
    };
    return `
      <div class="resumen">
        <div class="cifra"><b>${abiertas.length}</b><span>ISA abiertas</span></div>
        <div class="cifra"><b>${promDias != null ? duracion(promDias) : '—'}</b><span>tiempo abierta en promedio</span></div>
        <div class="cifra"><b>${masAntigua.length ? duracion(diasAbierta(masAntigua[0], h)) : '—'}</b><span>la más antigua</span></div>
        <div class="cifra"><b>${prom(abiertas, (f) => f.prorrogas) ?? '—'}</b><span>prórrogas en promedio</span></div>
      </div>
      <div class="panel"><h3>⏳ Cuánto tiempo llevan abiertas</h3>
        ${tramos.map(([t, n]) => `<div class="barra-dato"><span>${t}</span><div class="fondo-barra"><div class="relleno" style="width:${(100 * n) / max}%"></div></div><span class="num">${n}</span></div>`).join('')}
      </div>
      <div class="panel"><h3>🏆 Ranking de tiempo abierta</h3>
        <p class="ayuda">De la que lleva más tiempo abierta a la más nueva. Toque una para ver su ficha.</p>
        <div class="tabla-scroll"><table class="ranking"><thead><tr><th>N.°</th><th>Afectado / ISA</th><th>Abierta</th><th>Desde</th><th>Tipo</th><th>Prórrogas</th><th>Fiscal</th><th>Vence</th></tr></thead><tbody>
        ${masAntigua.map((f, i) => { const sm = semaforo(f, h); return `<tr data-id="${f.id}"><td class="puesto">${i + 1}</td><td><b>${esc(f.titulo)}</b></td><td><b>${duracion(diasAbierta(f, h))}</b><br><small class="ayuda">${plural(diasAbierta(f, h), 'día')}</small></td><td>${fechaMilitar(f.inicio)}</td><td>${esc(f.tipoGrupo)}</td><td>${f.prorrogas ?? '—'}</td><td>${esc(f.fiscal)}</td><td>${SEMAFORO[sm].punto} ${f.vence ? fechaMilitar(f.vence) : ''}</td></tr>`; }).join('')}
        </tbody></table></div>
        ${abiertas.length > conDias.length ? `<p class="ayuda">${plural(abiertas.length - conDias.length, 'ISA abierta')} sin fecha de inicio no entra${abiertas.length - conDias.length === 1 ? '' : 'n'} al ranking.</p>` : ''}
      </div>
      ${tabla('📂 Por tipo de ISA', (f) => f.tipoGrupo)}
      ${tabla('🏥 Por etapa en la CSE', (f) => Plazos.etapaCse(f, h).etapa || 'Sin trámite en la CSE')}
      ${tabla('👤 Por fiscal', (f) => f.fiscal)}
      ${tabla('🏢 Por unidad', (f) => f.unidad)}`;
  }
  function verEstadistica() { $('estadistica').innerHTML = estadisticaHtml(); }
  function imprimirEstadistica() {
    $('impresion').innerHTML = `<h1>Investigaciones sumarias ESCINF — estadística al ${fechaMilitar(hoy())}</h1><p>${esc(estado.datos.archivo)} · solo ISA abiertas</p>` + estadisticaHtml();
    window.print();
  }

  // ---------- Ficha ----------
  // Las próximas diligencias, en un recuadro destacado al comienzo de la ficha
  function proximasDestacadas(f, h) {
    if (f.estadoGrupo === 'Cerrada' && !f.proximas.length) return '';
    const filas = f.proximas.map((p) => {
      let cuandoTxt = '';
      if (p.fecha) {
        const d = diasEntre(h, p.fecha);
        cuandoTxt = d < 0 ? `<span class="cuando-prox pasada">fecha ya pasada</span>` : `<span class="cuando-prox">${d === 0 ? 'hoy' : 'en ' + plural(d, 'día')}</span>`;
      }
      return `<li><div class="fecha-prox">${p.fecha ? fechaMilitar(p.fecha) : 'Sin fecha'}${cuandoTxt}</div><div class="texto-prox">${esc(p.fecha ? p.resto || p.texto : p.texto)}<span class="etiqueta">${esc(p.tipo)}</span></div></li>`;
    }).join('');
    return `<div class="proximas-destacadas${f.proximas.length ? '' : ' vacia'}"><h3>➡️ PRÓXIMAS DILIGENCIAS</h3>${filas ? `<ul>${filas}</ul>` : '<p>⚠️ No hay próximas diligencias anotadas.</p>'}</div>`;
  }

  // modo: 'ficha' (pantalla), 'expo' (lámina: últimas 4 diligencias) o 'foto' (para enviar: sin avisos)
  function contenidoFicha(f, modo) {
    const enExpo = modo === 'expo';
    const h = hoy(), s = semaforo(f, h);
    const dato = (t, v) => `<div class="dato"><small>${t}</small><b>${esc(vacio(v) ? '—' : v)}</b></div>`;
    const casilla = (t, v) => `<span class="casilla">${t}:<b>${esc(vacio(v) ? '—' : v)}</b></span>`;
    const tiempo = (lista, clase) => lista.length ? `<ul class="tiempo ${clase}">${lista.map((d) => {
      const pasada = clase === 'proximas' && d.fecha && d.fecha < h;
      return `<li class="${pasada ? 'pasada' : ''}">${d.fecha ? `<span class="fecha">${fechaMilitar(d.fecha)}</span>` : ''}<span class="etiqueta">${esc(d.tipo)}</span> ${esc(d.fecha ? d.resto || d.texto : d.texto)}${pasada ? ' <span class="txt-amarillo">(fecha ya pasada)</span>' : ''}</li>`;
    }).join('')}</ul>` : '<p class="ayuda">— (no hay)</p>';
    let venceTxt = '—';
    if (f.vence) {
      const d = diasEntre(h, Plazos.venceEfectivo(f));
      venceTxt = fechaMilitar(f.vence) + (s === 'cerrada' ? '' : Plazos.venceEfectivo(f) !== f.vence ? ` (venció antes del ${INICIO()}; se cuenta desde ese día${d < 0 ? ': atrasado ' + plural(-d, 'día') : ''})` : d < 0 ? ` (venció hace ${plural(-d, 'día')})` : d === 0 ? ' (hoy)' : ` (en ${plural(d, 'día')})`);
    }
    const cse = Plazos.etapaCse(f, h).etapa;
    const tareas = Plazos.tareas(f, h);
    const ultimas = enExpo ? f.ultimas.slice(-4) : f.ultimas;
    const av = avisos(f, h);
    return `
      <span class="estado-ficha ${s}">${SEMAFORO[s].punto} ${esc(SEMAFORO[s].texto)}</span>
      <h2>${esc(f.titulo)}</h2>
      <p class="ayuda">${numero(f)}${cse ? ` · CSE: <b>${esc(cse)}</b>` : ''}</p>
      <div class="datos-ficha">
        ${dato('Unidad', f.unidad)}${dato('Tipo de ISA', f.tipo)}${dato('Estado', f.estado)}
        ${dato('Fecha de inicio', f.inicio ? fechaMilitar(f.inicio) + (f.inicioSinAnio ? ' (sin año)' : '') : f.inicioTexto)}
        ${dato('Fiscal', f.fiscal)}${dato('Secretario', f.secretario)}
        ${dato('Prórrogas', f.prorrogas)}${dato('Vence la prórroga', venceTxt)}${dato('Lleva abierta', diasAbierta(f, h) != null ? duracion(diasAbierta(f, h)) + ' (' + plural(diasAbierta(f, h), 'día') + ')' : '')}
      </div>
      <div class="casillas">${casilla('CS DIVEDUC', f.csDiveduc)}${casilla('CSE', f.cse)}${casilla('Dictamen', f.dictamen)}</div>
      ${proximasDestacadas(f, h)}
      ${tareas.length ? `<div class="plazos-ficha"><h3>📅 Plazos</h3>${tareas.map((t) => `<div class="tarea-linea ${t.estado}">${PLAZO[t.estado]} <b>${fechaMilitar(t.limite)}</b> <small>(${cuando(t.dias)})</small> — ${esc(t.que)} <small class="ayuda">· ${esc(t.regla)}</small></div>`).join('')}</div>` : ''}
      <div class="seccion-ficha">
        <h3>📝 Resumen del hecho</h3>${f.resumen.split('\n').map((p) => `<p>${esc(p)}</p>`).join('') || '<p class="ayuda">—</p>'}
        <h3>✅ Últimas diligencias${enExpo && f.ultimas.length > 4 ? ' (las 4 más recientes)' : ''}</h3>${tiempo(ultimas, 'ultimas')}
        ${f.notas ? `<h3>🗒 Notas</h3><p>${esc(f.notas)}</p>` : ''}
      </div>
      ${modo === 'ficha' && av.length ? `<div class="avisos-ficha"><b>⚠️ Para revisar</b><ul>${av.map((a) => `<li>${esc(a.texto)}</li>`).join('')}</ul></div>` : ''}`;
  }
  function abrirFicha(id) {
    const f = estado.datos.fichas.find((x) => x.id === id);
    if (!f) return;
    $('fichaContenido').innerHTML = contenidoFicha(f, 'ficha');
    $('ficha').showModal();
    $('ficha').scrollTop = 0;
    prepararFoto(f);
  }

  // ---------- Enviar la ficha como foto (WhatsApp, correo…) ----------
  // La imagen se arma apenas se abre la ficha: así, al tocar el botón, el teléfono deja compartirla de inmediato.
  const foto = { id: null, archivo: null, listo: null };
  function prepararFoto(f) {
    const boton = $('enviarFoto');
    foto.id = f.id; foto.archivo = null;
    boton.disabled = true; boton.textContent = '📷 Preparando foto…';
    if (typeof html2canvas === 'undefined') { boton.hidden = true; return; }
    const caja = document.createElement('div');
    caja.className = 'foto-ficha';
    caja.innerHTML = `<div class="foto-encabezado">📁 ISA ESCINF · al ${fechaMilitar(hoy())}</div>` + contenidoFicha(f, 'foto');
    document.body.appendChild(caja);
    const id = f.id;
    foto.listo = html2canvas(caja, { scale: 2, backgroundColor: '#ffffff', logging: false })
      .then((lienzo) => new Promise((ok) => lienzo.toBlob(ok, 'image/png')))
      .then((blob) => {
        if (foto.id !== id) return;
        const nombre = 'ISA ' + f.titulo.replace(/[\\/:*?"<>|.]+/g, ' ').replace(/\s+/g, ' ').trim() + '.png';
        foto.archivo = new File([blob], nombre, { type: 'image/png' });
        boton.disabled = false; boton.textContent = '📷 Enviar como foto';
      })
      .catch(() => { boton.textContent = '📷 No se pudo armar la foto'; })
      .finally(() => caja.remove());
  }
  async function enviarFoto() {
    const a = foto.archivo;
    if (!a) return;
    // Teléfono (y Safari del Mac): menú «Compartir» → WhatsApp
    if (navigator.canShare && navigator.canShare({ files: [a] })) {
      try { await navigator.share({ files: [a], title: a.name }); return; } catch (e) { if (e.name === 'AbortError') return; }
    }
    // Si no se puede compartir: copiar (para pegar en WhatsApp con ⌘V) y descargar
    let copiada = false;
    try { await navigator.clipboard.write([new ClipboardItem({ 'image/png': a })]); copiada = true; } catch (e) { /* sin permiso */ }
    const url = URL.createObjectURL(a), enlace = document.createElement('a');
    enlace.href = url; enlace.download = a.name; document.body.appendChild(enlace); enlace.click(); enlace.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    alert((copiada ? 'La foto quedó copiada: en WhatsApp pégela con ⌘V (o Ctrl+V).\n\n' : '') + `También se guardó «${a.name}» en Descargas, para adjuntarla.`);
  }

  // ---------- Imprimir ----------
  function imprimirResumen() {
    const d = estado.datos, h = hoy();
    const orden = d.fichas.slice().sort(ordenVence);
    const por = (s) => d.fichas.filter((f) => semaforo(f, h) === s).length;
    const ultimaDe = (f) => f.ultimas.filter((u) => u.fecha).sort((a, b) => a.fecha.localeCompare(b.fecha)).pop() || f.ultimas[f.ultimas.length - 1];
    $('impresion').innerHTML = `
      <h1>Investigaciones sumarias ESCINF — resumen al ${fechaMilitar(h)}</h1>
      <p>${esc(d.archivo)} · ${d.fichas.length} ISA · ${por('vencida')} vencidas · ${por('por_vencer')} vencen en 7 días · ${por('al_dia')} al día · ${por('cerrada')} cerradas</p>
      <table><thead><tr><th></th><th>Afectado / ISA</th><th>Unidad</th><th>Tipo</th><th>Estado</th><th>Fiscal</th><th>Prór.</th><th>Vence</th><th>Abierta</th><th>Última diligencia</th><th>Próxima diligencia</th><th>Plazos atrasados</th></tr></thead>
      <tbody>${orden.map((f) => { const s = semaforo(f, h), u = ultimaDe(f); return `<tr><td>${SEMAFORO[s].punto}</td><td>${esc(f.titulo)}</td><td>${esc(f.unidad)}</td><td>${esc(f.tipoGrupo)}</td><td>${esc(f.estadoGrupo)}</td><td>${esc(f.fiscal)}</td><td>${f.prorrogas ?? ''}</td><td>${f.vence ? fechaMilitar(f.vence) : ''}</td><td>${diasAbierta(f, h) != null ? plural(diasAbierta(f, h), 'día') : ''}</td><td>${u ? (u.fecha ? fechaMilitar(u.fecha) + ' ' : '') + esc(u.resto || u.texto) : ''}</td><td>${esc(f.proximas.map((p) => p.texto).join('\n'))}</td><td>${esc(atrasos(f, h).map((t) => t.que).join('\n'))}</td></tr>`; }).join('')}</tbody></table>`;
    window.print();
  }
  function imprimirAgenda() {
    const h = hoy();
    const grupos = agendaDe($('agendaFiscal').value, $('agendaEstado').value === 'atrasado' ? 'atrasado' : 'urgente');
    if (!grupos.length) { alert('No hay plazos atrasados ni de esta semana para imprimir.'); return; }
    $('impresion').innerHTML = grupos.map((g) => `
      <div class="pagina">
        <h1>Agenda de ISA — ${esc(g.nombre)}</h1>
        <p>Al ${fechaMilitar(h)} · plazos atrasados y de los próximos 7 días · Reglas: prórroga hasta 20 días; CSE cada 30 días hasta que lo citen; demás gestiones cada 15 días.</p>
        <table><thead><tr><th>✔</th><th>Plazo</th><th>ISA</th><th>Qué hacer</th><th>Regla</th></tr></thead>
        <tbody>${g.tareas.map((t) => `<tr><td>☐</td><td>${PLAZO[t.estado]} ${fechaMilitar(t.limite)}<br>${cuando(t.dias)}</td><td>${esc(t.f.titulo)}</td><td>${esc(t.que)}</td><td>${esc(t.regla)}</td></tr>`).join('')}</tbody></table>
        <h2>Sus ISA</h2>
        <table><thead><tr><th></th><th>Afectado / ISA</th><th>Estado</th><th>Prórrogas</th><th>Vence</th><th>Abierta</th><th>CSE</th></tr></thead>
        <tbody>${g.fichas.slice().sort(ordenVence).map((f) => { const s = semaforo(f, h); return `<tr><td>${SEMAFORO[s].punto}</td><td>${esc(f.titulo)}</td><td>${esc(f.estadoGrupo)}</td><td>${f.prorrogas ?? ''}</td><td>${f.vence ? fechaMilitar(f.vence) : ''}</td><td>${diasAbierta(f, h) != null ? plural(diasAbierta(f, h), 'día') : ''}</td><td>${esc(Plazos.etapaCse(f, h).etapa)}</td></tr>`; }).join('')}</tbody></table>
      </div>`).join('');
    window.print();
  }

  // ---------- Pasar a Excel ----------
  function pasarExcel() {
    const bytes = ExcelIsa.escribir(ExcelIsa.armarLibro(estado.datos));
    const url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    const a = document.createElement('a');
    a.href = url;
    a.download = 'ISA ESCINF.xlsx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    alert('Listo: se guardó «ISA ESCINF.xlsx» en Descargas.\n\nÁbralo, revise la columna «Revisar» de la hoja «Diligencias» y desde ahora cargue ese Excel.');
  }

  // ---------- Exponer ----------
  function laminasExpo() {
    const d = estado.datos, h = hoy();
    const por = (s) => d.fichas.filter((f) => semaforo(f, h) === s).length;
    const lams = [`<div class="portada"><h1>INVESTIGACIONES SUMARIAS</h1><p class="fecha-portada">${fechaMilitar(h)}</p>
      <div class="resumen">${[[d.fichas.length, 'ISA'], [d.fichas.filter((f) => f.estadoGrupo !== 'Cerrada').length, 'abiertas'], [por('vencida'), 'vencidas'], [por('por_vencer'), 'vencen en 7 días'], [d.fichas.filter((f) => atrasos(f, h).length).length, 'con plazos atrasados']]
        .map(([n, t]) => `<div class="cifra"><b>${n}</b><span>${t}</span></div>`).join('')}</div></div>`];
    lams.push('<h2>📈 Estadística de las ISA abiertas</h2>' + estadisticaHtml());
    for (const f of d.fichas.slice().sort((a, b) => a.lamina - b.lamina)) lams.push(contenidoFicha(f, 'expo'));
    for (const t of d.tablas) {
      lams.push(`<h2>${esc(t.titulo)}</h2><table><thead><tr>${t.titulos.map((c) => `<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>${t.filas.map((fila) => `<tr>${t.titulos.map((c, i) => `<td>${esc(fila[i])}</td>`).join('')}</tr>`).join('')}</tbody></table>`);
    }
    return lams;
  }
  function verExpo() {
    const lams = laminasExpo();
    estado.expo = Math.max(0, Math.min(estado.expo, lams.length - 1));
    $('expoLamina').innerHTML = lams[estado.expo];
    $('expoLamina').scrollTop = 0;
    $('expoNumero').textContent = `${estado.expo + 1} / ${lams.length}`;
  }
  function exponer() {
    estado.expo = 0;
    $('expo').hidden = false;
    verExpo();
    const el = $('expo');
    if (el.requestFullscreen) el.requestFullscreen().catch(() => {});
    else if (el.webkitRequestFullscreen) el.webkitRequestFullscreen();
  }
  function salirExpo() {
    $('expo').hidden = true;
    if (document.fullscreenElement && document.exitFullscreen) document.exitFullscreen().catch(() => {});
    else if (document.webkitFullscreenElement && document.webkitExitFullscreen) document.webkitExitFullscreen();
  }
  const moverExpo = (n) => { estado.expo += n; verExpo(); };

  // ---------- Eventos ----------
  $('archivo').addEventListener('change', (e) => { cargarArchivo(e.target.files[0]); e.target.value = ''; });
  $('imprimir').addEventListener('click', imprimirResumen);
  $('imprimirAgenda').addEventListener('click', imprimirAgenda);
  $('imprimirEstadistica').addEventListener('click', imprimirEstadistica);
  $('enviarFoto').addEventListener('click', (e) => { e.preventDefault(); enviarFoto(); });
  $('pasarExcel').addEventListener('click', pasarExcel);
  $('exponer').addEventListener('click', exponer);
  $('expoSalir').addEventListener('click', salirExpo);
  $('expoAtras').addEventListener('click', () => moverExpo(-1));
  $('expoAdelante').addEventListener('click', () => moverExpo(1));
  document.addEventListener('keydown', (e) => {
    if ($('expo').hidden) return;
    if (e.key === 'ArrowRight' || e.key === ' ' || e.key === 'PageDown') { e.preventDefault(); moverExpo(1); }
    else if (e.key === 'ArrowLeft' || e.key === 'PageUp') { e.preventDefault(); moverExpo(-1); }
    else if (e.key === 'Escape') salirExpo();
  });
  // Al salir de pantalla completa con Esc del navegador, cerrar también la exposición
  ['fullscreenchange', 'webkitfullscreenchange'].forEach((ev) => document.addEventListener(ev, () => {
    if (!document.fullscreenElement && !document.webkitFullscreenElement) $('expo').hidden = true;
  }));
  $('olvidar').addEventListener('click', () => {
    if (!confirm('¿Borrar los datos guardados en este computador? La próxima vez habrá que cargar el archivo de nuevo.')) return;
    olvidar();
    location.reload();
  });
  $('pestanas').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) cambiarVista(b.dataset.vista); });
  ['busqueda', 'filtroEstado', 'filtroTipo', 'filtroFiscal', 'filtroSemaforo', 'orden'].forEach((id) => $(id).addEventListener('input', verLista));
  ['agendaFiscal', 'agendaEstado'].forEach((id) => $(id).addEventListener('input', verAgenda));
  $('busquedaJudicial').addEventListener('input', verJudiciales);

  document.addEventListener('click', (e) => {
    if (!$('expo').hidden) return;
    const isa = e.target.closest('[data-id]');
    if (isa) { abrirFicha(isa.dataset.id); return; }
    const cifra = e.target.closest('.cifra');
    if (cifra && cifra.dataset.filtro != null) {
      $('busqueda').value = ''; $('filtroEstado').value = ''; $('filtroTipo').value = ''; $('filtroFiscal').value = '';
      $('filtroSemaforo').value = cifra.dataset.filtro;
      cambiarVista('buscar');
      return;
    }
    const fiscal = e.target.closest('[data-fiscal]');
    if (fiscal) {
      $('busqueda').value = ''; $('filtroEstado').value = ''; $('filtroTipo').value = ''; $('filtroSemaforo').value = '';
      $('filtroFiscal').value = fiscal.dataset.fiscal;
      cambiarVista('buscar');
    }
  });
  // Cerrar la ficha tocando fuera de ella
  $('ficha').addEventListener('click', (e) => { if (e.target === $('ficha')) $('ficha').close(); });

  // Arrastrar el archivo a la ventana
  const zona = $('zonaCarga');
  ['dragenter', 'dragover'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); zona.classList.add('encima'); }));
  ['dragleave', 'drop'].forEach((ev) => document.addEventListener(ev, (e) => { e.preventDefault(); zona.classList.remove('encima'); }));
  document.addEventListener('drop', (e) => { const a = e.dataTransfer && e.dataTransfer.files[0]; if (a) cargarArchivo(a); });

  // Al abrir: lo que quedó guardado (si el usuario pidió recordarlo)
  const guardado = recuperar();
  $('recordar').checked = !!guardado;
  if (guardado && guardado.fichas) {
    guardado.fichas.forEach((f) => Plazos.completar(f, +(guardado.fechaExpo || hoy()).slice(0, 4)));
    mostrarDatos(guardado);
  }

  // Instalable y sin internet (solo publicada; abierta con doble clic no hace falta)
  if ('serviceWorker' in navigator && location.protocol !== 'file:') navigator.serviceWorker.register('sw.js').catch(() => {});
})();
