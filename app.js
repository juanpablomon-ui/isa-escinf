// ISA ESCINF: consulta rápida de las investigaciones sumarias desde el PowerPoint de la exposición.
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

  const $ = (id) => document.getElementById(id);
  const estado = { datos: null, vista: 'tablero' };
  const { normal, fechaMilitar, diasEntre, semaforo, revisar } = LeerPpt;
  const esc = (t) => String(t == null ? '' : t).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const hoy = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
  const vacio = (t) => !String(t || '').trim() || /^-\.?-$/.test(String(t).trim());

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

  // ---------- Cargar el PowerPoint ----------
  async function inflar(datos) {
    const flujo = new Blob([datos]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(flujo).arrayBuffer());
  }
  async function cargarArchivo(archivo) {
    if (!archivo) return;
    if (!/\.pptx$/i.test(archivo.name)) { alert('Ese archivo no es un PowerPoint .pptx.\n\nSi está en .ppt (antiguo), ábralo en PowerPoint y guárdelo como .pptx.'); return; }
    if (typeof DecompressionStream === 'undefined') { alert('Este navegador es muy antiguo para leer el PowerPoint. Use Chrome, Edge o Safari actualizados.'); return; }
    try {
      const datos = await LeerPpt.leer(new Uint8Array(await archivo.arrayBuffer()), inflar, archivo.name);
      if (!datos.fichas.length) throw new Error('No encontré láminas de ISA (con casillas «TIPO DE ISA» o «FISCAL»).');
      if ($('recordar').checked) guardar(datos); else olvidar();
      mostrarDatos(datos);
    } catch (e) {
      alert('No pude leer el PowerPoint.\n\n' + e.message);
    }
  }

  // ---------- Mostrar ----------
  function mostrarDatos(datos) {
    estado.datos = datos;
    $('bienvenida').hidden = true;
    $('pestanas').hidden = false;
    $('imprimir').hidden = false;
    $('olvidar').hidden = !recuperar();
    $('fuente').textContent = (datos.archivo || 'PowerPoint cargado') + (datos.fechaExpo ? ' · exposición del ' + fechaMilitar(datos.fechaExpo) : '');
    $('pestanaJudiciales').hidden = !datos.tablas.length;

    const opciones = (sel, titulo, valores) => {
      sel.innerHTML = `<option value="">${titulo}</option>` + valores.map(([v, t]) => `<option value="${esc(v)}">${esc(t)}</option>`).join('');
    };
    const unicos = (fn) => [...new Set(datos.fichas.map(fn).filter(Boolean))].sort();
    opciones($('filtroEstado'), 'Estado: todos', unicos((f) => f.estadoGrupo).map((v) => [v, v]));
    opciones($('filtroTipo'), 'Tipo: todos', unicos((f) => f.tipoGrupo).map((v) => [v, v]));
    opciones($('filtroFiscal'), 'Fiscal: todos', fiscales().map((g) => [g.clave, g.nombre]));

    const nRev = datos.fichas.filter((f) => revisar(f, hoy()).some((a) => a.nivel !== 'info')).length;
    $('nRevisar').textContent = nRev || '';
    cambiarVista(estado.vista);
  }

  function cambiarVista(v) {
    estado.vista = v;
    document.querySelectorAll('.pestanas button').forEach((b) => b.classList.toggle('activa', b.dataset.vista === v));
    document.querySelectorAll('.vista').forEach((s) => { s.hidden = s.id !== 'vista-' + v; });
    ({ tablero: verTablero, buscar: verLista, fiscales: verFiscales, judiciales: verJudiciales, revisar: verRevisar }[v] || (() => {}))();
    window.scrollTo(0, 0);
  }

  // Una ISA en una línea de la lista
  function tarjeta(f, resaltar) {
    const s = semaforo(f, hoy());
    const marcar = (t) => {
      if (!resaltar) return esc(t);
      let html = esc(t);
      for (const p of resaltar) {
        const re = new RegExp('(' + p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&').split('').map((c) => c + '[̀-ͯ]?').join('') + ')', 'gi');
        html = html.replace(re, '<mark>$1</mark>');
      }
      return html;
    };
    let derecha;
    if (s === 'cerrada') derecha = '<b>Cerrada</b>';
    else if (!f.vence) derecha = '<b>Sin fecha</b><small>de vencimiento</small>';
    else {
      const d = diasEntre(hoy(), f.vence);
      const cuando = d < 0 ? `<span class="txt-rojo">venció hace ${-d} d</span>` : d === 0 ? '<span class="txt-amarillo">vence hoy</span>' : d <= 7 ? `<span class="txt-amarillo">vence en ${d} d</span>` : `vence en ${d} d`;
      derecha = `<b>${SEMAFORO[s].punto} ${fechaMilitar(f.vence)}</b><small>${cuando}</small>`;
    }
    const prox = f.proximas[0];
    return `<button class="item ${s}" data-id="${f.id}">
      <span class="nombre">${marcar(f.titulo)}</span>
      <span class="derecha">${derecha}${f.prorrogas != null ? `<small>${f.prorrogas} prórroga${f.prorrogas === 1 ? '' : 's'}</small>` : ''}</span>
      <span class="linea">${marcar(f.unidad)} · <span class="etiqueta">${esc(f.tipoGrupo)}</span> <span class="etiqueta">${esc(f.estadoGrupo)}</span> · Fiscal: ${marcar(f.fiscal || '—')}</span>
      ${prox ? `<span class="proxima">➡️ ${marcar(prox.texto)}${f.proximas.length > 1 ? ` <small class="ayuda">(+${f.proximas.length - 1})</small>` : ''}</span>` : ''}
    </button>`;
  }

  function barras(destino, pares) {
    const max = Math.max(1, ...pares.map((p) => p[1]));
    $(destino).innerHTML = pares.map(([t, n]) => `<div class="barra-dato"><span>${esc(t)}</span><div class="fondo-barra"><div class="relleno" style="width:${(100 * n) / max}%"></div></div><span class="num">${n}</span></div>`).join('');
  }
  const contar = (fn) => {
    const c = {};
    for (const f of estado.datos.fichas) { const k = fn(f); c[k] = (c[k] || 0) + 1; }
    return Object.entries(c).sort((a, b) => b[1] - a[1]);
  };

  // ---------- Tablero ----------
  function verTablero() {
    const fs = estado.datos.fichas, h = hoy();
    const por = (s) => fs.filter((f) => semaforo(f, h) === s).length;
    const abiertas = fs.filter((f) => f.estadoGrupo !== 'Cerrada').length;
    const cifras = [
      ['', fs.length, 'ISA en total', ''],
      ['abiertas', abiertas, 'abiertas o en trámite', ''],
      ['vencida', por('vencida'), 'vencidas', 'rojo'],
      ['por_vencer', por('por_vencer'), 'vencen en 7 días', 'amarillo'],
      ['al_dia', por('al_dia'), 'al día', ''],
      ['cerrada', por('cerrada'), 'cerradas', '']
    ];
    $('cifras').innerHTML = cifras.map(([filtro, n, t, c]) => `<button class="cifra ${c}" data-filtro="${filtro}"><b>${n}</b><span>${t}</span></button>`).join('');
    $('hoyTexto').textContent = `Al día de hoy, ${fechaMilitar(h)}. Las ISA abiertas, de la que vence primero a la última.`;

    const abiertasOrden = fs.filter((f) => f.estadoGrupo !== 'Cerrada').sort((a, b) => (a.vence || '9999').localeCompare(b.vence || '9999'));
    $('listaVencimientos').innerHTML = abiertasOrden.map((f) => tarjeta(f)).join('') || '<p class="ayuda">No hay ISA abiertas.</p>';
    barras('barrasTipo', contar((f) => f.tipoGrupo));
    barras('barrasEstado', contar((f) => f.estadoGrupo));
    $('listaProrrogas').innerHTML = fs.filter((f) => f.prorrogas != null && f.estadoGrupo !== 'Cerrada')
      .sort((a, b) => b.prorrogas - a.prorrogas).slice(0, 5)
      .map((f) => `<button class="item ${semaforo(f, h)}" data-id="${f.id}"><span class="nombre">${esc(f.titulo)}</span><span class="derecha"><b>${f.prorrogas}</b><small>prórrogas</small></span><span class="linea">Desde ${f.inicio ? fechaMilitar(f.inicio) : esc(f.inicioTexto)} · ${esc(f.fiscal)}</span></button>`).join('');
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
      if (fsem === 'abiertas' ? f.estadoGrupo === 'Cerrada' : fsem && semaforo(f, h) !== fsem) return false;
      if (!palabras.length) return true;
      const texto = normal([f.titulo, f.unidad, f.tipo, f.estado, f.fiscal, f.secretario, f.resumen, ...f.ultimas.map((u) => u.texto), ...f.proximas.map((p) => p.texto)].join(' '));
      return palabras.every((p) => texto.includes(p));
    });
    const orden = $('orden').value;
    const cmp = {
      vence: (a, b) => (a.estadoGrupo === 'Cerrada') - (b.estadoGrupo === 'Cerrada') || (a.vence || '9999').localeCompare(b.vence || '9999'),
      lamina: (a, b) => a.lamina - b.lamina,
      prorrogas: (a, b) => (b.prorrogas || 0) - (a.prorrogas || 0),
      inicio: (a, b) => (a.inicio || '9999').localeCompare(b.inicio || '9999'),
      nombre: (a, b) => a.titulo.localeCompare(b.titulo, 'es')
    }[orden];
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
      const pronto = abiertas.filter((f) => semaforo(f, h) === 'por_vencer').length;
      const filas = g.fichas.slice().sort((a, b) => (a.vence || '9999').localeCompare(b.vence || '9999'))
        .map((f) => { const s = semaforo(f, h); return `<div class="fila"><span>${esc(f.titulo)}</span><span>${SEMAFORO[s].punto} ${f.vence ? fechaMilitar(f.vence) : esc(SEMAFORO[s].texto)}</span></div>`; }).join('');
      return `<button class="tarjeta" data-fiscal="${esc(g.clave)}">
        <h3>${esc(g.nombre)}</h3>
        <div class="cuenta">${g.fichas.length} ISA (${abiertas.length} abierta${abiertas.length === 1 ? '' : 's'})${venc ? ` · <span class="txt-rojo">${venc} vencida${venc === 1 ? '' : 's'}</span>` : ''}${pronto ? ` · <span class="txt-amarillo">${pronto} por vencer</span>` : ''}</div>
        ${filas}
      </button>`;
    }).join('');
  }

  // ---------- Procesos judiciales ----------
  function verJudiciales() {
    const palabras = normal($('busquedaJudicial').value).split(/\s+/).filter(Boolean);
    $('tablasJudiciales').innerHTML = estado.datos.tablas.map((t) => {
      const filas = t.filas.filter((fila) => { const tx = normal(fila.join(' ')); return palabras.every((p) => tx.includes(p)); });
      return `<h3>${esc(t.titulo)} <small class="ayuda">(lámina ${t.lamina})</small></h3>
        <div class="tabla-scroll"><table><thead><tr>${t.titulos.map((c) => `<th>${esc(c)}</th>`).join('')}</tr></thead>
        <tbody>${filas.map((fila) => `<tr>${fila.map((c) => `<td>${esc(c)}</td>`).join('')}</tr>`).join('') || `<tr><td colspan="${t.titulos.length}">Nada con esa búsqueda.</td></tr>`}</tbody></table></div>`;
    }).join('');
  }

  // ---------- Revisar planilla ----------
  function verRevisar() {
    const h = hoy();
    const conAvisos = estado.datos.fichas.map((f) => ({ f, avisos: revisar(f, h) })).filter((x) => x.avisos.length);
    const peso = (x) => (x.avisos.some((a) => a.nivel === 'rojo') ? 0 : x.avisos.some((a) => a.nivel === 'amarillo') ? 1 : 2);
    conAvisos.sort((a, b) => peso(a) - peso(b) || a.f.lamina - b.f.lamina);
    const icono = { rojo: '🔴', amarillo: '🟡', info: 'ℹ️' };
    $('listaRevisar').innerHTML = conAvisos.length ? conAvisos.map(({ f, avisos }) => `
      <div class="bloque ${['rojo', 'amarillo', ''][peso({ avisos })]}">
        <h4 data-id="${f.id}">Lámina ${f.lamina} · ${esc(f.titulo)}</h4>
        <ul>${avisos.map((a) => `<li>${icono[a.nivel]} ${esc(a.texto)}</li>`).join('')}</ul>
      </div>`).join('') : '<p class="panel">✅ No hay nada que revisar. ¡Bien!</p>';
  }

  // ---------- Ficha ----------
  function abrirFicha(id) {
    const f = estado.datos.fichas.find((x) => x.id === id);
    if (!f) return;
    const h = hoy(), s = semaforo(f, h);
    const dato = (t, v) => `<div class="dato"><small>${t}</small><b>${esc(vacio(v) ? '—' : v)}</b></div>`;
    const casilla = (t, v) => `<span class="casilla">${t}:<b>${esc(vacio(v) ? '—' : v)}</b></span>`;
    const tiempo = (lista, clase) => lista.length ? `<ul class="tiempo ${clase}">${lista.map((d) => {
      const pasada = clase === 'proximas' && d.fecha && d.fecha < h;
      return `<li class="${pasada ? 'pasada' : ''}">${d.fecha ? `<span class="fecha">${fechaMilitar(d.fecha)}</span>` : ''}${esc(d.fecha ? d.resto : d.texto)}${pasada ? ' <span class="txt-amarillo">(fecha ya pasada)</span>' : ''}</li>`;
    }).join('')}</ul>` : '<p class="ayuda">— (no hay en la lámina)</p>';
    let venceTxt = '—';
    if (f.vence) {
      const d = diasEntre(h, f.vence);
      venceTxt = fechaMilitar(f.vence) + (s === 'cerrada' ? '' : d < 0 ? ` (venció hace ${-d} días)` : d === 0 ? ' (hoy)' : ` (en ${d} días)`);
    }
    const avisos = revisar(f, h);
    $('fichaContenido').innerHTML = `
      <span class="estado-ficha ${s}">${SEMAFORO[s].punto} ${esc(SEMAFORO[s].texto)}</span>
      <h2>${esc(f.titulo)}</h2>
      <p class="ayuda">Lámina ${f.lamina} de la presentación</p>
      <div class="datos-ficha">
        ${dato('Unidad', f.unidad)}${dato('Tipo de ISA', f.tipo)}${dato('Estado', f.estado)}
        ${dato('Fecha de inicio', f.inicio ? fechaMilitar(f.inicio) + (f.inicioSinAnio ? ' (sin año en la lámina)' : '') : f.inicioTexto)}
        ${dato('Fiscal', f.fiscal)}${dato('Secretario', f.secretario)}
        ${dato('Prórrogas', f.prorrogas)}${dato('Vence', venceTxt)}
      </div>
      <div class="casillas">${casilla('CS DIVEDUC', f.csDiveduc)}${casilla('CSE', f.cse)}${casilla('Dictamen', f.dictamen)}</div>
      <div class="seccion-ficha">
        <h3>📝 Resumen del hecho</h3>${f.resumen.split('\n').map((p) => `<p>${esc(p)}</p>`).join('') || '<p class="ayuda">—</p>'}
        <h3>✅ Últimas diligencias</h3>${tiempo(f.ultimas, 'ultimas')}
        <h3>➡️ Próximas diligencias</h3>${tiempo(f.proximas, 'proximas')}
      </div>
      ${avisos.length ? `<div class="avisos-ficha"><b>⚠️ Para revisar en la lámina</b><ul>${avisos.map((a) => `<li>${esc(a.texto)}</li>`).join('')}</ul></div>` : ''}`;
    $('ficha').showModal();
    $('ficha').scrollTop = 0;
  }

  // ---------- Imprimir resumen ----------
  function imprimir() {
    const d = estado.datos, h = hoy();
    const orden = d.fichas.slice().sort((a, b) => (a.estadoGrupo === 'Cerrada') - (b.estadoGrupo === 'Cerrada') || (a.vence || '9999').localeCompare(b.vence || '9999'));
    const por = (s) => d.fichas.filter((f) => semaforo(f, h) === s).length;
    const ultimaDe = (f) => f.ultimas.slice().sort((a, b) => (a.fecha || '').localeCompare(b.fecha || '')).pop();
    $('impresion').innerHTML = `
      <h1>Investigaciones sumarias ESCINF — resumen al ${fechaMilitar(h)}</h1>
      <p>${esc(d.archivo)} · ${d.fichas.length} ISA · ${por('vencida')} vencidas · ${por('por_vencer')} vencen en 7 días · ${por('al_dia')} al día · ${por('cerrada')} cerradas</p>
      <table><thead><tr><th></th><th>Afectado / ISA</th><th>Unidad</th><th>Tipo</th><th>Estado</th><th>Fiscal</th><th>Prór.</th><th>Vence</th><th>Última diligencia</th><th>Próxima diligencia</th></tr></thead>
      <tbody>${orden.map((f) => { const s = semaforo(f, h), u = ultimaDe(f); return `<tr><td>${SEMAFORO[s].punto}</td><td>${esc(f.titulo)}</td><td>${esc(f.unidad)}</td><td>${esc(f.tipoGrupo)}</td><td>${esc(f.estadoGrupo)}</td><td>${esc(f.fiscal)}</td><td>${f.prorrogas ?? ''}</td><td>${f.vence ? fechaMilitar(f.vence) : ''}</td><td>${esc(u ? u.texto : '')}</td><td>${esc(f.proximas.map((p) => p.texto).join('\n'))}</td></tr>`; }).join('')}</tbody></table>`;
    window.print();
  }

  // ---------- Eventos ----------
  $('archivo').addEventListener('change', (e) => { cargarArchivo(e.target.files[0]); e.target.value = ''; });
  $('imprimir').addEventListener('click', imprimir);
  $('olvidar').addEventListener('click', () => {
    if (!confirm('¿Borrar los datos guardados en este computador? La próxima vez habrá que cargar el PowerPoint de nuevo.')) return;
    olvidar();
    location.reload();
  });
  $('pestanas').addEventListener('click', (e) => { const b = e.target.closest('button'); if (b) cambiarVista(b.dataset.vista); });
  ['busqueda', 'filtroEstado', 'filtroTipo', 'filtroFiscal', 'filtroSemaforo', 'orden'].forEach((id) => $(id).addEventListener('input', verLista));
  $('busquedaJudicial').addEventListener('input', verJudiciales);

  document.addEventListener('click', (e) => {
    const isa = e.target.closest('[data-id]');
    if (isa) { abrirFicha(isa.dataset.id); return; }
    const cifra = e.target.closest('.cifra');
    if (cifra) {
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
  if (guardado && guardado.fichas) mostrarDatos(guardado);
})();
