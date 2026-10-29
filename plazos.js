// Plazos de las ISA (reglas del S-1) y tipo de cada diligencia.
//  · La prórroga dura hasta 20 días.
//  · Mientras no lo citen, a la CSE se le pide o reitera cada 30 días.
//  · Todo lo demás (también pedir el informe a la CSE y pedir prórroga) cada 15 días.
(function (raiz) {
  'use strict';
  const L = typeof module !== 'undefined' && module.exports ? require('./leer_ppt.js') : raiz.LeerPpt;
  const { normal, diasEntre, fechaMilitar } = L;

  const REGLAS = { prorrogaMax: 20, cse: 30, gestion: 15, semana: 7 };
  const TIPOS = ['Solicitud CSE', 'Reiteración CSE', 'Citación CSE', 'Asistencia CSE', 'Pide informe CSE', 'Informe CSE recibido', 'Prórroga', 'Otra gestión'];
  const PIDE_CSE = ['Solicitud CSE', 'Reiteración CSE'];

  const sumarDias = (iso, n) => {
    const [a, m, d] = iso.split('-').map(Number);
    return new Date(Date.UTC(a, m - 1, d + n)).toISOString().slice(0, 10);
  };

  // Tipo de una diligencia según sus palabras (solo para lo que viene del PowerPoint; en el Excel se elige de la lista)
  function clasificar(texto) {
    const t = normal(texto);
    if (/PRORROGA/.test(t)) return 'Prórroga';
    const cse = t.search(/\bCSE\b|\bECSO\b|SANIDAD DEL EJERCITO/);
    if (cse < 0) return 'Otra gestión';
    // «Asistió a la Comisión de Sanidad Divisionaria… para pedir hora a la CSE»: lo hecho fue en la CSD
    const otra = t.search(/\bCSD\b|\bCSS\b|DIVISIONARIA/);
    if (otra >= 0 && otra < cse) return 'Otra gestión';
    if (/INFORME/.test(t)) return /^(RECEPCION|RECIBE|SE RECIBIO|LLEGA)/.test(t) ? 'Informe CSE recibido' : 'Pide informe CSE';
    if (/NO ASISTI|INASISTENCIA/.test(t)) return 'Solicitud CSE'; // no fue: hay que pedir hora de nuevo
    if (/ASISTI|ASISTENCIA|CONCURREN/.test(t)) return 'Asistencia CSE';
    const sinFecha = t.replace(/^\d{1,2}[A-Z]{3}(\d{4})?[\s.,]*/, '');
    if (/PARA EL|^HORA CSE|^CITACION|CON HORA|CITADO/.test(sinFecha)) return 'Citación CSE';
    if (/REITER/.test(t)) return 'Reiteración CSE';
    if (/SOLICIT|PETICION|PIDE|ENVI|EVALUACION/.test(t)) return 'Solicitud CSE';
    return 'Otra gestión';
  }

  // Fecha de la diligencia: si el texto trae «para el …», la citación es esa fecha y el trámite es la otra.
  function fechaDiligencia(texto, tipo, anioRef) {
    const t = normal(texto);
    const todas = [];
    let resto = t, corrido = 0;
    for (;;) {
      const f = L.leerFecha(resto, anioRef);
      if (!f) break;
      const i = resto.indexOf(f.escrita);
      todas.push({ ...f, pos: corrido + i });
      corrido += i + f.escrita.length;
      resto = resto.slice(i + f.escrita.length);
    }
    if (!todas.length) return null;
    if (todas.length > 1 && /PARA EL/.test(t)) {
      const pos = t.indexOf('PARA EL');
      const despues = todas.find((f) => f.pos > pos);
      if (tipo === 'Citación CSE' && despues) return despues;
      return todas.filter((f) => f !== despues).pop() || todas[0];
    }
    return todas[0];
  }

  // Deja cada ISA con su lista de diligencias {fecha, tipo, hecha, texto, resto}
  function completar(f, anioRef) {
    if (!f.diligencias) {
      const de = (lista, hecha) => lista.map((d) => {
        const tipo = clasificar(d.texto);
        const fe = fechaDiligencia(d.texto, tipo, anioRef);
        return { ...d, tipo, hecha, fecha: fe ? fe.iso : '', sinAnio: !!(fe && fe.sinAnio), adivinado: true };
      });
      f.diligencias = [...de(f.ultimas || [], true), ...de(f.proximas || [], false)];
    }
    f.ultimas = f.diligencias.filter((d) => d.hecha);
    f.proximas = f.diligencias.filter((d) => !d.hecha);
    return f;
  }

  // ¿En qué va con la CSE?
  function etapaCse(f, hoy) {
    const h = f.diligencias.filter((d) => d.hecha || (d.tipo === 'Citación CSE' && d.fecha));
    const ultimaDe = (tipos) => h.filter((d) => tipos.includes(d.tipo)).map((d) => d.fecha || '0000').sort().pop();
    if (ultimaDe(['Informe CSE recibido'])) return { etapa: 'Informe recibido' };
    // Si ya se pidió el informe, es porque asistió
    const asistio = ultimaDe(['Asistencia CSE', 'Pide informe CSE']);
    const citado = ultimaDe(['Citación CSE']);
    if (asistio && (!citado || asistio >= citado || citado === '0000')) return { etapa: 'Asistió, espera informe', desde: asistio };
    if (citado && citado !== '0000' && citado >= hoy) return { etapa: 'Citado', desde: citado };
    const pidio = ultimaDe(PIDE_CSE);
    if (citado && citado !== '0000') return { etapa: 'Sin citar', desde: [citado, pidio || ''].sort().pop(), noAsistio: true };
    if (pidio) return { etapa: 'Sin citar', desde: pidio === '0000' ? '' : pidio };
    if (citado) return { etapa: 'Citado' };
    return { etapa: '' };
  }

  function estadoPlazo(limite, hoy) {
    const d = diasEntre(hoy, limite);
    return { dias: d, estado: d < 0 ? 'atrasado' : d <= REGLAS.semana ? 'semana' : 'al_dia' };
  }

  // Lo que le toca a la ISA: [{que, limite, estado, dias, regla}]
  function tareas(f, hoy) {
    if (f.estadoGrupo === 'Cerrada') return [];
    const lista = [];
    const hechas = f.diligencias.filter((d) => d.hecha && d.fecha);
    // Si la próxima prórroga ya está anotada como pendiente, basta con esa
    if (f.vence && !f.proximas.some((p) => p.tipo === 'Prórroga' && p.fecha && p.fecha <= f.vence && p.fecha >= hoy)) {
      lista.push({ que: `Pedir prórroga${f.prorrogas != null ? ' N.° ' + (f.prorrogas + 1) : ''} (antes de que venza)`, limite: f.vence, regla: 'prórroga', ...estadoPlazo(f.vence, hoy) });
    }
    const cse = etapaCse(f, hoy);
    if (cse.etapa === 'Sin citar') {
      const base = cse.desde || f.inicio;
      if (base) {
        const limite = sumarDias(base, REGLAS.cse);
        lista.push({ que: cse.noAsistio ? 'Pedir nueva hora a la CSE (la citación ya pasó)' : 'Reiterar la solicitud a la CSE', limite, regla: 'CSE cada 30 días', desde: base, ...estadoPlazo(limite, hoy) });
      }
    }
    const gestiones = hechas.filter((d) => !PIDE_CSE.includes(d.tipo) && d.tipo !== 'Citación CSE').map((d) => d.fecha).sort();
    const base = gestiones.pop() || f.inicio;
    if (base) {
      const limite = sumarDias(base, REGLAS.gestion);
      const sugerida = f.proximas.find((p) => !PIDE_CSE.includes(p.tipo));
      const que = cse.etapa === 'Asistió, espera informe' ? 'Pedir el informe a la CSE u otra gestión' : 'Próxima gestión';
      lista.push({ que: que + (sugerida ? ': ' + (sugerida.resto || sugerida.texto) : ''), limite, regla: 'gestión cada 15 días', desde: base, ...estadoPlazo(limite, hoy) });
    }
    for (const p of f.proximas) {
      if (p.fecha && p.fecha >= hoy) lista.push({ que: 'Programada: ' + (p.resto || p.texto), limite: p.fecha, regla: 'programada', ...estadoPlazo(p.fecha, hoy) });
    }
    return lista.sort((a, b) => a.limite.localeCompare(b.limite));
  }

  // Avisos de las reglas para «Revisar planilla»
  function revisar(f, hoy) {
    if (f.estadoGrupo === 'Cerrada') return [];
    const avisos = [];
    const ultimaPr = f.diligencias.filter((d) => d.hecha && d.tipo === 'Prórroga' && d.fecha).map((d) => d.fecha).sort().pop();
    if (f.vence) {
      const desde = ultimaPr && ultimaPr <= f.vence ? ultimaPr : hoy;
      const largo = diasEntre(desde, f.vence);
      if (largo > REGLAS.prorrogaMax) {
        avisos.push({ nivel: 'amarillo', texto: ultimaPr && desde === ultimaPr
          ? `La prórroga pedida el ${fechaMilitar(ultimaPr)} vence el ${fechaMilitar(f.vence)}: son ${largo} días (máximo ${REGLAS.prorrogaMax}).`
          : `Vence el ${fechaMilitar(f.vence)}, en ${largo} días: una prórroga dura hasta ${REGLAS.prorrogaMax}. Revisar la fecha.` });
      }
    }
    for (const t of tareas(f, hoy)) {
      if (t.estado === 'atrasado' && t.regla !== 'prórroga' && t.regla !== 'programada') {
        avisos.push({ nivel: 'rojo', texto: `${t.que}: atrasado ${-t.dias} día${t.dias === -1 ? '' : 's'} (${t.regla}, desde el ${fechaMilitar(t.desde)}).` });
      }
    }
    const sinFecha = f.diligencias.filter((d) => d.hecha && !d.fecha).length;
    if (sinFecha) avisos.push({ nivel: 'info', texto: `${sinFecha} diligencia${sinFecha === 1 ? '' : 's'} hecha${sinFecha === 1 ? '' : 's'} sin fecha: no se pueden contar los plazos con ella${sinFecha === 1 ? '' : 's'}.` });
    return avisos;
  }

  const Plazos = { REGLAS, TIPOS, clasificar, fechaDiligencia, completar, etapaCse, tareas, revisar, sumarDias };
  if (typeof module !== 'undefined' && module.exports) module.exports = Plazos;
  else raiz.Plazos = Plazos;
})(this);
