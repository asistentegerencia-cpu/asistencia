/**
 * RECORDATORIOS DE SALIDA — Asistencia Grupo Robles
 *
 * Archivo independiente: no modifica nada del resto del script. Solo LEE las
 * hojas Personal, Config, Feriados, Programacion y Marcajes.
 *
 * Qué hace:
 *  1. Todas las mañanas crea, en un calendario propio ("Asistencia · Marcar salida"),
 *     un evento a la hora de salida de cada colaborador que trabaja ese día, con su
 *     correo como invitado. El evento aparece en SU Google Calendar y el celular le
 *     avisa, aunque la app de asistencia esté cerrada.
 *     La hora sale de su programación publicada si la tiene; si no, de su turno.
 *  2. (Opcional) Cada 15 minutos revisa quién marcó entrada y no salida, pasada su
 *     hora de salida, y le manda un correo recordándole marcarla.
 *  3. (Opcional) En la misma revisión, a quien ya pasó su hora de entrada (más
 *     unos minutos de margen) y no marcó, le manda un correo para que la marque.
 *     No avisa a quien tiene un permiso del día o de tardanza, ni a quien está en
 *     una ausencia registrada (comisión, licencia, vacaciones).
 *
 * Cómo instalarlo (una sola vez):
 *  1. En el editor de Apps Script: Archivo › Nuevo › Script, nómbralo "Recordatorios"
 *     y pega todo este contenido.
 *  2. Elige la función  instalarRecordatorios  y pulsa Ejecutar. Acepta los permisos
 *     (Calendar, y Gmail si activas el correo).
 *  3. Listo. Para quitarlo: ejecuta  desinstalarRecordatorios.
 *  Si ya lo tenías instalado, vuelve a ejecutar  instalarRecordatorios  una vez para
 *  que empiece a revisar también las entradas.
 *
 * Importante:
 *  - Solo reciben aviso quienes tienen correo en la hoja Personal.
 *  - Google Calendar avisa según la configuración de notificaciones de cada persona
 *    (por defecto, unos minutos antes). Que tengan la app de Google Calendar en el
 *    celular con esa cuenta.
 */

const REC_SALIDA = {
  calendario:     'Asistencia · Marcar salida',
  etiqueta:       'asistencia-salida',
  diasAdelante:   1,     // crea los de hoy y mañana
  correo:         true,  // aviso por correo a quien no marcó salida
  margenCorreoMin: 20,   // minutos después de su hora de salida
  entrada:        true,  // aviso por correo a quien no marcó entrada
  margenEntradaMin: 15   // minutos después de su hora de entrada
};

/* ---------------- Instalación ---------------- */

function instalarRecordatorios(){
  desinstalarRecordatorios();
  ScriptApp.newTrigger('crearRecordatoriosSalida').timeBased().everyDays(1).atHour(5).create();
  if(REC_SALIDA.correo || REC_SALIDA.entrada)
    ScriptApp.newTrigger('revisarPendientes').timeBased().everyMinutes(15).create();
  crearRecordatoriosSalida();
}

function desinstalarRecordatorios(){
  ScriptApp.getProjectTriggers()
    .filter(t => ['crearRecordatoriosSalida','avisarSalidasPendientes','revisarPendientes'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

/* ---------------- 1. Eventos en Google Calendar ---------------- */

function crearRecordatoriosSalida(){
  const ctx = recSal_contexto();
  const cal = recSal_calendario();
  const ahora = new Date();

  for(let i = 0; i <= REC_SALIDA.diasAdelante; i++){
    const f = Utilities.formatDate(new Date(ahora.getTime() + i*86400000), ctx.tz, 'yyyy-MM-dd');
    const hechos = {};
    cal.getEventsForDay(recSal_aFecha(f, '12:00', ctx.tz)).forEach(ev => {
      const t = ev.getTag(REC_SALIDA.etiqueta);
      if(t) hechos[t] = true;
    });

    ctx.personal.forEach(p => {
      if(!p.email || !recSal_activo(p) || p.modalidad === 'EXENTO') return;
      const clave = p.id + '|' + f;
      if(hechos[clave]) return;                               // ya creado
      const salida = recSal_horaSalida(p, f, ctx);
      if(!salida) return;                                     // ese día no trabaja
      const ini = recSal_aFecha(f, salida, ctx.tz);
      if(ini.getTime() < ahora.getTime()) return;             // ya pasó
      const ev = cal.createEvent('⏰ Marca tu salida', ini, new Date(ini.getTime() + 5*60000), {
        description: 'Antes de irte, marca tu SALIDA en la app de asistencia.\n' +
                     'Si ya estás fuera de la oficina, igual puedes marcarla.',
        guests: p.email,
        sendInvites: false
      });
      ev.setTag(REC_SALIDA.etiqueta, clave);
      ev.removeAllReminders();
      ev.addPopupReminder(10);
      ev.addPopupReminder(0);
    });
  }
}

function recSal_calendario(){
  const ya = CalendarApp.getCalendarsByName(REC_SALIDA.calendario);
  if(ya.length) return ya[0];
  return CalendarApp.createCalendar(REC_SALIDA.calendario, {
    summary: 'Recordatorios automáticos para marcar la salida', color: CalendarApp.Color.ORANGE
  });
}

/* ---------------- 2 y 3. Correos a quien no marcó ---------------- */

function revisarPendientes(){
  if(REC_SALIDA.correo) avisarSalidasPendientes();
  if(REC_SALIDA.entrada) avisarEntradasPendientes();
}

function avisarEntradasPendientes(){
  const ctx = recSal_contexto();
  const ahora = new Date();
  const hoy = Utilities.formatDate(ahora, ctx.tz, 'yyyy-MM-dd');
  const minAhora = Number(Utilities.formatDate(ahora, ctx.tz, 'H'))*60 + Number(Utilities.formatDate(ahora, ctx.tz, 'm'));
  const props = PropertiesService.getScriptProperties();

  const entrada = {};
  recSal_leerTabla(ctx.ss.getSheetByName('Marcajes'), 400).forEach(m => {
    if(recSal_fechaISO(m.fecha, ctx.tz) === hoy && m.tipo === 'ENTRADA') entrada[String(m.emp_id)] = true;
  });

  // Quién no tiene que marcar hoy: permiso del día o de tardanza (aprobado o por
  // responder) y ausencias registradas que cubren hoy.
  const libre = {};
  recSal_leerTabla(ctx.ss.getSheetByName('Permisos'), 600).forEach(x => {
    if(recSal_fechaISO(x.fecha, ctx.tz) !== hoy || String(x.estado) === 'RECHAZADO') return;
    if(['DIA','TARDANZA','JORNADA'].indexOf(String(x.tipo)) >= 0) libre[String(x.emp_id)] = true;
  });
  recSal_leerTabla(ctx.ss.getSheetByName('Ausencias')).forEach(a => {
    const d = recSal_fechaISO(a.desde, ctx.tz), h = recSal_fechaISO(a.hasta, ctx.tz) || d;
    if(d && d <= hoy && hoy <= h) libre[String(a.emp_id)] = true;
  });

  ctx.personal.forEach(p => {
    if(!p.email || !recSal_activo(p) || p.modalidad === 'EXENTO' || entrada[p.id] || libre[p.id]) return;
    const hor = recSal_horario(p, hoy, ctx);
    if(!hor || !hor.in) return;
    const ini = recSal_aMin(hor.in);
    // Solo dentro de su jornada: pasado el margen y antes de su hora de salida
    if(minAhora < ini + REC_SALIDA.margenEntradaMin) return;
    if(hor.out && recSal_aMin(hor.out) > ini && minAhora >= recSal_aMin(hor.out)) return;
    const clave = 'avisoEntrada|' + p.id + '|' + hoy;
    if(props.getProperty(clave)) return;                      // uno por día
    MailApp.sendEmail(p.email, 'No has marcado tu entrada de hoy',
      'Hola ' + recSal_nombreCorto(p.nombre) + ',\n\n' +
      'Tu horario de hoy empezaba a las ' + hor.in + ' y todavía no marcas tu entrada.\n' +
      'Si ya llegaste, márcala ahora en la app de asistencia.\n' +
      'Si hoy no vienes o vas a llegar tarde, pide el permiso desde la app.\n\n' +
      'Asistencia ' + (ctx.cfg.empresa || ''));
    props.setProperty(clave, '1');
  });
}

function avisarSalidasPendientes(){
  const ctx = recSal_contexto();
  const ahora = new Date();
  const hoy = Utilities.formatDate(ahora, ctx.tz, 'yyyy-MM-dd');
  const minAhora = Number(Utilities.formatDate(ahora, ctx.tz, 'H'))*60 + Number(Utilities.formatDate(ahora, ctx.tz, 'm'));
  const props = PropertiesService.getScriptProperties();

  // Marcajes de hoy: quién tiene entrada y quién salida
  const mar = recSal_leerTabla(ctx.ss.getSheetByName('Marcajes'), 400);
  const entrada = {}, salida = {};
  mar.forEach(m => {
    if(recSal_fechaISO(m.fecha, ctx.tz) !== hoy) return;
    if(m.tipo === 'ENTRADA') entrada[m.emp_id] = true;
    if(m.tipo === 'SALIDA')  salida[m.emp_id]  = true;
  });

  ctx.personal.forEach(p => {
    if(!p.email || !recSal_activo(p) || !entrada[p.id] || salida[p.id]) return;
    const hs = recSal_horaSalida(p, hoy, ctx);
    if(!hs || minAhora < recSal_aMin(hs) + REC_SALIDA.margenCorreoMin) return;
    const clave = 'avisoSalida|' + p.id + '|' + hoy;
    if(props.getProperty(clave)) return;                      // uno por día
    MailApp.sendEmail(p.email, 'No has marcado tu salida de hoy',
      'Hola ' + recSal_nombreCorto(p.nombre) + ',\n\n' +
      'Tu horario de hoy terminaba a las ' + hs + ' y todavía no marcas tu salida.\n' +
      'Ábrela y márcala ahora; si ya saliste de la oficina, igual puedes hacerlo.\n\n' +
      'Asistencia ' + (ctx.cfg.empresa || ''));
    props.setProperty(clave, '1');
  });
}

/* ---------------- Horario del día (mismo criterio que la app) ---------------- */

function recSal_horaSalida(p, f, ctx){
  const h = recSal_horario(p, f, ctx);
  return h ? h.out : null;
}

/* {in, out} del día, o null si ese día no trabaja. */
function recSal_horario(p, f, ctx){
  // 1) Programación publicada o validada: manda sobre todo
  const tramos = ctx.prog.filter(a => a.emp_id === p.id && a.fecha === f && a.estado !== 'BORRADOR' && a.salida);
  if(tramos.length) return {in: tramos.map(a => a.entrada).filter(String).sort()[0] || '',
                            out: tramos.map(a => a.salida).sort().pop()};

  // 2) Horario propio de la persona (columna "dias" de Personal), si lo tiene
  if(ctx.feriados[f]) return null;
  const propios = recSal_json(p.dias, null);
  if(Array.isArray(propios) && propios.length === 7){
    const v = propios[recSal_aFecha(f, '12:00', ctx.tz).getDay()];
    return (v && v.out) ? {in: v.in || '', out: v.out} : null;
  }

  // 3) Turno, según el calendario de la empresa
  const t = ctx.turnos.filter(x => x.id === p.turno)[0] || ctx.turnos[0];
  if(!t) return null;
  const dia = recSal_aFecha(f, '12:00', ctx.tz).getDay();
  if(dia === 0) return (recSal_siNo(t.trabajaDom) && ctx.dom.indexOf(f) >= 0) ? {in:t.domEntrada, out:t.domSalida} : null;
  if(dia === 6) return (recSal_siNo(t.trabajaSab) && ctx.sab.indexOf(f) >= 0) ? {in:t.sabEntrada, out:t.sabSalida} : null;
  return {in:t.entrada, out:t.salida};
}

/* ---------------- Lectura de la hoja ---------------- */

function recSal_contexto(){
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const cfg = {};
  recSal_leerTabla(ss.getSheetByName('Config')).forEach(r => { cfg[r.clave] = r.valor; });
  const feriados = {};
  recSal_leerTabla(ss.getSheetByName('Feriados')).forEach(r => { feriados[recSal_fechaISO(r.fecha, tz)] = true; });
  const prog = recSal_leerTabla(ss.getSheetByName('Programacion')).map(r => ({
    emp_id: String(r.emp_id), fecha: recSal_fechaISO(r.fecha, tz),
    entrada: recSal_horaTxt(r.entrada, tz), salida: recSal_horaTxt(r.salida, tz), estado: String(r.estado || '')
  }));
  return {
    ss, tz, cfg, feriados, prog,
    personal: recSal_leerTabla(ss.getSheetByName('Personal')).map(r => Object.assign(r, {
      id: String(r.id), email: String(r.email || '').trim(), turno: String(r.turno || ''),
      modalidad: String(r.modalidad || '')
    })),
    turnos: recSal_json(cfg.turnos, []),
    sab: recSal_json(cfg.sabadosTrabajados, []),
    dom: recSal_json(cfg.domingosTrabajados, [])
  };
}

/* Filas como objetos con el encabezado. ultimas: leer solo las N últimas filas. */
function recSal_leerTabla(hoja, ultimas){
  if(!hoja) return [];
  const nf = hoja.getLastRow(), nc = hoja.getLastColumn();
  if(nf < 2 || nc < 1) return [];
  const cab = hoja.getRange(1, 1, 1, nc).getValues()[0].map(String);
  const desde = ultimas ? Math.max(2, nf - ultimas + 1) : 2;
  return hoja.getRange(desde, 1, nf - desde + 1, nc).getValues().map(fila => {
    const o = {};
    cab.forEach((c, i) => { if(c) o[c] = fila[i]; });
    return o;
  });
}

function recSal_fechaISO(v, tz){
  if(v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}
function recSal_horaTxt(v, tz){
  if(v instanceof Date) return Utilities.formatDate(v, tz, 'HH:mm');
  const p = String(v || '').split(':');
  return p.length >= 2 ? ('0' + Number(p[0])).slice(-2) + ':' + p[1].slice(0, 2) : '';
}
function recSal_aFecha(f, hhmm, tz){ return Utilities.parseDate(f + ' ' + hhmm, tz, 'yyyy-MM-dd HH:mm'); }
function recSal_aMin(hhmm){ const p = String(hhmm).split(':'); return Number(p[0])*60 + Number(p[1]); }
function recSal_siNo(v){ return v === true || v === '1' || v === 1 || v === 'true' || v === 'SI'; }
function recSal_activo(p){ return p.activo === true || String(p.activo).toUpperCase() === 'SI'; }
function recSal_json(v, d){ try{ return v ? JSON.parse(v) : d; }catch(e){ return d; } }
function recSal_nombreCorto(n){ n = String(n || '').toLowerCase().split(' ')[0]; return n.charAt(0).toUpperCase() + n.slice(1); }
