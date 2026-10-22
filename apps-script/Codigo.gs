/**
 * Control de Asistencia — Grupo Robles
 * Servidor en Google Apps Script, con una hoja de cálculo como base de datos.
 *
 * PUESTA EN MARCHA (una sola vez):
 *   1. Crea una hoja de cálculo nueva en Google Drive
 *   2. Extensiones → Apps Script
 *   3. Borra lo que haya y pega TODO este archivo
 *   4. Guarda (💾)
 *   5. Implementar → Nueva implementación → Aplicación web
 *        Ejecutar como: Yo
 *        Quién tiene acceso: Cualquier usuario
 *   6. Copia la URL que termina en /exec y pégala en el panel de la app
 *
 * Ver docs/INSTALAR-SERVIDOR.md para el paso a paso con más detalle.
 */

var TZ = 'America/Lima';

var HOJA = {
  mar: 'Marcajes',
  emp: 'Personal',
  cfg: 'Config',
  aus: 'Ausencias',
  fer: 'Feriados',
  dev: 'Dispositivos',
  ale: 'Alertas',
  per: 'Permisos',
  pro: 'Programacion',
  not: 'NotasProg',
  jef: 'Jefes',
  usu: 'Usuarios'
};

var COLS = {
  mar: ['ts_servidor','id','cid','emp_id','dni','nombre','area','proyecto','fecha','hora','tipo',
        'origen','estado','tardanza','anticipada','lat','lng','dist','fuera_zona','gps',
        'motivo','registrado_por','dispositivo','equipo','sede'],
  emp: ['id','nombre','dni','cargo','area','proyecto','tel','email','turno','activo','modalidad','planilla',
        'sede','vac_dias','vac_desde','dias'],
  cfg: ['clave','valor'],
  aus: ['id','emp_id','tipo','desde','hasta','lugar','actividad','motivo','autorizado_por','creado'],
  fer: ['fecha','nombre'],
  /* Un celular por persona. Vive en su propia hoja a propósito: si estuviera en
     Personal, cada vez que el panel sube el padrón se borrarían los vínculos. */
  dev: ['emp_id','nombre','dispositivo','desde','liberado_por'],
  /* Cada señal de suplantación queda escrita aquí, se haya podido mandar el correo
     o no. El correo puede fallar por mil motivos —permisos, cuota, spam— y sin este
     registro un intento no dejaba ningún rastro: se rechazaba y desaparecía. */
  ale: ['ts_servidor','fecha','hora','tipo','emp_id','nombre','detalle','dispositivo','equipo','correo'],
  /* Solicitudes de permiso. La persona las pide desde la app y administración las
     resuelve en el panel. Un permiso aprobado cambia cómo se lee ese día. */
  /* REGLA: una columna nueva va SIEMPRE al final de la lista, nunca en medio.
     La hoja ya existe con su orden; si se inserta en medio, asegurarColumnas la
     agrega igual al final y a partir de ahí los valores se escriben corridos una
     posición. Ya pasó dos veces: con 'equipo' en Marcajes y con 'hora_fin' aquí. */
  per: ['ts_servidor','id','emp_id','nombre','tipo','fecha','hora','motivo',
        'estado','resuelto_por','resuelto_el','comentario','hora_fin',
        'adjunto','visto_jefe','visto_jefe_por','visto_jefe_el','resuelto_usuario'],
  /* Horario asignado a una persona para un día concreto. Manda sobre su turno:
     el equipo de ventas se reparte la jornada y cada quien cubre una franja
     distinta, así que el turno fijo no alcanza. */
  pro: ['ts_servidor','id','emp_id','nombre','fecha','entrada','salida','modalidad','nota',
        'estado','validado_por','validado_el'],
  /* Las notas al pie del horario. La cuadrícula dice quién cubre qué hora, pero no
     por qué: "Fátima falta el jueves porque trabaja todo el domingo" no es un
     tramo de nadie y sin eso el horario no se entiende. Van por semana y área. */
  not: ['ts_servidor','semana','area','texto','autor'],
  /* Jefes de área: dan el primer visto bueno a los permisos de su gente antes que
     Administración. Viven en su propia hoja, como los celulares, para que subir el
     padrón no los borre. El PIN se guarda solo como huella (hash), nunca tal cual. */
  jef: ['emp_id','nombre','areas','pin_hash','actualizado'],
  // Quienes entran al panel con nombre propio. La clave se guarda como huella, nunca tal cual.
  usu: ['usuario','nombre','rol','pass_hash','activo','actualizado','areas','emp_id']
};

/* ══════════════════════════════════════════════════════════════
   ENTRADA HTTP
   ══════════════════════════════════════════════════════════════ */

/**
 * PRUEBA DEL CORREO — se ejecuta A MANO desde el editor de Apps Script.
 *
 * Selecciona "probarCorreo" en el desplegable de arriba y dale a Ejecutar.
 *
 * Sirve para dos cosas:
 *   1. Dice si hay un correo configurado y a dónde va a enviar.
 *   2. Si Apps Script nunca pidió permiso para enviar correo, al ejecutarla lo pide.
 *      Ese permiso es aparte: si no se concedió, los avisos automáticos no salen
 *      y no queda ningún rastro, porque el envío está hecho para no interrumpir
 *      un marcaje aunque el correo falle.
 *
 * El resultado sale en el registro de ejecución, abajo.
 */
function probarCorreo() {
  var cfg = leerConfig();
  var crudo = String(cfg.correoAlertas || '').trim();
  if (!crudo) {
    throw new Error('No hay ningún correo configurado. Llena la fila "correoAlertas" ' +
                    'en la pestaña Config de la hoja y vuelve a ejecutar.');
  }

  var dest = crudo.split(/[,;\s]+/)
    .map(function (x) { return x.trim(); })
    .filter(function (x) { return x && x.indexOf('@') > 0 && x.indexOf('.') > 0; });

  if (!dest.length) {
    throw new Error('Hay algo escrito en correoAlertas ("' + crudo + '") pero no parece ' +
                    'un correo válido. Revisa que esté bien escrito.');
  }

  MailApp.sendEmail({
    to: dest.join(','),
    subject: '[Asistencia Grupo Robles] Prueba de aviso',
    body: 'Si estás leyendo esto, los avisos automáticos funcionan.\n\n' +
          'A partir de ahora te llegará un correo cuando:\n' +
          '  · alguien intente marcar desde un celular que no es el suyo\n' +
          '  · un mismo celular marque por varias personas el mismo día\n\n' +
          '—\nPrueba lanzada a mano desde el editor.'
  });

  var msg = 'Enviado a: ' + dest.join(', ') +
            '  |  Correos que aún puedes enviar hoy: ' + MailApp.getRemainingDailyQuota();
  Logger.log(msg);
  return msg;
}

function doGet(e)  { return manejar(e, (e && e.parameter) || {}); }

function doPost(e) {
  var p = (e && e.parameter) || {};
  if (e && e.postData && e.postData.contents) {
    try { p = mezclar(p, JSON.parse(e.postData.contents)); } catch (err) { /* queda el parameter */ }
  }
  return manejar(e, p);
}

function manejar(e, p) {
  var callback = (e && e.parameter && e.parameter.callback) || p.callback || '';
  var r;
  try {
    r = enrutar(p);
  } catch (err) {
    r = { ok: false, error: String((err && err.message) || err) };
  }
  return responder(r, callback);
}

/**
 * Los Apps Script no permiten cabeceras CORS propias. Para una petición simple
 * (GET, o POST con text/plain) el navegador la deja pasar. Si aun así la red de
 * la oficina la bloquea, el cliente reintenta por JSONP y esto responde envuelto
 * en la función que pida.
 */
function responder(obj, callback) {
  var txt = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + txt + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(txt)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Qué puede hacer quien llama, según la clave que traiga.
 * '' = anónimo (el celular que va a marcar), 'vista' = Gerencia, 'admin' = Administración.
 */
function rolDe(p) {
  var c = String(p.clave || '');
  if (!c) return '';
  // Con usuario, manda su propio rol (y solo su clave): las claves compartidas
  // no sirven con un nombre de usuario delante.
  var u = normUsuario(p.usuario);
  if (u) { var x = usuarioValido(u, c); return x ? (x.rol === 'admin' ? 'admin' : x.rol === 'area' ? 'area' : 'vista') : ''; }
  var cfg = leerConfig();
  // Con usuarios administradores, la clave general de administración ya no abre
  // nada: cada uno entra con lo suyo. Si se desactivan o borran todos, vuelve a
  // servir (es la puerta de emergencia: editar la hoja Usuarios).
  if (c === String(cfg.adminPass || '')) return hayAdminUsuario() ? '' : 'admin';
  if (cfg.vistaPass && c === String(cfg.vistaPass)) return 'vista';
  return '';
}

var NO_AUTORIZADO = { ok: false, motivo: 'no_autorizado',
                      error: 'Necesitas la clave de administrador para esto.' };

/**
 * El celular de cada persona solo necesita tres cosas: saber la hora, preguntar
 * si un DNI existe, y registrar el marcaje. Nunca se lleva el padrón completo.
 *
 * Así la dirección del servidor puede ir dentro de la app —que es lo que hace que
 * funcione instalada, sin instalar, desde el QR o desde donde sea— sin que eso
 * exponga los datos de nadie. Los reportes y el padrón exigen clave.
 */
function enrutar(p) {
  var accion = p.accion || 'ping';
  var rol;

  switch (accion) {
    // ---- Abiertas: las usa el celular que marca ----
    case 'calentar':    return calentar();
    case 'ping':        return { ok: true, servidor: 'asistencia-grupo-robles', version: 23,
                                 hora: ahoraISO().hora, fecha: ahoraISO().fecha };
    case 'arranque':    return arranque(rolDe(p), p);
    case 'identificar': return identificar(p);
    case 'marcar':      return marcar(p);
    case 'mios':        return mios(p);
    case 'permiso':     return pedirPermiso(p);
    // Jefes de área: se identifican con su DNI y su PIN en cada pedido
    case 'jefepermisos': return jefePermisos(p);
    case 'jeferesolver': return jefeResolver(p);
    case 'jefeprog':     return jefeProgramacion(p);
    case 'jefeprogramar': return jefeProgramar(p);

    // ---- Con clave ----
    case 'entrar':      if (normUsuario(p.usuario)) return entrarUsuario(p);
                        rol = rolDe(p);
                        if (!rol && hayAdminUsuario() && String(p.clave || '') === String(leerConfig().adminPass || ''))
                          return { ok: false, error: 'La clave general de administración ya no se usa: entra con tu usuario.' };
                        return rol ? { ok: true, rol: rol }
                                   : { ok: false, error: 'Clave incorrecta' };
    case 'reporte':     rol = rolDe(p);
                        return rol ? reporte(p, rol === 'area' ? ambitoDe(p) : null) : NO_AUTORIZADO;
    // El resumen de la semana para n8n: basta la clave de solo vista
    // Con tipo=dia, el resumen de hoy (el del mediodía, para WhatsApp)
    case 'resumen':     rol = rolDe(p);
                        if (!rol || rol === 'area') return NO_AUTORIZADO;
                        return p.tipo === 'dia' ? resumenDia(p) : resumenSemana(p);

    // ---- Solo administración ----
    case 'instalar':    return rolDe(p) === 'admin' ? instalar(p)            : NO_AUTORIZADO;
    case 'ausencia':    return rolDe(p) === 'admin' ? guardarAusencia(p)     : NO_AUTORIZADO;
    case 'borrar':      rol = rolDe(p);
                        if (rol === 'admin') return borrarMarcajes(p, null);
                        // El jefe de área corrige marcas de su gente, una por una (nunca por rango)
                        if (rol === 'area' && (p.ids || p.id)) return borrarMarcajes(mezclar(p, { desde: '', hasta: '' }), ambitoDe(p));
                        return NO_AUTORIZADO;
    case 'liberar':     return rolDe(p) === 'admin' ? liberarDispositivo(p)  : NO_AUTORIZADO;
    case 'resolver':    return puedeResolver(p) ? resolverPermiso(p)          : NO_AUTORIZADO;
    case 'programar':   return puedeProgramar(p) ? guardarProgramacion(p)    : NO_AUTORIZADO;
    case 'notasprog':   return rolDe(p) === 'admin' ? guardarNotasProg(p)     : NO_AUTORIZADO;
    case 'reparar':     return rolDe(p) === 'admin' ? repararTodo()           : NO_AUTORIZADO;
    case 'borrarperm':  return rolDe(p) === 'admin' ? borrarPermiso(p)        : NO_AUTORIZADO;
    case 'personal':    return rolDe(p) === 'admin' ? guardarPersonal(p)     : NO_AUTORIZADO;
    case 'config':      return rolDe(p) === 'admin' ? guardarConfig(p)       : NO_AUTORIZADO;
    case 'guardarjefe': return rolDe(p) === 'admin' ? guardarJefe(p)         : NO_AUTORIZADO;
    case 'usuarios':    return rolDe(p) === 'admin' ? listarUsuarios()        : NO_AUTORIZADO;
    case 'guardarusuario': return rolDe(p) === 'admin' ? guardarUsuario(p)    : NO_AUTORIZADO;

    default:            return { ok: false, error: 'Acción desconocida: ' + accion };
  }
}

/** Identifica a UNA persona. Devuelve solo lo justo para saludarla y calcular su turno. */
function identificar(p) {
  var emps = leerPersonal();
  var emp = buscarEmp(emps, p.ident || p.dni || '');
  if (!emp)        return { ok: false, motivo: 'no_encontrado',
                            error: 'No encontramos ese dato en el registro de personal.' };
  if (!emp.activo) return { ok: false, motivo: 'inactivo',
                            error: 'Ese colaborador ya no está activo.' };
  return {
    ok: true,
    emp: { id: emp.id, nombre: emp.nombre, cargo: emp.cargo, area: emp.area,
           turno: emp.turno, modalidad: emp.modalidad, activo: true,
           dni: emp.dni, proyecto: emp.proyecto, sede: emp.sede,
           vacDias: emp.vacDias, vacDesde: emp.vacDesde, esJefe: !!jefeDe(emp.id), dias: emp.dias }
  };
}

/**
 * Los marcajes de UNA persona, sin clave de administración.
 *
 * Existe por un problema concreto: el celular guarda una copia de sus propios
 * marcajes, y cuando administración borraba uno desde el panel ese celular no
 * tenía forma de enterarse. Seguía creyendo que la jornada estaba completa y no
 * dejaba volver a marcar, sin manera de arreglarlo desde el propio aparato.
 *
 * Pide el documento igual que 'identificar', y solo devuelve los marcajes de esa
 * persona. Nadie puede listar los de otro probando ids al azar.
 */
function mios(p) {
  var emps = leerPersonal();
  var emp = buscarEmp(emps, p.ident || p.dni || '');
  if (!emp) return { ok: false, motivo: 'no_encontrado',
                     error: 'No encontramos ese dato en el registro de personal.' };
  if (p.empId && String(p.empId) !== String(emp.id)) {
    return { ok: false, motivo: 'no_coincide',
             error: 'El documento no corresponde a ese colaborador.' };
  }

  var desde = p.desde || '0000-01-01';
  var hasta = p.hasta || '9999-12-31';
  var out = filasDesde(HOJA.mar, COLS.mar, desde)
    .filter(function (m) {
      if (String(m.emp_id) !== String(emp.id)) return false;
      var f = fechaTexto(m.fecha);
      return f >= desde && f <= hasta;
    })
    .map(filaAObjeto);

  // Además de sus marcajes, su programación y sus permisos.
  //
  // Hace falta porque la app del celular decide si pedir ubicación o no, y para eso
  // necesita saber si a esa persona la programaron desde casa ese día. Antes esos
  // datos solo se enviaban a quien entraba con clave de administración, así que el
  // celular nunca se enteraba: el servidor habría aceptado el marcaje, pero la app
  // deshabilitaba el botón antes de intentarlo.
  //
  // Solo lo suyo, nunca lo de los demás: es el mismo criterio que con los marcajes.
  var miProg = leerProgramacion().filter(function (x) {
    return String(x.empId) === String(emp.id) && x.fecha >= desde && x.fecha <= hasta;
  });
  var misPermisos = leerPermisos().filter(function (x) {
    return String(x.empId) === String(emp.id);
  });
  // Y sus ausencias registradas: sin ellas, «Mis marcas» mostraría como falta
  // un día de vacaciones o de comisión.
  var misAusencias = leerAusencias().filter(function (x) {
    return String(x.empId) === String(emp.id) && x.hasta >= desde && x.desde <= hasta;
  });

  // Sus datos de ficha, por si administración le cambió la sede, el saldo de
  // vacaciones o lo nombró jefe después de que su celular lo guardara.
  return { ok: true, empId: emp.id, marcajes: out, total: out.length,
           programacion: miProg, permisos: misPermisos, ausencias: misAusencias,
           emp: { sede: emp.sede, vacDias: emp.vacDias, vacDesde: emp.vacDesde,
                  esJefe: !!jefeDe(emp.id), modalidad: emp.modalidad, turno: emp.turno,
                  dias: emp.dias, avisosJefe: avisosJefe(emp, emps) } };
}

/* ══════════════════════════════════════════════════════════════
   HOJAS
   ══════════════════════════════════════════════════════════════ */

function libro() { return SpreadsheetApp.getActiveSpreadsheet(); }

function hoja(nombre, cols) {
  var ss = libro();
  var h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.appendRow(cols);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    recordarCab(nombre, cols);
  }
  return h;
}

/* ── Escribir cada dato bajo SU título ─────────────────────────────────────

   Durante meses las filas se armaron en el orden de COLS, dando por sentado que
   la hoja tenía las columnas en ese mismo orden. Cuando dejó de ser cierto —una
   hoja creada antes de que existiera una columna— todo lo que venía después se
   escribió corrido una posición: el nombre cayó bajo 'fecha', la entrada bajo
   'salida', la modalidad bajo 'nota'.

   Pasó tres veces: con 'equipo' en Marcajes, con 'hora_fin' en Permisos y con
   'id' en Programación, donde arruinó el horario de una semana entera. La regla
   de "columna nueva siempre al final" ayuda, pero depende de que alguien la
   recuerde. Esto no: cada valor se busca por su título, así que el orden real de
   la hoja deja de importar. */

var CAB_ = {};   // los títulos reales de cada hoja, dentro de esta ejecución

function recordarCab(nombre, cab) {
  CAB_[nombre] = cab.map(function (x) { return String(x == null ? '' : x).trim(); });
  return CAB_[nombre];
}

/** Los títulos de la hoja, tal como están hoy. */
function cabDe(nombre, cols) {
  if (CAB_[nombre]) return CAB_[nombre];
  var h = hoja(nombre, cols);
  var ancho = Math.max(h.getLastColumn(), 1);
  return recordarCab(nombre, h.getRange(1, 1, 1, ancho).getValues()[0]);
}

/** Los valores de obj ordenados como están los títulos. Lo que no tenga título
    en la hoja se queda fuera; la columna sin dato queda vacía, no corrida. */
function filaSegunCab(cab, obj) {
  var out = [];
  for (var j = 0; j < cab.length; j++) {
    var n = String(cab[j] || '').trim();
    out.push(n && obj[n] !== undefined ? obj[n] : '');
  }
  return out;
}

/** Agrega una fila poniendo cada dato en su columna. */
function anexar(h, nombre, cols, obj) {
  h.appendRow(filaSegunCab(cabDe(nombre, cols), obj));
}

/**
 * Agrega al final las columnas nuevas que falten, sin tocar los datos que ya hay.
 * Permite ampliar una hoja en uso sin rehacerla ni perder marcajes.
 */
function asegurarColumnas(nombre, cols) {
  var h = hoja(nombre, cols);
  var ancho = Math.max(h.getLastColumn(), cols.length);
  var cab = h.getRange(1, 1, 1, ancho).getValues()[0].map(function (x) {
    return String(x == null ? '' : x).trim();
  });

  // Las filas se escriben SIEMPRE en el orden de cols. Así que un título que falta
  // hay que ponerlo en su sitio, no al final: si se agrega al final queda el dato en
  // una columna y el título en la siguiente, y al releerlo aparece vacío.
  // Pasó exactamente eso con 'equipo': el modelo del celular se guardaba bien pero
  // se leía en blanco, porque el título había quedado una columna corrida.
  var alineadas = true;
  for (var i = 0; i < cols.length; i++) {
    if (cab[i] && cab[i] !== cols[i]) { alineadas = false; break; }
  }

  if (alineadas) {
    var falta = false;
    for (var j = 0; j < cols.length; j++) if (cab[j] !== cols[j]) falta = true;
    if (falta) {
      h.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
      for (var q = 0; q < cols.length; q++) cab[q] = cols[q];
    }
    // Un título suelto más a la derecha, de un intento anterior, ya no corresponde
    for (var k = cols.length; k < ancho; k++) {
      if (cab[k] && cols.indexOf(cab[k]) >= 0) { h.getRange(1, k + 1).clearContent(); cab[k] = ''; }
    }
    recordarCab(nombre, cab);
    return h;
  }

  // Los títulos están en otro orden del esperado. No se toca ninguno: reescribir la
  // fila 1 aquí renombraría columnas que tienen datos de otra cosa. Solo se agrega
  // al final lo que falte, y queda a la vista para revisarlo a mano.
  var faltan = cols.filter(function (c) { return cab.indexOf(c) < 0; });
  if (faltan.length) {
    h.getRange(1, ancho + 1, 1, faltan.length).setValues([faltan]).setFontWeight('bold');
    cab = cab.concat(faltan);
  }
  recordarCab(nombre, cab);
  return h;
}

/**
 * Aviso inmediato por correo cuando algo huele mal. Solo para señales fuertes:
 * si llegan diez correos al día, nadie los va a mirar.
 */
function avisarPorCorreo(cfg, asunto, cuerpo) {
  // Admite varios destinatarios separados por coma, punto y coma o espacios.
  // Se filtran los que no parezcan correo para que uno mal escrito no impida
  // que el aviso llegue a los demás.
  var dest = String(cfg.correoAlertas || '')
    .split(/[,;\s]+/)
    .map(function (x) { return x.trim(); })
    .filter(function (x) { return x && x.indexOf('@') > 0 && x.indexOf('.') > 0; })
    .join(',');
  if (!dest) return 'sin destinatario';
  try {
    MailApp.sendEmail({
      to: dest,
      subject: '[Asistencia Grupo Robles] ' + asunto,
      body: cuerpo + '\n\n—\nAviso automático del control de asistencia.\n' +
            'Para dejar de recibirlos, vacía el campo de correo en Configuración → Servidor.'
    });
    return 'enviado';
  } catch (e) {
    // Un fallo de correo NUNCA debe impedir registrar un marcaje. Pero sí se devuelve
    // el motivo, para que quede escrito en la hoja de Alertas y se vea en el panel.
    return 'falló: ' + (e && e.message ? e.message : e);
  }
}

/* Deja constancia de una señal de suplantación, pase lo que pase con el correo. */
function registrarAlerta(tipo, emp, detalle, p, correo) {
  try {
    var h = asegurarColumnas(HOJA.ale, COLS.ale);
    var ahora = ahoraISO();
    anexar(h, HOJA.ale, COLS.ale, {
      ts_servidor: new Date(), fecha: ahora.fecha, hora: ahora.hora, tipo: tipo,
      emp_id: (emp && emp.id) || '', nombre: (emp && emp.nombre) || '', detalle: detalle,
      dispositivo: String((p && p.deviceId) || ''), equipo: String((p && p.equipo) || ''),
      correo: String(correo || '')
    });
  } catch (e) { /* tampoco esto puede tumbar un marcaje */ }
}

/** Devuelve las filas como objetos, usando la primera fila como encabezado. */
/**
 * Las últimas n filas, sin traerse la hoja entera.
 *
 * Al marcar hay que mirar solo lo de hoy —si ya marcó, y si ese celular marcó por
 * otro—, pero se leía la hoja completa cada vez, y eso ocurre dentro del candado
 * que serializa a todo el mundo. Con el equipo entrando a la misma hora, la cola
 * se hacía de medio minuto. Y empeoraba solo: la hoja crece unas 23 filas al día.
 */
/**
 * Las filas desde una fecha, sin traerse la hoja entera.
 *
 * La hoja de marcajes crece unas veinticinco filas por día. Leerla completa para
 * devolverle a una persona sus dos marcajes de hoy costaba cinco segundos —medidos,
 * el 17 de septiembre con 434 filas— y cada día un poco más. Eso es lo que la gente
 * sentía como "la página está lenta": el celular esperaba al servidor.
 *
 * Primero se lee solo la columna de fechas, que es una lectura estrecha y barata, y
 * con eso se sabe en qué fila empieza el periodo. Después se lee únicamente ese
 * trozo. Dos lecturas en vez de una, pero la pesada se vuelve pequeña.
 *
 * El orden de la hoja es el de llegada, no el del calendario: un marcaje hecho sin
 * conexión se escribe después con su fecha original. Se busca la PRIMERA fila cuya
 * fecha alcanza el periodo y se lee de ahí al final, así que una fila desordenada
 * puede hacer que se lea un poco más de lo necesario, nunca menos.
 */
function filasDesde(nombre, cols, desde) {
  if (!desde || desde <= '0000-01-01') return filas(nombre, cols);
  var h = asegurarColumnas(nombre, cols);
  var ultima = h.getLastRow();
  if (ultima < 2) return [];

  var cab = cabDe(nombre, cols);
  var cFecha = cab.indexOf('fecha');
  if (cFecha < 0) return filas(nombre, cols);

  var fechas = h.getRange(2, cFecha + 1, ultima - 1, 1).getValues();
  var ini = 0;
  for (var i = 0; i < fechas.length; i++) {
    if (fechaTexto(fechas[i][0]) >= desde) { ini = i + 2; break; }
  }
  if (!ini) return [];                       // ninguna fila alcanza el periodo

  var ancho = Math.max(h.getLastColumn(), cab.length);
  var datos = h.getRange(ini, 1, ultima - ini + 1, ancho).getValues();
  var out = [];
  for (var j = 0; j < datos.length; j++) {
    var o = {};
    for (var k = 0; k < cab.length; k++) o[cab[k]] = datos[j][k];
    out.push(o);
  }
  return out;
}

function filasUltimas(nombre, cols, n) {
  var h = hoja(nombre, cols);
  var total = h.getLastRow();
  if (total < 2) return [];
  var desde = Math.max(2, total - n + 1);
  var ancho = Math.max(h.getLastColumn(), cols.length);
  var cab = h.getRange(1, 1, 1, ancho).getValues()[0];
  var datos = h.getRange(desde, 1, total - desde + 1, ancho).getValues();
  var out = [];
  for (var i = 0; i < datos.length; i++) {
    var o = {};
    for (var j = 0; j < cab.length; j++) o[cab[j]] = datos[i][j];
    out.push(o);
  }
  return out;
}

function filas(nombre, cols) {
  var h = hoja(nombre, cols);
  var datos = h.getDataRange().getValues();
  if (datos.length < 2) return [];
  var cab = datos[0];
  var out = [];
  for (var i = 1; i < datos.length; i++) {
    var o = {};
    for (var j = 0; j < cab.length; j++) o[cab[j]] = datos[i][j];
    out.push(o);
  }
  return out;
}

function limpiarHoja(nombre, cols) {
  var h = hoja(nombre, cols);
  if (h.getLastRow() > 1) h.getRange(2, 1, h.getLastRow() - 1, h.getLastColumn()).clearContent();
  return h;
}

function escribir(nombre, cols, registros, mapear) {
  var h = limpiarHoja(nombre, cols);
  if (!registros || !registros.length) return 0;
  var cab = cabDe(nombre, cols);
  var matriz = registros.map(function (r) { return filaSegunCab(cab, mapear(r)); });
  h.getRange(2, 1, matriz.length, cab.length).setValues(matriz);
  return matriz.length;
}

/* ══════════════════════════════════════════════════════════════
   INSTALACIÓN — el panel empuja el padrón y la configuración
   ══════════════════════════════════════════════════════════════ */

function instalar(p) {
  olvidar();               // lo que se escriba aquí tiene que verse en el próximo pedido
  var datos = p.datos ? (typeof p.datos === 'string' ? JSON.parse(p.datos) : p.datos) : {};
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    asegurarColumnas(HOJA.mar, COLS.mar);   // agrega columnas nuevas sin borrar nada
    asegurarColumnas(HOJA.emp, COLS.emp);

    var nEmp = escribir(HOJA.emp, COLS.emp, datos.emp || [], function (e) {
      return { id: e.id, nombre: e.nombre, dni: "'" + (e.dni || ''), cargo: e.cargo || '',
               area: e.area || '', proyecto: e.proyecto || '', tel: "'" + (e.tel || ''),
               email: e.email || '', turno: e.turno || '', activo: e.activo ? 'SI' : 'NO',
               modalidad: e.modalidad || 'PRESENCIAL', planilla: e.planilla === false ? 'NO' : 'SI',
               sede: e.sede || '',
               vac_dias: (e.vacDias === null || e.vacDias === undefined || e.vacDias === '') ? '' : Number(e.vacDias),
               vac_desde: e.vacDesde ? "'" + fechaTexto(e.vacDesde) : '',
               dias: diasValidos(e.dias) ? JSON.stringify(diasValidos(e.dias)) : '' };
    });

    var cfg = datos.cfg || {};
    var pares = [];
    ['empresa','sigla','lat','lng','radio','gpsModo','tolerancia','umbralGrave',
     'inicioOperacion','cierreMargenMin','refrigerioMin','refrigerioDesdeMin','adminPass','vistaPass','exigirDispositivo','correoAlertas',
     'whatsappSoporte','resumenPara'].forEach(function (k) {
      if (cfg[k] !== undefined) pares.push([k, String(cfg[k])]);
    });
    pares.push(['turnos', JSON.stringify(cfg.turnos || [])]);
    pares.push(['proyectos', JSON.stringify(cfg.proyectos || [])]);
    pares.push(['sabadosTrabajados', JSON.stringify(cfg.sabadosTrabajados || [])]);
    pares.push(['domingosTrabajados', JSON.stringify(cfg.domingosTrabajados || [])]);
    // Sedes adicionales (fundo, almacén…). La principal sigue siendo lat/lng/radio.
    pares.push(['nombreSedePrincipal', String(cfg.nombreSedePrincipal || '')]);
    pares.push(['sedesExtra', JSON.stringify(listaSedesExtra(cfg.sedesExtra))]);
    var hc = limpiarHoja(HOJA.cfg, COLS.cfg);
    if (pares.length) hc.getRange(2, 1, pares.length, 2).setValues(pares);

    var nFer = escribir(HOJA.fer, COLS.fer, cfg.feriados || [], function (f) {
      return { fecha: f.fecha, nombre: f.nombre };
    });

    var nAus = escribir(HOJA.aus, COLS.aus, datos.aus || [], function (a) {
      return { id: a.id, emp_id: a.empId, tipo: a.tipo, desde: a.desde, hasta: a.hasta,
               lugar: a.lugar || '', actividad: a.actividad || '', motivo: a.motivo || '',
               autorizado_por: a.autorizadoPor || '', creado: a.creado || '' };
    });

    return { ok: true, personal: nEmp, feriados: nFer, ausencias: nAus, config: pares.length };
  } finally {
    lock.releaseLock();
  }
}

/* ══════════════════════════════════════════════════════════════
   LECTURA DE CONFIGURACIÓN Y PADRÓN
   ══════════════════════════════════════════════════════════════ */

/* ── Memoria de corta duración ─────────────────────────────────────────────

   Un pedido del celular abre cuatro o cinco hojas distintas: personal para saber de
   quién es el DNI, marcajes, programación y permisos. Cada apertura le cuesta a
   Apps Script entre medio segundo y uno, y ESO es lo que se siente al abrir la app.

   Medido el 18 de setiembre: un `ping` que no lee ninguna hoja tarda 1,8 s —ese es
   el piso, no se puede bajar—; el mismo pedido leyendo las hojas, 4,4 s. Acortar la
   lectura de marcajes ayudó a que no empeore con los años, pero el gasto de hoy no
   está en el tamaño de una lectura: está en la cantidad de lecturas.

   Estas tres hojas son chicas y cambian poco —el padrón, el horario, los permisos—,
   así que se guardan unos minutos. En la hora punta de la mañana, cuando veinticinco
   personas abren la app casi a la vez, la primera paga la lectura y las demás no.

   Eran un minuto; desde el 2 de octubre son cinco, porque en la llegada de la mañana
   la memoria vencía a cada rato. Igual se borra en cuanto algo se escribe.
   Importa: si la jefatura publica un horario y alguien marca enseguida, tiene que
   ver el horario nuevo o el sistema le exigirá ubicación estando en su casa. Con el
   borrado al escribir, ese caso no existe; los cinco minutos solo aplican a una
   edición hecha a mano directamente en la hoja. */
var MEMO_SEG = 900;   // quince minutos: lo escrito desde la app la borra al instante, y
                      // «calentar» (cada 10 min, desde GitHub) la renueva antes de que venza
var MEMO_ = {};                  // dentro de una misma ejecución, para no repetir
var MEMO_ESCRIBE_ = false;       // esta ejecución escribe: no deja nada en la memoria compartida
var MEMO_CLAVES = ['memo_cfg', 'memo_emp', 'memo_pro', 'memo_per', 'memo_jef', 'memo_usu'];

function memo(clave, calcular) {
  if (MEMO_.hasOwnProperty(clave)) return MEMO_[clave];

  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) { cache = null; }
  if (cache) {
    var guardado = null;
    try { guardado = cache.get(clave); } catch (e) { guardado = null; }
    if (guardado) {
      try { return (MEMO_[clave] = JSON.parse(guardado)); } catch (e) { /* mal guardado: se recalcula */ }
    }
  }

  var valor = calcular();
  MEMO_[clave] = valor;
  // Si no entra (100 KB por clave) se sigue sin memoria: esto es una mejora de
  // velocidad, no puede ser el motivo de que algo falle.
  // Quien escribe lee ANTES de escribir: si dejara eso en la memoria compartida,
  // los demás verían durante cinco minutos la hoja sin lo que acaba de escribir.
  if (cache && !MEMO_ESCRIBE_) { try { cache.put(clave, JSON.stringify(valor), MEMO_SEG); } catch (e) { } }
  return valor;
}

/** Se llama al escribir cualquier cosa: la próxima lectura vuelve a la hoja. */
function olvidar() {
  MEMO_ = {};
  MEMO_ESCRIBE_ = true;
  try {
    var c = CacheService.getScriptCache();
    c.removeAll(MEMO_CLAVES);
    c.put('memo_ver', String(new Date().getTime()) + Math.random(), 21600);   // ver calentar()
  } catch (e) { }
}

/* ══════════════════════════════════════════════════════════════
   CALENTAR
   Cuando la memoria compartida vence, el siguiente pedido tiene que leer seis
   hojas: así se medían arranques de 10 s, y con el panel pidiendo varias cosas a
   la vez se pasaba de los 15 s que la app espera antes de decir «sin conexión».
   El workflow «Vigilar servidor» de GitHub llama a esto cada 10 minutos en
   horario de oficina: relee las hojas y deja la memoria lista, así ningún
   celular ni el panel encuentran al servidor «dormido».
   No devuelve datos. Si algo se escribe mientras relee, no guarda lo leído (ya
   estaría viejo): para eso olvidar() deja una marca nueva en memo_ver.
   ══════════════════════════════════════════════════════════════ */
function calentar() {
  var cache;
  try { cache = CacheService.getScriptCache(); } catch (e) { return { ok: false, error: 'sin memoria' }; }
  if (cache.get('calentado')) return { ok: true, saltado: true };          // otro lo hizo hace menos de 1 min
  cache.put('calentado', '1', 60);
  var ver = cache.get('memo_ver');
  var fuentes = { memo_cfg: leerConfigDeHoja, memo_emp: leerPersonalDeHoja, memo_pro: leerProgramacionDeHoja,
                  memo_per: leerPermisosDeHoja, memo_jef: leerJefesDeHoja, memo_usu: leerUsuariosDeHoja };
  var t0 = new Date().getTime(), hechos = 0;
  for (var k in fuentes) {
    try {
      var v = fuentes[k]();
      if (cache.get('memo_ver') !== ver) return { ok: true, interrumpido: true, calentados: hechos };
      cache.put(k, JSON.stringify(v), MEMO_SEG);
      hechos++;
    } catch (e) { /* una que no entre en memoria no impide las demás */ }
  }
  return { ok: true, calentados: hechos, ms: new Date().getTime() - t0 };
}

function leerConfig()       { return memo('memo_cfg', leerConfigDeHoja); }
function leerPersonal()     { return memo('memo_emp', leerPersonalDeHoja); }
function leerProgramacion() { return memo('memo_pro', leerProgramacionDeHoja); }
function leerPermisos()     { return memo('memo_per', leerPermisosDeHoja); }

function leerConfigDeHoja() {
  var cfg = {};
  filas(HOJA.cfg, COLS.cfg).forEach(function (f) {
    var k = String(f.clave || '').trim();
    if (!k) return;
    var v = f.valor;
    if (k === 'turnos' || k === 'proyectos' || k === 'sabadosTrabajados' || k === 'domingosTrabajados' ||
        k === 'sedesExtra') { try { v = JSON.parse(v); } catch (e) { v = []; } }
    else if (['lat','lng','radio','tolerancia','umbralGrave','cierreMargenMin','refrigerioMin','refrigerioDesdeMin'].indexOf(k) >= 0) v = Number(v);
    // Si la hoja convirtió la fecha en fecha/hora, dejarla como texto AAAA-MM-DD.
    // Un "2026-08-24T05:00:00.000Z" rompe las comparaciones y esconde las faltas del día.
    else if (k === 'inicioOperacion') v = fechaTexto(v);
    cfg[k] = v;
  });
  if (!cfg.turnos || !cfg.turnos.length) {
    cfg.turnos = [{ id:'T1', nombre:'Administrativo', entrada:'09:00', salida:'18:00', trabajaSab:true, sabEntrada:'09:00', sabSalida:'13:00', trabajaDom:false, domEntrada:'10:00', domSalida:'14:00' }];
  }
  if (cfg.tolerancia == null || isNaN(cfg.tolerancia)) cfg.tolerancia = 10;
  if (cfg.umbralGrave == null || isNaN(cfg.umbralGrave)) cfg.umbralGrave = 30;
  cfg.feriados = filas(HOJA.fer, COLS.fer)
    .filter(function (f) { return f.fecha; })
    .map(function (f) { return { fecha: fechaTexto(f.fecha), nombre: String(f.nombre || '') }; });
  return cfg;
}

function leerPersonalDeHoja() {
  return filas(HOJA.emp, COLS.emp)
    .filter(function (f) { return f.id; })
    .map(function (f) {
      return {
        id: String(f.id), nombre: String(f.nombre || ''), dni: soloDigitos(f.dni),
        cargo: String(f.cargo || ''), area: String(f.area || ''), proyecto: String(f.proyecto || ''),
        tel: soloDigitos(f.tel), email: String(f.email || '').toLowerCase(),
        turno: String(f.turno || 'T1'),
        activo: String(f.activo).toUpperCase() !== 'NO',
        modalidad: String(f.modalidad || 'PRESENCIAL').toUpperCase(),
        exento: String(f.modalidad || '').toUpperCase() === 'EXENTO',
        planilla: String(f.planilla).toUpperCase() !== 'NO',
        sede: String(f.sede || ''),
        vacDias: (f.vac_dias === '' || f.vac_dias === null || f.vac_dias === undefined) ? null : Number(f.vac_dias),
        vacDesde: fechaTexto(f.vac_desde),
        dias: diasValidos(f.dias)
      };
    });
}

function leerAusencias() {
  return filas(HOJA.aus, COLS.aus)
    .filter(function (f) { return f.id; })
    .map(function (f) {
      return {
        id: String(f.id), empId: String(f.emp_id), tipo: String(f.tipo),
        desde: fechaTexto(f.desde), hasta: fechaTexto(f.hasta),
        lugar: String(f.lugar || ''), actividad: String(f.actividad || ''),
        motivo: String(f.motivo || ''), autorizadoPor: String(f.autorizado_por || '')
      };
    });
}

/**
 * Sin clave devuelve solo las reglas: turnos, tolerancias, feriados, ubicación.
 * Nada de eso identifica a nadie, y es lo que el celular necesita para funcionar.
 * Con clave, además, el padrón y las ausencias — eso es para el panel.
 */
function arranque(rol, p) {
  var cfg = leerConfig();

  // Las claves NUNCA salen del servidor
  var publica = {};
  for (var k in cfg) if (cfg.hasOwnProperty(k)) publica[k] = cfg[k];
  delete publica.adminPass;
  delete publica.vistaPass;
  delete publica.correoAlertas;
  publica.tieneVistaPass = !!cfg.vistaPass;
  publica.claveAdminActiva = !hayAdminUsuario();
  // El correo no sale nunca, pero hace falta poder comprobar desde el panel
  // que hay uno puesto: si se borra, los avisos dejan de llegar sin que nadie
  // se entere, porque el envío falla en silencio a propósito.
  publica.tieneCorreoAlertas = !!String(cfg.correoAlertas || '').trim();

  var r = {
    ok: true,
    servidor: true,
    fecha: ahoraISO().fecha,
    hora: ahoraISO().hora,
    cfg: publica
  };

  if (rol === 'admin' || rol === 'vista') {
    r.emp = leerPersonal();
    r.aus = leerAusencias();
    r.dispositivos = leerDispositivos();
    // leerDispositivos devuelve un mapa por persona, así que una fila repetida se
    // perdería sin que nadie la viera. Se manda también el total de filas: si no
    // cuadra con el número de personas, el panel lo avisa.
    try { r.dispositivosFilas = filas(HOJA.dev, COLS.dev).length; } catch (e) { r.dispositivosFilas = null; }
    r.alertas = leerAlertas(60);
    r.permisos = leerPermisos();
    r.programacion = leerProgramacion();
    r.notasProg = leerNotasProg();
    r.jefes = leerJefes().map(function (j) {             // sin la huella del PIN
      return { empId: j.empId, nombre: j.nombre, areas: j.areas, tienePin: !!j.pinHash };
    });
    r.rol = rol;
  }
  // El jefe de área recibe lo mismo, pero solo de la gente de sus áreas
  if (rol === 'area') {
    var amb = ambitoDe(p || {});
    if (!amb) return r;
    var mio = function (x) { return !!amb.ids[String(x.empId)]; };
    r.emp = leerPersonal().filter(function (e) { return !!amb.ids[e.id]; });
    r.aus = leerAusencias().filter(mio);
    var disp = leerDispositivos(), dsp = {};
    for (var id in disp) if (disp.hasOwnProperty(id) && amb.ids[id]) dsp[id] = disp[id];
    r.dispositivos = dsp;
    r.dispositivosFilas = null;
    r.alertas = leerAlertas(60).filter(mio);
    r.permisos = leerPermisos().filter(mio);
    r.programacion = leerProgramacion().filter(mio);
    r.notasProg = [];
    r.jefes = [];
    r.rol = 'area';
    r.ambito = { areas: amb.areas, empId: amb.empId };
  }
  return r;
}

/* ══════════════════════════════════════════════════════════════
   VINCULACIÓN POR DISPOSITIVO
   Cada persona queda atada al primer celular desde el que marca. Desde otro,
   el servidor rechaza. Evita que alguien marque por un compañero.
   No evita que le presten el celular desbloqueado: eso no lo resuelve el software.
   ══════════════════════════════════════════════════════════════ */

/* Las alertas recientes, para que el panel las muestre. Se leen al revés: lo
   último que pasó es lo que interesa ver primero. */
function leerAlertas(limite) {
  limite = limite || 60;
  var todas;
  try { todas = filas(HOJA.ale, COLS.ale); } catch (e) { return []; }
  var out = todas.slice(-limite).map(function (a) {
    return {
      fecha: fechaTexto(a.fecha), hora: horaTexto(a.hora),
      tipo: String(a.tipo || ''), empId: String(a.emp_id || ''),
      nombre: String(a.nombre || ''), detalle: String(a.detalle || ''),
      dispositivo: String(a.dispositivo || ''), equipo: String(a.equipo || ''),
      correo: String(a.correo || '')
    };
  });
  return out.reverse();
}

/* ══════════════════════════════════════════════════════════════
   PROGRAMACIÓN DE HORARIOS
   Un horario puesto a mano para una persona y un día. Manda sobre su turno.
   Nace de una necesidad concreta: los domingos la tienda abre todo el día y el
   equipo de ventas se reparte las franjas, unos temprano y otros tarde, algunos
   desde casa. Con un turno fijo por persona eso no se puede expresar.
   ══════════════════════════════════════════════════════════════ */

function leerProgramacionDeHoja() {
  var out;
  try { out = filas(HOJA.pro, COLS.pro); } catch (e) { return []; }
  return out.filter(function (f) { return f.emp_id && f.fecha; }).map(function (f) {
    var ent = horaTexto(f.entrada);
    return {
      // Las filas cargadas antes de que existieran los tramos no traen id. Se les
      // arma uno estable a partir de persona, día y hora de entrada, para poder
      // seguir editándolas y borrándolas como a las nuevas.
      id: String(f.id || ('G' + f.emp_id + fechaTexto(f.fecha) + ent)),
      empId: String(f.emp_id), nombre: String(f.nombre || ''),
      fecha: fechaTexto(f.fecha), entrada: ent, salida: horaTexto(f.salida),
      modalidad: String(f.modalidad || 'PRESENCIAL').toUpperCase(), nota: String(f.nota || ''),
      // Vacío cuenta como PUBLICADO: lo que ya estaba cargado antes de que
      // existieran los estados seguía en pie, y no debe desaparecer del cuadro.
      estado: String(f.estado || 'PUBLICADO').toUpperCase(),
      validadoPor: String(f.validado_por || ''), validadoEl: fechaTexto(f.validado_el)
    };
  });
}

/** Las notas al pie del horario, todas. Son pocas: una semana tiene tres o cuatro. */
function leerNotasProg() {
  var out;
  try { out = filas(HOJA.not, COLS.not); } catch (e) { return []; }
  return out.filter(function (f) { return f.semana && String(f.texto || '').trim(); })
            .map(function (f) {
    return { semana: fechaTexto(f.semana), area: String(f.area || ''),
             texto: String(f.texto || ''), autor: String(f.autor || '') };
  });
}

/**
 * Reemplaza las notas de una semana y un área. Llegan todas juntas desde el panel,
 * así que se borra lo que había y se escribe lo que vino: es la única forma de que
 * borrar una nota en el panel la borre de verdad aquí.
 */
function guardarNotasProg(p) {
  olvidar();               // lo que se escriba aquí tiene que verse en el próximo pedido
  var semana = fechaTexto(p.semana), area = String(p.area || '');
  if (!semana) return { ok: false, error: 'Falta la semana.' };
  var lista = p.datos ? (typeof p.datos === 'string' ? JSON.parse(p.datos) : p.datos) : [];

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var h = asegurarColumnas(HOJA.not, COLS.not);
    var cab = cabDe(HOJA.not, COLS.not);
    var datos = h.getDataRange().getValues();
    var cSem = cab.indexOf('semana'), cAre = cab.indexOf('area');

    var borrar = [];
    for (var i = 1; i < datos.length; i++) {
      if (fechaTexto(datos[i][cSem]) === semana && String(datos[i][cAre] || '') === area) {
        borrar.push(i + 1);
      }
    }
    borrar.sort(function (x, y) { return y - x; });
    for (var b = 0; b < borrar.length; b++) h.deleteRow(borrar[b]);

    var n = 0;
    for (var k = 0; k < lista.length; k++) {
      var texto = String(lista[k] == null ? '' : lista[k]).trim();
      if (!texto) continue;
      anexar(h, HOJA.not, COLS.not, {
        ts_servidor: new Date(), semana: "'" + semana, area: area,
        texto: texto, autor: String(p.autor || '')
      });
      n++;
    }
    return { ok: true, notas: n };
  } finally { lock.releaseLock(); }
}

/**
 * Endereza los permisos escritos mientras 'hora_fin' estuvo en medio de la lista.
 *
 * Mismo error que en Programación, distinta hoja. Durante esas horas la hora de fin
 * se escribió bajo 'motivo', el motivo bajo 'estado' y el estado bajo 'resuelto_por'.
 * Un permiso así pedido queda invisible para administración: nunca figura como
 * PENDIENTE, así que nadie lo aprueba ni lo rechaza y la persona se queda esperando
 * una respuesta que no va a llegar.
 *
 * Se reconoce porque en 'estado' hay algo que no es PENDIENTE, APROBADO ni RECHAZADO.
 * Ninguna fila sana puede tener otra cosa ahí.
 */
function repararPermisos() {
  var h = hoja(HOJA.per, COLS.per);
  var datos = h.getDataRange().getValues();
  if (datos.length < 2) return { reparados: 0, revisados: 0 };

  var cab = datos[0].map(function (x) { return String(x == null ? '' : x).trim(); });
  var cEstado = cab.indexOf('estado');
  if (cEstado < 0) return { reparados: 0, revisados: 0 };

  // El orden exacto en que se escribieron esas filas.
  var ORDEN_VIEJO = ['ts_servidor', 'id', 'emp_id', 'nombre', 'tipo', 'fecha', 'hora',
                     'hora_fin', 'motivo', 'estado', 'resuelto_por', 'resuelto_el', 'comentario'];
  var SANOS = ['PENDIENTE', 'APROBADO', 'RECHAZADO', ''];

  var reparados = 0, ejemplo = null;
  for (var i = 1; i < datos.length; i++) {
    var v = datos[i];
    if (!String(v[1] || '').trim()) continue;                 // fila sin id: no es un permiso
    var est = String(v[cEstado] || '').trim().toUpperCase();
    if (SANOS.indexOf(est) >= 0) continue;

    var obj = {};
    for (var k = 0; k < ORDEN_VIEJO.length && k < v.length; k++) obj[ORDEN_VIEJO[k]] = v[k];
    obj.fecha    = "'" + fechaTexto(obj.fecha);
    obj.hora     = "'" + horaTexto(obj.hora);
    obj.hora_fin = "'" + horaTexto(obj.hora_fin);
    if (!obj.estado) obj.estado = 'PENDIENTE';

    h.getRange(i + 1, 1, 1, cab.length).setValues([filaSegunCab(cab, obj)]);
    if (!ejemplo) ejemplo = String(obj.nombre) + ' · ' + String(obj.tipo) + ' · ' + fechaTexto(v[5]);
    reparados++;
  }
  return { reparados: reparados, revisados: datos.length - 1, ejemplo: ejemplo };
}

/**
 * Endereza las filas de Programación que quedaron corridas una columna.
 *
 * La hoja se creó antes de que los tramos tuvieran id. Al agregarse esa columna
 * quedó al final —es lo que hace asegurarColumnas cuando no reconoce el orden—
 * pero las filas se siguieron armando en el orden de COLS. Resultado: desde la
 * segunda columna en adelante cada valor cayó bajo el título de al lado. El
 * nombre quedó bajo 'fecha', la entrada bajo 'salida', la modalidad bajo 'nota'.
 * En la práctica el horario de la semana se leía como un disparate, y a quien
 * estaba programado desde casa el sistema le seguía exigiendo ubicación.
 *
 * Se reconoce una fila corrida porque bajo 'emp_id' hay un id de tramo —G y una
 * marca de tiempo— en vez de un código de persona. Los valores están en el orden
 * de COLS empezando en la primera columna, así que se releen de ahí y se
 * reescriben por título. Correrlo dos veces no hace daño: la segunda no encuentra
 * ninguna.
 */
function repararProgramacion() {
  var h = hoja(HOJA.pro, COLS.pro);
  var datos = h.getDataRange().getValues();
  if (datos.length < 2) return { ok: true, reparadas: 0, revisadas: 0 };

  var cab = datos[0].map(function (x) { return String(x == null ? '' : x).trim(); });
  var cEmp = cab.indexOf('emp_id');
  if (cEmp < 0) return { ok: false, error: 'La hoja Programacion no tiene columna emp_id.' };

  var esIdDeTramo = /^G\d{13,}$/;
  var reparadas = 0, ejemplo = null;

  for (var i = 1; i < datos.length; i++) {
    var v = datos[i];
    if (!esIdDeTramo.test(String(v[cEmp] || '').trim())) continue;

    var obj = {};
    for (var k = 0; k < COLS.pro.length && k < v.length; k++) obj[COLS.pro[k]] = v[k];
    // Fechas y horas vuelven como texto para que la hoja no las reinterprete.
    obj.fecha   = "'" + fechaTexto(obj.fecha);
    obj.entrada = "'" + horaTexto(obj.entrada);
    obj.salida  = "'" + horaTexto(obj.salida);

    h.getRange(i + 1, 1, 1, cab.length).setValues([filaSegunCab(cab, obj)]);
    if (!ejemplo) ejemplo = obj.nombre + ' ' + fechaTexto(v[4]) + ' ' + horaTexto(v[5]);
    reparadas++;
  }
  return { ok: true, reparadas: reparadas, revisadas: datos.length - 1, ejemplo: ejemplo };
}

/** Las dos reparaciones juntas: es una sola llamada desde el panel. */
function repararTodo() {
  olvidar();               // lo que se escriba aquí tiene que verse en el próximo pedido
  var pro = repararProgramacion();
  var per = repararPermisos();
  return { ok: true, programacion: pro, permisos: per };
}

/**
 * Los tramos de una persona ese día, en orden.
 *
 * Son varios a propósito: el equipo de ventas parte la jornada. Brunella entra a
 * las 6 desde su casa y a las 9 se viene a la oficina hasta las 3. Eso son dos
 * tramos del mismo día con modalidades distintas, y con un solo horario por día
 * no había forma de expresarlo ni de saber cuándo exigirle ubicación.
 */
function segmentosDe(prog, empId, fecha) {
  var out = [];
  for (var i = 0; i < prog.length; i++) {
    var a = prog[i];
    if (String(a.empId) === String(empId) && a.fecha === fecha && a.entrada && a.salida) out.push(a);
  }
  out.sort(function (x, y) { return x.entrada < y.entrada ? -1 : 1; });
  return out;
}

/** El tramo que cubre esa hora. Si no cae en ninguno, el más cercano. */
function segmentoEn(prog, empId, fecha, hora, soloPublicados) {
  var segs = segmentosDe(prog, empId, fecha).filter(function (s) {
    return !soloPublicados || s.estado !== 'BORRADOR';
  });
  if (!segs.length) return null;
  var m = aMin(hora);
  for (var i = 0; i < segs.length; i++) {
    if (m >= aMin(segs[i].entrada) && m <= aMin(segs[i].salida)) return segs[i];
  }
  // Fuera de todos: se toma el más cercano, para no exigir ubicación a quien marca
  // unos minutos antes de empezar su tramo desde casa.
  var mejor = segs[0], dist = Math.abs(m - aMin(segs[0].entrada));
  for (var j = 1; j < segs.length; j++) {
    var d = Math.min(Math.abs(m - aMin(segs[j].entrada)), Math.abs(m - aMin(segs[j].salida)));
    if (d < dist) { dist = d; mejor = segs[j]; }
  }
  return mejor;
}

/** La jornada completa del día: de la primera entrada a la última salida. */
function jornadaProgramada(prog, empId, fecha) {
  var segs = segmentosDe(prog, empId, fecha).filter(function (s) { return s.estado !== 'BORRADOR'; });
  if (!segs.length) return null;
  var ent = segs[0].entrada, sal = segs[0].salida;
  for (var i = 1; i < segs.length; i++) {
    if (segs[i].entrada < ent) ent = segs[i].entrada;
    if (segs[i].salida > sal) sal = segs[i].salida;
  }
  return { in: ent, out: sal, tramos: segs };
}

/**
 * Guarda o borra asignaciones. Llega una lista de {empId, fecha, entrada, salida,
 * modalidad, nota}; si entrada o salida vienen vacías, esa asignación se borra y
 * la persona vuelve a su turno normal ese día.
 * Solo se toca lo que llega: no se reescribe la hoja entera, para que dos personas
 * programando a la vez no se pisen el trabajo.
 */
function guardarProgramacion(p) {
  olvidar();               // lo que se escriba aquí tiene que verse en el próximo pedido
  var lista = p.datos ? (typeof p.datos === 'string' ? JSON.parse(p.datos) : p.datos) : [];
  if (!lista.length) return { ok: true, guardados: 0, borrados: 0 };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var h = asegurarColumnas(HOJA.pro, COLS.pro);
    var datos = h.getDataRange().getValues();
    var cab = datos[0];
    var cId = cab.indexOf('id'), cEmp = cab.indexOf('emp_id');
    var cFecha = cab.indexOf('fecha'), cEnt = cab.indexOf('entrada');

    // Cada tramo se localiza por su id. Las filas viejas no lo tienen, así que
    // también se indexan por persona+día+hora de entrada, que es como se les
    // fabrica el id al leerlas.
    var donde = {};
    for (var i = 1; i < datos.length; i++) {
      var idFila = String(datos[i][cId] || '');
      var alt = 'G' + datos[i][cEmp] + fechaTexto(datos[i][cFecha]) + horaTexto(datos[i][cEnt]);
      if (idFila) donde[idFila] = i + 1;
      donde[alt] = i + 1;
    }

    var guardados = 0, borrar = [];
    for (var k = 0; k < lista.length; k++) {
      var a = lista[k];
      var id = String(a.id || '');
      var fila = id ? donde[id] : null;
      var vacia = !a.entrada || !a.salida;

      if (vacia) { if (fila) borrar.push(fila); continue; }

      // Un tramo nuevo estrena id; uno que ya existía conserva el suyo.
      if (!id) id = 'G' + new Date().getTime() + Math.floor(Math.random() * 900 + 100);

      var valores = filaSegunCab(cab, {
        ts_servidor: new Date(), id: id, emp_id: String(a.empId), nombre: String(a.nombre || ''),
        fecha: "'" + fechaTexto(a.fecha), entrada: "'" + horaTexto(a.entrada),
        salida: "'" + horaTexto(a.salida),
        modalidad: String(a.modalidad || 'PRESENCIAL').toUpperCase(), nota: String(a.nota || ''),
        estado: String(a.estado || 'BORRADOR').toUpperCase(),
        validado_por: String(a.validadoPor || ''),
        validado_el: a.validadoEl ? "'" + fechaTexto(a.validadoEl) : ''
      });
      if (fila) h.getRange(fila, 1, 1, valores.length).setValues([valores]);
      else      h.appendRow(valores);
      guardados++;
    }

    // De abajo hacia arriba: borrar filas no corre las que faltan por borrar
    borrar.sort(function (x, y) { return y - x; });
    for (var b = 0; b < borrar.length; b++) h.deleteRow(borrar[b]);

    return { ok: true, guardados: guardados, borrados: borrar.length };
  } finally { lock.releaseLock(); }
}

/* ══════════════════════════════════════════════════════════════
   PERMISOS
   La persona pide desde la app; administración aprueba o rechaza en el panel.
   Un permiso aprobado no borra lo que pasó: la hora real se sigue registrando.
   Lo que cambia es cómo se cuenta —una tardanza justificada deja de sumar— y
   que quede escrito quién lo autorizó.
   ══════════════════════════════════════════════════════════════ */

var TIPOS_PERMISO = ['TARDANZA', 'SALIDA_ANTES', 'DIA', 'CASA', 'JORNADA', 'SALIDA_OLVIDO', 'ENTRADA_FALLIDA',
                     'ENTRADA_OLVIDO'];

function leerPermisosDeHoja() {
  var out;
  try { out = filas(HOJA.per, COLS.per); } catch (e) { return []; }
  return out.filter(function (f) { return f.id; }).map(function (f) {
    return {
      id: String(f.id), empId: String(f.emp_id), nombre: String(f.nombre || ''),
      tipo: String(f.tipo || ''), fecha: fechaTexto(f.fecha),
      hora: horaTexto(f.hora), horaFin: horaTexto(f.hora_fin),
      motivo: String(f.motivo || ''), estado: String(f.estado || 'PENDIENTE'),
      resueltoPor: String(f.resuelto_por || ''), resueltoEl: fechaTexto(f.resuelto_el),
      comentario: String(f.comentario || ''),
      adjunto: String(f.adjunto || ''),
      vistoJefe: String(f.visto_jefe || '').toUpperCase() === 'SI',
      vistoJefePor: String(f.visto_jefe_por || ''), vistoJefeEl: fechaTexto(f.visto_jefe_el),
      resueltoUsuario: String(f.resuelto_usuario || '')
    };
  });
}

/** Un permiso aprobado de cierto tipo para esa persona y esa fecha, o null. */
function permisoAprobado(empId, fecha, tipo) {
  var todos = leerPermisos();
  for (var i = 0; i < todos.length; i++) {
    var p = todos[i];
    if (p.estado === 'APROBADO' && String(p.empId) === String(empId) &&
        p.fecha === fecha && p.tipo === tipo) return p;
  }
  return null;
}

/** La pide el propio colaborador desde la app, sin clave. */
function pedirPermiso(p) {
  olvidar();               // lo que se escriba aquí tiene que verse en el próximo pedido
  var emps = leerPersonal();
  var emp = buscarEmp(emps, p.ident || p.dni || '');
  if (!emp)        return { ok: false, motivo: 'no_encontrado',
                            error: 'No encontramos ese dato en el registro de personal.' };
  if (!emp.activo) return { ok: false, motivo: 'inactivo', error: 'Ese colaborador ya no está activo.' };

  var tipo = String(p.tipo || '').toUpperCase();
  if (TIPOS_PERMISO.indexOf(tipo) < 0) return { ok: false, error: 'Tipo de permiso no válido.' };

  var fecha = fechaTexto(p.fecha || '');
  var hoy = ahoraISO().fecha;
  if (!fecha) return { ok: false, error: 'Indica la fecha del permiso.' };
  // Los permisos se piden antes o el mismo día. La excepción es JORNADA, que es
  // justamente para el día que ya pasó y no se marcó: sirve para regularizarlo
  // sin que administración tenga que cargarlo a mano. Se limita a una semana para
  // que no se conviertan en un cajón de sastre para justificar faltas viejas.
  // El descanso médico se avisa cuando ya empezó: se admite hasta 3 días atrás.
  // Llega como DIA con la marca "[Descanso médico]" al inicio del motivo.
  var esMedico = tipo === 'DIA' && String(p.motivo || '').indexOf('[Descanso médico]') === 0;
  if (esMedico && fecha < hoy) {
    var tope3 = new Date(new Date(hoy + 'T12:00:00').getTime() - 3 * 86400000);
    if (fecha < Utilities.formatDate(tope3, TZ, 'yyyy-MM-dd')) {
      return { ok: false, motivo: 'muy_viejo',
               error: 'El descanso médico se puede registrar hasta 3 días después. Habla con Administración.' };
    }
  }
  var haciaAtras = (tipo === 'JORNADA' || tipo === 'SALIDA_OLVIDO' || tipo === 'ENTRADA_FALLIDA' ||
                    tipo === 'ENTRADA_OLVIDO');
  if (!haciaAtras && !esMedico && fecha < hoy) {
    return { ok: false, motivo: 'fecha_pasada',
             error: 'No se puede pedir un permiso para un día que ya pasó. Habla con Administración.' };
  }
  if (tipo === 'SALIDA_OLVIDO') {
    // La salida que no se marcó: la persona solo la PIDE; se registra al aprobarse.
    if (fecha > hoy) return { ok: false, error: 'Solo se puede corregir un día que ya pasó.' };
    var tope2 = new Date(new Date(hoy + 'T12:00:00').getTime() - 7 * 86400000);
    if (fecha < Utilities.formatDate(tope2, TZ, 'yyyy-MM-dd')) {
      return { ok: false, motivo: 'muy_viejo',
               error: 'Solo se pueden corregir salidas de los últimos 7 días. Habla con Administración.' };
    }
    if (!p.hora) return { ok: false, error: 'Indica a qué hora saliste.' };
  }
  /* La entrada que no se marcó a tiempo: llegó a las 9:00 y marcó a las 11:00, o
     todavía no marca. Se PIDE con la hora real de llegada y se corrige al aprobarse. */
  if (tipo === 'ENTRADA_OLVIDO') {
    var e = validarEntradaOlvido(emp, fecha, horaTexto(p.hora || ''), hoy);
    if (e) return { ok: false, motivo: e.motivo || '', error: e.error };
  }
  /* "No pude marcar mi entrada": la persona está en el trabajo y la app no la deja
     marcar (GPS, celular cambiado…) y nadie le responde. Deja constancia de la
     hora de llegada —la del servidor, no una que escriba— y con eso ya puede
     marcar su salida. Al aprobarse se registra la entrada. */
  if (tipo === 'ENTRADA_FALLIDA') {
    if (fecha !== hoy) return { ok: false, error: 'Solo se puede registrar la llegada de hoy.' };
    p.hora = ahoraISO().hora;
    var yaPedida = leerPermisos().some(function (x) {
      return String(x.empId) === emp.id && x.tipo === 'ENTRADA_FALLIDA' && x.fecha === hoy && x.estado !== 'RECHAZADO';
    });
    if (yaPedida) return { ok: false, error: 'Ya registraste tu llegada de hoy. Ahora puedes marcar tu salida.' };
  }
  if (tipo === 'JORNADA') {
    if (fecha > hoy) return { ok: false, error: 'Solo se puede registrar una jornada ya trabajada.' };
    var tope = new Date(new Date(hoy + 'T12:00:00').getTime() - 7 * 86400000);
    if (fecha < Utilities.formatDate(tope, TZ, 'yyyy-MM-dd')) {
      return { ok: false, motivo: 'muy_viejo',
               error: 'Solo se pueden registrar jornadas de los últimos 7 días. Habla con Administración.' };
    }
    if (!p.hora || !p.horaFin) {
      return { ok: false, error: 'Indica desde qué hora hasta qué hora trabajaste.' };
    }
    if (aMin(horaTexto(p.horaFin)) <= aMin(horaTexto(p.hora))) {
      return { ok: false, error: 'La hora de salida debe ser posterior a la de entrada.' };
    }
  }

  var motivo = String(p.motivo || '').trim();
  if (motivo.length < 4) return { ok: false, error: 'Escribe el motivo del permiso.' };

  // El descanso médico exige la foto del certificado. Llega como imagen en el
  // primer día del pedido; los demás días traen el enlace que se devolvió.
  var adjuntoUrl = '';
  if (p.adjuntoUrl && /^https:\/\/(drive|docs)\.google\.com\//.test(String(p.adjuntoUrl))) {
    adjuntoUrl = String(p.adjuntoUrl);
  }
  if (esMedico && !adjuntoUrl && !p.adjunto) {
    return { ok: false, motivo: 'sin_certificado', error: 'Adjunta la foto del certificado médico.' };
  }

  // Los correos se mandan DESPUÉS de soltar el candado: cada uno tarda alrededor de
  // un segundo, y mientras el candado está tomado nadie puede marcar.
  var correos = [];
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Una sola solicitud viva por persona, día y tipo: si no, se acumulan
    // duplicados de quien toca el botón dos veces y administración no sabe cuál resolver.
    var yaHay = leerPermisos().filter(function (x) {
      return String(x.empId) === emp.id && x.fecha === fecha &&
             x.tipo === tipo && x.estado === 'PENDIENTE';
    });
    if (yaHay.length) {
      return { ok: true, duplicado: true, permiso: yaHay[0],
               aviso: 'Ya tienes una solicitud pendiente para ese día.' };
    }

    if (p.adjunto && !adjuntoUrl) {
      try { adjuntoUrl = guardarAdjunto(emp, fecha, String(p.adjunto)); }
      catch (e) { return { ok: false, error: 'No se pudo guardar la foto: ' + (e && e.message ? e.message : e) }; }
    }

    var h = asegurarColumnas(HOJA.per, COLS.per);
    var id = 'P' + new Date().getTime() + Math.floor(Math.random() * 900 + 100);
    anexar(h, HOJA.per, COLS.per, {
      ts_servidor: new Date(), id: id, emp_id: emp.id, nombre: emp.nombre, tipo: tipo,
      fecha: "'" + fecha, hora: "'" + horaTexto(p.hora || ''),
      hora_fin: "'" + horaTexto(p.horaFin || ''), motivo: motivo,
      estado: 'PENDIENTE', resuelto_por: '', resuelto_el: '', comentario: '',
      adjunto: adjuntoUrl, visto_jefe: '', visto_jefe_por: '', visto_jefe_el: ''
    });

    // Un pedido de varios días llega como una solicitud por día: se avisa solo con la primera.
    if (String(p.sinCorreo) !== 'true') correos.push(avisarPorCorreo.bind(null, leerConfig(), 'Nueva solicitud de permiso',
      emp.nombre + ' solicitó un permiso.\n\n' +
      'Tipo: ' + tipo + '\nFecha: ' + fecha + (p.hora ? '\nHora: ' + p.hora : '') + '\n' +
      'Motivo: ' + motivo + '\n\n' +
      'Apruébalo o recházalo en el panel, pestaña Permisos.'));

    // Y a su jefe de área, que es quien da el primer visto bueno.
    if (String(p.sinCorreo) !== 'true') correos.push(avisarJefes.bind(null, emp, tipo, fecha, motivo));

    // Olvidos de salida: desde el tercero del mes se avisa al jefe y a Administración
    // y lo mismo con los de entrada, contados aparte.
    if (tipo === 'SALIDA_OLVIDO' || tipo === 'ENTRADA_OLVIDO') {
      var mes = fecha.slice(0, 7);
      var olvidos = leerPermisos().filter(function (x) {
        return String(x.empId) === emp.id && x.tipo === tipo && x.fecha.slice(0, 7) === mes;
      }).length + 1;
      if (olvidos >= 3) correos.push(avisarOlvidos.bind(null, emp, olvidos, mes, tipo === 'ENTRADA_OLVIDO' ? 'entrada' : 'salida'));
    }

    return { ok: true, permiso: { id: id, empId: emp.id, nombre: emp.nombre, tipo: tipo,
                                  fecha: fecha, motivo: motivo, estado: 'PENDIENTE',
                                  hora: horaTexto(p.hora || ''), horaFin: horaTexto(p.horaFin || ''),
                                  adjunto: adjuntoUrl } };
  } finally {
    lock.releaseLock();
    correos.forEach(function (enviar) { try { enviar(); } catch (e) { } });
  }
}

/**
 * Escribe la entrada y la salida de una jornada declarada y aprobada.
 *
 * Quedan como MANUAL y con el motivo del permiso, para que en el reporte se
 * distinga de un marcaje hecho por la persona en el momento. No se pisa lo que
 * ya exista: si ese día ya tenía entrada, se respeta la que hay.
 */
function registrarJornada(empId, fecha, entrada, salida, motivo, quien, notaFija) {
  if (!empId || !fecha || (!entrada && !salida)) return 0;
  var emps = leerPersonal();
  var emp = null;
  for (var i = 0; i < emps.length; i++) if (emps[i].id === String(empId)) emp = emps[i];
  if (!emp) return 0;

  var cfg = leerConfig();
  var prog = leerProgramacion();
  var h = asegurarColumnas(HOJA.mar, COLS.mar);
  var yaHay = filasUltimas(HOJA.mar, COLS.mar, 800).filter(function (m) {
    return String(m.emp_id) === String(empId) && fechaTexto(m.fecha) === fecha;
  });
  var tiene = function (tipo) {
    for (var k = 0; k < yaHay.length; k++) if (String(yaHay[k].tipo) === tipo) return true;
    return false;
  };
  // Una jornada completa (entrada y salida) es otro tramo del día si ninguno de
  // los marcajes que ya hay cae dentro de ella: Dolly marca de 6 a 9 y declara
  // el tramo de 19 a 23:59 que no pudo marcar. Ese no es un duplicado.
  var tramoAparte = !!(entrada && salida) && yaHay.every(function (m) {
    var x = aMin(horaTexto(m.hora));
    return x < aMin(entrada) || x > aMin(salida);
  });

  var nota = (entrada && salida ? 'Jornada declarada y aprobada'
              : entrada ? 'Llegada registrada sin poder marcar, aprobada'
              : 'Salida olvidada, corrección aprobada') + (motivo ? ': ' + motivo : '');
  if (notaFija) nota = notaFija;
  var creados = 0;
  var poner = function (tipo, hora) {
    if (tiene(tipo) && !tramoAparte) return;
    var ev = tipo === 'ENTRADA' ? evaluarEntrada(emp, fecha, hora, cfg, prog)
                                : evaluarSalida(emp, fecha, hora, cfg, prog);
    anexar(h, HOJA.mar, COLS.mar, {
      ts_servidor: new Date(),
      id: 'M' + new Date().getTime() + Math.floor(Math.random() * 900 + 100),
      cid: '', emp_id: emp.id, dni: "'" + (emp.dni || ''), nombre: emp.nombre,
      area: emp.area, proyecto: emp.proyecto, fecha: "'" + fecha, hora: "'" + hora,
      tipo: tipo, origen: 'MANUAL', estado: ev.estado,
      tardanza: ev.tardanza || 0, anticipada: ev.anticipada || 0,
      lat: '', lng: '', dist: '', fuera_zona: 'NO', gps: '',
      motivo: nota, registrado_por: quien, dispositivo: '', equipo: ''
    });
    creados++;
  };
  if (entrada) poner('ENTRADA', entrada);
  if (salida) poner('SALIDA', salida);
  return creados;
}

/* Las entradas de esa persona en ese día, con su fila en la hoja (sin las de la
   noche que abren el turno de madrugada del día siguiente). */
function entradasDelDia(empId, fecha) {
  var h = hoja(HOJA.mar, COLS.mar);
  var total = h.getLastRow();
  if (total < 2) return { h: h, cab: cabDe(HOJA.mar, COLS.mar), filas: [] };
  var ancho = Math.max(h.getLastColumn(), COLS.mar.length);
  var cab = h.getRange(1, 1, 1, ancho).getValues()[0];
  var desde = Math.max(2, total - 2000 + 1);
  var datos = h.getRange(desde, 1, total - desde + 1, ancho).getValues();
  var c = { emp: cab.indexOf('emp_id'), fecha: cab.indexOf('fecha'), tipo: cab.indexOf('tipo'), hora: cab.indexOf('hora') };
  var out = [], salidas = [];
  for (var i = 0; i < datos.length; i++) {
    if (String(datos[i][c.emp]) !== String(empId) || fechaTexto(datos[i][c.fecha]) !== fecha) continue;
    var t = String(datos[i][c.tipo]), hr = horaTexto(datos[i][c.hora]);
    if (t === 'ENTRADA') out.push({ fila: desde + i, hora: hr, motivo: String(datos[i][cab.indexOf('motivo')] || '') });
    if (t === 'SALIDA') salidas.push(hr);
  }
  return { h: h, cab: cab, filas: out, salidas: salidas };
}

/* ¿Se puede pedir corregir la entrada de ese día a esa hora? null si sí. */
function validarEntradaOlvido(emp, fecha, hora, hoy) {
  if (fecha > hoy) return { error: 'Solo se puede corregir un día que ya pasó o el de hoy.' };
  var tope = new Date(new Date(hoy + 'T12:00:00').getTime() - 7 * 86400000);
  if (fecha < Utilities.formatDate(tope, TZ, 'yyyy-MM-dd')) {
    return { motivo: 'muy_viejo', error: 'Solo se pueden corregir entradas de los últimos 7 días. Habla con Administración.' };
  }
  if (!hora) return { error: 'Indica a qué hora llegaste.' };
  if (fecha === hoy && aMin(hora) > aMin(ahoraISO().hora)) {
    return { error: 'La hora de llegada no puede ser posterior a la hora actual.' };
  }
  var d = entradasDelDia(emp.id, fecha);
  if (d.filas.length && aMin(hora) >= aMin(d.filas[0].hora)) {
    return { error: 'La hora debe ser antes de tu entrada marcada (' + d.filas[0].hora + ').' };
  }
  if (!d.filas.length && fecha < hoy && !d.salidas.length) {
    return { motivo: 'usar_jornada', error: 'Ese día no tiene ningún marcaje. Pide «Trabajé y no marqué» con tu hora de entrada y de salida.' };
  }
  for (var k = 0; k < d.salidas.length; k++) {
    if (aMin(hora) >= aMin(d.salidas[k])) return { error: 'La hora debe ser antes de tu salida (' + d.salidas[k] + ').' };
  }
  return null;
}

/**
 * Corrige la entrada de un día con la hora de llegada aprobada.
 *
 * Si ese día ya hay entrada (la que se marcó tarde), se cambia su hora y se vuelve
 * a medir la tardanza; la hora que se marcó queda escrita en el motivo, para que
 * se sepa qué se corrigió. Si no hay, se escribe una entrada nueva. El origen
 * queda CORREGIDO, así el reporte la distingue de una marcada en el momento.
 */
function corregirEntrada(empId, fecha, hora, motivo, quien) {
  if (!empId || !fecha || !hora) return 0;
  var d = entradasDelDia(empId, fecha);
  var obs = motivo ? ': ' + motivo : '';
  if (!d.filas.length) {
    return registrarJornada(empId, fecha, hora, '', motivo, quien,
                            'Entrada olvidada, corrección aprobada' + obs);
  }
  var emp = null, emps = leerPersonal();
  for (var i = 0; i < emps.length; i++) if (emps[i].id === String(empId)) emp = emps[i];
  if (!emp) return 0;
  var ev = evaluarEntrada(emp, fecha, hora, leerConfig(), leerProgramacion());
  var f = d.filas[0], col = function (n) { return d.cab.indexOf(n) + 1; };
  var poner = function (n, v) { if (col(n) > 0) d.h.getRange(f.fila, col(n)).setValue(v); };
  poner('hora', "'" + hora);
  poner('estado', ev.estado);
  poner('tardanza', ev.tardanza || 0);
  poner('origen', 'CORREGIDO');
  poner('motivo', 'Entrada olvidada (marcó ' + f.hora + '), corrección aprobada' + obs);
  poner('registrado_por', quien);
  return 1;
}

/**
 * Borra una solicitud de permiso.
 *
 * Hasta ahora un permiso solo se podía aprobar o rechazar, y eso dejaba sin salida
 * a los que nacieron mal: una prueba, un duplicado de quien tocó dos veces, alguien
 * que se equivocó de día. Rechazarlos no es lo mismo —queda escrito que se le dijo
 * que no a una persona que en realidad nunca pidió eso— y quedaban ahí para siempre.
 *
 * Borrar el permiso NO borra los marcajes que se hayan creado al aprobarlo. Son dos
 * hechos distintos: uno es la autorización y el otro es la asistencia registrada.
 * Llevarse la asistencia por delante al limpiar una solicitud sería mucho peor que
 * el desorden que se está limpiando. Los marcajes se borran aparte, desde Corregir.
 */
function borrarPermiso(p) {
  olvidar();               // lo que se escriba aquí tiene que verse en el próximo pedido
  var ids = [];
  if (p.ids)     ids = (typeof p.ids === 'string') ? JSON.parse(p.ids) : p.ids;
  else if (p.id) ids = [String(p.id)];
  if (!ids.length) return { ok: false, error: 'Indica qué permiso borrar.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var h = hoja(HOJA.per, COLS.per);
    var datos = h.getDataRange().getValues();
    if (datos.length < 2) return { ok: true, borrados: 0 };
    var cab = datos[0];
    var cId = cab.indexOf('id'), cNom = cab.indexOf('nombre'), cFecha = cab.indexOf('fecha');

    var borrados = 0, detalle = [];
    // De abajo hacia arriba: borrar una fila corre las que faltan por revisar.
    for (var i = datos.length - 1; i >= 1; i--) {
      if (ids.indexOf(String(datos[i][cId])) < 0) continue;
      detalle.push(fechaTexto(datos[i][cFecha]) + ' ' + String(datos[i][cNom] || ''));
      h.deleteRow(i + 1);
      borrados++;
    }
    return { ok: true, borrados: borrados, detalle: detalle };
  } finally { lock.releaseLock(); }
}

/** Administración aprueba o rechaza. */
function resolverPermiso(p) {
  olvidar();               // lo que se escriba aquí tiene que verse en el próximo pedido
  var id = String(p.id || '');
  var decision = String(p.decision || '').toUpperCase();
  if (!id) return { ok: false, error: 'Indica qué permiso resolver.' };
  if (decision !== 'APROBADO' && decision !== 'RECHAZADO') {
    return { ok: false, error: 'La decisión debe ser APROBADO o RECHAZADO.' };
  }

  // Quien entra con su usuario firma con su nombre (no con lo que se escriba) y
  // queda su usuario, para el historial de lo que resolvió cada uno.
  var quien = usuarioDe(p);
  if (quien) p.resueltoPor = quien.nombre;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var h = asegurarColumnas(HOJA.per, COLS.per);
    var datos = h.getDataRange().getValues();
    var cab = datos[0];
    var cId = cab.indexOf('id'), cEstado = cab.indexOf('estado');
    var cUsu = cab.indexOf('resuelto_usuario');
    var cPor = cab.indexOf('resuelto_por'), cEl = cab.indexOf('resuelto_el');
    var cCom = cab.indexOf('comentario');

    var cTipo = cab.indexOf('tipo'), cEmp = cab.indexOf('emp_id');
    var cNom = cab.indexOf('nombre'), cFecha = cab.indexOf('fecha');
    var cHora = cab.indexOf('hora'), cFin = cab.indexOf('hora_fin');
    var cMot = cab.indexOf('motivo');

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][cId]) === id) {
        h.getRange(i + 1, cEstado + 1).setValue(decision);
        h.getRange(i + 1, cPor + 1).setValue(String(p.resueltoPor || 'Administración'));
        h.getRange(i + 1, cEl + 1).setValue("'" + ahoraISO().fecha);
        h.getRange(i + 1, cCom + 1).setValue(String(p.comentario || ''));
        if (cUsu >= 0) h.getRange(i + 1, cUsu + 1).setValue(quien ? quien.usuario : '');

        // Aprobar una jornada declarada no es solo cambiar un estado: hay que dejar
        // registrada la asistencia de ese día. Si no, el permiso diría "aprobado" y
        // el reporte seguiría contando falta.
        var creados = 0;
        if (decision === 'APROBADO' && String(datos[i][cTipo]) === 'JORNADA') {
          creados = registrarJornada(
            String(datos[i][cEmp]), fechaTexto(datos[i][cFecha]),
            horaTexto(datos[i][cHora]), horaTexto(datos[i][cFin]),
            String(datos[i][cMot] || ''), String(p.resueltoPor || 'Administración'));
        }
        // La llegada que no se pudo marcar: al aprobarla se escribe la entrada
        if (decision === 'APROBADO' && String(datos[i][cTipo]) === 'ENTRADA_FALLIDA') {
          creados = registrarJornada(
            String(datos[i][cEmp]), fechaTexto(datos[i][cFecha]),
            horaTexto(datos[i][cHora]), '',
            String(datos[i][cMot] || ''), String(p.resueltoPor || 'Administración'));
        }
        // La entrada olvidada: al aprobarla se corrige la hora de la entrada de ese día
        if (decision === 'APROBADO' && String(datos[i][cTipo]) === 'ENTRADA_OLVIDO') {
          creados = corregirEntrada(
            String(datos[i][cEmp]), fechaTexto(datos[i][cFecha]), horaTexto(datos[i][cHora]),
            String(datos[i][cMot] || ''), String(p.resueltoPor || 'Administración'));
        }
        // La salida olvidada: al aprobarla se escribe solo la salida de ese día
        if (decision === 'APROBADO' && String(datos[i][cTipo]) === 'SALIDA_OLVIDO') {
          creados = registrarJornada(
            String(datos[i][cEmp]), fechaTexto(datos[i][cFecha]),
            '', horaTexto(datos[i][cHora]),
            String(datos[i][cMot] || ''), String(p.resueltoPor || 'Administración'));
        }
        // La persona se entera por correo de la respuesta, con el comentario.
        // Un pedido de varios días se resuelve día por día: el panel pide el aviso solo una vez.
        var avisado = '';
        if (String(p.sinCorreo) !== 'true') {
          avisado = avisarAlColaborador(String(datos[i][cEmp]), decision, String(datos[i][cTipo]),
            fechaTexto(datos[i][cFecha]), String(datos[i][cMot] || ''), String(p.comentario || ''),
            String(p.resueltoPor || 'Administración'));
        }
        return { ok: true, id: id, estado: decision, marcajesCreados: creados, avisado: avisado };
      }
    }
    return { ok: false, error: 'No se encontró ese permiso.' };
  } finally { lock.releaseLock(); }
}

/** Correo al colaborador con la respuesta a su permiso. Nunca interrumpe nada si falla. */
function avisarAlColaborador(empId, decision, tipo, fecha, motivo, comentario, quien) {
  var emp = null, emps = leerPersonal();
  for (var i = 0; i < emps.length; i++) if (emps[i].id === String(empId)) emp = emps[i];
  if (!emp || !emp.email || emp.email.indexOf('@') < 0) return 'sin correo';
  var NOMBRES = { TARDANZA: 'Llegar tarde', SALIDA_ANTES: 'Salir antes', DIA: 'Día de permiso',
                  CASA: 'Trabajar desde casa', JORNADA: 'Trabajé y no marqué',
                  SALIDA_OLVIDO: 'Olvidé marcar mi salida', ENTRADA_FALLIDA: 'No pude marcar mi entrada',
                  ENTRADA_OLVIDO: 'Olvidé marcar mi entrada' };
  var que = NOMBRES[tipo] || tipo;
  if (motivo.indexOf('[Vacaciones]') === 0) que = 'Vacaciones';
  if (motivo.indexOf('[Descanso médico]') === 0) que = 'Descanso médico';
  try {
    MailApp.sendEmail({
      to: emp.email,
      subject: '[Asistencia] Tu solicitud fue ' + (decision === 'APROBADO' ? 'APROBADA' : 'RECHAZADA'),
      body: 'Hola ' + emp.nombre.split(' ')[0] + ',\n\n' +
            'Tu solicitud de «' + que + '» para el ' + fecha + ' fue ' +
            (decision === 'APROBADO' ? 'APROBADA' : 'RECHAZADA') + ' por ' + quien + '.\n' +
            (comentario ? '\nComentario: ' + comentario + '\n' : '') +
            '\nMotivo que indicaste: ' + motivo.replace(/^\[[^\]]+\]\s*/, '') + '\n\n' +
            '—\nAviso automático del control de asistencia.'
    });
    return 'enviado';
  } catch (e) { return 'falló: ' + (e && e.message ? e.message : e); }
}

/* ══════════════════════════════════════════════════════════════
   CERTIFICADOS MÉDICOS
   La foto llega del celular como imagen en base64 y se guarda en una carpeta de
   Drive del dueño del script. En la hoja queda solo el enlace.
   ══════════════════════════════════════════════════════════════ */
var CARPETA_CERT = 'Asistencia – Certificados médicos';

function guardarAdjunto(emp, fecha, dataUrl) {
  var m = String(dataUrl).match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
  if (!m) throw new Error('El archivo debe ser una foto (JPG o PNG).');
  var bytes = Utilities.base64Decode(m[2]);
  if (bytes.length > 6 * 1024 * 1024) throw new Error('La foto pesa demasiado.');
  var carpetas = DriveApp.getFoldersByName(CARPETA_CERT);
  var carpeta = carpetas.hasNext() ? carpetas.next() : DriveApp.createFolder(CARPETA_CERT);
  var ext = m[1] === 'image/png' ? 'png' : (m[1] === 'image/webp' ? 'webp' : 'jpg');
  var nombre = fecha + ' ' + emp.nombre + ' (' + emp.id + ').' + ext;
  var archivo = carpeta.createFile(Utilities.newBlob(bytes, m[1], nombre));
  return archivo.getUrl();
}

/** Ejecútala una vez desde el editor para conceder el permiso de Drive. */
function autorizarDrive() {
  var c = DriveApp.getFoldersByName(CARPETA_CERT);
  Logger.log(c.hasNext() ? 'La carpeta ya existe: ' + c.next().getUrl()
                         : 'Permiso concedido. La carpeta se crea con el primer certificado.');
}

/* ══════════════════════════════════════════════════════════════
   JEFES DE ÁREA
   ══════════════════════════════════════════════════════════════ */

function leerJefes() { return memo('memo_jef', leerJefesDeHoja); }

function leerJefesDeHoja() {
  var out;
  try { out = filas(HOJA.jef, COLS.jef); } catch (e) { return []; }
  return out.filter(function (f) { return f.emp_id; }).map(function (f) {
    return { empId: String(f.emp_id), nombre: String(f.nombre || ''),
             areas: String(f.areas || '').split('|').map(function (a) { return a.trim(); })
                                         .filter(function (a) { return a; }),
             pinHash: String(f.pin_hash || '') };
  });
}

function jefeDe(empId) {
  var js = leerJefes();
  for (var i = 0; i < js.length; i++) if (js[i].empId === String(empId)) return js[i];
  return null;
}

function huellaPin(empId, pin) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'asistencia|' + empId + '|' + pin);
  return d.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

/** Administración nombra, cambia o quita a un jefe. El PIN solo se cambia si llega uno. */
function guardarJefe(p) {
  olvidar();
  var empId = String(p.empId || '');
  var emp = null, emps = leerPersonal();
  for (var i = 0; i < emps.length; i++) if (emps[i].id === empId) emp = emps[i];
  if (!emp) return { ok: false, error: 'No se encontró a ese colaborador.' };
  var areas = p.areas ? (typeof p.areas === 'string' ? JSON.parse(p.areas) : p.areas) : [];
  var pin = String(p.pin || '');
  var borrar = String(p.borrar) === 'true' || p.borrar === true;
  if (!borrar && !areas.length) return { ok: false, error: 'Elige al menos un área.' };
  if (pin && !/^\d{4,6}$/.test(pin)) return { ok: false, error: 'El PIN debe tener de 4 a 6 números.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var h = asegurarColumnas(HOJA.jef, COLS.jef);
    var datos = h.getDataRange().getValues();
    var cab = cabDe(HOJA.jef, COLS.jef);
    var cEmp = cab.indexOf('emp_id'), cPin = cab.indexOf('pin_hash');
    var fila = 0, hashPrevio = '';
    for (var k = 1; k < datos.length; k++) {
      if (String(datos[k][cEmp]) === empId) { fila = k + 1; hashPrevio = String(datos[k][cPin] || ''); }
    }
    if (borrar) { if (fila) h.deleteRow(fila); return { ok: true, borrado: !!fila }; }
    var hash = pin ? huellaPin(empId, pin) : hashPrevio;
    if (!hash) return { ok: false, error: 'Ponle un PIN al jefe para que pueda entrar.' };
    var valores = filaSegunCab(cab, { emp_id: empId, nombre: emp.nombre, areas: areas.join('|'),
                                      pin_hash: hash, actualizado: new Date() });
    if (fila) h.getRange(fila, 1, 1, valores.length).setValues([valores]);
    else h.appendRow(valores);
    return { ok: true, jefe: { empId: empId, nombre: emp.nombre, areas: areas, tienePin: true } };
  } finally { lock.releaseLock(); }
}

/* ══════════════════════════════════════════════════════════════
   USUARIOS DEL PANEL
   Cada uno entra con su usuario y su clave, y tiene su rol: 'admin' (puede
   cambiar cosas) o 'vista' (solo mira). Las claves compartidas de siempre
   (adminPass, vistaPass) siguen sirviendo; n8n usa la de solo vista.
   ══════════════════════════════════════════════════════════════ */
function normUsuario(u) { return String(u || '').trim().toLowerCase(); }

/* admin: todo · area: jefe de área, solo su equipo (ve, programa, aprueba y
   corrige marcas) · permisos: ve todo y además aprueba o rechaza permisos ·
   vista: solo mira */
var ROLES_USUARIO = ['admin', 'area', 'permisos', 'vista'];

/* Lo que puede ver y tocar un jefe de área: la gente de sus áreas (activa o no),
   sin él mismo para lo que cambia algo. null = sin límite (los demás roles). */
function ambitoDe(p) {
  var x = usuarioDe(p);
  if (!x || x.rol !== 'area') return null;
  var ids = {};
  var areas = x.areas || [];             // una copia en memoria de antes de esta versión no las trae
  leerPersonal().forEach(function (e) { if (areas.indexOf(e.area) >= 0) ids[e.id] = true; });
  return { areas: areas, empId: x.empId || '', ids: ids, nombre: x.nombre };
}
/* ¿Puede cambiar algo de esta persona? Su gente, nunca él mismo. */
function enAmbito(amb, empId) {
  return !amb || (!!amb.ids[String(empId)] && String(empId) !== String(amb.empId));
}

/* El usuario con el que vino el pedido, si vino con uno válido. */
function usuarioDe(p) {
  var u = normUsuario(p.usuario);
  return u ? usuarioValido(u, String(p.clave || '')) : null;
}

/* Aprobar o rechazar permisos: administración, un usuario con rol permisos, o el
   jefe de área para la gente de su equipo (nunca los suyos). */
function puedeResolver(p) {
  if (rolDe(p) === 'admin') return true;
  var x = usuarioDe(p);
  if (x && x.rol === 'permisos') return true;
  if (x && x.rol === 'area') {
    var amb = ambitoDe(p), id = String(p.id || '');
    var per = leerPermisos().filter(function (q) { return q.id === id; })[0];
    return !!(per && enAmbito(amb, per.empId));
  }
  return false;
}

/* Programar: administración, o el jefe de área solo con tramos de su gente. */
function puedeProgramar(p) {
  if (rolDe(p) === 'admin') return true;
  var amb = ambitoDe(p);
  if (!amb) return false;
  var lista = p.datos ? (typeof p.datos === 'string' ? JSON.parse(p.datos) : p.datos) : [];
  var dueno = {};
  leerProgramacion().forEach(function (x) { dueno[x.id] = String(x.empId); });
  for (var i = 0; i < lista.length; i++) {
    if (!enAmbito(amb, lista[i].empId)) return false;
    if (lista[i].id && dueno[lista[i].id] && !enAmbito(amb, dueno[lista[i].id])) return false;
  }
  return true;
}

function huellaUsuario(usuario, clave) {
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'usuario|' + usuario + '|' + clave);
  return d.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

function leerUsuarios() { return memo('memo_usu', leerUsuariosDeHoja); }

function leerUsuariosDeHoja() {
  var out;
  try { out = filas(HOJA.usu, COLS.usu); } catch (e) { return []; }
  return out.filter(function (f) { return f.usuario; }).map(function (f) {
    return { usuario: normUsuario(f.usuario), nombre: String(f.nombre || ''),
             rol: ROLES_USUARIO.indexOf(String(f.rol || '').toLowerCase()) >= 0 ? String(f.rol).toLowerCase() : 'vista',
             passHash: String(f.pass_hash || ''), activo: String(f.activo).toUpperCase() !== 'NO',
             areas: String(f.areas || '').split('|').map(function (a) { return a.trim(); }).filter(function (a) { return a; }),
             empId: String(f.emp_id || '') };
  });
}

/* ¿Hay al menos un administrador con usuario propio que pueda entrar? */
function hayAdminUsuario() {
  return leerUsuarios().some(function (x) { return x.activo && x.rol === 'admin' && x.passHash; });
}

function usuarioValido(usuario, clave) {
  var us = leerUsuarios(), h = null;
  for (var i = 0; i < us.length; i++) {
    var x = us[i];
    if (x.usuario !== usuario || !x.activo || !x.passHash) continue;
    h = h || huellaUsuario(usuario, clave);
    if (x.passHash === h) return x;
  }
  return null;
}

/* Entrar con usuario: cinco intentos fallidos lo bloquean 15 minutos. */
function entrarUsuario(p) {
  var u = normUsuario(p.usuario);
  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) { cache = null; }
  var claveFallos = 'usufallos_' + u;
  var fallos = cache ? Number(cache.get(claveFallos) || 0) : 0;
  if (fallos >= 5) return { ok: false, error: 'Demasiados intentos. Espera 15 minutos e inténtalo de nuevo.' };
  var x = usuarioValido(u, String(p.clave || ''));
  if (!x) {
    if (cache) { try { cache.put(claveFallos, String(fallos + 1), 900); } catch (e) { } }
    return { ok: false, error: 'Usuario o clave incorrectos' };
  }
  if (cache) { try { cache.remove(claveFallos); } catch (e) { } }
  return { ok: true, rol: x.rol === 'admin' ? 'admin' : x.rol === 'area' ? 'area' : 'vista', usuario: x.usuario, nombre: x.nombre,
           aprueba: x.rol !== 'vista', rolUsuario: x.rol, areas: x.areas || [] };
}

function listarUsuarios() {
  return { ok: true, usuarios: leerUsuarios().map(function (x) {
    return { usuario: x.usuario, nombre: x.nombre, rol: x.rol, activo: x.activo, tieneClave: !!x.passHash,
             areas: x.areas, empId: x.empId };
  }) };
}

/** Crea, cambia o quita un usuario. La clave solo se cambia si llega una. */
function guardarUsuario(p) {
  olvidar();
  var u = normUsuario(p.nuevoUsuario || p.cual);
  if (!/^[a-z0-9._-]{3,30}$/.test(u))
    return { ok: false, error: 'El usuario debe tener de 3 a 30 letras o números, sin espacios ni tildes.' };
  var borrar = String(p.borrar) === 'true' || p.borrar === true;
  var nombre = String(p.nombre || '').trim().replace(/\s+/g, ' ');
  var rolNuevo = ROLES_USUARIO.indexOf(String(p.rol || '').toLowerCase()) >= 0 ? String(p.rol).toLowerCase() : 'vista';
  var clave = String(p.nuevaClave || '');
  var activo = !(String(p.activo) === 'false' || p.activo === false);
  var areas = p.areas ? (typeof p.areas === 'string' ? (p.areas.charAt(0) === '[' ? JSON.parse(p.areas) : p.areas.split('|')) : p.areas) : [];
  areas = areas.map(function (a) { return String(a).trim(); }).filter(function (a) { return a; });
  var empVinc = String(p.empId || '');
  if (!borrar && nombre.length < 3) return { ok: false, error: 'Escribe el nombre de quien usará este acceso.' };
  if (!borrar && rolNuevo === 'area' && !areas.length) return { ok: false, error: 'Elige al menos un área para el jefe de área.' };
  if (rolNuevo !== 'area') { areas = []; }
  if (clave && clave.length < 8) return { ok: false, error: 'La clave debe tener al menos 8 caracteres.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var h = asegurarColumnas(HOJA.usu, COLS.usu);
    var datos = h.getDataRange().getValues();
    var cab = cabDe(HOJA.usu, COLS.usu);
    var cU = cab.indexOf('usuario'), cP = cab.indexOf('pass_hash');
    var fila = 0, hashPrevio = '';
    for (var k = 1; k < datos.length; k++) {
      if (normUsuario(datos[k][cU]) === u) { fila = k + 1; hashPrevio = String(datos[k][cP] || ''); }
    }
    if (borrar) { if (fila) h.deleteRow(fila); return { ok: true, borrado: !!fila }; }
    if (p.nuevoUsuario && fila && String(p.crear) === 'true') return { ok: false, error: 'Ya existe un usuario «' + u + '».' };
    var hash = clave ? huellaUsuario(u, clave) : hashPrevio;
    if (!hash) return { ok: false, error: 'Ponle una clave al usuario para que pueda entrar.' };
    var valores = filaSegunCab(cab, { usuario: u, nombre: nombre, rol: rolNuevo, pass_hash: hash,
                                      activo: activo ? 'SI' : 'NO', actualizado: new Date(),
                                      areas: areas.join('|'), emp_id: empVinc });
    if (fila) h.getRange(fila, 1, 1, valores.length).setValues([valores]);
    else h.appendRow(valores);
    return { ok: true, usuario: { usuario: u, nombre: nombre, rol: rolNuevo, activo: activo, tieneClave: true,
                                  areas: areas, empId: empVinc } };
  } finally { lock.releaseLock(); }
}

/* Comprueba DNI + PIN del jefe. Cinco intentos fallidos lo bloquean 15 minutos.
   Si el sistema exige celular vinculado, tiene que ser el suyo. */
function autenticarJefe(p) {
  var emps = leerPersonal();
  var emp = buscarEmp(emps, p.ident || '');
  if (!emp || !emp.activo) return { error: 'No encontramos ese dato en el registro de personal.' };
  var jefe = jefeDe(emp.id);
  if (!jefe) return { error: 'No figuras como jefe de área.' };

  var cache = null;
  try { cache = CacheService.getScriptCache(); } catch (e) { cache = null; }
  var claveFallos = 'pinfallos_' + emp.id;
  var fallos = cache ? Number(cache.get(claveFallos) || 0) : 0;
  if (fallos >= 5) return { error: 'Demasiados intentos. Espera 15 minutos e inténtalo de nuevo.' };

  if (!jefe.pinHash || huellaPin(emp.id, String(p.pin || '')) !== jefe.pinHash) {
    if (cache) { try { cache.put(claveFallos, String(fallos + 1), 900); } catch (e) { } }
    return { error: 'PIN incorrecto.' };
  }
  if (cache) { try { cache.remove(claveFallos); } catch (e) { } }

  var cfg = leerConfig();
  if (String(cfg.exigirDispositivo) === 'true') {
    var suyo = leerDispositivos()[emp.id] || '';
    if (suyo && suyo !== String(p.deviceId || '')) {
      return { error: 'Entra desde tu propio celular, el que usas para marcar.' };
    }
  }
  return { emp: emp, jefe: jefe, emps: emps };
}

/* Los permisos de la gente de sus áreas. Los suyos propios no: esos van directo
   a Administración, nadie se aprueba a sí mismo. */
function jefePermisos(p) {
  var a = autenticarJefe(p);
  if (a.error) return { ok: false, error: a.error };
  var areaDe = {};
  a.emps.forEach(function (e) { areaDe[e.id] = e.area; });
  var hoy = ahoraISO().fecha;
  var hace30 = diaDespues(hoy, -30);
  var mios = leerPermisos().filter(function (x) {
    return x.empId !== a.emp.id && a.jefe.areas.indexOf(areaDe[x.empId]) >= 0;
  });
  return {
    ok: true, jefe: { nombre: a.emp.nombre, areas: a.jefe.areas },
    pendientes: mios.filter(function (x) { return x.estado === 'PENDIENTE' && !x.vistoJefe; }),
    recientes: mios.filter(function (x) {
      return (x.vistoJefe || x.estado !== 'PENDIENTE') && x.fecha >= hace30;
    }).slice(-60)
  };
}

/* La gente que un jefe puede programar: la activa de sus áreas, sin él mismo
   (igual que con los permisos, nadie se arma su propio horario) y sin exentos. */
function equipoDeJefe(a) {
  return a.emps.filter(function (e) {
    return e.activo && e.id !== a.emp.id && a.jefe.areas.indexOf(e.area) >= 0 &&
           String(e.modalidad || '') !== 'EXENTO';
  });
}

/* La programación de su equipo en un rango de fechas (una semana, un mes). */
function jefeProgramacion(p) {
  var a = autenticarJefe(p);
  if (a.error) return { ok: false, error: a.error };
  var hoy = ahoraISO().fecha;
  var desde = fechaTexto(p.desde || '') || diaDespues(hoy, -7);
  var hasta = fechaTexto(p.hasta || '') || diaDespues(hoy, 21);
  var equipo = equipoDeJefe(a);
  var ids = equipo.map(function (e) { return e.id; });
  return {
    ok: true, jefe: { nombre: a.emp.nombre, areas: a.jefe.areas },
    equipo: equipo.map(function (e) {
      return { id: e.id, nombre: e.nombre, area: e.area, turno: e.turno,
               modalidad: e.modalidad, dias: e.dias, activo: true };
    }),
    programacion: leerProgramacion().filter(function (x) {
      return ids.indexOf(String(x.empId)) >= 0 && x.fecha >= desde && x.fecha <= hasta;
    })
  };
}

/* El jefe guarda tramos de su gente. Solo de su gente: cada tramo que llega, y el
   que ya estaba en la hoja con ese id, tienen que ser de alguien de sus áreas.
   Lo que valida queda firmado con su nombre. Si publica, se avisa a Administración. */
function jefeProgramar(p) {
  var a = autenticarJefe(p);
  if (a.error) return { ok: false, error: a.error };
  var lista = p.datos ? (typeof p.datos === 'string' ? JSON.parse(p.datos) : p.datos) : [];
  if (!lista.length) return { ok: true, guardados: 0, borrados: 0 };

  var equipo = equipoDeJefe(a), nombres = {};
  equipo.forEach(function (e) { nombres[e.id] = e.nombre; });
  var dueno = {};
  leerProgramacion().forEach(function (x) { dueno[x.id] = String(x.empId); });
  var hoy = ahoraISO().fecha, publicados = 0;

  for (var i = 0; i < lista.length; i++) {
    var t = lista[i];
    var empId = String(t.empId || '');
    if (!nombres[empId]) return { ok: false, error: 'Solo puedes programar a la gente de tu equipo.' };
    if (t.id && dueno[t.id] && !nombres[dueno[t.id]]) {
      return { ok: false, error: 'Solo puedes programar a la gente de tu equipo.' };
    }
    t.nombre = nombres[empId];
    var est = String(t.estado || 'BORRADOR').toUpperCase();
    if (['BORRADOR', 'PUBLICADO', 'VALIDADO'].indexOf(est) < 0) est = 'BORRADOR';
    t.estado = est;
    if (est === 'VALIDADO') {
      t.validadoPor = a.emp.nombre + ' (jefe de área)';
      t.validadoEl = t.validadoEl || hoy;
    } else {
      t.validadoPor = ''; t.validadoEl = '';
    }
    if (est === 'PUBLICADO' && t.entrada && t.salida) publicados++;
  }

  var r = guardarProgramacion({ datos: lista });
  if (r.ok && publicados && String(p.sinCorreo) !== 'true') {
    avisarPorCorreo(leerConfig(), 'Programación publicada por ' + a.emp.nombre,
      a.emp.nombre + ' (jefe de ' + a.jefe.areas.join(', ') + ') publicó ' + publicados +
      ' turno(s) de su equipo.\n\nPuedes verlos en el panel, pestaña Programación.');
  }
  return r;
}

/* Lo que espera al jefe, para avisarle apenas abre la app (sin pedirle el PIN):
   solo cuántos, nunca de quién ni el detalle, que eso queda detrás del PIN.
   - permisos de su gente sin su visto bueno (agrupados como los ve él: un permiso
     de varios días cuenta una vez)
   - gente de su equipo sin programar: la que tuvo turnos publicados en las cuatro
     semanas anteriores y no tiene ninguno en la semana que toca. De lunes a
     miércoles mira la semana en curso; desde el jueves, la siguiente.
   - turnos en borrador de esa semana, que todavía no cuentan para nadie. */
function avisosJefe(emp, emps) {
  var jefe = jefeDe(emp.id);
  if (!jefe) return null;
  var equipo = equipoDeJefe({ emp: emp, jefe: jefe, emps: emps }), enEquipo = {};
  equipo.forEach(function (e) { enEquipo[e.id] = true; });
  var areaDe = {};
  emps.forEach(function (e) { areaDe[e.id] = e.area; });
  var grupos = {};
  leerPermisos().forEach(function (x) {
    if (x.estado !== 'PENDIENTE' || x.vistoJefe || x.empId === emp.id) return;
    if (jefe.areas.indexOf(areaDe[x.empId]) < 0) return;
    grupos[x.empId + '|' + x.tipo + '|' + (x.motivo || '') + '|' + x.estado + '|' + (x.hora || '')] = true;
  });

  var hoy = ahoraISO().fecha;
  var dow = (new Date(hoy + 'T12:00:00').getDay() + 6) % 7;            // 0 = lunes
  var lunes = diaDespues(hoy, -dow);
  var semana = dow >= 3 ? diaDespues(lunes, 7) : lunes;
  var finSemana = diaDespues(semana, 6), desde4 = diaDespues(semana, -28);
  var usa = {}, tiene = {}, borradores = 0;
  leerProgramacion().forEach(function (x) {
    if (!enEquipo[x.empId] || !x.entrada || !x.salida) return;
    if (x.fecha >= desde4 && x.fecha < semana && x.estado !== 'BORRADOR') usa[x.empId] = true;
    if (x.fecha >= semana && x.fecha <= finSemana) {
      tiene[x.empId] = true;
      if (x.estado === 'BORRADOR') borradores++;
    }
  });
  var sin = equipo.filter(function (e) { return usa[e.id] && !tiene[e.id]; }).length;
  return { permisos: Object.keys(grupos).length, sinProgramar: sin, borradores: borradores,
           semana: semana, proxima: semana > lunes };
}

/* El jefe da el visto bueno (queda para Administración) o rechaza (y ahí termina). */
function jefeResolver(p) {
  var a = autenticarJefe(p);
  if (a.error) return { ok: false, error: a.error };
  olvidar();
  var ids = p.ids ? (typeof p.ids === 'string' ? JSON.parse(p.ids) : p.ids) : [];
  var decision = String(p.decision || '').toUpperCase();
  if (decision !== 'VISTO' && decision !== 'RECHAZADO') return { ok: false, error: 'Decisión no válida.' };
  var comentario = String(p.comentario || '').trim();
  if (decision === 'RECHAZADO' && comentario.length < 4) return { ok: false, error: 'Escribe el motivo del rechazo.' };

  var areaDe = {};
  a.emps.forEach(function (e) { areaDe[e.id] = e.area; });
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var h = asegurarColumnas(HOJA.per, COLS.per);
    var datos = h.getDataRange().getValues();
    var cab = cabDe(HOJA.per, COLS.per);
    var c = {}; ['id','emp_id','estado','tipo','fecha','motivo','comentario','resuelto_por','resuelto_el',
                 'visto_jefe','visto_jefe_por','visto_jefe_el'].forEach(function (n) { c[n] = cab.indexOf(n); });
    var hechos = 0, avisados = {};
    for (var i = 1; i < datos.length; i++) {
      var fila = datos[i];
      if (ids.indexOf(String(fila[c.id])) < 0) continue;
      var empId = String(fila[c.emp_id]);
      if (empId === a.emp.id || a.jefe.areas.indexOf(areaDe[empId]) < 0) continue;   // no es de su gente
      if (String(fila[c.estado]) !== 'PENDIENTE') continue;
      if (decision === 'VISTO') {
        h.getRange(i + 1, c.visto_jefe + 1).setValue('SI');
        h.getRange(i + 1, c.visto_jefe_por + 1).setValue(a.emp.nombre);
        h.getRange(i + 1, c.visto_jefe_el + 1).setValue("'" + ahoraISO().fecha);
      } else {
        h.getRange(i + 1, c.estado + 1).setValue('RECHAZADO');
        h.getRange(i + 1, c.resuelto_por + 1).setValue(a.emp.nombre + ' (jefe de área)');
        h.getRange(i + 1, c.resuelto_el + 1).setValue("'" + ahoraISO().fecha);
        h.getRange(i + 1, c.comentario + 1).setValue(comentario);
        var clave = empId + '|' + String(fila[c.motivo]);
        if (!avisados[clave]) {
          avisados[clave] = true;
          avisarAlColaborador(empId, 'RECHAZADO', String(fila[c.tipo]), fechaTexto(fila[c.fecha]),
                              String(fila[c.motivo] || ''), comentario, a.emp.nombre + ' (jefe de área)');
        }
      }
      hechos++;
    }
    if (decision === 'VISTO' && hechos) {
      avisarPorCorreo(leerConfig(), 'Permiso con visto bueno del jefe',
        a.emp.nombre + ' dio su visto bueno a ' + hechos + ' día(s) de permiso.\n\n' +
        'Falta tu aprobación final en el panel, pestaña Permisos.');
    }
    return { ok: true, resueltos: hechos };
  } finally { lock.releaseLock(); }
}

/* Desde el tercer olvido del mes (de salida o de entrada): correo al jefe de área y a Administración. */
function avisarOlvidos(emp, n, mes, que) {
  que = que || 'salida';
  var cuerpo = emp.nombre + ' (' + emp.area + ') ya lleva ' + n + ' olvidos de marcar su ' + que + ' en ' + mes + '.\n\n' +
               'Cada olvido queda como solicitud de corrección para aprobar. Conviene conversarlo con esa persona.';
  avisarPorCorreo(leerConfig(), emp.nombre + ': ' + n + ' olvidos de ' + que + ' este mes', cuerpo);
  var emps = leerPersonal();
  leerJefes().forEach(function (j) {
    if (j.empId === emp.id || j.areas.indexOf(emp.area) < 0) return;
    for (var i = 0; i < emps.length; i++) {
      if (emps[i].id === j.empId && emps[i].email) {
        try { MailApp.sendEmail({ to: emps[i].email, subject: '[Asistencia] ' + emp.nombre + ': ' + n + ' olvidos de ' + que,
                                  body: cuerpo }); } catch (e) { }
      }
    }
  });
}

/* Correo al jefe o jefes del área de esa persona cuando pide un permiso. */
function avisarJefes(emp, tipo, fecha, motivo) {
  var emps = leerPersonal();
  leerJefes().forEach(function (j) {
    if (j.empId === emp.id || j.areas.indexOf(emp.area) < 0) return;
    var yo = null;
    for (var i = 0; i < emps.length; i++) if (emps[i].id === j.empId) yo = emps[i];
    if (!yo || !yo.email) return;
    try {
      MailApp.sendEmail({ to: yo.email, subject: '[Asistencia] ' + emp.nombre + ' pidió un permiso',
        body: emp.nombre + ' (' + emp.area + ') pidió un permiso.\n\nTipo: ' + tipo + '\nFecha: ' + fecha +
              '\nMotivo: ' + motivo + '\n\nRevísalo en la app de asistencia: "Permisos de mi equipo".' });
    } catch (e) { }
  });
}

function leerDispositivos() {
  var m = {};
  filas(HOJA.dev, COLS.dev).forEach(function (f) {
    if (f.emp_id) m[String(f.emp_id)] = String(f.dispositivo || '');
  });
  return m;
}

function registrarDispositivo(emp, deviceId) {
  var h = hoja(HOJA.dev, COLS.dev);
  var datos = h.getDataRange().getValues();
  var cab = datos[0], cEmp = cab.indexOf('emp_id');
  var valores = filaSegunCab(cab, { emp_id: emp.id, nombre: emp.nombre,
                                    dispositivo: deviceId, desde: new Date(), liberado_por: '' });
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][cEmp]) === emp.id) {
      h.getRange(i + 1, 1, 1, valores.length).setValues([valores]);
      return;
    }
  }
  h.appendRow(valores);
}

/** Libera el vínculo para que la persona pueda registrar otro celular. */
function liberarDispositivo(p) {
  olvidar();               // lo que se escriba aquí tiene que verse en el próximo pedido
  var ids = [];
  if (p.ids)         ids = (typeof p.ids === 'string') ? JSON.parse(p.ids) : p.ids;
  else if (p.id)     ids = [p.id];
  else if (p.empId)  ids = [p.empId];   // se admite el otro nombre: una app en caché
                                        // podría mandarlo así, y quedarse sin liberar
                                        // deja a esa persona sin poder marcar
  if (!ids.length) return { ok: false, error: 'Indica de quién liberar el celular.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var h = hoja(HOJA.dev, COLS.dev);
    var datos = h.getDataRange().getValues();
    var cab = datos[0], cEmp = cab.indexOf('emp_id');
    var n = 0;
    for (var i = datos.length - 1; i >= 1; i--) {
      if (ids.indexOf(String(datos[i][cEmp])) >= 0) { h.deleteRow(i + 1); n++; }
    }
    return { ok: true, liberados: n };
  } finally { lock.releaseLock(); }
}

/* ══════════════════════════════════════════════════════════════
   MARCAJE — el servidor decide la hora y el estado, nunca el celular
   ══════════════════════════════════════════════════════════════ */

function marcar(p) {
  // ── FUERA DEL CANDADO ──────────────────────────────────────────────────────
  // Todo esto solo lee. Tenerlo dentro obligaba a cada persona a esperar a que la
  // anterior terminara SEIS consultas a la hoja, y con el equipo entrando a la
  // misma hora la cola llegaba a medio minuto: la app cortaba y los marcajes se
  // iban a la cola del celular. Dentro del candado queda únicamente lo que de
  // verdad necesita exclusividad: comprobar duplicados y escribir.
  var cfg  = leerConfig();
  var emps = leerPersonal();
  var emp  = buscarEmp(emps, p.id || p.dni || p.ident || '');

  // Solo las últimas filas: los controles que las usan miran el día de hoy, y
  // traerse la hoja entera se vuelve más caro cada día que pasa.
  var _mar = null;
  var marTodos = function () {
    if (_mar === null) _mar = filasUltimas(HOJA.mar, COLS.mar, 400);
    return _mar;
  };
    if (!emp)        return { ok: false, error: 'No encontramos ese dato en el registro de personal.', motivo: 'no_encontrado' };
    if (!emp.activo) return { ok: false, error: 'Ese colaborador está inactivo.', motivo: 'inactivo' };

    // --- Vinculación por dispositivo ---
    // El registro manual se salta el candado del celular y además fija la hora que
    // le manden. Eso solo puede hacerlo administración: sin esta comprobación,
    // cualquiera que supiera un DNI podía mandar origen MANUAL y marcar por otra
    // persona, a la hora que quisiera, esquivando todo el control.
    var manualAdmin = (p.origen === 'MANUAL');
    // El jefe de área también registra a mano, pero solo a la gente de su equipo
    var rolManual = manualAdmin ? rolDe(p) : '';
    if (manualAdmin && rolManual === 'area' && !enAmbito(ambitoDe(p), emp.id)) {
      return { ok: false, motivo: 'no_autorizado', error: 'Solo puedes registrar marcas de la gente de tu equipo.' };
    }
    if (manualAdmin && rolManual !== 'admin' && rolManual !== 'area') {
      return { ok: false, motivo: 'no_autorizado',
               error: 'El registro manual solo lo puede hacer administración.' };
    }
    if (!manualAdmin && String(cfg.exigirDispositivo) === 'true') {
      var deviceId = String(p.deviceId || '');
      if (!deviceId) {
        return { ok: false, motivo: 'sin_dispositivo',
                 error: 'No se pudo identificar este celular. Prueba a recargar la página.' };
      }
      var devs = leerDispositivos();
      var suyo = devs[emp.id] || '';
      if (!suyo) {
        // Todavía no tiene celular propio, así que este marcaje se lo va a asignar.
        // Antes hay que mirar de quién es ese celular: si ya es el de otra persona,
        // esto es justo el caso de "marco por mi compañero desde mi teléfono", y sin
        // esta comprobación quedaba amarrado en silencio y sin aviso a nadie.
        var dueno = '';
        for (var eid in devs) {
          if (devs.hasOwnProperty(eid) && devs[eid] === deviceId && eid !== emp.id) dueno = eid;
        }
        if (dueno) {
          var otroNombre = '';
          for (var q = 0; q < emps.length; q++) if (emps[q].id === dueno) otroNombre = emps[q].nombre;
          otroNombre = otroNombre || dueno;

          if (!p.confirmar) {
            return { ok: false, motivo: 'celular_compartido', requiereConfirmar: true,
                     esDeOtro: true, otros: [otroNombre],
                     error: 'Este celular es el de ' + otroNombre + '.' };
          }
          var envio2 = avisarPorCorreo(cfg, 'Alguien marcó desde el celular de otra persona',
            emp.nombre + ' marcó desde el celular vinculado a ' + otroNombre + '.\n\n' +
            'Fecha: ' + ahoraISO().fecha + ' ' + ahoraISO().hora + '\n' +
            'Equipo usado: ' + (p.equipo || 'desconocido') + '\n\n' +
            'El marcaje SE REGISTRÓ, porque la persona confirmó en pantalla, y a partir\n' +
            'de ahora ese celular queda vinculado a ' + emp.nombre + '.\n' +
            'Si no corresponde, libera el vínculo en el panel, pestaña Celulares.');
          registrarAlerta('MARCÓ DESDE EL CELULAR DE OTRO', emp,
            'Marcó desde el celular vinculado a ' + otroNombre + '. Confirmó en pantalla y se registró.',
            p, envio2);
        }
        registrarDispositivo(emp, deviceId);          // primer marcaje: queda vinculado
      } else if (suyo !== deviceId) {
        var envio1 = avisarPorCorreo(cfg, 'Intento de marcar desde otro celular',
          emp.nombre + ' intentó marcar desde un celular que no es el suyo.\n\n' +
          'Fecha: ' + ahoraISO().fecha + ' ' + ahoraISO().hora + '\n' +
          'Equipo usado: ' + (p.equipo || 'desconocido') + '\n\n' +
          'El marcaje fue RECHAZADO. Si cambió de teléfono, libera su vínculo en\n' +
          'Configuración → Un celular por persona.');
        registrarAlerta('MARCAJE RECHAZADO · celular ajeno', emp,
          'Intentó marcar desde un celular que no es el suyo. El marcaje fue rechazado.',
          p, envio1);
        return { ok: false, motivo: 'otro_dispositivo',
                 error: 'Este celular no está vinculado a tu registro. Si cambiaste de teléfono, pídele a Administración que lo libere.',
                 nombre: emp.nombre };
      }
    }

    // --- Un mismo celular marcando por varias personas el mismo día ---
    // Es la señal más clara de que alguien está marcando por un compañero.
    // No se bloquea (a veces es legítimo: un celular prestado de verdad), pero
    // se pide confirmar y queda constancia, con aviso inmediato por correo.
    if (!manualAdmin && p.deviceId) {
      var hoyF = ahoraISO().fecha;
      var otros = {};
      marTodos().forEach(function (m) {
        if (String(m.dispositivo) === String(p.deviceId) &&
            fechaTexto(m.fecha) === hoyF &&
            String(m.emp_id) !== emp.id) {
          otros[String(m.emp_id)] = String(m.nombre);
        }
      });
      var nombres = [];
      for (var k in otros) if (otros.hasOwnProperty(k)) nombres.push(otros[k]);

      if (nombres.length) {
        if (!p.confirmar) {
          return { ok: false, motivo: 'celular_compartido',
                   requiereConfirmar: true,
                   otros: nombres,
                   error: 'Este celular ya marcó hoy por ' + nombres.join(', ') + '.' };
        }
        var envio3 = avisarPorCorreo(cfg, 'Un celular marcó por varias personas',
          'Se registró un marcaje de ' + emp.nombre + ' desde un celular que hoy ya había\n' +
          'marcado por: ' + nombres.join(', ') + '\n\n' +
          'Fecha: ' + hoyF + ' ' + ahoraISO().hora + '\n' +
          'Equipo: ' + (p.equipo || 'desconocido') + '\n\n' +
          'La persona confirmó que quería hacerlo. Puede ser legítimo —un celular\n' +
          'prestado— o puede ser que alguien esté marcando por un compañero.\n' +
          'Revísalo en el panel, pestaña Análisis.');
        registrarAlerta('UN CELULAR, VARIAS PERSONAS', emp,
          'Marcó desde un celular que hoy ya había marcado por: ' + nombres.join(', ') +
          '. Confirmó en pantalla y se registró.', p, envio3);
      }
    }

    var t = ahoraISO();
    var fecha = p.fecha || t.fecha;
    var hora  = p.hora  || t.hora;
    // El celular NO decide la hora de un marcaje normal: la pone el servidor.
    // Se respeta la que llega solo en dos casos, ambos de un día ya pasado:
    //   MANUAL    → lo carga administración con un motivo
    //   DECLARADO → la persona declara al día siguiente la salida que olvidó
    // 'diferido' es un marcaje que la persona hizo en su celular pero que no pudo
    // enviarse en el momento por falta de red. La hora buena es la de entonces, no
    // la de ahora: si se la pusiéramos ahora, alguien que marcó a las 9 y recuperó
    // señal a mediodía aparecería con tres horas de tardanza que no cometió.
    // Se acepta acotada —nunca al futuro, nunca más de quince días atrás— porque es
    // la única vía por la que el celular puede influir en la hora de un marcaje.
    // Fueron tres días hasta octubre de 2026: el servidor estuvo caído del 3 al 7,
    // la gente abrió la app recién el 9 y las entradas del lunes 5 quedaron
    // registradas con la fecha y la hora del viernes.
    var diferido = (String(p.diferido) === 'true' || p.diferido === true) && !!p.fecha && !!p.hora;
    var conFechaPropia = (p.origen === 'MANUAL' || p.origen === 'DECLARADO' || diferido);
    if (!conFechaPropia) { fecha = t.fecha; hora = t.hora; }
    if (conFechaPropia && fecha > t.fecha) { fecha = t.fecha; hora = t.hora; }   // nunca al futuro
    if (diferido) {
      var limite = new Date(new Date(t.fecha + 'T12:00:00').getTime() - 15 * 86400000);
      if (fecha < Utilities.formatDate(limite, TZ, 'yyyy-MM-dd')) { fecha = t.fecha; hora = t.hora; }
    }

    var tipo = String(p.tipo || '').toUpperCase() === 'SALIDA' ? 'SALIDA' : 'ENTRADA';
    // asegurarColumnas y no hoja(): si el modelo de datos ganó una columna nueva,
    // la hoja se pone al día sola antes de escribir. Sin esto la fila se guarda
    // con un dato en una columna sin cabecera, y al releerla ese dato no aparece.
    // Fue lo que pasó con 'equipo': se enviaba, se escribía, y se leía vacío.
    // Evaluar el horario y validar la ubicación son solo lecturas, y si el GPS
    // rechaza no hace falta ni abrir el candado. Van antes, para que la parte
    // que hace cola sea lo más corta posible.
    var prog = leerProgramacion();
    var ev = tipo === 'ENTRADA' ? evaluarEntrada(emp, fecha, hora, cfg, prog)
                                : evaluarSalida(emp, fecha, hora, cfg, prog);
    // Se mide contra la sede más cercana de las que le corresponden a la persona
    // (oficina, fundo o cualquiera). Sin sedes adicionales, es la oficina de siempre.
    var ubic = ubicacionEnSedes(emp, cfg, p.lat, p.lng);
    var dist = ubic.dist === null ? '' : ubic.dist;
    var fueraZona = (dist !== '') ? !ubic.dentro : false;

    // La ubicación se exigía SOLO en el celular. Aquí se guardaba la distancia pero
    // el marcaje se aceptaba igual, viniera de donde viniera. Daba lo mismo abrir por
    // el QR, por el enlace o desde la app instalada: la regla vivía en el aparato, y
    // el aparato es justo lo que no se puede dar por bueno. Con el modo "bloquear"
    // la decisión tiene que tomarse acá.
    //   - Teletrabajo y exentos quedan fuera: marcan legítimamente desde su casa.
    //   - MANUAL lo carga administración, que ya se identificó con su clave.
    //   - DECLARADO es la salida de un día anterior que la persona reconoce después;
    //     exigirle GPS a eso no tendría sentido, y queda marcado como declarado.
    // Un permiso aprobado de trabajo desde casa exime del GPS ese día: es
    // exactamente para lo que sirve, y sin esto la persona no podría marcar.
    // Programado para trabajar desde casa ese día: no se le exige ubicación.
    // Cuál de sus tramos cubre la hora a la que está marcando: si en ese momento
    // le toca estar en casa no se le exige ubicación, y si le toca oficina sí.
    // La misma persona el mismo día puede tener las dos cosas.
    var tramoAhora = segmentoEn(prog, emp.id, fecha, hora, true);
    var remotoProgramado = !!(tramoAhora && tramoAhora.modalidad === 'REMOTO');
    // El permiso solo se consulta si el GPS fuera a bloquear de verdad: leer la hoja
    // de permisos en cada marcaje cuesta tiempo dentro del candado, y casi nunca hay uno.
    // La SALIDA no se bloquea por ubicación: quien ya salió y olvidó marcar la
    // declaraba de memoria al día siguiente. Mejor la hora real desde la calle,
    // con su distancia a la vista, que una hora inventada. Queda como fuera de zona.
    var vaAValidarGps = tipo === 'ENTRADA' &&
                        !manualAdmin && p.origen !== 'DECLARADO' && !remotoProgramado &&
                        String(cfg.gpsModo) === 'bloquear' &&
                        String(emp.modalidad || 'PRESENCIAL') === 'PRESENCIAL';
    var permisoCasa = vaAValidarGps ? permisoAprobado(emp.id, fecha, 'CASA') : null;
    if (vaAValidarGps && !permisoCasa) {
      if (p.lat === undefined || p.lat === '' || p.lat === null ||
          p.lng === undefined || p.lng === '' || p.lng === null) {
        return { ok: false, motivo: 'sin_ubicacion',
                 error: 'Sin ubicación activada no se puede registrar tu asistencia. ' +
                        'Actívala y vuelve a intentar.' };
      }
      if (fueraZona) {
        return { ok: false, motivo: 'fuera_zona',
                 error: 'Estás a ' + dist + ' m de ' + (ubic.sedeNombre || 'la oficina') +
                        '. Acércate e inténtalo de nuevo.' };
      }
    }

    var id = 'M' + new Date().getTime() + Math.floor(Math.random() * 900 + 100);

    // ── DESDE AQUÍ SÍ, CON CANDADO ────────────────────────────────────────────
    // Comprobar que no haya marcado ya y escribir tienen que ser una sola cosa:
    // si dos peticiones suyas llegaran a la vez, ambas verían "no ha marcado".
    var lock = LockService.getScriptLock();
    // La hoja y sus columnas se preparan antes: es una lectura de la cabecera y
    // no tiene por qué hacer cola.
    var hMar = asegurarColumnas(HOJA.mar, COLS.mar);

    lock.waitLock(30000);
    try {
    // Relectura fresca a propósito: si el control de duplicados usara la foto
    // tomada antes del candado, dos peticiones simultáneas de la misma persona
    // verían las dos "todavía no marcó" y se escribirían dos veces.
    var todos = filasUltimas(HOJA.mar, COLS.mar, 400);

    // Deduplicación: mismo cid, o misma persona + tipo + fecha
    // Excepción: la entrada de la noche (desde las 21:00) que abre un turno de
    // madrugada del día siguiente no choca con la entrada normal de ese día.
    var nocturna = tipo === 'ENTRADA' && esEntradaNocturna(emp, fecha, hora, cfg, prog);
    var cid = String(p.cid || '');
    var repetido = null, delDia = [];
    for (var i = 0; i < todos.length; i++) {
      var m = todos[i];
      var mismaFecha = fechaTexto(m.fecha) === fecha;
      if (cid && String(m.cid) === cid) {
        return { ok: true, duplicado: true, marcaje: filaAObjeto(m) };
      }
      if (String(m.emp_id) !== emp.id || !mismaFecha) continue;
      var mNoche = aMin(horaTexto(m.hora)) >= 21 * 60;
      // La entrada de la noche que abre el turno de mañana no es de este día
      if (String(m.tipo) === 'ENTRADA' && mNoche && esEntradaNocturna(emp, fecha, horaTexto(m.hora), cfg, prog)) {
        if (tipo === 'ENTRADA' && nocturna) repetido = m;
        continue;
      }
      delDia.push(m);
      if (String(m.tipo) === tipo && !(tipo === 'ENTRADA' && nocturna)) repetido = m;
    }
    /* Jornada partida: de 6 a 9 desde casa y de 19 a 23:59, o de 6 a 9 y de 10:30
       a 18 en la oficina. Cada tramo tiene su entrada y su salida. Una segunda
       entrada se admite si lo último del día fue una salida y lo que toca a esta
       hora es un tramo que empieza después de esa salida; una segunda salida, si
       lo último fue una entrada. Lo demás sigue siendo un doble toque. */
    if (repetido && !nocturna) {
      delDia.sort(function (a, b) { return aMin(horaTexto(a.hora)) - aMin(horaTexto(b.hora)); });
      var ult = delDia[delDia.length - 1];
      var permitido = false;
      if (ult && aMin(hora) > aMin(horaTexto(ult.hora))) {
        if (tipo === 'ENTRADA' && String(ult.tipo) === 'SALIDA') {
          var tramoSig = segmentoEn(prog, emp.id, fecha, hora, true);
          // Administración, que carga con su clave, puede registrar el otro tramo aunque no esté programado
          permitido = manualAdmin || !!(tramoSig && aMin(tramoSig.entrada) >= aMin(horaTexto(ult.hora)) - 60);
        }
        if (tipo === 'SALIDA' && String(ult.tipo) === 'ENTRADA') permitido = true;
      }
      if (!permitido) {
        return { ok: true, duplicado: true, marcaje: filaAObjeto(repetido),
                 aviso: tipo === 'ENTRADA' ? 'Ya registraste tu entrada hoy.' : 'Ya registraste tu salida hoy.' };
      }
    } else if (repetido) {
      return { ok: true, duplicado: true, marcaje: filaAObjeto(repetido), aviso: 'Ya registraste tu entrada hoy.' };
    }

    if (tipo === 'SALIDA') {
      var ayer = diaDespues(fecha, -1);
      var tieneEntrada = todos.some(function (m) {
        if (String(m.emp_id) !== emp.id || String(m.tipo) !== 'ENTRADA') return false;
        if (fechaTexto(m.fecha) === fecha) return true;
        // Turno de madrugada: la entrada se marcó anoche, pasadas las 21:00
        return fechaTexto(m.fecha) === ayer && aMin(horaTexto(m.hora)) >= 21 * 60;
      });
      // Quien registró su llegada porque no pudo marcarla puede marcar la salida
      // aunque todavía nadie la haya aprobado.
      if (!tieneEntrada) tieneEntrada = leerPermisos().some(function (x) {
        return String(x.empId) === emp.id && (x.tipo === 'ENTRADA_FALLIDA' || x.tipo === 'ENTRADA_OLVIDO') &&
               x.fecha === fecha && x.estado !== 'RECHAZADO';
      });
      if (!tieneEntrada) return { ok: false, error: 'No hay una entrada registrada hoy.', motivo: 'sin_entrada' };
    }

    // La programación del día manda sobre el turno: sin esto, alguien programado
    // de 6 a 12 un domingo saldría con una tardanza enorme contra su turno normal.
    anexar(hMar, HOJA.mar, COLS.mar, {
      ts_servidor: new Date(), id: id, cid: cid, emp_id: emp.id, dni: "'" + (emp.dni || ''),
      nombre: emp.nombre, area: emp.area, proyecto: emp.proyecto,
      fecha: fecha, hora: hora, tipo: tipo, origen: p.origen || 'APP', estado: ev.estado,
      tardanza: ev.tardanza || 0, anticipada: ev.anticipada || 0,
      lat: p.lat !== undefined ? p.lat : '', lng: p.lng !== undefined ? p.lng : '',
      dist: dist, fuera_zona: fueraZona ? 'SI' : 'NO', gps: p.gps || '',
      motivo: p.motivo || (diferido ? 'Enviado con retraso: el celular no tenía conexión al marcar' : ''),
      registrado_por: p.registradoPor || '', dispositivo: p.deviceId || '', equipo: p.equipo || '',
      sede: ubic.sede || ''
    });

    return {
      ok: true,
      marcaje: {
        id: id, cid: cid, empId: emp.id, nombre: emp.nombre, dni: emp.dni, area: emp.area,
        proyecto: emp.proyecto, fecha: fecha, hora: hora, tipo: tipo,
        origen: p.origen || 'APP', estado: ev.estado,
        tardanza: ev.tardanza || 0, anticipada: ev.anticipada || 0,
        lat: p.lat || null, lng: p.lng || null, dist: dist === '' ? null : dist,
        fueraZona: fueraZona, gps: p.gps || '', sede: ubic.sede || '', srv: true
      },
      emp: { id: emp.id, nombre: emp.nombre, cargo: emp.cargo, area: emp.area, turno: emp.turno,
             modalidad: emp.modalidad, sede: emp.sede }
    };
  } finally {
    lock.releaseLock();
  }
}

function reporte(p, amb) {
  var desde = p.desde || '0000-01-01';
  var hasta = p.hasta || '9999-12-31';
  var out = filasDesde(HOJA.mar, COLS.mar, desde)
    .filter(function (m) {
      var f = fechaTexto(m.fecha);
      return f >= desde && f <= hasta && (!amb || amb.ids[String(m.emp_id)]);
    })
    .map(filaAObjeto);
  return { ok: true, marcajes: out, total: out.length };
}

/**
 * Borra marcajes: por id concreto, o todos los de un rango de fechas.
 * Sirve para limpiar los marcajes de prueba antes de arrancar en serio, y para
 * corregir un registro equivocado sin tener que buscar la fila a mano.
 */
function borrarMarcajes(p, amb) {
  olvidar();               // lo que se escriba aquí tiene que verse en el próximo pedido
  var ids = [];
  if (p.ids)      ids = (typeof p.ids === 'string') ? JSON.parse(p.ids) : p.ids;
  else if (p.id)  ids = [p.id];
  var desde = p.desde || '', hasta = p.hasta || '';
  if (!ids.length && !(desde && hasta)) {
    return { ok: false, error: 'Indica los marcajes a borrar (id) o un rango de fechas (desde y hasta).' };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var h = hoja(HOJA.mar, COLS.mar);
    var datos = h.getDataRange().getValues();
    if (datos.length < 2) return { ok: true, borrados: 0 };
    var cab = datos[0];
    var cId = cab.indexOf('id'), cFecha = cab.indexOf('fecha');
    var borrados = 0, detalle = [];
    // De abajo hacia arriba: borrar filas no corre las que faltan por revisar
    for (var i = datos.length - 1; i >= 1; i--) {
      var id = String(datos[i][cId]);
      var f  = fechaTexto(datos[i][cFecha]);
      var coincide = ids.length ? (ids.indexOf(id) >= 0) : (f >= desde && f <= hasta);
      if (coincide && amb && !enAmbito(amb, datos[i][cab.indexOf('emp_id')])) coincide = false;   // no es de su gente
      if (coincide) {
        detalle.push(f + ' ' + String(datos[i][cab.indexOf('nombre')]) + ' ' + String(datos[i][cab.indexOf('tipo')]));
        h.deleteRow(i + 1);
        borrados++;
      }
    }
    return { ok: true, borrados: borrados, detalle: detalle.slice(0, 50) };
  } finally { lock.releaseLock(); }
}

function filaAObjeto(m) {
  return {
    id: String(m.id), cid: String(m.cid || ''), empId: String(m.emp_id),
    dni: soloDigitos(m.dni), nombre: String(m.nombre || ''), area: String(m.area || ''),
    proyecto: String(m.proyecto || ''),
    fecha: fechaTexto(m.fecha), hora: horaTexto(m.hora), tipo: String(m.tipo),
    origen: String(m.origen || 'APP'), estado: String(m.estado || ''),
    tardanza: Number(m.tardanza) || 0, anticipada: Number(m.anticipada) || 0,
    lat: m.lat === '' ? null : Number(m.lat), lng: m.lng === '' ? null : Number(m.lng),
    dist: m.dist === '' ? null : Number(m.dist),
    fueraZona: String(m.fuera_zona).toUpperCase() === 'SI',
    gps: String(m.gps || ''), motivo: String(m.motivo || ''),
    registradoPor: String(m.registrado_por || ''),
    dispositivo: String(m.dispositivo || ''),
    equipo: String(m.equipo || ''),
    sede: String(m.sede || ''),
    ts: new Date(m.ts_servidor).getTime() || 0,
    srv: true
  };
}

/* ══════════════════════════════════════════════════════════════
   AUSENCIAS Y PADRÓN — el panel los sincroniza al servidor
   ══════════════════════════════════════════════════════════════ */

function guardarAusencia(p) {
  olvidar();               // lo que se escriba aquí tiene que verse en el próximo pedido
  var lista = p.datos ? (typeof p.datos === 'string' ? JSON.parse(p.datos) : p.datos) : [];
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var n = escribir(HOJA.aus, COLS.aus, lista, function (a) {
      return { id: a.id, emp_id: a.empId, tipo: a.tipo, desde: a.desde, hasta: a.hasta,
               lugar: a.lugar || '', actividad: a.actividad || '', motivo: a.motivo || '',
               autorizado_por: a.autorizadoPor || '', creado: a.creado || '' };
    });
    return { ok: true, ausencias: n };
  } finally { lock.releaseLock(); }
}

/**
 * Escribir el padrón REEMPLAZA la hoja entera. Si llega una lista vacía o mucho
 * más corta que la actual, casi seguro es un error del que llama —un navegador
 * desactualizado, una sincronización a medias— y no una baja masiva real.
 * En ese caso se rechaza. Perder el padrón deja a todos sin poder marcar.
 */
function guardarPersonal(p) {
  olvidar();               // lo que se escriba aquí tiene que verse en el próximo pedido
  var lista = p.datos ? (typeof p.datos === 'string' ? JSON.parse(p.datos) : p.datos) : [];
  var actual = leerPersonal().length;

  if (!lista.length && actual > 0) {
    return { ok: false, motivo: 'padron_vacio',
             error: 'Se rechazó: llegó un padrón vacío y en el servidor hay ' + actual +
                    ' colaboradores. Si de verdad quieres vaciarlo, hazlo desde la hoja de cálculo.' };
  }
  if (actual >= 5 && lista.length < actual / 2 && !p.forzar) {
    return { ok: false, motivo: 'padron_sospechoso',
             error: 'Se rechazó: llegaron ' + lista.length + ' colaboradores y en el servidor hay ' +
                    actual + '. Parece una sincronización incompleta. Recarga el panel y vuelve a intentar.' };
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    asegurarColumnas(HOJA.emp, COLS.emp);   // por si la columna 'planilla' aún no existe
    var n = escribir(HOJA.emp, COLS.emp, lista, function (e) {
      return { id: e.id, nombre: e.nombre, dni: "'" + (e.dni || ''), cargo: e.cargo || '',
               area: e.area || '', proyecto: e.proyecto || '', tel: "'" + (e.tel || ''),
               email: e.email || '', turno: e.turno || '', activo: e.activo ? 'SI' : 'NO',
               modalidad: e.modalidad || 'PRESENCIAL', planilla: e.planilla === false ? 'NO' : 'SI',
               sede: e.sede || '',
               vac_dias: (e.vacDias === null || e.vacDias === undefined || e.vacDias === '') ? '' : Number(e.vacDias),
               vac_desde: e.vacDesde ? "'" + fechaTexto(e.vacDesde) : '',
               dias: diasValidos(e.dias) ? JSON.stringify(diasValidos(e.dias)) : '' };
    });
    return { ok: true, personal: n };
  } finally { lock.releaseLock(); }
}

function guardarConfig(p) {
  olvidar();               // lo que se escriba aquí tiene que verse en el próximo pedido
  var cfg = p.datos ? (typeof p.datos === 'string' ? JSON.parse(p.datos) : p.datos) : {};
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // Guardar reescribe la hoja Config entera. Y estas tres nunca salen del servidor
    // —arranque las quita a propósito, para que no viajen a los celulares—, así que
    // el panel no las tiene y las mandaría en blanco. Sin esta guarda, cambiar una
    // tolerancia borraría la clave de administración y nadie podría volver a entrar.
    // Si llegan vacías se conserva lo que ya había. Para cambiarlas de verdad hay que
    // mandar un valor; para dejarlas en blanco, editar la hoja a mano.
    var previo = leerConfig();
    ['adminPass', 'vistaPass', 'correoAlertas'].forEach(function (k) {
      if (cfg[k] === undefined || String(cfg[k]) === '') cfg[k] = previo[k] || '';
    });

    var pares = [];
    ['empresa','sigla','lat','lng','radio','gpsModo','tolerancia','umbralGrave',
     'inicioOperacion','cierreMargenMin','refrigerioMin','refrigerioDesdeMin','adminPass','vistaPass','exigirDispositivo','correoAlertas',
     'whatsappSoporte','resumenPara'].forEach(function (k) {
      if (cfg[k] !== undefined) pares.push([k, String(cfg[k])]);
    });
    pares.push(['turnos', JSON.stringify(cfg.turnos || [])]);
    pares.push(['proyectos', JSON.stringify(cfg.proyectos || [])]);
    pares.push(['sabadosTrabajados', JSON.stringify(cfg.sabadosTrabajados || [])]);
    pares.push(['domingosTrabajados', JSON.stringify(cfg.domingosTrabajados || [])]);
    // Sedes adicionales (fundo, almacén…). La principal sigue siendo lat/lng/radio.
    pares.push(['nombreSedePrincipal', String(cfg.nombreSedePrincipal || '')]);
    pares.push(['sedesExtra', JSON.stringify(listaSedesExtra(cfg.sedesExtra))]);
    var hc = limpiarHoja(HOJA.cfg, COLS.cfg);
    if (pares.length) hc.getRange(2, 1, pares.length, 2).setValues(pares);
    var nFer = escribir(HOJA.fer, COLS.fer, cfg.feriados || [], function (f) { return { fecha: f.fecha, nombre: f.nombre }; });
    return { ok: true, config: pares.length, feriados: nFer };
  } finally { lock.releaseLock(); }
}

/* ══════════════════════════════════════════════════════════════
   REGLAS DE NEGOCIO — espejo exacto de las del cliente
   ══════════════════════════════════════════════════════════════ */

function buscarEmp(emps, txt) {
  var v = String(txt || '').trim();
  if (!v) return null;
  var d = soloDigitos(v);
  var i;
  if (d.length >= 7 && d.length <= 9) {
    for (i = 0; i < emps.length; i++) if (emps[i].dni && emps[i].dni === d) return emps[i];
  }
  for (i = 0; i < emps.length; i++) if (emps[i].id === v) return emps[i];
  if (d.length >= 8) {
    for (i = 0; i < emps.length; i++) if (emps[i].tel && emps[i].tel === d) return emps[i];
  }
  var mail = v.toLowerCase();
  if (mail.indexOf('@') > 0) {
    for (i = 0; i < emps.length; i++) if (emps[i].email && emps[i].email === mail) return emps[i];
  }
  return null;
}

function esFeriado(cfg, fechaISO) {
  for (var i = 0; i < (cfg.feriados || []).length; i++) {
    if (cfg.feriados[i].fecha === fechaISO) return cfg.feriados[i];
  }
  return null;
}

function turnoDe(cfg, id) {
  for (var i = 0; i < cfg.turnos.length; i++) if (cfg.turnos[i].id === id) return cfg.turnos[i];
  return cfg.turnos[0];
}

function enCalendario(cfg, clave, fechaISO) {
  var lista = (cfg && cfg[clave]) || [];
  return lista.indexOf(fechaISO) >= 0;
}

function horarioDe(emp, fechaISO, cfg, prog) {
  // Un horario asignado a mano manda sobre todo lo demás, incluidos feriados y
  // domingos: si administración programó a alguien ese día, es porque trabaja.
  if (prog) {
    // Un borrador no cambia la jornada de nadie: mientras la jefa lo esté armando,
    // esa persona sigue con su turno normal. Solo lo publicado manda.
    var j = jornadaProgramada(prog, emp.id, fechaISO);
    if (j) return { in: j.in, out: j.out, programado: j.tramos[0], tramos: j.tramos };
  }
  if (esFeriado(cfg, fechaISO)) return null;
  var d = new Date(fechaISO + 'T12:00:00').getDay();
  // Horario propio de la persona (columna "dias" de Personal): manda sobre su turno
  var propios = diasValidos(emp.dias);
  if (propios) {
    var v = propios[d];
    return (v && v.in && v.out) ? { in: v.in, out: v.out } : null;
  }
  var t = turnoDe(cfg, emp.turno);
  // El servidor consultaba solo el turno y se saltaba el calendario: daba por
  // laborable cualquier sábado del turno, marcara o no la empresa ese día concreto.
  // Eso hacía que el estado de un marcaje se calculara distinto aquí y en el panel.
  if (d === 0) return (enCalendario(cfg, 'domingosTrabajados', fechaISO) && siNo(t.trabajaDom))
                      ? { in: t.domEntrada || '10:00', out: t.domSalida || '14:00' } : null;
  if (d === 6) return (enCalendario(cfg, 'sabadosTrabajados', fechaISO) && siNo(t.trabajaSab))
                      ? { in: t.sabEntrada, out: t.sabSalida } : null;
  return { in: t.entrada, out: t.salida };
}

function evaluarEntrada(emp, fechaISO, hora, cfg, prog) {
  var m = aMin(hora);
  // Turno de madrugada marcado la noche anterior (23:58 para el de las 00:00):
  // llegó 2 minutos antes, no 1.438 tarde.
  if (esEntradaNocturna(emp, fechaISO, hora, cfg, prog)) {
    var hs = horarioDe(emp, diaDespues(fechaISO, 1), cfg, prog);
    return clasificarEntrada(m - 1440 - aMin(hs.in), cfg);
  }
  var h = horarioDe(emp, fechaISO, cfg, prog);
  if (!h) return { estado: 'FUERA DE HORARIO', tardanza: 0 };
  return clasificarEntrada(m - aMin(inicioParaEntrada(h, m)), cfg);
}

/* Con varios tramos en el día (de 6 a 9 y de 19 a 23:59) la entrada se mide
   contra el primer tramo que todavía no terminó: 18:57 para el de las 19:00 es
   puntual, no 777 minutos tarde por el de las 6. */
function inicioParaEntrada(h, m) {
  if (!h.tramos || h.tramos.length < 2) return h.in;
  for (var i = 0; i < h.tramos.length; i++) {
    if (aMin(h.tramos[i].salida) > m) return h.tramos[i].entrada;
  }
  return h.tramos[h.tramos.length - 1].entrada;
}

/* Con la jornada partida, cada salida se mide contra el fin de SU tramo: el
   último que ya empezó a esa hora. Salir a las 9:08 del tramo de 6 a 9 es una
   salida completa, no una anticipada de 14 horas contra el fin del día. */
function finParaSalida(h, m) {
  if (!h.tramos || h.tramos.length < 2) return h.out;
  var orden = h.tramos.slice().sort(function (a, b) { return aMin(a.entrada) - aMin(b.entrada); });
  var t = orden[0];
  for (var i = 0; i < orden.length; i++) if (aMin(orden[i].entrada) <= m) t = orden[i];
  return t.salida;
}

/* ¿Es una entrada de la noche (desde las 21:00) para un turno que empieza entre
   las 00:00 y las 02:00 del día siguiente? */
function esEntradaNocturna(emp, fechaISO, hora, cfg, prog) {
  if (aMin(hora) < 21 * 60) return false;
  var hs = horarioDe(emp, diaDespues(fechaISO, 1), cfg, prog);
  return !!(hs && aMin(hs.in) <= 120);
}

function clasificarEntrada(diff, cfg) {
  if (diff <= 0)                 return { estado: 'PUNTUAL',        tardanza: 0 };
  if (diff <= cfg.tolerancia)    return { estado: 'EN TOLERANCIA',  tardanza: 0 };
  if (diff <= cfg.umbralGrave)   return { estado: 'TARDANZA',       tardanza: diff };
  return                                { estado: 'TARDANZA GRAVE', tardanza: diff };
}

/* Siete casillas (0 = domingo … 6 = sábado) con {in, out} o null; si no llega eso, null. */
function diasValidos(v) {
  if (!v) return null;
  if (typeof v === 'string') { try { v = JSON.parse(v); } catch (e) { return null; } }
  if (Object.prototype.toString.call(v) !== '[object Array]' || v.length !== 7) return null;
  return v.map(function (x) {
    return (x && x.in && x.out) ? { in: horaTexto(x.in), out: horaTexto(x.out) } : null;
  });
}

function siNo(v) { return v === true || v === 1 || v === '1' || v === 'true' || v === 'SI'; }

function diaDespues(fechaISO, dias) {
  var d = new Date(new Date(fechaISO + 'T12:00:00').getTime() + dias * 86400000);
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

/* ── Sedes ─────────────────────────────────────────────────────────────── */

function listaSedesExtra(v) {
  if (typeof v === 'string') { try { v = JSON.parse(v); } catch (e) { v = []; } }
  return (v || []).filter(function (s) {
    return s && s.id && s.lat !== null && s.lat !== '' && !isNaN(Number(s.lat)) &&
           s.lng !== null && s.lng !== '' && !isNaN(Number(s.lng));
  }).map(function (s) {
    return { id: String(s.id), nombre: String(s.nombre || 'Sede'), lat: Number(s.lat),
             lng: Number(s.lng), radio: Number(s.radio) || 300 };
  });
}

/* Las sedes donde puede marcar esa persona. Columna "sede" de Personal:
   vacía o PRINCIPAL = oficina, TODAS = cualquiera, o el id de una sede. */
function sedesDe(emp, cfg) {
  var todas = [{ id: 'PRINCIPAL', nombre: String(cfg.nombreSedePrincipal || 'la oficina'),
                 lat: Number(cfg.lat), lng: Number(cfg.lng), radio: Number(cfg.radio) || 150 }]
              .concat(listaSedesExtra(cfg.sedesExtra));
  var s = String(emp.sede || '');
  if (!s || s === 'PRINCIPAL') return [todas[0]];
  if (s === 'TODAS') return todas;
  var una = todas.filter(function (x) { return x.id === s; });
  return una.length ? una : [todas[0]];
}

/* La sede más cercana de las suyas: {sede, sedeNombre, dist, dentro}. */
function ubicacionEnSedes(emp, cfg, lat, lng) {
  if (lat === undefined || lat === '' || lat === null || lng === undefined || lng === '' || lng === null ||
      isNaN(Number(lat)) || isNaN(Number(lng)) || !cfg.lat) {
    return { sede: '', sedeNombre: '', dist: null, dentro: false };
  }
  var mejor = null;
  sedesDe(emp, cfg).forEach(function (x) {
    var d = Math.round(haversine(Number(lat), Number(lng), x.lat, x.lng));
    if (!mejor || d < mejor.dist) mejor = { sede: x.id, sedeNombre: x.nombre, dist: d, dentro: d <= x.radio };
  });
  return mejor;
}

function evaluarSalida(emp, fechaISO, hora, cfg, prog) {
  var h = horarioDe(emp, fechaISO, cfg, prog);
  if (!h) return { estado: '—', anticipada: 0 };
  var diff = aMin(finParaSalida(h, aMin(hora))) - aMin(hora);
  if (diff <= 0) return { estado: 'COMPLETA', anticipada: 0 };
  return { estado: 'SALIDA ANTICIPADA', anticipada: diff };
}

/* ══════════════════════════════════════════════════════════════
   UTILIDADES
   ══════════════════════════════════════════════════════════════ */

function ahoraISO() {
  var d = new Date();
  return {
    fecha: Utilities.formatDate(d, TZ, 'yyyy-MM-dd'),
    hora:  Utilities.formatDate(d, TZ, 'HH:mm')
  };
}

function fechaTexto(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v).replace(/^'/, '').slice(0, 10);
}

function horaTexto(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'HH:mm');
  var s = String(v);
  return s.length >= 5 ? s.slice(0, 5) : s;
}

function soloDigitos(v) { return String(v == null ? '' : v).replace(/\D/g, ''); }

function aMin(hm) {
  var p = String(hm || '').split(':');
  return (Number(p[0]) || 0) * 60 + (Number(p[1]) || 0);
}

function haversine(la1, lo1, la2, lo2) {
  var R = 6371000, r = Math.PI / 180;
  var dLa = (la2 - la1) * r, dLo = (lo2 - lo1) * r;
  var a = Math.sin(dLa / 2) * Math.sin(dLa / 2) +
          Math.cos(la1 * r) * Math.cos(la2 * r) * Math.sin(dLo / 2) * Math.sin(dLo / 2);
  return 2 * R * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function mezclar(a, b) {
  var o = {}, k;
  for (k in a) if (a.hasOwnProperty(k)) o[k] = a[k];
  for (k in b) if (b.hasOwnProperty(k)) o[k] = b[k];
  return o;
}

/* ══════════════════════════════════════════════════════════════
   PRUEBA MANUAL — ejecútala desde el editor para verificar que todo responde
   ══════════════════════════════════════════════════════════════ */

function probar() {
  Logger.log('ping     → ' + JSON.stringify(enrutar({ accion: 'ping' })));
  Logger.log('arranque → ' + JSON.stringify(enrutar({ accion: 'arranque' })).slice(0, 400));
  Logger.log('reporte  → ' + JSON.stringify(enrutar({ accion: 'reporte' })).slice(0, 400));
}

/* ══════════════════════════════════════════════════════════════
   RESUMEN SEMANAL
   Lo pide n8n cada lunes (con la clave de solo vista) y lo reparte por correo
   o WhatsApp. Se calcula aquí, con las mismas reglas que el panel —mismo horario,
   mismos permisos, mismas ausencias—, para que el número del correo y el de la
   pestaña Reportes sean el mismo. Sin fechas, toma la semana pasada (lunes a
   domingo) y la compara con la anterior.
   ══════════════════════════════════════════════════════════════ */
function resumenSemana(p) {
  var cfg = leerConfig(), prog = leerProgramacion();
  var hoy = ahoraISO().fecha;
  var desde = fechaTexto(p.desde || ''), hasta = fechaTexto(p.hasta || '');
  var dow = (new Date(hoy + 'T12:00:00').getDay() + 6) % 7;          // 0 = lunes
  var lunesEsta = diaDespues(hoy, -dow);
  if (!desde || !hasta) { desde = diaDespues(lunesEsta, -7); hasta = diaDespues(lunesEsta, -1); }
  if (hasta >= hoy) hasta = diaDespues(hoy, -1);                     // hoy todavía no cerró
  var largo = Math.round((new Date(hasta + 'T12:00:00') - new Date(desde + 'T12:00:00')) / 86400000) + 1;
  var antDesde = diaDespues(desde, -largo), antHasta = diaDespues(desde, -1);

  var emps = leerPersonal().filter(function (e) {
    return e.activo && String(e.modalidad || '') !== 'EXENTO';
  });
  var porDia = {};
  filasDesde(HOJA.mar, COLS.mar, diaDespues(antDesde, -1)).map(filaAObjeto).forEach(function (m) {
    var k = m.empId + '|' + m.fecha;
    (porDia[k] = porDia[k] || []).push(m);
  });
  var permisos = leerPermisos(), aprob = {}, llegada = {}, pendientes = [];
  permisos.forEach(function (x) {
    if (x.estado === 'APROBADO') aprob[x.empId + '|' + x.fecha + '|' + x.tipo] = x;
    if ((x.tipo === 'ENTRADA_FALLIDA' || x.tipo === 'ENTRADA_OLVIDO') && x.estado !== 'RECHAZADO') llegada[x.empId + '|' + x.fecha] = x;
    if (x.estado === 'PENDIENTE') pendientes.push(x);
  });
  var aus = leerAusencias();
  var ausDe = function (id, f) {
    for (var i = 0; i < aus.length; i++) if (aus[i].empId === id && aus[i].desde <= f && aus[i].hasta >= f) return aus[i];
    return null;
  };

  var jornada = function (emp, f) { return jornadaResumen(emp, f, porDia, cfg, prog); };

  var acumular = function (d1, d2) {
    var tot = { programados: 0, asistidos: 0, faltas: 0, justificadas: 0, tardanzas: 0, minutos: 0,
                puntuales: 0, olvidos: 0, trabajados: 0, diasConJornada: 0 };
    var personas = emps.map(function (emp) {
      var a = { id: emp.id, nombre: emp.nombre, area: emp.area, programados: 0, asistidos: 0, faltas: 0,
                justificadas: 0, tardanzas: 0, minutos: 0, puntuales: 0, olvidos: 0, trabajados: 0,
                diasConJornada: 0, diasFalta: [] };
      for (var f = d1; f <= d2; f = diaDespues(f, 1)) {
        var hor = horarioDe(emp, f, cfg, prog);
        if (!hor) continue;                                           // no laborable o feriado
        var j = jornada(emp, f);
        if (cfg.inicioOperacion && f < fechaTexto(cfg.inicioOperacion) && !j.entrada) continue;
        if (aprob[emp.id + '|' + f + '|DIA'] && !j.entrada) { a.justificadas++; continue; }
        var au = ausDe(emp.id, f);
        if (au && !j.entrada) {
          if (au.tipo === 'COMISION') { a.programados++; a.asistidos++; } else a.justificadas++;
          continue;
        }
        a.programados++;
        if (!j.entrada) {
          if (llegada[emp.id + '|' + f]) { a.asistidos++; a.puntuales++; continue; }
          a.faltas++; a.diasFalta.push(f); continue;
        }
        a.asistidos++;
        var ev = evaluarEntrada(emp, j.entrada.fecha, j.entrada.hora, cfg, prog);
        var tard = ev.tardanza > 0 && !aprob[emp.id + '|' + f + '|TARDANZA'] ? ev.tardanza : 0;
        if (tard > 0) { a.tardanzas++; a.minutos += tard; } else a.puntuales++;
        if (!j.salida || j.salida.origen === 'DECLARADO' || aprob[emp.id + '|' + f + '|SALIDA_OLVIDO']) a.olvidos++;
        // Como el panel: cuentan los tramos cerrados, aunque el último del día quedara sin salida
        if (j.min > 0) { a.trabajados += j.min; a.diasConJornada++; }
      }
      for (var k in tot) tot[k] += a[k];
      return a;
    });
    return { tot: tot, personas: personas };
  };
  var pct = function (n, d) { return d ? Math.round(1000 * n / d) / 10 : null; };
  var indicadores = function (t) {
    return { asistencia: pct(t.asistidos, t.programados),
             puntualidad: pct(t.puntuales, t.puntuales + t.tardanzas),
             faltas: t.faltas, tardanzas: t.tardanzas, minutosTardanza: t.minutos,
             olvidosSalida: t.olvidos, justificadas: t.justificadas,
             horasPromedio: t.diasConJornada ? Math.round(t.trabajados / t.diasConJornada / 6) / 10 : null };
  };

  var actual = acumular(desde, hasta), anterior = acumular(antDesde, antHasta);
  var areas = {};
  actual.personas.forEach(function (x) {
    var g = areas[x.area] = areas[x.area] || { area: x.area, programados: 0, asistidos: 0, puntuales: 0, tardanzas: 0 };
    g.programados += x.programados; g.asistidos += x.asistidos; g.puntuales += x.puntuales; g.tardanzas += x.tardanzas;
  });
  var listaAreas = Object.keys(areas).sort().map(function (k) {
    var g = areas[k];
    return { area: k, asistencia: pct(g.asistidos, g.programados), puntualidad: pct(g.puntuales, g.puntuales + g.tardanzas) };
  }).filter(function (g) { return g.asistencia !== null; });
  var conDias = actual.personas.filter(function (x) { return x.programados > 0; });
  var perfectos = conDias.filter(function (x) { return x.asistidos >= 3 && x.faltas === 0 && x.tardanzas === 0; })
                         .map(function (x) { return x.nombre; });
  var tardes = conDias.filter(function (x) { return x.minutos > 0; })
                      .sort(function (a, b) { return b.minutos - a.minutos; }).slice(0, 5)
                      .map(function (x) { return { nombre: x.nombre, tardanzas: x.tardanzas, minutos: x.minutos }; });
  var faltan = conDias.filter(function (x) { return x.faltas > 0; })
                      .sort(function (a, b) { return b.faltas - a.faltas; })
                      .map(function (x) { return { nombre: x.nombre, faltas: x.faltas, dias: x.diasFalta }; });
  var olvidadizos = conDias.filter(function (x) { return x.olvidos >= 2; })
                           .sort(function (a, b) { return b.olvidos - a.olvidos; })
                           .map(function (x) { return { nombre: x.nombre, olvidos: x.olvidos }; });

  // ¿Está programada la semana que empieza? Quien tuvo tramos la semana pasada y
  // no tiene ninguno en esta, probablemente quedó sin programar.
  var tenia = {}, tiene = {};
  prog.forEach(function (x) {
    if (x.estado === 'BORRADOR') return;
    if (x.fecha >= diaDespues(lunesEsta, -7) && x.fecha < lunesEsta) tenia[x.empId] = x.nombre || x.empId;
    if (x.fecha >= lunesEsta && x.fecha <= diaDespues(lunesEsta, 6)) tiene[x.empId] = true;
  });
  var sinProgramar = Object.keys(tenia).filter(function (id) {
    return !tiene[id] && emps.some(function (e) { return e.id === id; });
  }).map(function (id) { return tenia[id]; });
  var nombreDe = {}; emps.forEach(function (e) { nombreDe[e.id] = e.nombre; });
  var pend = pendientes.filter(function (x) { return nombreDe[x.empId]; }).map(function (x) {
    return { nombre: x.nombre || nombreDe[x.empId], tipo: NOMBRE_PERMISO[x.tipo] || x.tipo, fecha: x.fecha };
  });

  var r = { ok: true, periodo: { desde: desde, hasta: hasta }, anterior: { desde: antDesde, hasta: antHasta },
            indicadores: indicadores(actual.tot), indicadoresAnterior: indicadores(anterior.tot),
            areas: listaAreas, puntualesPerfectos: perfectos, masTardanzas: tardes, faltas: faltan,
            olvidosSalida: olvidadizos, permisosPendientes: pend, semanaSinProgramar: sinProgramar,
            personas: actual.personas, para: destinatariosResumen(cfg) };
  r.texto = textoResumenSemana(r, cfg);
  r.html = htmlResumenSemana(r, cfg);
  r.asunto = 'Resumen semanal de asistencia · ' + fechaCortaSrv(desde) + ' al ' + fechaCortaSrv(hasta);
  return r;
}

/* La jornada de un día, como la arma el panel: la entrada de la noche anterior
   que abre un turno de madrugada es de este día; la de esta noche, del siguiente.
   porDia agrupa los marcajes por 'empId|fecha'. */
function jornadaResumen(emp, f, porDia, cfg, prog) {
  var ms = (porDia[emp.id + '|' + f] || []).filter(function (m) {
    return !(m.tipo === 'ENTRADA' && esEntradaNocturna(emp, f, m.hora, cfg, prog));
  }).sort(function (a, b) { return aMin(a.hora) - aMin(b.hora) || a.ts - b.ts; });
  var pares = [], previa = null;
  if (!ms.some(function (m) { return m.tipo === 'ENTRADA'; })) {
    var ayer = diaDespues(f, -1);
    (porDia[emp.id + '|' + ayer] || []).forEach(function (m) {
      if (m.tipo === 'ENTRADA' && esEntradaNocturna(emp, ayer, m.hora, cfg, prog)) previa = m;
    });
    if (previa) pares.push({ entrada: previa, salida: null, ini: aMin(previa.hora) - 1440 });
  }
  ms.forEach(function (m) {
    var u = pares[pares.length - 1];
    if (m.tipo === 'ENTRADA') { if (!u || u.salida) pares.push({ entrada: m, salida: null, ini: aMin(m.hora) }); }
    else if (u) u.salida = m;
    else pares.push({ entrada: null, salida: m, ini: null });
  });
  var ent = null;
  for (var i = 0; i < pares.length; i++) if (pares[i].entrada) { ent = pares[i]; break; }
  var ult = pares[pares.length - 1];
  var min = 0, hueco = 0, finPrevio = null;
  pares.forEach(function (x) {
    if (!(x.entrada && x.salida)) return;
    if (finPrevio !== null) hueco = Math.max(hueco, x.ini - finPrevio);
    min += Math.max(0, aMin(x.salida.hora) - x.ini);
    finPrevio = aMin(x.salida.hora);
  });
  var refri = refrigerioDe(min, hueco, cfg);
  return { entrada: ent ? ent.entrada : null, salida: ult ? ult.salida : null, min: min - refri, bruto: min, refrigerio: refri };
}

/* El refrigerio no es trabajo: en una jornada continua de más de
   refrigerioDesdeMin (6 h) se descuentan refrigerioMin (60) minutos. Si la
   persona marcó salida y regreso para almorzar (un hueco de 30 min o más), ese
   tiempo ya quedó fuera y no se descuenta otra vez. Igual que el panel. */
function refrigerioDe(bruto, huecoMax, cfg) {
  var ref = (cfg.refrigerioMin === undefined || cfg.refrigerioMin === '' || cfg.refrigerioMin === null) ? 60 : (Number(cfg.refrigerioMin) || 0);
  var desde = (cfg.refrigerioDesdeMin === undefined || cfg.refrigerioDesdeMin === '' || cfg.refrigerioDesdeMin === null) ? 360 : (Number(cfg.refrigerioDesdeMin) || 0);
  if (!ref || bruto <= desde || huecoMax >= 30) return 0;
  return Math.min(ref, bruto);
}

/* ══════════════════════════════════════════════════════════════
   RESUMEN DEL DÍA
   El mismo resumen que el panel arma con el botón «Resumen del día» (cuántos
   vinieron, quién falta, quién llegó tarde, qué queda pendiente), calculado aquí
   para que n8n lo mande solo por WhatsApp a mediodía, sin que nadie abra el
   panel. Lo pide con  accion=resumen&tipo=dia  y la clave de solo vista.
   ══════════════════════════════════════════════════════════════ */
var NOMBRE_AUSENCIA = { COMISION: 'comisión de servicio', VACACIONES: 'vacaciones', MEDICO: 'descanso médico',
  PERMISO: 'permiso', LICENCIA: 'licencia', JUSTIFICADA: 'falta justificada' };
var DIAS_SEMANA = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
var MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre',
  'octubre', 'noviembre', 'diciembre'];

function resumenDia(p) {
  var cfg = leerConfig(), prog = leerProgramacion();
  var ah = ahoraISO(), hoy = ah.fecha, ayer = diaDespues(hoy, -1), ahoraMin = aMin(ah.hora);
  var personal = leerPersonal();
  var emps = personal.filter(function (e) { return e.activo && e.modalidad !== 'EXENTO'; });
  var porDia = {};
  filasDesde(HOJA.mar, COLS.mar, diaDespues(hoy, -2)).map(filaAObjeto).forEach(function (m) {
    var k = m.empId + '|' + m.fecha;
    (porDia[k] = porDia[k] || []).push(m);
  });
  var permisos = leerPermisos();
  var buscar = function (id, f, cond) {
    for (var i = 0; i < permisos.length; i++) {
      var x = permisos[i];
      if (x.empId === id && x.fecha === f && cond(x)) return x;
    }
    return null;
  };
  var aprobado = function (id, f, tipo) {
    return buscar(id, f, function (x) { return x.estado === 'APROBADO' && x.tipo === tipo; });
  };
  var llegadaDe = function (id, f) {
    return buscar(id, f, function (x) {
      return (x.tipo === 'ENTRADA_FALLIDA' || x.tipo === 'ENTRADA_OLVIDO') && x.estado !== 'RECHAZADO';
    });
  };
  var aus = leerAusencias();
  var ausDe = function (id, f) {
    for (var i = 0; i < aus.length; i++) if (aus[i].empId === id && aus[i].desde <= f && aus[i].hasta >= f) return aus[i];
    return null;
  };
  var tol = Number(cfg.tolerancia) || 0;

  var presentes = 0, programados = 0, minutos = 0;
  var faltan = [], tardes = [], llegadas = [], conPermiso = [], sinSalidaAyer = [];
  emps.forEach(function (emp) {
    var n = nombreCorto(emp.nombre);

    // Ayer: marcó entrada (o registró su llegada) y no marcó salida ni pidió corregirla
    var jAyer = jornadaResumen(emp, ayer, porDia, cfg, prog);
    if (!jAyer.salida && (jAyer.entrada || llegadaDe(emp.id, ayer)) &&
        !buscar(emp.id, ayer, function (x) { return x.tipo === 'SALIDA_OLVIDO' && x.estado === 'PENDIENTE'; }))
      sinSalidaAyer.push({ n: n, d: '' });

    var hor = horarioDe(emp, hoy, cfg, prog);
    if (!hor) return;                                                 // no le toca hoy
    var j = jornadaResumen(emp, hoy, porDia, cfg, prog);
    if (cfg.inicioOperacion && hoy < fechaTexto(cfg.inicioOperacion) && !j.entrada) return;
    if (aprobado(emp.id, hoy, 'DIA') && !j.entrada) { conPermiso.push({ n: n, d: 'permiso' }); return; }
    var au = ausDe(emp.id, hoy);
    if (au && !j.entrada) {
      if (au.tipo === 'COMISION') { programados++; presentes++; }
      conPermiso.push({ n: n, d: NOMBRE_AUSENCIA[au.tipo] || 'ausencia' });
      return;
    }
    programados++;
    var lleg = !j.entrada ? llegadaDe(emp.id, hoy) : null;
    if (lleg) {
      presentes++;
      if (lleg.tipo === 'ENTRADA_FALLIDA' && lleg.estado === 'PENDIENTE') llegadas.push({ n: n, d: lleg.hora || '' });
      return;
    }
    if (j.entrada) {
      presentes++;
      var ev = evaluarEntrada(emp, j.entrada.fecha, j.entrada.hora, cfg, prog);
      if (ev.tardanza > 0 && !aprobado(emp.id, hoy, 'TARDANZA')) {
        tardes.push({ n: n, d: ev.tardanza + ' min', min: ev.tardanza });
        minutos += ev.tardanza;
      }
      return;
    }
    if (ahoraMin > aMin(hor.in) + tol) faltan.push({ n: n, d: 'entra ' + hor.in });
  });
  tardes.sort(function (a, b) { return b.min - a.min; });
  var desde7 = diaDespues(hoy, -7);
  var pendientes = permisos.filter(function (x) { return x.estado === 'PENDIENTE' && x.fecha >= desde7; }).length;

  // Los viernes, quién tiene programación esta semana y todavía no la de la próxima
  var sinProgramar = [];
  var dow = new Date(hoy + 'T12:00:00').getDay();
  if (dow === 5) {
    var lunes = diaDespues(hoy, -4), lunesSig = diaDespues(hoy, 3);
    var tiene = {}, tieneSig = {};
    prog.forEach(function (x) {
      if (x.estado === 'BORRADOR') return;
      if (x.fecha >= lunes && x.fecha < lunesSig) tiene[x.empId] = true;
      if (x.fecha >= lunesSig && x.fecha <= diaDespues(lunesSig, 6)) tieneSig[x.empId] = true;
    });
    sinProgramar = emps.filter(function (e) { return tiene[e.id] && !tieneSig[e.id]; })
                       .map(function (e) { return nombreCorto(e.nombre); });
  }

  var d = new Date(hoy + 'T12:00:00');
  var fecha = DIAS_SEMANA[d.getDay()].replace(/^./, function (c) { return c.toUpperCase(); }) + ' ' +
              Number(hoy.slice(8, 10)) + ' de ' + MESES[Number(hoy.slice(5, 7)) - 1];
  var r = {
    ok: true, tipo: 'dia', fecha: hoy, hora: ah.hora, fechaTexto: fecha,
    asistencia: programados ? Math.round(presentes / programados * 100) : null,
    presentes: presentes, programados: programados, minutosTardanza: minutos,
    secciones: [
      { titulo: 'No marcaron entrada', items: faltan },
      { titulo: 'Llegaron tarde', items: tardes.map(function (x) { return { n: x.n, d: x.d }; }) },
      { titulo: 'Llegada registrada sin poder marcar (por aprobar)', items: llegadas },
      { titulo: 'Con permiso, comisión o ausencia', items: conPermiso },
      { titulo: 'Ayer no marcaron salida', items: sinSalidaAyer }
    ].filter(function (s) { return s.items.length; }),
    permisosPendientes: pendientes, semanaSiguienteSinProgramar: sinProgramar,
    para: destinatariosResumen(cfg, personal)
  };
  r.texto = textoResumenDia(r, faltan.length, tardes.length);
  r.asunto = 'Resumen de asistencia · ' + fecha;
  return r;
}

/* El texto con el formato de WhatsApp (*negrita*), igual al del botón del panel. */
function textoResumenDia(r, falta, tarde) {
  var t = '📊 *Resumen de asistencia*\n' + r.fechaTexto + ' · corte ' + r.hora + '\n\n' +
    '✅ Asistencia: *' + (r.asistencia === null ? '—' : r.asistencia + '%') + '* (' + r.presentes + ' de ' + r.programados + ')\n' +
    '❌ Faltan: *' + falta + '*\n' +
    '⏰ Tardanzas: *' + tarde + '*' + (r.minutosTardanza ? ' (' + r.minutosTardanza + ' min en total)' : '') + '\n';
  r.secciones.forEach(function (sec) {
    t += '\n*' + sec.titulo + ' (' + sec.items.length + ')*\n' +
         sec.items.map(function (it) { return '• ' + it.n + (it.d ? ' — ' + it.d : ''); }).join('\n') + '\n';
  });
  if (!r.secciones.length) t += '\n✓ Todo en orden: nadie falta y nadie llegó tarde.\n';
  if (r.permisosPendientes) t += '\n📝 ' + r.permisosPendientes + ' permiso' + (r.permisosPendientes === 1 ? '' : 's') + ' por responder en el panel.';
  if (r.semanaSiguienteSinProgramar.length)
    t += '\n🗓️ Falta programar la próxima semana: ' + r.semanaSiguienteSinProgramar.join(', ') + '.';
  return t.trim();
}

/* Quienes reciben los resúmenes por WhatsApp: los marcados en Configuración →
   «Resumen del día» (resumenPara, ids de Personal) que tengan celular. El número
   sale con el código de Perú delante, como lo pide WhatsApp. */
function destinatariosResumen(cfg, personal) {
  personal = personal || leerPersonal();
  return String(cfg.resumenPara || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean)
    .map(function (id) {
      for (var i = 0; i < personal.length; i++) if (personal[i].id === id) return personal[i];
      return null;
    })
    .filter(function (e) { return e && e.tel; })
    .map(function (e) {
      var t = soloDigitos(e.tel);
      if (t.length === 9) t = '51' + t;
      return { nombre: nombreCorto(e.nombre), numero: t };
    });
}

function fechaCortaSrv(f) { return f.slice(8, 10) + '/' + f.slice(5, 7); }
var NOMBRE_PERMISO = { TARDANZA: 'llegar tarde', SALIDA_ANTES: 'salir antes', DIA: 'día de permiso',
  CASA: 'desde casa', JORNADA: 'trabajé y no marqué', SALIDA_OLVIDO: 'olvidó la salida',
  ENTRADA_FALLIDA: 'no pudo marcar', ENTRADA_OLVIDO: 'olvidó la entrada' };
function nombreCorto(n) {
  return String(n || '').toLowerCase().split(' ').slice(0, 2)
    .map(function (w) { return w.charAt(0).toUpperCase() + w.slice(1); }).join(' ');
}
function textoResumenSemana(r, cfg) {
  var i = r.indicadores, a = r.indicadoresAnterior, L = [];
  L.push('📊 *Resumen semanal de asistencia*');
  L.push((cfg.empresa || 'Grupo Robles') + ' · ' + fechaCortaSrv(r.periodo.desde) + ' al ' + fechaCortaSrv(r.periodo.hasta));
  L.push('');
  L.push('✅ Asistencia: *' + i.asistencia + '%*' + (a.asistencia !== null ? ' (semana anterior ' + a.asistencia + '%)' : ''));
  L.push('⏰ Puntualidad: *' + i.puntualidad + '%*' + (a.puntualidad !== null ? ' (antes ' + a.puntualidad + '%)' : ''));
  L.push('❌ Faltas: ' + i.faltas + ' · Tardanzas: ' + i.tardanzas + ' (' + i.minutosTardanza + ' min)');
  L.push('🚪 Salidas sin marcar: ' + i.olvidosSalida + (i.horasPromedio ? ' · Promedio ' + i.horasPromedio + ' h/día' : ''));
  if (r.areas.length) {
    L.push(''); L.push('*Por área* (asistencia · puntualidad)');
    r.areas.forEach(function (g) { L.push('• ' + g.area + ': ' + g.asistencia + '% · ' + (g.puntualidad === null ? '—' : g.puntualidad + '%')); });
  }
  if (r.puntualesPerfectos.length) { L.push(''); L.push('🏅 *Sin faltas ni tardanzas:* ' + r.puntualesPerfectos.map(nombreCorto).join(', ')); }
  if (r.masTardanzas.length) {
    L.push(''); L.push('🐢 *Más minutos de tardanza*');
    r.masTardanzas.forEach(function (x) { L.push('• ' + nombreCorto(x.nombre) + ': ' + x.minutos + ' min en ' + x.tardanzas + ' día(s)'); });
  }
  if (r.faltas.length) {
    L.push(''); L.push('❌ *Faltas*');
    r.faltas.forEach(function (x) { L.push('• ' + nombreCorto(x.nombre) + ': ' + x.faltas + ' (' + x.dias.map(fechaCortaSrv).join(', ') + ')'); });
  }
  if (r.olvidosSalida.length) { L.push(''); L.push('🚪 *Olvidan la salida:* ' + r.olvidosSalida.map(function (x) { return nombreCorto(x.nombre) + ' (' + x.olvidos + ')'; }).join(', ')); }
  var hay = r.permisosPendientes.length || r.semanaSinProgramar.length;
  if (hay) {
    L.push(''); L.push('⚠️ *Para atender esta semana*');
    if (r.permisosPendientes.length) L.push('• Permisos sin resolver: ' + r.permisosPendientes.map(function (x) {
      return nombreCorto(x.nombre) + ' (' + x.tipo + ', ' + fechaCortaSrv(x.fecha) + ')'; }).join('; '));
    if (r.semanaSinProgramar.length) L.push('• Sin programación esta semana: ' + r.semanaSinProgramar.map(nombreCorto).join(', '));
  }
  return L.join('\n');
}

function htmlResumenSemana(r, cfg) {
  var e = function (s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  var i = r.indicadores, a = r.indicadoresAnterior;
  var tarjeta = function (t, v, ant, malo) {
    var dif = (v !== null && ant !== null) ? Math.round((v - ant) * 10) / 10 : null;
    var bien = dif === null ? null : (malo ? dif < 0 : dif > 0);
    // Bloques que se acomodan solos: cuatro en fila en la computadora, dos y dos en
    // el celular. Con una tabla de cuatro columnas el celular cortaba la cuarta.
    return '<div style="display:inline-block;vertical-align:top;box-sizing:border-box;width:22%;min-width:132px;' +
             'margin:0 6px 8px 0;padding:10px 12px;border:1px solid #e3e8e6;border-radius:10px">' +
      '<div style="font-size:12px;color:#5b6b66">' + t + '</div>' +
      '<div style="font-size:24px;font-weight:800;color:#1b2b4a">' + (v === null ? '—' : v) + '</div>' +
      (dif === null || dif === 0 ? '<div style="font-size:11px;color:#8a9692">igual que la semana anterior</div>'
        : '<div style="font-size:11px;color:' + (bien ? '#1f7a4d' : '#b3261e') + '">' + (dif > 0 ? '▲ +' : '▼ ') + dif + ' vs. semana anterior</div>') +
      '</div>';
  };
  var lista = function (titulo, items) {
    if (!items.length) return '';
    return '<h3 style="font-size:15px;margin:18px 0 6px;color:#1b2b4a">' + titulo + '</h3><ul style="margin:0;padding-left:18px">' +
      items.map(function (x) { return '<li style="margin:2px 0">' + x + '</li>'; }).join('') + '</ul>';
  };
  // lang="es": sin esto Gmail ofrecía traducirlo «del inglés»
  var h = '<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">' +
    '<meta http-equiv="Content-Language" content="es"><meta name="viewport" content="width=device-width,initial-scale=1"></head>' +
    '<body style="margin:0"><div lang="es" style="font-family:Arial,Helvetica,sans-serif;color:#22302c;max-width:640px">' +
    '<h2 style="margin:0;color:#1b2b4a">Resumen semanal de asistencia</h2>' +
    '<div style="color:#5b6b66;margin:4px 0 14px">' + e(cfg.empresa || 'Grupo Robles') + ' · del ' +
      fechaCortaSrv(r.periodo.desde) + ' al ' + fechaCortaSrv(r.periodo.hasta) + '</div>' +
    '<div>' +
      tarjeta('Asistencia %', i.asistencia, a.asistencia) + tarjeta('Puntualidad %', i.puntualidad, a.puntualidad) +
      tarjeta('Faltas', i.faltas, a.faltas, true) + tarjeta('Salidas sin marcar', i.olvidosSalida, a.olvidosSalida, true) +
    '</div>' +
    '<div style="font-size:13px;color:#5b6b66;margin:4px 6px">Tardanzas: ' + i.tardanzas + ' (' + i.minutosTardanza + ' min)' +
      (i.horasPromedio ? ' · Promedio trabajado: ' + i.horasPromedio + ' h por día' : '') + ' · Días justificados: ' + i.justificadas + '</div>';
  if (r.areas.length) {
    h += '<h3 style="font-size:15px;margin:18px 0 6px;color:#1b2b4a">Por área</h3><table style="border-collapse:collapse;font-size:13px;width:100%;max-width:640px;table-layout:fixed">' +
      '<tr style="background:#f2f5f4"><th style="text-align:left;padding:6px;width:50%">Área</th><th style="text-align:right;padding:6px">Asist.</th><th style="text-align:right;padding:6px">Puntual.</th></tr>' +
      r.areas.map(function (g) {
        return '<tr><td style="padding:6px;border-top:1px solid #e3e8e6;word-wrap:break-word">' + e(g.area) + '</td>' +
               '<td style="padding:6px;border-top:1px solid #e3e8e6;text-align:right">' + g.asistencia + '%</td>' +
               '<td style="padding:6px;border-top:1px solid #e3e8e6;text-align:right">' + (g.puntualidad === null ? '—' : g.puntualidad + '%') + '</td></tr>';
      }).join('') + '</table>';
  }
  h += lista('🏅 Sin faltas ni tardanzas', r.puntualesPerfectos.length ? [r.puntualesPerfectos.map(function (n) { return e(nombreCorto(n)); }).join(', ')] : []);
  h += lista('Más minutos de tardanza', r.masTardanzas.map(function (x) { return e(nombreCorto(x.nombre)) + ': <b>' + x.minutos + ' min</b> en ' + x.tardanzas + ' día(s)'; }));
  h += lista('Faltas', r.faltas.map(function (x) { return e(nombreCorto(x.nombre)) + ': <b>' + x.faltas + '</b> (' + x.dias.map(fechaCortaSrv).join(', ') + ')'; }));
  h += lista('Olvidan marcar la salida', r.olvidosSalida.map(function (x) { return e(nombreCorto(x.nombre)) + ': ' + x.olvidos + ' día(s)'; }));
  var at = [];
  if (r.permisosPendientes.length) at.push('<b>' + r.permisosPendientes.length + ' permiso(s) sin resolver</b>: ' +
    r.permisosPendientes.map(function (x) { return e(nombreCorto(x.nombre)) + ' (' + e(x.tipo) + ', ' + fechaCortaSrv(x.fecha) + ')'; }).join('; '));
  if (r.semanaSinProgramar.length) at.push('<b>Sin programación esta semana</b> (la tenían la pasada): ' + r.semanaSinProgramar.map(function (n) { return e(nombreCorto(n)); }).join(', '));
  if (at.length) h += '<div style="background:#fff6e5;border:1px solid #f0d9a8;border-radius:10px;padding:10px 12px;margin-top:18px">' +
    '<b>⚠️ Para atender esta semana</b><ul style="margin:6px 0 0;padding-left:18px">' + at.map(function (x) { return '<li>' + x + '</li>'; }).join('') + '</ul></div>';
  h += '<p style="font-size:11px;color:#8a9692;margin-top:18px">Calculado con las mismas reglas del panel (pestaña Reportes). Aviso automático del control de asistencia.</p></div></body></html>';
  return h;
}
