/**
 * FASES 4 Y 5 — funciones de apoyo para el servidor (Asistencia Grupo Robles)
 *
 * Archivo independiente para agregar al proyecto de Apps Script. Por sí solo no
 * cambia nada: son funciones que el código actual tiene que LLAMAR. Dónde
 * llamarlas se ve mañana, con el código del servidor a la vista:
 *
 *  1. Config  → al guardar ('config') y al devolverla ('arranque'), incluir las
 *     claves  sedesExtra  y  nombreSedePrincipal  (se guardan como texto JSON,
 *     igual que sabadosTrabajados).
 *  2. Personal → nueva columna  sede  (vacía = oficina principal).
 *  3. marcar  → reemplazar el cálculo de distancia por  f45_validarUbicacion(...)
 *     y guardar la sede en una columna nueva  sede  de Marcajes.
 *     Para SALIDA: NO rechazar por fuera_zona; guardar y marcar fuera_zona = SI.
 *  4. marcar  → la tardanza con  f45_tardanza(...)  (turnos de madrugada y
 *     días con varios tramos), igual que la app.
 *  5. resolver → guardar  resueltoPor  tal como llega (hoy la app ya manda el
 *     nombre real de quien aprueba).
 *
 * Todas las funciones empiezan con f45_ para no chocar con las existentes.
 */

/* ---------- Sedes ---------- */

/* Todas las sedes: la principal (lat/lng/radio de Config) y las adicionales. */
function f45_sedes(cfg){
  const principal = {id:'PRINCIPAL', nombre: cfg.nombreSedePrincipal || 'Oficina',
                     lat:Number(cfg.lat), lng:Number(cfg.lng), radio:Number(cfg.radio) || 150};
  let extra = cfg.sedesExtra || [];
  if(typeof extra === 'string'){ try{ extra = JSON.parse(extra); }catch(e){ extra = []; } }
  return [principal].concat((extra || [])
    .filter(s => s && s.id && !isNaN(Number(s.lat)) && !isNaN(Number(s.lng)))
    .map(s => ({id:String(s.id), nombre:s.nombre || 'Sede', lat:Number(s.lat), lng:Number(s.lng),
                radio:Number(s.radio) || 300})));
}

/* Las sedes donde puede marcar esa persona (columna sede de Personal). */
function f45_sedesDe(emp, cfg){
  const todas = f45_sedes(cfg);
  const s = String(emp.sede || '');
  if(!s || s === 'PRINCIPAL') return [todas[0]];
  if(s === 'TODAS') return todas;
  const una = todas.filter(x => x.id === s);
  return una.length ? una : [todas[0]];
}

/* Distancia en metros entre dos puntos. */
function f45_distancia(la1, lo1, la2, lo2){
  const R = 6371000, r = Math.PI/180;
  const dLa = (la2-la1)*r, dLo = (lo2-lo1)*r;
  const a = Math.pow(Math.sin(dLa/2),2) + Math.cos(la1*r)*Math.cos(la2*r)*Math.pow(Math.sin(dLo/2),2);
  return Math.round(2*R*Math.asin(Math.sqrt(a)));
}

/* Valida la ubicación de un marcaje contra las sedes de la persona.
   Devuelve {sede, sedeNombre, dist, dentro}. Sin coordenadas: dentro = false, dist = null. */
function f45_validarUbicacion(emp, cfg, lat, lng){
  if(lat === '' || lat == null || lng === '' || lng == null || isNaN(Number(lat)))
    return {sede:'', sedeNombre:'', dist:null, dentro:false};
  const cand = f45_sedesDe(emp, cfg).map(s => Object.assign({}, s, {
    dist: f45_distancia(Number(lat), Number(lng), s.lat, s.lng)
  })).sort((a, b) => a.dist - b.dist)[0];
  return {sede:cand.id, sedeNombre:cand.nombre, dist:cand.dist, dentro: cand.dist <= cand.radio};
}

/* ¿Se rechaza este marcaje por ubicación?
   La SALIDA nunca se rechaza por estar lejos: queda registrada y señalada. */
function f45_rechazarPorUbicacion(tipo, gpsModo, ubic, esRemoto){
  if(esRemoto || gpsModo !== 'bloquear') return false;
  if(tipo === 'SALIDA') return false;
  return !ubic.dentro;
}

/* ---------- Tardanza (mismo criterio que la app) ---------- */

function f45_min(hhmm){ const p = String(hhmm).split(':'); return Number(p[0])*60 + Number(p[1]); }

/* tramos: [{entrada:'06:00', salida:'09:00'}, …] del día (programación publicada),
   o un solo tramo con el horario del turno.
   Devuelve minutos de diferencia (negativo = llegó antes). */
function f45_tardanza(horaHHMM, tramos){
  if(!tramos || !tramos.length) return 0;
  const m = f45_min(horaHHMM);
  const orden = tramos.slice().sort((a, b) => f45_min(a.entrada) - f45_min(b.entrada));
  // Se mide contra el primer tramo que todavía no terminó
  const vigente = orden.filter(t => f45_min(t.salida) > m)[0] || orden[orden.length - 1];
  const inicio = f45_min(vigente.entrada);
  let diff = m - inicio;
  // Turno que empieza pasada la medianoche, marcado la noche anterior (23:58 → 00:00)
  if(inicio <= 120 && m >= 21*60) diff -= 1440;
  return diff;
}

/* Estado de la entrada a partir de la diferencia en minutos. */
function f45_estadoEntrada(diff, tolerancia, umbralGrave){
  if(diff <= 0)            return {estado:'PUNTUAL',        tardanza:0};
  if(diff <= tolerancia)   return {estado:'EN TOLERANCIA',  tardanza:0};
  if(diff <= umbralGrave)  return {estado:'TARDANZA',       tardanza:diff};
  return                          {estado:'TARDANZA GRAVE', tardanza:diff};
}

/* ---------- Prueba rápida (ejecutar desde el editor) ---------- */
function f45_probar(){
  const cfg = {lat:-12.101999, lng:-77.021853, radio:150, nombreSedePrincipal:'Oficina Lima',
               sedesExtra: JSON.stringify([{id:'SFUNDO', nombre:'Fundo Jayanca', lat:-6.3910, lng:-79.8215, radio:400}])};
  Logger.log(f45_validarUbicacion({sede:'SFUNDO'}, cfg, -6.3905, -79.8210));   // dentro del fundo
  Logger.log(f45_validarUbicacion({sede:''},      cfg, -6.3905, -79.8210));   // oficina: fuera
  Logger.log(f45_tardanza('23:58', [{entrada:'00:00', salida:'06:00'}]));        // -2
  Logger.log(f45_tardanza('18:57', [{entrada:'06:00', salida:'09:00'}, {entrada:'19:00', salida:'23:59'}])); // -3
}
