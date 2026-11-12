/**
 * AGENDA PERSONAL - Backend (Apps Script ligado al Google Sheet)
 * 1) Ejecuta setup() una vez. 2) Implementar > Aplicación web (Ejecutar como: Yo, Acceso: Cualquier persona).
 * 3) Opcional IA real: Configuración del proyecto > Propiedades del script > ANTHROPIC_API_KEY
 */
const TZ = Session.getScriptTimeZone();
const HEAD = {
  Agenda_Diaria: ['Fecha','Actividad','Tipo','Hora_Plan_Inicio','Hora_Plan_Fin','Hora_Real_Inicio','Hora_Real_Fin','Estado'],
  Pendientes: ['ID','Tarea','Fecha_Objetivo','Prioridad','Categoria','Completado'],
  Notas_Diarias: ['Fecha','Hora','Nota_o_Idea'],
  Registro_Diario: ['Fecha','Energia','Enfoque','Frase_IA','Versiculo_IA']
};

function setup() {
  const ss = SpreadsheetApp.getActive();
  for (const n in HEAD) {
    const s = ss.getSheetByName(n) || ss.insertSheet(n);
    s.getRange(1, 1, 1, HEAD[n].length).setValues([HEAD[n]]).setFontWeight('bold');
    s.setFrozenRows(1);
  }
}

/* ---------- utilidades ---------- */
const sh_ = n => SpreadsheetApp.getActive().getSheetByName(n);
const S = (v, f) => v instanceof Date ? Utilities.formatDate(v, TZ, f || 'yyyy-MM-dd') : String(v == null ? '' : v);
const json_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function rows_(n) {
  return sh_(n).getDataRange().getValues().slice(1)
    .map((r, i) => ({ r, row: i + 2 })).filter(x => x.r[0] !== '');
}
function find_(n, col, val) {
  const x = rows_(n).find(x => S(x.r[col]) === val);
  return x ? x.row : 0;
}
// Agrega filas al final (según columna A). cb = columna de checkbox (0 = ninguna). Devuelve la fila inicial.
function append_(n, data, cb) {
  const s = sh_(n), w = data[0].length;
  const colA = s.getRange(1, 1, s.getMaxRows(), 1).getValues();
  let last = colA.length;
  while (last > 1 && colA[last - 1][0] === '') last--;
  const start = last + 1;
  if (start + data.length - 1 > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), data.length + 50);
  const rg = s.getRange(start, 1, data.length, w);
  rg.setNumberFormat('@').setValues(data);
  if (cb) s.getRange(start, cb, data.length, 1).setNumberFormat('General').insertCheckboxes().setValues(data.map(r => [r[cb - 1]]));
  return start;
}

/* ---------- IA: frase + versículo (con respaldo local) ---------- */
function ia_() {
  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (key) {
    try {
      const r = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-api-key': key, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({
          model: 'claude-haiku-5-5', max_tokens: 300,
          messages: [{ role: 'user', content: 'Responde SOLO un JSON {"frase":"...","versiculo":"texto (Libro cap:vers)"}. La frase: disciplina y constancia, breve, original. El versículo: bíblico, al azar, en español.' }]
        })
      });
      const t = JSON.parse(r.getContentText()).content[0].text;
      return JSON.parse(t.replace(/```json|```/g, '').trim());
    } catch (e) {}
  }
  const F = ['La disciplina es elegir lo que quieres más por encima de lo que quieres ahora.', 'Hazlo hoy, aunque sea pequeño. La constancia gana al talento.', 'Cumple la promesa que te hiciste esta mañana.'];
  const V = ['Todo lo puedo en Cristo que me fortalece. (Filipenses 4:13)', 'Encomienda a Jehová tus obras, y tus pensamientos serán afirmados. (Proverbios 16:3)', 'Esfuérzate y sé valiente; no temas ni desmayes. (Josué 1:9)'];
  return { frase: F[Math.floor(Math.random() * F.length)], versiculo: V[Math.floor(Math.random() * V.length)] };
}

/* ---------- GET: cargar el día ---------- */
function doGet(e) {
  const l = LockService.getScriptLock(); l.waitLock(20000);
  try {
    if (e.parameter.action !== 'getDay') return json_({ ok: false, error: 'acción inválida' });
    return json_(getDay_(e.parameter.fecha));
  } catch (err) { return json_({ ok: false, error: String(err) }); }
  finally { l.releaseLock(); }
}

function getDay_(f) {
  const toA = x => ({ row: x.row, fecha: S(x.r[0]), actividad: S(x.r[1]), tipo: S(x.r[2]), pi: S(x.r[3], 'HH:mm'), pf: S(x.r[4], 'HH:mm'), ri: S(x.r[5], 'HH:mm'), rf: S(x.r[6], 'HH:mm'), estado: x.r[7] === true });
  let all = rows_('Agenda_Diaria').map(toA), hoy = all.filter(x => x.fecha === f);
  if (!hoy.length) { // copia la rutina del día anterior más reciente
    const prev = all.filter(x => x.tipo === 'Rutina' && x.fecha < f).map(x => x.fecha).sort().pop();
    if (prev) {
      const nuevas = all.filter(x => x.fecha === prev && x.tipo === 'Rutina').map(x => [f, x.actividad, 'Rutina', x.pi, x.pf, '', '', false]);
      append_('Agenda_Diaria', nuevas, 8);
      hoy = rows_('Agenda_Diaria').map(toA).filter(x => x.fecha === f);
    }
  }
  const pendientes = rows_('Pendientes').map(x => ({ id: S(x.r[0]), tarea: S(x.r[1]), fecha: S(x.r[2]), prioridad: S(x.r[3]), categoria: S(x.r[4]), completado: x.r[5] === true }))
    .filter(p => !p.completado || p.fecha === f);
  const notas = rows_('Notas_Diarias').filter(x => S(x.r[0]) === f).map(x => ({ hora: S(x.r[1], 'HH:mm'), nota: S(x.r[2]) }));
  let rr = rows_('Registro_Diario').find(x => S(x.r[0]) === f);
  if (!rr) {
    const ia = ia_();
    append_('Registro_Diario', [[f, '', '', ia.frase, ia.versiculo]], 0);
    rr = rows_('Registro_Diario').find(x => S(x.r[0]) === f);
  }
  const registro = { energia: rr.r[1] === '' ? '' : Number(rr.r[1]), enfoque: rr.r[2] === '' ? '' : Number(rr.r[2]), frase: S(rr.r[3]), versiculo: S(rr.r[4]) };
  return { ok: true, agenda: hoy, pendientes, notas, registro };
}

/* ---------- POST: guardar (el cliente envía text/plain para evitar preflight CORS) ---------- */
function doPost(e) {
  const l = LockService.getScriptLock(); l.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents); let out = { ok: true };
    switch (d.action) {
      case 'addBloque':
        out.row = append_('Agenda_Diaria', [[d.fecha, d.actividad, d.tipo || 'Evento', d.pi, d.pf, '', '', false]], 8); break;
      case 'setBloque':
        sh_('Agenda_Diaria').getRange(d.row, 6, 1, 3).setValues([[d.ri, d.rf, d.estado]]); break;
      case 'addPendiente':
        append_('Pendientes', [[d.id, d.tarea, d.fecha, d.prioridad, d.categoria || '', false]], 6); break;
      case 'setPendiente': {
        const r = find_('Pendientes', 0, d.id), s = sh_('Pendientes');
        if (r) { if (d.completado !== undefined) s.getRange(r, 6).setValue(d.completado); if (d.fecha) s.getRange(r, 3).setValue(d.fecha); }
        break;
      }
      case 'addNota':
        append_('Notas_Diarias', [[d.fecha, d.hora, d.nota]], 0); break;
      case 'setRegistro': {
        let r = find_('Registro_Diario', 0, d.fecha);
        if (!r) r = append_('Registro_Diario', [[d.fecha, '', '', '', '']], 0);
        sh_('Registro_Diario').getRange(r, 2, 1, 2).setValues([[d.energia, d.enfoque]]); break;
      }
      default: out = { ok: false, error: 'acción inválida' };
    }
    return json_(out);
  } catch (err) { return json_({ ok: false, error: String(err) }); }
  finally { l.releaseLock(); }
}
