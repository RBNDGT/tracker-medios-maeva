// ─── CONFIGURACIÓN ────────────────────────────────────────────────────────────
// Reemplaza con el ID de tu Google Sheet
const SHEET_ID   = 'TU_SHEET_ID_AQUI';
const HOJA_DATOS = 'Registros';

const MEDIOS = [
  'ya_nos_conocia','facebook','instagram','x','tiktok','whatsapp',
  'recomendacion','tv','radio','prensa','otro','agencia_viajes',
  'google','eventos','membresias'
];

// ─── RECIBIR REGISTRO (POST desde la app) ─────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const hoja = obtenerHoja();

    const fila = [
      data.fecha,
      new Date().toISOString(),
      data.agente,
      ...MEDIOS.map(m => Number(data.conteos[m]?.cot || 0)),
      ...MEDIOS.map(m => Number(data.conteos[m]?.res || 0))
    ];
    hoja.appendRow(fila);

    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: err.message });
  }
}

// ─── SERVIR DATOS AL DASHBOARD (GET) ──────────────────────────────────────────
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';

  if (action === 'ping') {
    return json({ ok: true, msg: 'Tracker Medios Maeva activo' });
  }

  try {
    const hoja = obtenerHoja();
    if (hoja.getLastRow() < 2) return json({ ok: true, rows: [] });

    const vals    = hoja.getDataRange().getValues();
    const headers = vals[0];

    const rows = vals.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        // UTC evita que la fecha quede un día antes por diferencia de zona horaria
        obj[h] = (row[i] instanceof Date)
          ? Utilities.formatDate(row[i], 'UTC', 'yyyy-MM-dd')
          : row[i];
      });
      return obj;
    });

    return json({ ok: true, rows });
  } catch (err) {
    return json({ ok: false, error: err.message, rows: [] });
  }
}

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function obtenerHoja() {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  let hoja   = ss.getSheetByName(HOJA_DATOS);

  if (!hoja) {
    hoja = ss.insertSheet(HOJA_DATOS);
    const headers = [
      'fecha', 'timestamp', 'agente',
      ...MEDIOS.map(m => `cot_${m}`),
      ...MEDIOS.map(m => `res_${m}`)
    ];
    hoja.appendRow(headers);
    // Formato visual del encabezado
    const rng = hoja.getRange(1, 1, 1, headers.length);
    rng.setFontWeight('bold').setBackground('#E6F7FE').setFontColor('#007DBF');
    hoja.setFrozenRows(1);
  }

  return hoja;
}

// ─── MIGRACIÓN (correr UNA sola vez desde el editor) ──────────────────────────
function migracion() {
  const ss   = SpreadsheetApp.openById(SHEET_ID);
  const hoja = ss.getSheetByName(HOJA_DATOS);

  if (!hoja) { Logger.log('Hoja no encontrada.'); return; }

  const esperados = [
    'fecha', 'timestamp', 'agente',
    ...MEDIOS.map(m => `cot_${m}`),
    ...MEDIOS.map(m => `res_${m}`)
  ];

  const actuales = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];

  if (actuales.length >= esperados.length) {
    Logger.log('Ya actualizado (' + actuales.length + ' columnas). Nada que hacer.');
    return;
  }

  // Leer todos los datos y remapear por nombre de columna
  const datos = hoja.getDataRange().getValues();
  const nuevaData = [esperados];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    nuevaData.push(esperados.map(h => {
      const idx = actuales.indexOf(h);
      return idx >= 0 ? fila[idx] : 0;
    }));
  }

  // Reescribir hoja con el nuevo formato
  hoja.clearContents();
  hoja.getRange(1, 1, nuevaData.length, esperados.length).setValues(nuevaData);

  // Re-aplicar formato de encabezado
  hoja.getRange(1, 1, 1, esperados.length)
    .setFontWeight('bold').setBackground('#E6F7FE').setFontColor('#007DBF');
  hoja.setFrozenRows(1);

  Logger.log('Migración OK — ' + (nuevaData.length - 1) + ' filas conservadas, ' + esperados.length + ' columnas.');
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
Add Code.gs backend
