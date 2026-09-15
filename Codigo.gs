// ============================================================
// Tu Barrio Más Seguro — Backend en Google Apps Script
// Este archivo va PEGADO en el editor de Apps Script (Extensiones > Apps Script)
// de la MISMA Google Sheet donde está la pestaña "Proyectos" e "Historial".
// No necesita ID de planilla: al estar "atado" a la Sheet, usa la Sheet activa.
// ============================================================

const SHEET_PROYECTOS = 'Proyectos';
const SHEET_HISTORIAL = 'Historial';

// Carpeta de Drive donde se guardarán los PDF de autorización de Modificación.
// Déjalo vacío ('') para usar tu carpeta raíz de Drive, o pega aquí el ID
// de una carpeta específica (lo ves en la URL de la carpeta en Drive).
const DRIVE_FOLDER_ID = '';

const ETAPA_PREFIX = {
  retiro_cheque: 'RetiroCheque',
  cobro_deposito: 'CobroDeposito',
  compra_insumos: 'CompraInsumos',
  agenda_instalacion: 'AgendaInstalacion',
  instalacion: 'Instalacion',
  contacto_agendamiento: 'ContactoAgendamiento',
  modificacion: 'Modificacion',
  rendicion: 'Rendicion',
  fotografias: 'Fotografias'
};

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Tu Barrio Más Seguro — Seguimiento')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('No se encontró la pestaña "' + name + '". Revisa que exista con ese nombre exacto.');
  return sheet;
}

function sheetToObjects_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(row => row[0] !== '' && row[0] !== null)
    .map(row => {
      const o = {};
      headers.forEach((h, i) => { o[h] = row[i]; });
      return o;
    });
}

// Devuelve todos los proyectos con su estado actual (para el Tablero y el Reportar)
function getProyectos() {
  return sheetToObjects_(getSheet_(SHEET_PROYECTOS));
}

// Devuelve el historial completo de un proyecto, más reciente primero
function getHistorial(id) {
  const rows = sheetToObjects_(getSheet_(SHEET_HISTORIAL));
  return rows.filter(r => String(r.ID_Proyecto) === String(id)).reverse();
}

function findRowIndex_(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // +2: fila 1 es encabezado, base 1
  }
  return -1;
}

// Guarda el avance de una etapa normal (no Modificación).
// payload: {Estado, Fecha, Obs, Por, Enlaces (solo fotografías, opcional)}
function guardarEtapa(id, etapaKey, payload) {
  const sheet = getSheet_(SHEET_PROYECTOS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowIdx = findRowIndex_(sheet, id);
  if (rowIdx === -1) throw new Error('No se encontró el proyecto con ID: ' + id);

  const prefix = ETAPA_PREFIX[etapaKey];
  if (!prefix) throw new Error('Etapa no reconocida: ' + etapaKey);

  Object.keys(payload).forEach(field => {
    const colName = prefix + '_' + field;
    const colIdx = headers.indexOf(colName);
    if (colIdx > -1) sheet.getRange(rowIdx, colIdx + 1).setValue(payload[field]);
  });

  const ultimaCol = headers.indexOf('UltimaActualizacion');
  if (ultimaCol > -1) sheet.getRange(rowIdx, ultimaCol + 1).setValue(new Date());

  getSheet_(SHEET_HISTORIAL).appendRow([
    new Date(), id, etapaKey, payload.Estado || '', payload.Fecha || '',
    payload.Obs || '', payload.Por || '', payload.Tipo || '', payload.Detalle || '',
    payload.Enlace || '', payload.ArchivoURL || ''
  ]);
  return true;
}

// Sube un PDF (recibido como base64 desde el formulario) a Drive y devuelve su URL pública (solo con el link)
function guardarArchivoModificacion(base64, filename, mimeType) {
  const folder = DRIVE_FOLDER_ID ? DriveApp.getFolderById(DRIVE_FOLDER_ID) : DriveApp.getRootFolder();
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType, filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// Agrega un proyecto nuevo individual desde el formulario (pestaña "Cargar proyectos")
function agregarProyecto(datos) {
  const sheet = getSheet_(SHEET_PROYECTOS);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rutDigits = String(datos.RUT || '').replace(/[^0-9kK]/gi, '').toUpperCase();
  let id = rutDigits || ('P' + Utilities.getUuid().slice(0, 6).toUpperCase());
  if (findRowIndex_(sheet, id) > -1) id = id + '-' + Utilities.getUuid().slice(0, 4);

  const row = headers.map(h => {
    if (h === 'ID') return id;
    if (h === 'UltimaActualizacion') return '';
    return datos[h] !== undefined ? datos[h] : '';
  });
  sheet.appendRow(row);
  return id;
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
