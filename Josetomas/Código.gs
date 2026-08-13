/**
 * BACKEND — "Play with Baby Josetomas" (con usuario y contraseña)
 * Este código va pegado en el Editor de Apps Script de tu Google Sheet.
 * Ver INSTRUCCIONES.md para el paso a paso.
 */

const SHEET_USUARIOS = 'Usuarios';
const SHEET_CONFIG = 'Config';

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = params.action;
  if (action === 'getWinner') return getWinner();
  if (action === 'getResponses') return getResponses();
  return jsonResponse({ error: 'Acción no reconocida' });
}

function doPost(e) {
  let data = {};
  if (e && e.postData && e.postData.contents) {
    try {
      data = JSON.parse(e.postData.contents);
    } catch (err) {
      return jsonResponse({ error: 'Datos inválidos en el cuerpo de la petición' });
    }
  }

  const action = data.action;

  if (action === 'loginOrRegister') return loginOrRegister(data);
  if (action === 'submitVote') return submitVote(data);
  if (action === 'setAnswer') return setCorrectAnswer(data);
  if (action === 'verifyAdmin') return verifyAdmin(data);
  return jsonResponse({ error: 'Acción no reconocida' });
}

// ---------- Hojas ----------

const CABECERA_USUARIOS = ['Usuario', 'PasswordHash', 'Piel', 'Cabello', 'Hora', 'Parto', 'Semana', 'Telefono', 'FechaRegistro', 'FechaVoto', 'VotoRegistrado'];
const NUM_COLUMNAS = 11;

function getUsersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_USUARIOS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_USUARIOS);
    sheet.appendRow(CABECERA_USUARIOS);
    return sheet;
  }
  
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  
  // Si la hoja ya tiene 11 columnas, está bien
  if (lastCol >= NUM_COLUMNAS) return sheet;
  
  // --- MIGRACIÓN de formato viejo (10 columnas, sin Telefono) a nuevo (11 columnas) ---
  // Formato viejo: Usuario(1), PasswordHash(2), Piel(3), Cabello(4), Hora(5), Parto(6), Semana(7), FechaRegistro(8), FechaVoto(9), VotoRegistrado(10)
  // Formato nuevo: Usuario(1), PasswordHash(2), Piel(3), Cabello(4), Hora(5), Parto(6), Semana(7), Telefono(8), FechaRegistro(9), FechaVoto(10), VotoRegistrado(11)
  
  // Paso 1: Escribir header completo en fila 1
  for (let c = 1; c <= NUM_COLUMNAS; c++) {
    sheet.getRange(1, c).setValue(CABECERA_USUARIOS[c - 1]);
  }
  
  // Paso 2: Migrar datos si existen
  if (lastRow >= 2) {
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      const rowNum = r + 2; // fila real en la hoja
      
      if (lastCol === 10) {
        // Formato 10 columnas: mover FechaRegistro de col8→col9, FechaVoto de col9→col10, VotoRegistrado de col10→col11
        if (row[7]) sheet.getRange(rowNum, 9).setValue(row[7]);  // FechaRegistro
        if (row[8]) sheet.getRange(rowNum, 10).setValue(row[8]); // FechaVoto
        if (row[9] !== undefined && row[9] !== '') sheet.getRange(rowNum, 11).setValue(row[9]); // VotoRegistrado
        // Col 8 (Telefono) queda vacío
      } else if (lastCol === 9) {
        // Formato 9 columnas: FechaRegistro en col8, FechaVoto+VotoRegistrado combinados
        if (row[7]) sheet.getRange(rowNum, 9).setValue(row[7]); // FechaRegistro
        if (row[8] !== undefined && row[8] !== '') sheet.getRange(rowNum, 11).setValue(row[8]); // VotoRegistrado
      }
    }
  }
  
  return sheet;
}

// Obtiene el índice (1-based) de la columna VotoRegistrado dinámicamente
function getVotoRegistradoColIndex() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_USUARIOS);
  if (!sheet) return 11; // default
  const header = sheet.getRange(1, 1, 1, NUM_COLUMNAS).getValues()[0];
  for (let i = 0; i < header.length; i++) {
    if (String(header[i]).trim().toLowerCase() === 'votoregistrado') return i + 1;
  }
  return 11; // fallback
}

// Obtiene el índice (0-based) de VotoRegistrado en un array de fila
function getVotoRegistradoIdx(row) {
  for (let i = 0; i < row.length; i++) {
    // Buscar el valor TRUE en las últimas columnas
  }
  // Según el formato: 11 cols = índice 10, 10 cols = índice 9
  if (row.length >= 11) return 10;
  return row.length - 1;
}

function getConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_CONFIG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_CONFIG);
    sheet.appendRow(['Piel', 'Cabello', 'Hora', 'Parto', 'Semana', 'Password']);
    // 👇 Cambia esta contraseña de administrador antes de compartir el link de admin
    sheet.appendRow(['', '', '', '', '', 'baby2026']);
  } else {
    // Asegurar que la fila 2 exista y tenga contraseña
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      sheet.appendRow(['', '', '', '', '', 'baby2026']);
    } else {
      const existingPass = sheet.getRange(2, 6).getValue();
      if (!existingPass || String(existingPass).trim() === '') {
        sheet.getRange(2, 6).setValue('baby2026');
      }
    }
  }
  return sheet;
}

// ---------- Utilidad de contraseña ----------

function hashPassword(pw) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8);
  return bytes.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0')).join('');
}

function findUserRow(sheet, usuario) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toLowerCase() === String(usuario).trim().toLowerCase()) {
      return { rowIndex: i + 1, row: values[i] }; // rowIndex es 1-indexado para la hoja
    }
  }
  return null;
}

// ---------- Acciones ----------

function loginOrRegister(data) {
  const usuario = String(data.usuario || '').trim();
  const password = String(data.password || '');

  if (!usuario || !password) {
    return jsonResponse({ error: 'Usuario y contraseña son obligatorios' });
  }
  if (password.length < 4) {
    return jsonResponse({ error: 'La contraseña debe tener al menos 4 caracteres' });
  }

  const sheet = getUsersSheet();
  const existing = findUserRow(sheet, usuario);
  const hash = hashPassword(password);

  if (existing) {
    if (existing.row[1] !== hash) {
      return jsonResponse({ error: 'Contraseña incorrecta para ese usuario' });
    }
    const voto = existing.row[10] === true || existing.row[10] === 'TRUE';
    const result = { success: true, isNew: false, voto: voto };
    if (voto) {
      result.respuestas = {
        piel: existing.row[2],
        cabello: existing.row[3],
        hora: existing.row[4],
        parto: existing.row[5],
        semana: existing.row[6]
      };
    }
    return jsonResponse(result);
  }

  // Usuario nuevo -> se registra automáticamente (incluye teléfono)
  sheet.appendRow([usuario, hash, '', '', '', '', '', data.telefono || '', new Date(), '', false]);
  return jsonResponse({ success: true, isNew: true, voto: false });
}

function submitVote(data) {
  const usuario = String(data.usuario || '').trim();
  const password = String(data.password || '');
  const sheet = getUsersSheet();
  const existing = findUserRow(sheet, usuario);

  if (!existing) {
    return jsonResponse({ error: 'Usuario no encontrado. Vuelve a iniciar sesión.' });
  }
  const hash = hashPassword(password);
  if (existing.row[1] !== hash) {
    return jsonResponse({ error: 'Contraseña incorrecta' });
  }
  const yaVoto = existing.row[10] === true || existing.row[10] === 'TRUE';
  if (yaVoto) {
    return jsonResponse({ error: 'Ya registraste tu predicción anteriormente. No se puede cambiar.' });
  }

  const valuesToSave = [
    data.piel || '',
    data.cabello || '',
    data.hora || '',
    data.parto || '',
    data.semana || ''
  ];

  sheet.getRange(existing.rowIndex, 3, 1, 5).setValues([valuesToSave]);
  sheet.getRange(existing.rowIndex, 8).setValue(data.telefono || existing.row[7] || '');
  sheet.getRange(existing.rowIndex, 10).setValue(new Date());
  sheet.getRange(existing.rowIndex, 11).setValue(true);

  return jsonResponse({ success: true });
}

function setCorrectAnswer(data) {
  const configSheet = getConfigSheet();
  const storedPassword = configSheet.getRange(2, 6).getValue();
  if (String(data.password) !== String(storedPassword)) {
    return jsonResponse({ error: 'Contraseña incorrecta' });
  }
  configSheet.getRange(2, 1, 1, 5).setValues([[data.piel, data.cabello, data.hora, data.parto, data.semana]]);
  return jsonResponse({ success: true });
}

function verifyAdmin(data) {
  const configSheet = getConfigSheet();
  const storedPassword = configSheet.getRange(2, 6).getValue();
  if (String(data.password) !== String(storedPassword)) {
    return jsonResponse({ success: false, error: 'Contraseña de administrador incorrecta' });
  }
  return jsonResponse({ success: true });
}

function getResponses() {
  const sheet = getUsersSheet();
  const values = sheet.getDataRange().getValues();
  values.shift();
  const rows = values
    .filter(r => {
      const idx = getVotoRegistradoIdx(r);
      return r[idx] === true || String(r[idx]).toUpperCase() === 'TRUE';
    })
    .map(r => ({
      usuario: r[0],
      piel: r[2],
      cabello: r[3],
      hora: r[4],
      parto: r[5],
      semana: r[6],
      telefono: r.length >= 11 ? (r[7] || '') : '',
      fechaVoto: r.length >= 11 ? r[9] : (r[7] || '')
    }));
  return jsonResponse({ data: rows });
}

function getWinner() {
  const configSheet = getConfigSheet();
  const config = configSheet.getRange(2, 1, 1, 5).getValues()[0];
  const [piel, cabello, hora, parto, semana] = config;

  if (!piel || !cabello || !hora || !parto || !semana) {
    return jsonResponse({ configured: false });
  }

  const sheet = getUsersSheet();
  const values = sheet.getDataRange().getValues();
  values.shift();

  const resultados = values
    .filter(r => {
      const idx = getVotoRegistradoIdx(r);
      return r[idx] === true || String(r[idx]).toUpperCase() === 'TRUE';
    })
    .map(r => {
      const votoIdx = getVotoRegistradoIdx(r);
      const usuario = r[0];
      const pPiel = r[2];
      const pCabello = r[3];
      const pHora = r[4];
      const pParto = r[5];
      const pSemana = r[6];
      const pTelefono = r.length >= 11 ? (r[7] || '') : '';
      let aciertos = 0;
      if (pPiel === piel) aciertos++;
      if (pCabello === cabello) aciertos++;
      if (pHora === hora) aciertos++;
      if (pParto === parto) aciertos++;
      if (pSemana === semana) aciertos++;
      return { usuario, piel: pPiel, cabello: pCabello, hora: pHora, parto: pParto, semana: pSemana, telefono: pTelefono, aciertos };
    });

  resultados.sort((a, b) => b.aciertos - a.aciertos);

  return jsonResponse({
    configured: true,
    respuestaCorrecta: { piel, cabello, hora, parto, semana },
    resultados
  });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}