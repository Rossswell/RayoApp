const SPREADSHEET_ID = '1O35ZWq2yEAW-5FOJsby2Kszo6WP3_0nqVGp4yd0q4Lk';
const EMP_SHEET      = '👤EMPLEADOS FLASH HIGH';
const HORARIOS_SHEET = '⌛HORARIOS';
const TIMEZONE       = 'America/Caracas';
 
function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  let p = e.parameter || {};
  
  // Obtener datos del cuerpo POST (JSON)
  if (e.postData && e.postData.contents) {
    try {
      const payload = JSON.parse(e.postData.contents);
      p = { ...p, ...payload };
    } catch(err) {
      // No es JSON, p se queda con e.parameter
    }
  }

  // Limpiamos la acción
  const action = (p.action || "").toString().trim();
  
  try {
    if (action === 'login')      return json(login(p));
    if (action === 'register')   return json(register(p));
    if (action === 'getRecords') return json(getRecords(p));
    if (action === 'verify2FA')  return json(verify2FA(p));
    
    // Si llegamos aquí, devolvemos qué acción detectó el script
    return json({ 
      success: false, 
      error: 'Acción desconocida: "' + action + '". Verifica que hayas desplegado una NUEVA VERSIÓN en Google Apps Script.' 
    });
  } catch(err) {
    return json({ success:false, error: 'Error interno: ' + err.message });
  }
}
 
function json(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
 
function getSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    // Intentar con un espacio después del emoji por si acaso
    const withSpace = name.replace(/^(.)/, "$1 ");
    sheet = ss.getSheetByName(withSpace);
  }
  return sheet;
}

function login(p) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheet(ss, EMP_SHEET);
  
  if (!sheet) return { success:false, error: 'No se encontró la hoja: ' + EMP_SHEET };
  
  try {
    const rows = sheet.getRange('A3:H100').getValues();
    for (const row of rows) {
      if (!row[0]) continue;
      const correo = (row[3]||'').toString().trim().toLowerCase();
      const cedula = (row[1]||'').toString().trim();
      
      if (correo === p.correo.trim().toLowerCase()) {
        if (cedula === p.password.trim()) {
          const emp = {
            nombre:row[0], cedula:row[1], telefono:row[2],
            correo:row[3], area:row[4],
            fecha_inicio:row[5], fecha_cumple:row[6]
          };
          
          const hSheet = getSheet(ss, HORARIOS_SHEET);
          let horario = "No asignado";
          let toleranceMins = 0;
          if (hSheet) {
            const hData = hSheet.getRange('A2:D100').getValues();
            for (const hRow of hData) {
              if (String(hRow[1]).trim() === String(emp.cedula).trim()) {
                horario = hRow[2] || "No asignado";
                toleranceMins = parseInt(hRow[3]) || 0;
                break;
              }
            }
          }
          emp.horario = horario;
          emp.tolerancia = toleranceMins;
          return { success:true, employee: emp };
        }
        return { success:false, reason:'password' };
      }
    }
    return { success:false, reason:'correo' };
  } catch(e) {
    return { success:false, error: 'Error en login: ' + e.message };
  }
}
 
function register(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hSheet = getSheet(ss, HORARIOS_SHEET);
  if (!hSheet) return { success:false, error:'Hoja de horarios no encontrada' };

  const hData = hSheet.getRange('A2:D100').getValues();
  let employeeSchedule = null;
  let toleranceMins = 0;

  for (let i = 0; i < hData.length; i++) {
    if (String(hData[i][1]).trim() === String(p.cedula).trim()) {
      employeeSchedule = hData[i][2];
      toleranceMins = parseInt(hData[i][3]) || 0;
      break;
    }
  }

  if (!employeeSchedule) return { success:false, error:'Horario no asignado aún' };

  const daysEs = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
  const now = new Date(new Date().toLocaleString("en-US", {timeZone: TIMEZONE}));
  const dayName = daysEs[now.getDay()];
  
  const dayRegex = new RegExp(dayName + " \\(([^)]+)\\) - \\(([^)]+)\\)", "i");
  const match = employeeSchedule.match(dayRegex);

  if (!match) return { success:false, error:'Hoy es tu día libre según tu horario' };

  const schEntryStr = match[1]; 
  const schExitStr = match[2];  

  const to24h = (str) => {
    let [time, ampm] = str.split(' ');
    let [h, m] = time.split(':').map(Number);
    if (ampm === 'pm' && h < 12) h += 12;
    if (ampm === 'am' && h === 12) h = 0;
    return h * 60 + m;
  };

  const schEntryMins = to24h(schEntryStr);
  const schExitMins = to24h(schExitStr);
  const exitToleranceTotal = 480;
  
  const nowH = now.getHours();
  const nowM = now.getMinutes();
  const nowTotalMins = nowH * 60 + nowM;

  // Final Server side validation
  if (p.tipo === 'ENTRADA') {
    if (nowTotalMins < (schEntryMins - toleranceMins)) {
      return { success:false, error: 'Aún no es hora de entrar.' };
    }
    if (nowTotalMins > schExitMins) {
      return { success:false, error: 'Tu jornada ya terminó.' };
    }
  } else if (p.tipo === 'SALIDA') {
    if (nowTotalMins < schEntryMins) {
      return { success:false, error: 'Aún no has empezado tu jornada.' };
    }
    if (nowTotalMins > (schExitMins + exitToleranceTotal)) {
      return { success:false, error: 'El tiempo límite de salida ha expirado.' };
    }
  }

  // VALIDACIÓN: 1 sola entrada y 1 sola salida por día
  const existing = getRecords(p).records;
  const alreadyHas = existing.some(r => r.fecha === p.fecha && r.tipo === p.tipo);
  if (alreadyHas) {
    return { success:false, error: 'Ya registraste tu ' + p.tipo.toLowerCase() + ' el día de hoy.' };
  }

  // ISO-8601 Week Number (Starts on Monday)
  const getWeekNumber = (d) => {
    const date = new Date(d.getTime());
    date.setHours(0, 0, 0, 0);
    // Jueves de la semana actual decide el año
    date.setDate(date.getDate() + 3 - (date.getDay() + 6) % 7);
    // 4 de Enero siempre está en la semana 1
    const week1 = new Date(date.getFullYear(), 0, 4);
    // Ajustar al jueves de la semana 1
    week1.setDate(week1.getDate() + 3 - (week1.getDay() + 6) % 7);
    return 1 + Math.round(((date.getTime() - week1.getTime()) / 86400000) / 7);
  };
  const weekNum = getWeekNumber(now);

  // Column mapping:
  // F(6) Fecha, G(7) Semana, H(8) Nombre, I(9) Cedula, J(10) Area, K(11) Tipo, L(12) Hora
  // M(13) vacío, N(14) MinTarde, O(15) Mensaje
  // P(16) Justif entrada tardía, Q(17) vacío, R(18) Justif salida anticipada, S(19) Justif salida tardía

  const msg = (p.mensajeSistema || '').toLowerCase();
  const justifP = (p.tipo === 'ENTRADA' && /tard/i.test(msg))       ? (p.justificativo || '') : '';
  const justifR = (p.tipo === 'SALIDA'  && /anticipada/i.test(msg)) ? (p.justificativo || '') : '';
  const justifS = (p.tipo === 'SALIDA'  && /fuera/i.test(msg))      ? (p.justificativo || '') : '';

  const recordData = [
    p.fecha,               // F (6)
    "Semana " + weekNum,   // G (7)
    p.nombre,              // H (8)
    p.cedula,              // I (9)
    p.area,                // J (10)
    p.tipo,                // K (11)
    p.hora,                // L (12)
    "",                    // M (13)
    p.minTarde || "",      // N (14)
    p.mensajeSistema,      // O (15)
    justifP,               // P (16) - justif entrada tardía
    "",                    // Q (17) - reservado
    justifR,               // R (18) - justif salida anticipada
    justifS                // S (19) - justif salida tardía
  ];

  // Anti-duplicate check (prevent 2 clicks within 10 seconds)
  // We use try-catch here to ensure any formatting error doesn't break the registration
  try {
    const lastRecord = hSheet.getRange("F2:L2").getValues()[0];
    if (lastRecord[0]) {
      const lastCedula = String(lastRecord[3]).trim(); // Col I
      const lastTipo = String(lastRecord[5]).trim();   // Col K
      const lastHora = lastRecord[6];                  // Col L (could be Date or String)

      const toSeconds = (val) => {
        if (!val) return 0;
        if (val instanceof Date) {
          return val.getHours() * 3600 + val.getMinutes() * 60 + val.getSeconds();
        }
        if (typeof val === 'string' && val.includes(':')) {
          const parts = val.split(':');
          return (+parts[0]) * 3600 + (+parts[1]) * 60 + (+parts[2] || 0);
        }
        return 0;
      };

      const nowSecs = toSeconds(p.hora);
      const lastSecs = toSeconds(lastHora);

      if (lastCedula === String(p.cedula).trim() && 
          lastTipo === p.tipo && 
          Math.abs(nowSecs - lastSecs) < 10) {
        return { success:true, warning: 'Duplicate ignored' };
      }
    }
  } catch (err) {
    console.error("Error in duplicate check:", err);
    // Continue anyway to avoid blocking the user
  }

  // Insertar celdas en F2:S2 preservando A-E y T+ (2FA en T2)
  // Pero necesitamos proteger Q2 (el código 2FA) para que no se ruede hacia abajo
  const oldCode2FA = hSheet.getRange("Q2").getValue();
  
  hSheet.getRange("F2:S2").insertCells(SpreadsheetApp.Dimension.ROWS);
  hSheet.getRange(2, 6, 1, recordData.length).setValues([recordData]);
  
  // Restauramos el código en Q2 para que no se haya "bajado" a Q3
  hSheet.getRange("Q2").setValue(oldCode2FA);
  
  SpreadsheetApp.flush(); 

  return { success:true };
}
 
function getRecords(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheet(ss, HORARIOS_SHEET);
  if (!sheet) return { success:true, records:[] };

  const data = sheet.getRange("F2:S2000").getValues(); // F-S range
  const records = [];
  const tz = Session.getScriptTimeZone();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (String(row[3]).trim() === String(p.cedula).trim()) {
      let fechaStr = row[0];
      if (row[0] instanceof Date) {
        fechaStr = Utilities.formatDate(row[0], tz, "dd/MM/yyyy");
      }

      let horaStr = row[6]; // Column L (index 6 in F-S)
      if (row[6] instanceof Date) {
        horaStr = Utilities.formatDate(row[6], tz, "HH:mm:ss");
      }

      records.push({
        fecha: fechaStr,
        semana: row[1],
        nombre: row[2],
        cedula: row[3],
        area: row[4],
        tipo: row[5],
        hora: horaStr,
        minTarde: row[8],
        mensajeSistema: row[9],
        justifEntrada:           row[10], // P
        justifSalidaAnticipada:  row[12], // R
        justifSalidaTardia:      row[13]  // S
      });
    }
  }
  return { success:true, records };
}

function verify2FA(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheet(ss, HORARIOS_SHEET);
  if (!sheet) return { success:false, error:'Hoja no encontrada' };
  
  const realCode = sheet.getRange('Q2').getValue().toString().trim();
  const userCode = (p.code || "").toString().trim();
  
  if (userCode === realCode) {
    return { success:true };
  }
  return { success:false, error:'Código de verificación incorrecto' };
}

function generarCodigo() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheet(ss, HORARIOS_SHEET);
  if (!sheet) return;

  const code = Math.floor(10000 + Math.random() * 90000).toString();
  sheet.getRange('Q2').setValue(code);
}
