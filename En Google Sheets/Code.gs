const SPREADSHEET_ID = '1O35ZWq2yEAW-5FOJsby2Kszo6WP3_0nqVGp4yd0q4Lk';
const EMP_SHEET      = '👤EMPLEADOS FLASH HIGH';
const HORARIOS_SHEET = '⌛HORARIOS';
 
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
  const now = new Date();
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
  const exitToleranceTotal = 60; 
  
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

  // Calculate Week Number
  const getWeekNumber = (d) => {
    d = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay()||7));
    var yearStart = new Date(Date.UTC(d.getUTCFullYear(),0,1));
    return Math.ceil((((d - yearStart) / 86400000) + 1)/7);
  };
  const weekNum = getWeekNumber(now);

  // Mapping as per LATEST USER requirements:
  // F: Fecha (6), G: Semana (7), H: Empleado (8), I: Cedula (9), J: Area (10), K: Tipo (11), L: Hora (12)
  // M: VACIO (13)
  // N: Minutos Tarde (14)
  // O: Mensaje (15)
  // P: Justificativo (16)

  const recordData = [
    p.fecha,               // F (6)
    "Semana " + weekNum,    // G (7)
    p.nombre,              // H (8)
    p.cedula,              // I (9)
    p.area,                // J (10)
    p.tipo,                // K (11)
    p.hora,                // L (12)
    "",                    // M (13)
    p.minTarde || "",      // N (14)
    p.mensajeSistema,      // O (15)
    p.justificativo || ""  // P (16)
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

  // Insertar celdas solo en el rango F2:P2 para preservar A-D intactas
  hSheet.getRange("F2:P2").insertCells(SpreadsheetApp.Dimension.ROWS);
  hSheet.getRange(2, 6, 1, recordData.length).setValues([recordData]);
  SpreadsheetApp.flush(); // Asegurar que se guarde antes de responder

  return { success:true };
}
 
function getRecords(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getSheet(ss, HORARIOS_SHEET);
  if (!sheet) return { success:true, records:[] };

  const data = sheet.getRange("F2:P2000").getValues(); // Read from row 2 upwards
  const records = [];
  const tz = Session.getScriptTimeZone();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (String(row[3]).trim() === String(p.cedula).trim()) { 
      // Column I is cedula (index 3 in F-P range)
      
      let fechaStr = row[0];
      if (row[0] instanceof Date) {
        fechaStr = Utilities.formatDate(row[0], tz, "dd/MM/yyyy");
      }

      let horaStr = row[6]; // Column L (Index 6 in F-P)
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
        justificativo: row[10]
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
