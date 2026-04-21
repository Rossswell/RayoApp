const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const { autoUpdater } = require('electron-updater');
const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');

const credsPath = app.isPackaged
    ? path.join(process.resourcesPath, 'credenciales.json')
    : path.join(__dirname, 'credenciales.json');
const creds = require(credsPath);

let mainWindow;

function createMainWindow() {
    mainWindow = new BrowserWindow({
        width: 1200,
        height: 800,
        frame: false, // Keeping it premium frameless
        show: false,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        }
    });

    mainWindow.loadFile('index.html');
    mainWindow.center();

    // Smoothly show when ready
    mainWindow.once('ready-to-show', () => {
        mainWindow.show();
    });

    mainWindow.setMenuBarVisibility(false);
    mainWindow.on('closed', () => (mainWindow = null));
}

app.on('ready', () => {
    createMainWindow();
    setupAutoUpdater();
});

function setupAutoUpdater() {
    autoUpdater.autoDownload = true;
    autoUpdater.autoInstallOnAppQuit = true;

    autoUpdater.on('update-available', (info) => {
        mainWindow.webContents.send('update-available', info.version);
    });

    autoUpdater.on('update-downloaded', () => {
        mainWindow.webContents.send('update-downloaded');
    });

    autoUpdater.on('error', (err) => {
        console.error('AutoUpdater error:', err.message);
    });

    autoUpdater.checkForUpdatesAndNotify();
}

ipcMain.on('install-update', () => {
    autoUpdater.quitAndInstall();
});

// Handle close/minimize for the custom frameless window
ipcMain.on('app-close', () => app.quit());
ipcMain.on('app-minimize', () => mainWindow.minimize());

// --- GOOGLE SHEETS INTEGRATION ---

const SPREADSHEET_ID = '1O35ZWq2yEAW-5FOJsby2Kszo6WP3_0nqVGp4yd0q4Lk';
const SHEET_NAME = '👤EMPLEADOS FLASH HIGH';

const serviceAccountAuth = new JWT({
    email: creds.client_email,
    key: creds.private_key,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const doc = new GoogleSpreadsheet(SPREADSHEET_ID, serviceAccountAuth);

ipcMain.handle('add-employee-to-sheet', async (event, employeeData) => {
    try {
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle[SHEET_NAME];

        if (!sheet) {
            throw new Error(`No se encontró la hoja con el nombre: ${SHEET_NAME}`);
        }

        // Buscar primera fila vacía en el rango A3:H30 (Limite solicitado por el usuario)
        const START_ROW = 3;
        const END_ROW = 30;
        await sheet.loadCells(`A${START_ROW}:H${END_ROW}`);
        
        let targetRowIndex = -1;
        // El offset en loadCells considera la fila START_ROW como índice relativo 0 si no se cargó toda la hoja
        // Pero doc.sheetsByTitle[SHEET_NAME].getCell(absoluteRow, absoluteCol) usa índices absolutos (0-indexed)
        for (let r = START_ROW - 1; r < END_ROW; r++) {
            const cellA = sheet.getCell(r, 0); // Columna A
            if (!cellA.value) {
                targetRowIndex = r;
                break;
            }
        }

        if (targetRowIndex === -1) {
            throw new Error(`No se encontraron filas vacías disponibles en el rango A${START_ROW}:H${END_ROW}.`);
        }

        // Mapeo de datos a las celdas de la fila encontrada (A-H)
        const dataArray = [
            employeeData.nombre,
            employeeData.cedula,
            employeeData.telefono,
            employeeData.correo,
            employeeData.area,
            employeeData.fecha_inicio,
            employeeData.fecha_cumple,
            employeeData.ubicacion || ''
        ];

        for (let col = 0; col < dataArray.length; col++) {
            const cell = sheet.getCell(targetRowIndex, col);
            cell.value = dataArray[col];
        }

        await sheet.saveUpdatedCells();
        console.log(`✅ Empleado guardado exitosamente en la fila ${targetRowIndex + 1}:`, employeeData.nombre);
        return { success: true };
    } catch (error) {
        console.error('❌ Error al guardar en Google Sheets:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('get-employees', async () => {
    try {
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle[SHEET_NAME];

        if (!sheet) {
            throw new Error(`No se encontró la hoja con el nombre: ${SHEET_NAME}`);
        }

        const START_ROW = 3;
        const END_ROW = 30;
        await sheet.loadCells(`A${START_ROW}:H${END_ROW}`);

        const employees = [];
        for (let r = START_ROW - 1; r < END_ROW; r++) {
            const rowData = [];
            const cellA = sheet.getCell(r, 0);
            
            // Si la columna A está vacía, ignoramos la fila
            if (!cellA.value) continue;

            for (let c = 0; c < 8; c++) {
                rowData.push(sheet.getCell(r, c).formattedValue || '');
            }
            
            employees.push({
                nombre: rowData[0],
                cedula: rowData[1],
                telefono: rowData[2],
                correo: rowData[3],
                area: rowData[4],
                fecha_inicio: rowData[5],
                fecha_cumple: rowData[6],
                ubicacion: rowData[7],
                rowIndex: r + 1 // Guardar el número de fila para poder eliminar
            });
        }

        // Devolver en orden inverso (más recientes primero)
        return { success: true, employees: employees.reverse() };
    } catch (error) {
        console.error('❌ Error al obtener empleados:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('delete-employee', async (event, { cedula }) => {
    try {
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle[SHEET_NAME];

        if (!sheet) {
            throw new Error(`No se encontró la hoja con el nombre: ${SHEET_NAME}`);
        }

        const START_ROW = 3;
        const END_ROW = 30;
        await sheet.loadCells(`A${START_ROW}:H${END_ROW}`);

        // Buscar la fila por cédula
        let targetRowIndex = -1;
        for (let r = START_ROW - 1; r < END_ROW; r++) {
            const cellB = sheet.getCell(r, 1); // Columna B = Cédula
            if (cellB.value && String(cellB.value).trim() === String(cedula).trim()) {
                targetRowIndex = r;
                break;
            }
        }

        if (targetRowIndex === -1) {
            throw new Error(`No se encontró un empleado con la cédula: ${cedula}`);
        }

        // Limpiar todas las celdas de la fila (columnas A-H)
        for (let col = 0; col < 8; col++) {
            const cell = sheet.getCell(targetRowIndex, col);
            cell.value = '';
        }

        await sheet.saveUpdatedCells();
        console.log(`✅ Empleado con cédula ${cedula} eliminado exitosamente de la fila ${targetRowIndex + 1}`);
        return { success: true };
    } catch (error) {
        console.error('❌ Error al eliminar empleado:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('update-employee', async (event, { originalCedula, employeeData }) => {
    try {
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle[SHEET_NAME];

        if (!sheet) {
            throw new Error(`No se encontró la hoja con el nombre: ${SHEET_NAME}`);
        }

        const START_ROW = 3;
        const END_ROW = 30;
        await sheet.loadCells(`A${START_ROW}:H${END_ROW}`);

        // Buscar la fila por la cédula original
        let targetRowIndex = -1;
        for (let r = START_ROW - 1; r < END_ROW; r++) {
            const cellB = sheet.getCell(r, 1); // Columna B = Cédula
            if (cellB.value && String(cellB.value).trim() === String(originalCedula).trim()) {
                targetRowIndex = r;
                break;
            }
        }

        if (targetRowIndex === -1) {
            throw new Error(`No se encontró un empleado con la cédula: ${originalCedula}`);
        }

        // Actualizar todas las celdas de la fila (columnas A-H)
        const dataArray = [
            employeeData.nombre,
            employeeData.cedula,
            employeeData.telefono,
            employeeData.correo,
            employeeData.area,
            employeeData.fecha_inicio,
            employeeData.fecha_cumple,
            employeeData.ubicacion || ''
        ];

        for (let col = 0; col < dataArray.length; col++) {
            const cell = sheet.getCell(targetRowIndex, col);
            cell.value = dataArray[col];
        }

        await sheet.saveUpdatedCells();
        console.log(`✅ Empleado actualizado exitosamente en la fila ${targetRowIndex + 1}:`, employeeData.nombre);
        return { success: true };
    } catch (error) {
        console.error('❌ Error al actualizar empleado:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('save-employee-schedule', async (event, { employee, scheduleSummary, tolerance }) => {
    try {
        console.log(`💾 Guardando horario para: ${employee.nombre} (${employee.cedula})`);
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle['⌛HORARIOS'];
        if (!sheet) throw new Error('No se encontró la hoja ⌛HORARIOS');

        // Load header and data to find columns by index to avoid duplicate header error
        const START_ROW = 1; // Start from row 1 (0-indexed) to skip headers if any, or row 0 if we want headers
        const END_ROW = 100;
        await sheet.loadCells(`A1:D${END_ROW}`);

        let targetRowIndex = -1;
        for (let r = 1; r < END_ROW; r++) {
            const cellCedula = sheet.getCell(r, 1); // Col B
            if (cellCedula.value && String(cellCedula.value).trim() === String(employee.cedula).trim()) {
                targetRowIndex = r;
                break;
            }
            if (!cellCedula.value && targetRowIndex === -1) {
                // Keep track of first empty row as backup if not found
            }
        }

        if (targetRowIndex !== -1) {
            console.log(`📝 Actualizando fila existente para ${employee.cedula} en fila ${targetRowIndex + 1}`);
            sheet.getCell(targetRowIndex, 2).value = scheduleSummary; // Col C
            sheet.getCell(targetRowIndex, 3).value = tolerance;       // Col D
        } else {
            // Find first empty row in Col B
            for (let r = 1; r < END_ROW; r++) {
                if (!sheet.getCell(r, 1).value) {
                    targetRowIndex = r;
                    break;
                }
            }
            if (targetRowIndex === -1) throw new Error('No hay espacio en la hoja de horarios.');
            
            console.log(`➕ Añadiendo nueva fila para ${employee.cedula} en fila ${targetRowIndex + 1}`);
            sheet.getCell(targetRowIndex, 0).value = employee.nombre;
            sheet.getCell(targetRowIndex, 1).value = employee.cedula;
            sheet.getCell(targetRowIndex, 2).value = scheduleSummary;
            sheet.getCell(targetRowIndex, 3).value = tolerance;
        }

        await sheet.saveUpdatedCells();
        console.log('✅ Horario guardado con éxito');
        return { success: true };
    } catch (error) {
        console.error('❌ Error al guardar horario:', error);
        return { success: false, error: error.message };
    }
});

function parseScheduleSummary(summary) {
    if (!summary || summary === 'No asignado') return null;
    
    const days = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo'];
    const result = {};
    
    // Inicializar todos los días como inactivos
    days.forEach(d => result[d] = { activo: false });
    
    // Dividir por punto y coma (o coma, por si acaso)
    const parts = summary.split(/[;]/);
    parts.forEach(part => {
        // Regex para capturar: Día (Entrada) - (Salida)
        // Ejemplo: "Lunes (8:00 am) - (5:00 pm)"
        const match = part.match(/(Lunes|Martes|Miércoles|Jueves|Viernes|Sábado|Domingo)\s*\((.*?)\)\s*-\s*\((.*?)\)/i);
        if (match) {
            let dayName = match[1].toLowerCase();
            // Normalizar acentos
            dayName = dayName.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
            
            const entrada = match[2].trim();
            const salida = match[3].trim();
            
            if (days.includes(dayName)) {
                result[dayName] = { activo: true, entrada, salida };
            }
        }
    });
    
    return result;
}

ipcMain.handle('get-employee-schedule', async (event, { cedula }) => {
    try {
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle['⌛HORARIOS'];
        if (!sheet) return { success: false, error: 'Hoja no encontrada' };

        const END_ROW = 100;
        await sheet.loadCells(`A1:D${END_ROW}`);

        for (let r = 1; r < END_ROW; r++) {
            const cellCedula = sheet.getCell(r, 1); // Col B
            if (cellCedula.value && String(cellCedula.value).trim() === String(cedula).trim()) {
                const summary = sheet.getCell(r, 2).value;
                const tolerance = sheet.getCell(r, 3).value;
                
                return { 
                    success: true, 
                    schedule: parseScheduleSummary(summary),
                    scheduleSummary: summary, 
                    tolerance: tolerance 
                };
            }
        }

        return { success: true, schedule: null };
    } catch (error) {
        console.error('Error al obtener horario:', error);
        return { success: false, error: error.message };
    }
});

// Fetch analytics data for Dashboard
ipcMain.handle('get-attendance-stats', async () => {
    try {
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle['⌛HORARIOS'];
        if (!sheet) throw new Error('Hoja no encontrada');

        const endRow = Math.min(sheet.rowCount, 5000); // Fetch up to 5000 rows
        await sheet.loadCells(`F1:S${endRow}`);

        const records = [];
        for (let r = 1; r < endRow; r++) {
            const fechaVal = sheet.getCell(r, 5).value; // Col F
            if (!fechaVal) continue;

            records.push({
                fecha: fechaVal,
                semana: sheet.getCell(r, 6).value,
                empleado: sheet.getCell(r, 7).value,
                cedula: sheet.getCell(r, 8).value,
                area: sheet.getCell(r, 9).value,
                tipo: sheet.getCell(r, 10).value,
                hora: sheet.getCell(r, 11).value,
                tarde: sheet.getCell(r, 12).value,
                minTarde: sheet.getCell(r, 13).value,
                mensaje: sheet.getCell(r, 14).value,
                justifEntrada:          sheet.getCell(r, 15).value, // P
                // col 16 = Q (reserved)
                justifSalidaAnticipada: sheet.getCell(r, 17).value, // R
                justifSalidaTardia:     sheet.getCell(r, 18).value  // S
            });
        }

        return { success: true, records };
    } catch (error) {
        console.error('❌ Error al obtener estadísticas:', error);
        return { success: false, error: error.message };
    }
});

// --- NÓMINA (PAYROLL) INTEGRATION ---

const NOMINA_SHEET_NAME = '👷NOMINA';

ipcMain.handle('get-payroll-data', async () => {
    try {
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle[NOMINA_SHEET_NAME];
        if (!sheet) throw new Error(`Hoja no encontrada: ${NOMINA_SHEET_NAME}`);

        const START_ROW = 3;
        const END_ROW = 100;
        await sheet.loadCells(`A${START_ROW}:J${END_ROW}`);

        const records = [];
        for (let r = START_ROW - 1; r < END_ROW; r++) {
            const cellA = sheet.getCell(r, 0); // Nombre
            if (!cellA.value) continue;

            records.push({
                nombre: cellA.formattedValue || '',
                cedula: sheet.getCell(r, 1).formattedValue || '',
                departamento: sheet.getCell(r, 2).formattedValue || '',
                pordia: sheet.getCell(r, 3).formattedValue || '0',
                dias: sheet.getCell(r, 4).formattedValue || '0',
                semana: sheet.getCell(r, 5).formattedValue || '0',
                mes: sheet.getCell(r, 6).formattedValue || '0',
                trimestral: sheet.getCell(r, 7).formattedValue || '0',
                semestral: sheet.getCell(r, 8).formattedValue || '0',
                anual: sheet.getCell(r, 9).formattedValue || '0',
                rowIndex: r + 1
            });
        }

        return { success: true, records: records.reverse() }; // newest first
    } catch (error) {
        console.error('❌ Error al obtener nómina:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('save-payroll-data', async (event, payrollData) => {
    try {
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle[NOMINA_SHEET_NAME];
        if (!sheet) throw new Error(`Hoja no encontrada: ${NOMINA_SHEET_NAME}`);

        const START_ROW = 3;
        const END_ROW = 100;
        // Only load A:E to prevent overwriting formulas in F:J
        await sheet.loadCells(`A${START_ROW}:E${END_ROW}`);

        let targetRowIndex = -1;
        for (let r = START_ROW - 1; r < END_ROW; r++) {
            const cellA = sheet.getCell(r, 0);
            if (!cellA.value) {
                targetRowIndex = r;
                break;
            }
        }

        if (targetRowIndex === -1) {
            throw new Error(`No hay espacio en la hoja ${NOMINA_SHEET_NAME}.`);
        }

        // Set values A-E (Indices 0-4)
        sheet.getCell(targetRowIndex, 0).value = payrollData.nombre;
        sheet.getCell(targetRowIndex, 1).value = payrollData.cedula;
        sheet.getCell(targetRowIndex, 2).value = payrollData.departamento;
        // Parse numbers specifically for D and E so they work properly in Google Sheets formulas
        sheet.getCell(targetRowIndex, 3).value = parseFloat(payrollData.pordia) || 0;
        sheet.getCell(targetRowIndex, 4).value = parseInt(payrollData.dias) || 0;

        await sheet.saveUpdatedCells();
        console.log(`✅ Nómina guardada exitosamente en la fila ${targetRowIndex + 1}:`, payrollData.nombre);
        return { success: true };
    } catch (error) {
        console.error('❌ Error al guardar nómina:', error);
        return { success: false, error: error.message };
    }
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});
