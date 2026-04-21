const { GoogleSpreadsheet } = require('google-spreadsheet');
const { JWT } = require('google-auth-library');
const fs = require('fs');
const path = require('path');

async function checkHeaders() {
    try {
        const creds = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'credenciales.json'), 'utf8'));
        
        const auth = new JWT({
            email: creds.client_email,
            key: creds.private_key,
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });

        const SPREADSHEET_ID = '1O35ZWq2yEAW-5FOJsby2Kszo6WP3_0nqVGp4yd0q4Lk';
        const doc = new GoogleSpreadsheet(SPREADSHEET_ID, auth);
        
        await doc.loadInfo();
        const sheet = doc.sheetsByTitle['⌛HORARIOS'];
        
        // Cargar celdas para leer cabeceras (F1 a P1) y algunos datos (F2 a P5)
        await sheet.loadCells('F1:P5');
        
        const headers = [];
        for (let col = 5; col <= 15; col++) {
            headers.push(sheet.getCell(0, col).value);
        }
        console.log('F:P Headers (indices 5-15):', headers);
        
        console.log('\nSample Row 2:');
        const row2 = [];
        for (let col = 5; col <= 15; col++) {
            row2.push(sheet.getCell(1, col).value);
        }
        console.log(row2);
        
    } catch (e) {
        console.error(e);
    }
}

checkHeaders();
