const xlsx = require('xlsx');
const path = require('path');

const filePath = 'C:\\Users\\USER\\OneDrive\\Anti-Gravity\\고용관련 세액공제\\2024 (주)평우서비스(본점)_사원등록_20260312.xlsx';
const workbook = xlsx.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Read with cellDates: true to avoid serial numbers for dates
const data = xlsx.utils.sheet_to_json(sheet, { cellDates: false, raw: false }); 

console.log('--- SAMPLE ROW ---');
if (data.length > 0) {
    const firstRow = data[0];
    console.log(JSON.stringify(firstRow, null, 2));
    
    console.log('--- KEYS ---');
    console.log(Object.keys(firstRow));
}

// Find P column 
const rowsRaw = xlsx.utils.sheet_to_json(sheet, { header: 1 });
console.log('--- RAW ROW 0 (Column Labels) ---');
console.log(rowsRaw[0]);

const pIndex = 15; // P is 16th column
console.log('--- COLUMN P VALUES (Rows 1-5) ---');
for (let i = 1; i < Math.min(6, rowsRaw.length); i++) {
    console.log(`Row ${i}:`, rowsRaw[i][pIndex]);
}

// Check resignation date field for highlighting logic
console.log('--- RESIGNATION DATA CHECK ---');
data.forEach((emp, i) => {
    if (emp.퇴사일자) {
        console.log(`Row ${i} (${emp.사원명}): 퇴사일자 = '${emp.퇴사일자}' (Type: ${typeof emp.퇴사일자})`);
    }
});
