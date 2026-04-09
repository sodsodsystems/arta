const XLSX = require('xlsx');
const path = require('path');

const filePath = path.join('c:', 'Users', 'admin', 'Documents', 'GIT PROJECTS', 'ARTA', 'SRS HR COPY.xlsx');

try {
    const workbook = XLSX.readFile(filePath);
    console.log('Sheet names:', workbook.SheetNames);
    workbook.SheetNames.forEach(sheetName => {
        console.log(`\n--- ${sheetName} ---`);
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 }).slice(0, 10);
        console.log(JSON.stringify(data, null, 2));
    });
} catch (error) {
    console.error('Error:', error.message);
}
