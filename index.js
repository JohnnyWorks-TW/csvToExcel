const Excel = require('exceljs');
const fsPromises = require('fs').promises;
const fileExists = async (path) => !!(await fsPromises.stat(path).catch(e => false));

let csvFilename = 'test.csv';
let xlsFilename = 'test.xlsx';

if (process.argv && process.argv.length > 2) {
    csvFilename = process.argv[2];
    if (csvFilename.includes('.csv') === false) {
        csvFilename = csvFilename + '.csv';
    }
} else {
    console.log('Convert comma-separated values (csv) to excel (xlsx) tool.\n\nUsage: node index.js <csv filename>');
    process.exit(1);
}

xlsFilename = csvFilename.slice(0, -4) + '.xlsx';

(async () => {
    if (await fileExists(csvFilename) === false) {
        console.log('Please input a csv file');
        process.exit();
    }
    console.log('Converting \'' + csvFilename + '\' to \'' + xlsFilename + '\'...');
    const workbook = new Excel.Workbook();
    const worksheet = await workbook.csv.readFile(csvFilename);
    await workbook.xlsx.writeFile(xlsFilename);
    console.log('Complete!');
    process.exit();
})();