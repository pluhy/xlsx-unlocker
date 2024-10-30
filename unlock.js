const fs = require('fs');
const os = require('os');
const path = require('path');
const AdmZip = require('adm-zip');
const xml2js = require('xml2js');

const UNLOCKED_FILE_SUFFIX = '_unlocked';

// Check the filename, that should be passed as argument (e.g. "node unlock.js file.xlsx")
if (process.argv.length < 3) {
    console.error("Error: No filename provided");
    process.exit(1);
}

const inputFileName = process.argv[2];

if (path.extname(inputFileName) !== '.xlsx') {
    console.error("Error: File is not of type .xlsx");
    process.exit(1);
}

// Extract the zip file to a temporary directory
const zip = new AdmZip(inputFileName);
const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'unlocked-'));
zip.extractAllTo(tempDir, true);

const parser = new xml2js.Parser();
const builder = new xml2js.Builder();

// Remove the sheetProtection tag from all sheets
const worksheetsPath = path.join(tempDir, 'xl/worksheets');
fs.readdirSync(worksheetsPath).forEach((sheetFilename) => {
    if (sheetFilename.startsWith('sheet') && path.extname(sheetFilename) === '.xml') {
        const sheetPath = path.join(worksheetsPath, sheetFilename);
        const sheet = fs.readFileSync(sheetPath, 'utf8');
        parser.parseString(sheet, (err, result) => {
            if (err) {
                console.error(`Error parsing XML from worksheet file ${sheetFilename}: ${err}`);
                return;
            }
            if (Object.prototype.hasOwnProperty.call(result.worksheet, 'sheetProtection')) {
                delete result.worksheet.sheetProtection;
                let newXml = builder.buildObject(result);
                fs.writeFileSync(sheetPath, newXml, 'utf8');
            }
        });
    }
});

// Write the modified files back to a new zip file
const newZip = new AdmZip();
newZip.addLocalFolder(tempDir);
let outputFileName = path.join(path.dirname(inputFileName), path.basename(inputFileName, '.xlsx') + `${UNLOCKED_FILE_SUFFIX}.xlsx`);
newZip.writeZip(outputFileName);

console.log(`Successfully written ${outputFileName}`);
process.exit(0);