# Excel files unlocker

This script is able to unlock MS Excel 2007 or higher.

## Prerequisities

Node.js must be installed on the system and in the path.

## Usage

```
node unlock.js <path_to_xlsx_file>
```

## How does it work

XLSX files are in fact ZIP archives and individual worksheets are stored in the xl/worksheets folder as XML documents.

Worksheet protection is added simply by using the `<sheetProtection>` element. By removing the element, protection is removed too.

The script just unzips the XLSX file, goes through it's worksheets XML and removes the `<sheetProtection>` elements if present. Then it puts everything back to an archive, and outputs the unlocked XLSX file. Output filename is written to the console.
