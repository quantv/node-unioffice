# Node Unioffice

Edit excel file using unioffice (https://github.com/unidoc/unioffice)

This package contains a very limited set of APIs that can be use for simple read/write excel file.

# Usage

Set environment variable to supply LICENSE file to unioffice: UNIOFFICE_LICENSE_PATH

NPM cli publish complains about Hard link: `415 Unsupported Media Type - PUT https://registry.npmjs.org/... - Hard link is not allowed`

I could not find the hard link or made it works. So, if you want to use the library:

```sh
npm install github:quantv/node-unioffice
```

Example:

```javascript
const office = require('unioffice');
let wb = office.spreadsheet.open("./filename.xlsx");
//let wb = office.spreadsheet.new();
let ws = wb.sheet('MOMO');
ws.addRows(10);
ws.cell('A1').setValue(123);
ws.cell('A2').setString("ABC");
wb.close();
```


# APIS

## Workbook

- save(filePath)
- close()
- sheet(sheetName)
- addSheet()

## Sheet

- insertRows(rowNumber, numberOfRows)
- addRows(numberOfRows)
- copyRows(source, dest, numberOfRows)
- recalculateFormulas()
- maxRowIndex(): 1-based index
- maxColumnIndex(): 0-based index.
- cell(cellReference)

## Cell

- setValue(val)
- setNumber(val)
- setString(val)
- setDate(val)
- getValue()
- getNumber()
- getString()
- getDate()
