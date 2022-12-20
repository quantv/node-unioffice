# Node Unioffice

Edit excel file using unioffice (https://github.com/unidoc/unioffice)

This package contains a very limited set of APIs that can be use for simple read/write excel file.

# Usage

Set environment variable to supply LICENSE file to unioffice: UNIOFFICE_LICENSE_PATH

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