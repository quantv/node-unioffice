const spreadsheet = require('bindings')('unispreadsheet');

class Base{
    constructor(wb){
        Object.defineProperty(this, 'wb', { value: wb, writable: false });
    }
}

class Workbook extends Base{
    constructor(wb){
        super(wb);
    }
    addSheet(){
        let s = spreadsheet.add_sheet(this.wb);
        return new Sheet(this, s);
    }
    save(file){
        spreadsheet.save(this.wb, file);
    }
    close(){
        spreadsheet.close(this.wb);
    }
    test(sheet, value, count){
        spreadsheet.test(this.wb, sheet, value, count);
    }
}

class Sheet extends Base{
    constructor(wb, s){
        super(wb.wb);
        this._number = s;
    }
    addRow(){
        let r = spreadsheet.add_row(this.wb, this._number);
        return new Row(this, r);
    }
    addRows(c){
        let r = spreadsheet.add_rows(this.wb, this._number, c);
        return new Row(this, r);
    }
    cell(ref){
        return new Cell(this, ref);
    }
}

class Row extends Base{
    constructor(ws, r){
        super(ws.wb);
        this.sheet = ws;
        this._number = r;
    }
    addCell(){
        let c = spreadsheet.add_cell(this.wb, this.sheet._number, this._number);
        return new Cell(this.sheet, c);
    }
}

class Cell extends Base{
    constructor(sheet, address){
        super(sheet.wb);
        this.sheet = sheet;
        this._address = address;
    }
    setString(val){
        return spreadsheet.set_cell_string(this.wb, this.sheet._number, this._address, val);
    }
}


module.exports = {
    new: function(){
        let wb = spreadsheet.new();
        return new Workbook(wb);
    },
    open: function(file){
        let wb = spreadsheet.open(file);
        return new Workbook(wb);
    }
}