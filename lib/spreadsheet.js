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
    sheet(name){
        spreadsheet.check_sheet(this.wb, name);
        return new Sheet(this, name);
    }
    save(file){
        spreadsheet.save(this.wb, file);
    }
    savePdf(sheet, file){
        spreadsheet.save_pdf(this.wb, sheet, file);
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
    insertRows(rownum, rows){
        return spreadsheet.insert_rows(this.wb, this._number, rownum, rows);
    }
    copyRows(source, dest, rows){
        return spreadsheet.copy_rows(this.wb, this._number, source, dest, rows);
    }
    recalculateFormulas(){
        return spreadsheet.sheet_recalculate_formulas(this.wb, this._number);
    }
    cell(ref){
        return new Cell(this, ref);
    }
    getRowValues(row){
        return spreadsheet.sheet_get_row_values(this.wb, this._number, row);
    }
    maxColumnIndex(){
        //zero-based
        return spreadsheet.sheet_max_column(this.wb, this._number);
    }
    maxRowIndex(){
        //1-based
        return spreadsheet.sheet_max_row(this.wb, this._number);
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
    setBool(val){
        return spreadsheet.set_cell_bool(this.wb, this.sheet._number, this._address, val);
    }
    setDate(val){
        return spreadsheet.set_cell_date(this.wb, this.sheet._number, this._address, val);
    }
    setFormula(val){
        return spreadsheet.set_cell_formula_raw(this.wb, this.sheet._number, this._address, val);
    }
    setFormulaArray(val){
        return spreadsheet.set_cell_formula_array(this.wb, this.sheet._number, this._address, val);
    }
    setFormulaShared(val, rows, cols){
        return spreadsheet.set_cell_formula_shared(this.wb, this.sheet._number, this._address, val, rows, cols);
    }
    setNumber(val){
        return spreadsheet.set_cell_number(this.wb, this.sheet._number, this._address, val);
    }
    setValue(val){
        switch (typeof val) {
            case "number":
                return this.setNumber(val);
            case "boolean":
                return this.setBool(val);
            case "string":
                return this.setString(val);
            default:
                return this.setString(val.toString());
        }
    }
    getValue(){
        let val = spreadsheet.cell_get_value(this.wb, this.sheet._number, this._address, 0);
        function is_date(format){
            if (format == 'General') return false;
            format = format.replace(/"[^"]*"/g, "")
                           .replace(/\\./, "")
            return /[YMDHS]/.test(format);
        }
        //return val;
        switch(val.type){
            case 1: 
                return val.value == "1";
            case 2: 
                if (is_date(val.format)){
                    return new Date((parseFloat(val.value) - 25569) * 86400000);
                }
                return parseFloat(val.value);
            case 0:
                if (is_date(val.format)){
                    return new Date((parseFloat(val.value) - 25569) * 86400000);
                }
            case 4:
            default:
                return val.value;
        }
    }
    getString(){
        return spreadsheet.cell_get_value(this.wb, this.sheet._number, this._address, 1);
    }
    getBool(){
        return spreadsheet.cell_get_value(this.wb, this.sheet._number, this._address, 4) == 1;
    }
    getNumber(){
        return spreadsheet.cell_get_value(this.wb, this.sheet._number, this._address, 2);
    }
    getDate(){
        return spreadsheet.cell_get_value(this.wb, this.sheet._number, this._address, 3);
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