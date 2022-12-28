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
        if (typeof name == 'number'){
            let n = spreadsheet.sheet_name(this.wb, name);
            this._active_sheet = n;
            return new Sheet(this, n);
        }else{
            spreadsheet.check_sheet(this.wb, name);
            this._active_sheet = name;
            return new Sheet(this, name);
        }
    }
    save(file, opts={}){
        if (opts.export_pdf){
            spreadsheet.save_pdf(this.wb, this._active_sheet, file);
        }else{
            spreadsheet.save(this.wb, file);
        }
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
        this.name = s;
    }
    addRow(){
        let r = spreadsheet.add_row(this.wb, this.name);
        return new Row(this, r);
    }
    addRows(c){
        if (isNaN(c)){
            throw new Error("Invalid params");
        }
        let r = spreadsheet.add_rows(this.wb, this.name, c);
        return c;
    }
    insertRows(rownum, rows){
        if (isNaN(rownum) || isNaN(rows)){
            throw new Error("Invalid args");
        }
        return spreadsheet.insert_rows(this.wb, this.name, rownum, rows);
    }
    copyRows(source, dest, rows){
        if (isNaN(source) || isNaN(rows) || isNaN(dest)){
            throw new Error("Invalid args");
        }
        return spreadsheet.copy_rows(this.wb, this.name, source, dest, rows);
    }
    recalculateFormulas(){
        return spreadsheet.sheet_recalculate_formulas(this.wb, this.name);
    }
    cell(ref){
        if (!ref){
            throw new Error("invalid cell ref")
        }
        return new Cell(this, ref);
    }
    getRowValues(row){
        return spreadsheet.sheet_get_row_values(this.wb, this.name, row);
    }
    maxColumnIndex(){
        //zero-based
        return spreadsheet.sheet_max_column(this.wb, this.name);
    }
    maxRowIndex(){
        //1-based
        return spreadsheet.sheet_max_row(this.wb, this.name);
    }
}

class Row extends Base{
    constructor(ws, r){
        super(ws.wb);
        this.sheet = ws;
        this._number = r;
    }
    addCell(){
        let c = spreadsheet.add_cell(this.wb, this.sheet.name, this._number);
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
        if (val == null || val == undefined){
            val = "";
        }else if (typeof val != 'string'){
            val = val.toString()
        }
        return spreadsheet.set_cell_string(this.wb, this.sheet.name, this._address, val);
    }
    setBool(val){
        return spreadsheet.set_cell_bool(this.wb, this.sheet.name, this._address, val);
    }
    setDate(val){
        return spreadsheet.set_cell_date(this.wb, this.sheet.name, this._address, val);
    }
    setFormula(val){
        if (!val || typeof val != "string"){
            throw new Error("invalid arg")
        }
        return spreadsheet.set_cell_formula_raw(this.wb, this.sheet.name, this._address, val);
    }
    setFormulaArray(val){
        if (!val || typeof val != "string"){
            throw new Error("invalid arg")
        }
        return spreadsheet.set_cell_formula_array(this.wb, this.sheet.name, this._address, val);
    }
    setFormulaShared(val, rows, cols){
        if (!val || typeof val != "string"){
            throw new Error("invalid arg")
        }
        if (isNaN(rows) || isNaN(cols)){
            throw new Error("invalid arg")
        }
        return spreadsheet.set_cell_formula_shared(this.wb, this.sheet.name, this._address, val, rows, cols);
    }
    setNumber(val){
        if (isNaN(val)){
            throw new Error("invalid arg");
        }
        return spreadsheet.set_cell_number(this.wb, this.sheet.name, this._address, val);
    }
    setValue(val){
        if (val ==null || val == undefined){
            val = "";
        }
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
        let val = spreadsheet.cell_get_value(this.wb, this.sheet.name, this._address, 0);
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
        return spreadsheet.cell_get_value(this.wb, this.sheet.name, this._address, 1);
    }
    getBool(){
        return spreadsheet.cell_get_value(this.wb, this.sheet.name, this._address, 4) == 1;
    }
    getNumber(){
        return spreadsheet.cell_get_value(this.wb, this.sheet.name, this._address, 2);
    }
    getDate(){
        return spreadsheet.cell_get_value(this.wb, this.sheet.name, this._address, 3);
    }

    getRef(){
        return utils.parseCellRef(this._address);
    }
}

const utils = {
    parseCellRef(ref){
        let idx = 0;
        for(let i=0; i<ref.length; i++){
            if (ref.charCodeAt(i) < 65){
                idx = i;
                break;
            }
        }
        let col = ref.substring(0, idx);
        let row = parseInt(ref.substring(idx));
        return {
            col: col,
            colIndex: utils.columnToIndex(col),
            row
        }
    },
    columnToIndex(col){
        let val = 0;
        for(let i=0; i<col.length; i++){
            val *= 26;
            val += (col.charCodeAt(i) - 64);
        }
        return val - 1;
    },
    indexToColumn(index){
        let digits = new Array(65);
        let i = 65;
        let div = index;
        const base = 26;
        while (div >= base) {
            i--;
            let _div = Math.floor(div / base);
            digits[i] = 65 + (div % base);
            div = _div - 1;
        }
        i--;
        digits[i] = 65 + div;
        digits = digits.slice(i);
        return String.fromCharCode(...digits);
    },
    indexToReference(col, row){
        return utils.indexToColumn(col) + row;
    }
}
const fs = require('fs');
module.exports = {
    new: function(){
        let wb = spreadsheet.new();
        return new Workbook(wb);
    },
    open: function(file){
        if (!fs.existsSync(file)){
            throw new Error(`Open ${file}: no such file`);
        }
        let wb = spreadsheet.open(file);
        return new Workbook(wb);
    },
    utils,
}