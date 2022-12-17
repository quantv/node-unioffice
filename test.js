const office = require('./index');

let rows = 100;
async function run2(){
    let wb = office.spreadsheet.open("./test/test2.xlsx");
    let ws = wb.sheet('MOMO');
    //ws.addRows(rows);
    ws.insertRows(3, rows);
    ws.copyRows(2, 3, rows);
    let cols = "ABCDEFGHIJKLMNOPQRSTUV".split("");
    for(let i=0; i<rows; i++){
        for(let c of cols){
            let a = `${c}${i+1}`
            let cell = ws.cell(a);
            //cell.setString(a);
        }
    }
    // console.time("save");
    wb.save("/tmp/unioffice/test3.xlsx");
    //wb.close();
    // console.timeEnd("save");
    check_mem();
}

async function set_value(){
    let wb = office.spreadsheet.open("./test/test2.xlsx");
    let ws = wb.sheet('MOMO');

    ws.cell("A1").setBool(true);
    ws.cell("A2").setString("abc");
    ws.cell("A3").setNumber(1234)
    ws.cell("A4").setNumber(1234.1222)
    ws.cell("A5").setFormula("A3+A4")
    
    ws.cell("A6").setFormulaShared("A3+1", 2, 3);
    
    ws.recalculateFormulas();

    wb.save("/tmp/unioffice/set_value.xlsx");
}

async function get_value(){
    let wb = office.spreadsheet.open("/tmp/unioffice/date.xlsx");
    let ws = wb.sheet('MOMO');

    let cols = "ABCDEFGH".split("")
    for(let i=1; i<10; i++){
        for(let col of cols){
            let cellName = `${col}${i}`;
            let v = ws.cell(cellName).getValue();
            console.log(cellName, v);
        }
    }
}

function check_mem(){
    const used = process.memoryUsage().rss / 1024 / 1024;
    console.log(`The script uses approximately ${Math.round(used * 100) / 100} MB`);
}

set_value();
//get_value();
//process.exit(0)

for(let i=0; i<10; i++){
    //run2();
}