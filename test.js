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

    ws.cell('B4').setBool(true)
    ws.cell('B5').setBool(false)
    ws.cell('B6').setBool(1)
    ws.cell('B7').setBool(0)

    wb.save("/tmp/unioffice/test3.xlsx");
    check_mem();
}


function check_mem(){
    const used = process.memoryUsage().rss / 1024 / 1024;
    console.log(`The script uses approximately ${Math.round(used * 100) / 100} MB`);
}

set_value();
//process.exit(0)

for(let i=0; i<10; i++){
    //run2();
}