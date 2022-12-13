const office = require('./index');

let rows = 10;
async function run2(){
    let wb = office.spreadsheet.open("/tmp/unioffice.xlsx");
    let ws = wb.sheet('Sheet 1');
    let cols = "ABCDEFGHIJKLMNOPQRSTUV".split("");
    for(let i=0; i<rows; i++){
        for(let c of cols){
            let a = `${c}${i+1}`
            let cell = ws.cell(a);
            cell.setString(a);
        }
    }
    // console.time("save");
    wb.save("/tmp/unioffice-2.xlsx");
    //wb.close();
    // console.timeEnd("save");
    check_mem();
}

function check_mem(){
    const used = process.memoryUsage().rss / 1024 / 1024;
    console.log(`The script uses approximately ${Math.round(used * 100) / 100} MB`);
}

for(let i=0; i<10; i++){
    run2();
}