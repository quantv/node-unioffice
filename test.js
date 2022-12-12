const office = require('./index');

async function run(){
    console.log("run");
    let wb = office.spreadsheet.new();
    let ws = wb.addSheet();
    ws.addRows(10000);
    for(let i=0; i<10000; i++){
        let cell = ws.cell(`A${i+1}`);
        cell.setString("value");
    }
    console.time("save");
    wb.save("./file2.xlsx");
    console.timeEnd("save");
    // uncomment this and the application will not leak
    //await new Promise(setImmediate);
}

async function run2(){
    let wb = office.spreadsheet.new();
    let ws = wb.addSheet();
    ws.addRows(10000);
    for(let i=0; i<10000; i++){
        let cell = ws.cell(`A${i+1}`);
        cell.setString("value");
    }
    check_mem();
    // console.time("save");
    //wb.save("/tmp/unioffice.xlsx");
    wb.close();
    await (new Promise(setImmediate));
    //delete wb;
    //delete ws;
    // console.timeEnd("save");
}

function check_mem(){
    const used = process.memoryUsage().rss / 1024 / 1024;
    console.log(`The script uses approximately ${Math.round(used * 100) / 100} MB`);
}
run2();
check_mem();

run2();
check_mem();