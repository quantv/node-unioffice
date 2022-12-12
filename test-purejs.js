
async function run2(){
    
}

function check_mem(){
    const used = process.memoryUsage().rss / 1024 / 1024;
    console.log(`The script uses approximately ${Math.round(used * 100) / 100} MB`);
}
// console.time('run');
// run();
// console.timeEnd('run');
let c = 0;
setInterval(() => {
    run2();
    c++;
    console.log("round", c);
    check_mem()
}, 1000);