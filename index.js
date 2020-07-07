const file_name = process.argv.slice(2)[0];
const output    = process.argv.slice(2)[1];
const fs = require('fs');
var dafne_process = require('./lib/ExcelParser').process;

dafne_process(file_name).then((jsonData) =>{
    const stringData = JSON.stringify(jsonData,null,4)
    fs.writeFile(output, stringData, (err) => {
    });
});
