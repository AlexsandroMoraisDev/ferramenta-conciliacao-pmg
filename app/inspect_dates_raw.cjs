const XLSX = require('xlsx');
function load(file) { return XLSX.utils.sheet_to_json(XLSX.readFile(file).Sheets[XLSX.readFile(file).SheetNames[0]], { defval: '', raw: true }); }

const sienge = load('../0.APROVADOS/TÍTULOS/TÍTULOS SIENGE.xlsx');
console.log("SIENGE with raw:true");
for(let i=0; i<3; i++) {
   let v = sienge[i]['Data competência'];
   console.log(v, typeof v);
}
