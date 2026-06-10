const XLSX = require('xlsx');
function load(file) { return XLSX.utils.sheet_to_json(XLSX.readFile(file).Sheets[XLSX.readFile(file).SheetNames[0]], { defval: '' }); }

const sienge = load('../0.APROVADOS/TÍTULOS/TÍTULOS SIENGE.xlsx');
const zepp = load('../0.APROVADOS/TÍTULOS/TÍTULOS ZEPP.xlsx');

console.log("SIENGE:");
for(let i=0; i<3; i++) {
   let v = sienge[i]['Data competência'];
   console.log(v, typeof v);
}

console.log("\nZEPP:");
for(let i=1; i<4; i++) {
   let v = zepp[i]['Dt. Vencto'];
   console.log(v, typeof v);
}
