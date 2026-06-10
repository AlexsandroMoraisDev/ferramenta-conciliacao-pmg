const XLSX = require('xlsx');
function load(file) { return XLSX.utils.sheet_to_json(XLSX.readFile(file).Sheets[XLSX.readFile(file).SheetNames[0]], { defval: '', raw: true }); }

const zepp = load('../0.APROVADOS/TÍTULOS/TÍTULOS ZEPP.xlsx');
console.log("ZEPP with raw:true");
for(let i=1; i<4; i++) {
   let v = zepp[i]['Dt. Vencto'];
   console.log(v, typeof v);
}
