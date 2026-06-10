const XLSX = require('xlsx');
function load(file) { return XLSX.utils.sheet_to_json(XLSX.readFile(file).Sheets[XLSX.readFile(file).SheetNames[0]], { defval: '' }); }

const sienge = load('../0.APROVADOS/TÍTULOS/TÍTULOS SIENGE.xlsx');
const romaneio = load('../0.APROVADOS/TÍTULOS/ROMANEIO.xlsx');

console.log("SIENGE:");
for(let i=0; i<3; i++) {
   let v = sienge[i]['Valor líquido'];
   console.log(v, typeof v);
}

console.log("\nROMANEIO:");
for(let i=0; i<3; i++) {
   let v = romaneio[i][' VALOR LIQUIDO+JUROS E MULTAS '];
   console.log(v, typeof v);
}
