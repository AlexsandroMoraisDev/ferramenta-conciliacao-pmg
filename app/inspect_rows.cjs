const XLSX = require('xlsx');

function dumpRows(file, num) {
  try {
    const workbook = XLSX.readFile(file);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true });
    return json.slice(0, num);
  } catch (e) {
    return `Error: ${e.message}`;
  }
}

console.log("--- MEDIÇÃO BASE ---");
console.log(dumpRows('../0.APROVADOS/MEDIÇÃO/BASE-MEDIÇÃO.xlsx', 2));

console.log("--- PEDIDOS BASE ---");
console.log(dumpRows('../0.APROVADOS/PEDIDOS/BASE-PEDIDOS.xlsx', 2));

console.log("--- CONTRATOS ZEPP ---");
console.log(dumpRows('../0.APROVADOS/CONTRATOS/ZEPP_Analítico_Processos_Aprovações.xlsx', 2));
