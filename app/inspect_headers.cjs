const XLSX = require('xlsx');

function getHeaders(file) {
  try {
    const workbook = XLSX.readFile(file);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true, header: 1 });
    return json[0] || [];
  } catch (e) {
    return `Error: ${e.message}`;
  }
}

console.log("--- CONTRATOS ---");
console.log("Sienge:", getHeaders('../0.APROVADOS/CONTRATOS/sienge_relatorio-CONTRATOS.xlsx'));
console.log("Zepp:", getHeaders('../0.APROVADOS/CONTRATOS/ZEPP_Analítico_Processos_Aprovações.xlsx'));
console.log("Base/Romaneio:", getHeaders('../0.APROVADOS/CONTRATOS/BASE-CONTRATOS.xlsx'));

console.log("\n--- PEDIDOS ---");
console.log("Sienge:", getHeaders('../0.APROVADOS/PEDIDOS/sienge_relatorio-PEDIDOS.xlsx'));
console.log("Zepp:", getHeaders('../0.APROVADOS/PEDIDOS/ZEPP_Analítico_Processos_Aprovações.xlsx'));
console.log("Base/Romaneio:", getHeaders('../0.APROVADOS/PEDIDOS/BASE-PEDIDOS.xlsx'));

console.log("\n--- MEDIÇÃO ---");
console.log("Sienge:", getHeaders('../0.APROVADOS/MEDIÇÃO/sienge_relatorio-MEDIÇÃO.xlsx'));
console.log("Zepp:", getHeaders('../0.APROVADOS/MEDIÇÃO/ZEPP_Analítico_Processos_Aprovações.xlsx'));
console.log("Base/Romaneio:", getHeaders('../0.APROVADOS/MEDIÇÃO/BASE-MEDIÇÃO.xlsx'));
