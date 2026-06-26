const XLSX = require('xlsx');
const wb = XLSX.readFile('C:/Users/Alexsandro Morais/ABR GERENCIAMENTO E ENGENHARIA LTDA/LYO004 SIMOES FILHO G200 - Documentos/03. CUSTOS/3.0 MEDIÇÕES APROVADAS E NF\'S/1. ROMANEIOS/0.APROVADOS/PEDIDOS/BASE-PEDIDOS.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws, {header: 1});
console.log('BASE-PEDIDOS:', data[0]);

const wbZ = XLSX.readFile('C:/Users/Alexsandro Morais/ABR GERENCIAMENTO E ENGENHARIA LTDA/LYO004 SIMOES FILHO G200 - Documentos/03. CUSTOS/3.0 MEDIÇÕES APROVADAS E NF\'S/1. ROMANEIOS/0.APROVADOS/PEDIDOS/ZEPP_Analítico_Processos_Aprovações.xlsx');
const wsZ = wbZ.Sheets[wbZ.SheetNames[0]];
const dataZ = XLSX.utils.sheet_to_json(wsZ, {header: 1});
console.log('ZEPP:', dataZ[0]);
