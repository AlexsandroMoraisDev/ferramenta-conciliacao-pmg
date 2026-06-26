const XLSX = require('xlsx');
const wb = XLSX.readFile('C:/Users/Alexsandro Morais/ABR GERENCIAMENTO E ENGENHARIA LTDA/LYO004 SIMOES FILHO G200 - Documentos/03. CUSTOS/3.0 MEDIÇÕES APROVADAS E NF\'S/1. ROMANEIOS/0.APROVADOS/PEDIDOS/sienge_relatorio-PEDIDOS.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws, {header: 1});
console.log(data[0]);
