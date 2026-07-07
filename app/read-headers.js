import * as XLSX from 'xlsx';
const workbook = XLSX.readFile("C:/Users/Alexsandro Morais/ABR GERENCIAMENTO E ENGENHARIA LTDA/LYO004 SIMOES FILHO G200 - Documentos/03. CUSTOS/3.0 MEDIÇÕES APROVADAS E NF'S/1. ROMANEIOS/0.APROVADOS/TÍTULOS/TÍTULOS SIENGE.xlsx");
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const json = XLSX.utils.sheet_to_json(sheet, {header: 1});
console.log(json[0]);
