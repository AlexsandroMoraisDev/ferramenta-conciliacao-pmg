const XLSX = require('xlsx');

function normalizeNum(val) {
  if (typeof val === 'number') return val;
  const str = String(val || '').replace(/[^0-9,-]/g, '').replace(',', '.');
  return parseFloat(str) || 0;
}

const file1 = "C:\\Users\\Alexsandro Morais\\ABR GERENCIAMENTO E ENGENHARIA LTDA\\LYO004 SIMOES FILHO G200 - Documentos\\03. CUSTOS\\3.0 MEDIÇÕES APROVADAS E NF'S\\1. ROMANEIOS\\7. JULHO\\LYO004-SF-Romaneio 0045.xlsm";

const wb1 = XLSX.readFile(file1);
const ws1 = wb1.Sheets[wb1.SheetNames[0]];
const aoa1 = XLSX.utils.sheet_to_json(ws1, { header: 1, defval: '' });

const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");
const extendedLetters = [...letters];
for (let i = 0; i < letters.length; i++) {
  for (let j = 0; j < letters.length; j++) {
    extendedLetters.push(letters[i] + letters[j]);
  }
}
const plan1Data = [];
for (let r = 0; r < aoa1.length; r++) {
  const row = aoa1[r];
  if (!row || row.length === 0) continue;
  let obj = {};
  for (let c = 0; c < row.length; c++) {
    obj[extendedLetters[c]] = row[c];
  }
  plan1Data.push(obj);
}

const normalizeText = (text) => String(text || '').trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');

const plan1Rows = plan1Data.map(r => ({
  nf: String(r['K'] || '').trim(),
  fornecedor: normalizeText(r['I']),
  cnpj: String(r['J'] || '').replace(/\D/g, '')
})).filter(r => r.nf !== '');

console.log("Looking for NF 694849");
console.log(plan1Rows.filter(r => r.nf === '694849'));
console.log("Looking for NF 3590");
console.log(plan1Rows.filter(r => r.nf === '3590'));
console.log("Looking for NF 2026331");
console.log(plan1Rows.filter(r => r.nf === '2026331'));
console.log("Looking for NF 202629339");
console.log(plan1Rows.filter(r => r.nf === '202629339'));

console.log("\nLooking for Supplier 'T & A'");
console.log(plan1Rows.filter(r => r.fornecedor.includes('t & a') || r.fornecedor.includes('t&a') || r.fornecedor.includes('t e a') || r.fornecedor.includes('construcao pre-fabricada')));

console.log("\nLooking for Supplier 'TK ELEVADORES'");
console.log(plan1Rows.filter(r => r.fornecedor.includes('tk') || r.fornecedor.includes('elevador')));
