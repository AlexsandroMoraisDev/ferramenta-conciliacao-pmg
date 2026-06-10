const XLSX = require('xlsx');

function load(file) {
  const workbook = XLSX.readFile(file);
  return XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: '', raw: false });
}

function normalizeNum(val) {
  if (!val) return 0;
  if (typeof val === 'number') return val;
  const numStr = String(val).replace(/R\$/g, '').replace(/\./g, '').replace(',', '.').trim();
  const num = parseFloat(numStr);
  return isNaN(num) ? 0 : num;
}

const sienge = load('../0.APROVADOS/TÍTULOS/TÍTULOS SIENGE.xlsx');
const romaneio = load('../0.APROVADOS/TÍTULOS/ROMANEIO.xlsx');
const zepp = load('../0.APROVADOS/TÍTULOS/TÍTULOS ZEPP.xlsx');

console.log("Total Sienge:", sienge.length);
console.log("Total Romaneio:", romaneio.length);
console.log("Total Zepp:", zepp.length);

console.log("\nSienge amostra:");
for(let i=0;i<2;i++) {
    console.log(sienge[i]['Credor'], normalizeNum(sienge[i]['Valor líquido']));
}

console.log("\nRomaneio amostra:");
for(let i=0;i<2;i++) {
    console.log(romaneio[i]['RAZÃO SOCIAL'], normalizeNum(romaneio[i][' VALOR LIQUIDO+JUROS E MULTAS '] || romaneio[i][' VALOR BRUTO ']));
}

console.log("\nZepp amostra:");
for(let i=1;i<3;i++) {
    console.log(zepp[i]['Credor'], normalizeNum(zepp[i]['Valor Origem']));
}

// Vamos testar o cruzamento!
let matchesRomaneio = 0;
let matchesZepp = 0;

sienge.forEach(s => {
    if(!s['Credor']) return;
    const credor = String(s['Credor']).toLowerCase().trim().substring(0, 15);
    const valor = normalizeNum(s['Valor líquido']).toFixed(2);
    
    // Busca no romaneio
    const foundR = romaneio.find(r => {
        const cR = String(r['RAZÃO SOCIAL'] || '').toLowerCase().trim().substring(0, 15);
        const vR = normalizeNum(r[' VALOR LIQUIDO+JUROS E MULTAS '] || r[' VALOR BRUTO ']).toFixed(2);
        return cR === credor && vR === valor;
    });
    if(foundR) matchesRomaneio++;

    const foundZ = zepp.find(z => {
        const cZ = String(z['Credor'] || '').toLowerCase().trim().substring(0, 15);
        const vZ = normalizeNum(z['Valor Origem']).toFixed(2);
        return cZ === credor && vZ === valor;
    });
    if(foundZ) matchesZepp++;
});

console.log(`\nMatches: Romaneio=${matchesRomaneio}/${sienge.length}, Zepp=${matchesZepp}/${sienge.length}`);

// Tentar cruzar pelo Nº Documento (Nota) e Valor
let matchesR_Nota = 0;
sienge.forEach(s => {
    if(!s['Credor']) return;
    const nota = String(s['Nº documento'] || '').trim().replace(/^0+/, ''); // remove zeros a esquerda
    const valor = normalizeNum(s['Valor líquido']).toFixed(2);
    
    const foundR = romaneio.find(r => {
        const notaR = String(r['NºNOTA'] || '').trim().replace(/^0+/, '');
        const vR = normalizeNum(r[' VALOR LIQUIDO+JUROS E MULTAS '] || r[' VALOR BRUTO ']).toFixed(2);
        // Tenta ver se a notaR contém a notaSienge ou vice-versa
        return (notaR === nota || notaR.includes(nota) || nota.includes(notaR)) && vR === valor;
    });
    if(foundR && nota !== '') matchesR_Nota++;
});
console.log(`Matches usando Nota+Valor (Romaneio): ${matchesR_Nota}`);

