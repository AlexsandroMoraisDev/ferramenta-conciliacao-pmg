const XLSX = require('xlsx');

function load(file) {
  return XLSX.utils.sheet_to_json(XLSX.readFile(file).Sheets[XLSX.readFile(file).SheetNames[0]], { defval: '', raw: true });
}

function normalizeNum(val) {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  const numStr = String(val).replace(/R\$/g, '').replace(/\./g, '').replace(',', '.').trim();
  const num = parseFloat(numStr);
  return isNaN(num) ? 0 : num;
}

const sienge = load('../0.APROVADOS/TÍTULOS/TÍTULOS SIENGE.xlsx');
const romaneio = load('../0.APROVADOS/TÍTULOS/ROMANEIO.xlsx');
const zepp = load('../0.APROVADOS/TÍTULOS/TÍTULOS ZEPP.xlsx');

console.log("--- AUDITORIA PROFUNDA DE CHAVES ---");
console.log(`Total Sienge: ${sienge.length}`);
console.log(`Total Romaneio: ${romaneio.length}`);
console.log(`Total Zepp: ${zepp.length}`);

// 1. Analisar presença de TÍTULOS no Sienge
let siengeWithTitulo = 0;
sienge.forEach(s => {
    if (s['Título']) siengeWithTitulo++;
});
console.log(`Sienge com Título: ${siengeWithTitulo}/${sienge.length}`);

// 2. Analisar presença de TÍTULOS no Romaneio
let romaneioWithTitulo = 0;
romaneio.forEach(r => {
    if (r['TÍTULOS SIENGE']) romaneioWithTitulo++;
});
console.log(`Romaneio com Título: ${romaneioWithTitulo}/${romaneio.length}`);

// 3. Analisar formato de Código Origem no Zepp
let zeppWithCodigo = 0;
zepp.forEach(z => {
    if (z['Código Origem']) zeppWithCodigo++;
});
console.log(`Zepp com Código Origem: ${zeppWithCodigo}/${zepp.length}`);

// 4. Testar Join apenas por TÍTULO
let matchTituloRomaneio = 0;
let matchTituloZepp = 0;

sienge.forEach(s => {
    if (!s['Título']) return;
    const tituloSienge = String(s['Título']).trim();

    const foundR = romaneio.find(r => {
        const tr = String(r['TÍTULOS SIENGE']).trim();
        return tr === tituloSienge;
    });
    if (foundR) matchTituloRomaneio++;

    const foundZ = zepp.find(z => {
        const tz = String(z['Código Origem']).trim();
        // ZEPP tem '96666 / 1', então a gente pega o que tá antes da barra
        const zeppTitulo = tz.split('/')[0].trim();
        return zeppTitulo === tituloSienge;
    });
    if (foundZ) matchTituloZepp++;
});

console.log(`\nJoin Sienge -> Romaneio apenas por Título: ${matchTituloRomaneio} matches`);
console.log(`Join Sienge -> Zepp apenas por Título: ${matchTituloZepp} matches`);

// 5. Vamos pegar os que FALHARAM no nosso método antigo de (Credor + Valor) pra ver por que falharam.
let failOldZepp = 0;
let failOldRomaneio = 0;

sienge.forEach(s => {
    if (!s['Credor']) return;
    const credor = String(s['Credor']).toLowerCase().trim().substring(0, 15);
    const valor = normalizeNum(s['Valor líquido']).toFixed(2);
    
    const foundR = romaneio.find(r => {
        const cR = String(r['RAZÃO SOCIAL'] || '').toLowerCase().trim().substring(0, 15);
        const vR = normalizeNum(r[' VALOR LIQUIDO+JUROS E MULTAS '] || r[' VALOR BRUTO ']).toFixed(2);
        return cR === credor && vR === valor;
    });
    if (!foundR) failOldRomaneio++;

    const foundZ = zepp.find(z => {
        const cZ = String(z['Credor'] || '').toLowerCase().trim().substring(0, 15);
        const vZ = normalizeNum(z['Valor Origem']).toFixed(2);
        return cZ === credor && vZ === valor;
    });
    if (!foundZ) failOldZepp++;
});

console.log(`\nFalhas no Join antigo (Credor+Valor):`);
console.log(`Falharam Romaneio: ${failOldRomaneio}`);
console.log(`Falharam Zepp: ${failOldZepp}`);

// O que é melhor? Título + Valor? 
// Imprimir exemplos das primeiras falhas do Sienge para entendermos:
let amostrasFailRomaneio = 0;
console.log("\n--- AMOSTRAS DE FALHA NO ROMANEIO (MÉTODO ANTIGO) ---");
sienge.forEach(s => {
    if (amostrasFailRomaneio >= 3 || !s['Credor']) return;
    const credor = String(s['Credor']).toLowerCase().trim().substring(0, 15);
    const valor = normalizeNum(s['Valor líquido']).toFixed(2);
    
    const foundR = romaneio.find(r => {
        const cR = String(r['RAZÃO SOCIAL'] || '').toLowerCase().trim().substring(0, 15);
        const vR = normalizeNum(r[' VALOR LIQUIDO+JUROS E MULTAS '] || r[' VALOR BRUTO ']).toFixed(2);
        return cR === credor && vR === valor;
    });
    
    if (!foundR) {
        console.log(`Sienge Título: ${s['Título']}, Credor: ${s['Credor']}, Valor: ${valor}`);
        // Tentar achar pelo título no romaneio para ver qual era o credor/valor lá
        const tituloSienge = String(s['Título']).trim();
        const foundRByTitulo = romaneio.find(r => String(r['TÍTULOS SIENGE']).trim() === tituloSienge);
        if (foundRByTitulo) {
            console.log(`  -> Achado no Romaneio via Título!`);
            console.log(`     Razão Social: ${foundRByTitulo['RAZÃO SOCIAL']}`);
            console.log(`     Valor Bruto: ${foundRByTitulo[' VALOR BRUTO ']}, Valor Liq: ${foundRByTitulo[' VALOR LIQUIDO+JUROS E MULTAS ']}`);
        } else {
             console.log(`  -> NÃO ACHADO no Romaneio nem via Título!`);
        }
        amostrasFailRomaneio++;
    }
});

let amostrasFailZepp = 0;
console.log("\n--- AMOSTRAS DE FALHA NO ZEPP (MÉTODO ANTIGO) ---");
sienge.forEach(s => {
    if (amostrasFailZepp >= 3 || !s['Credor']) return;
    const credor = String(s['Credor']).toLowerCase().trim().substring(0, 15);
    const valor = normalizeNum(s['Valor líquido']).toFixed(2);
    
    const foundZ = zepp.find(z => {
        const cZ = String(z['Credor'] || '').toLowerCase().trim().substring(0, 15);
        const vZ = normalizeNum(z['Valor Origem']).toFixed(2);
        return cZ === credor && vZ === valor;
    });
    
    if (!foundZ) {
        console.log(`Sienge Título: ${s['Título']}, Credor: ${s['Credor']}, Valor: ${valor}`);
        // Tentar achar pelo título
        const tituloSienge = String(s['Título']).trim();
        const foundZByTitulo = zepp.find(z => String(z['Código Origem']).trim().split('/')[0].trim() === tituloSienge);
        if (foundZByTitulo) {
            console.log(`  -> Achado no Zepp via Título!`);
            console.log(`     Credor Zepp: ${foundZByTitulo['Credor']}, Valor Zepp: ${foundZByTitulo['Valor Origem']}`);
        } else {
             console.log(`  -> NÃO ACHADO no Zepp nem via Título!`);
        }
        amostrasFailZepp++;
    }
});
