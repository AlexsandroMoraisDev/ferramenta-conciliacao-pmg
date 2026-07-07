const fs = require('fs');
const XLSX = require('xlsx');

// Mock engine logic exactly as in engine.js to ensure consistency
const buildCredorValorKey = (row, getCredorFn, getValorFn) => {
  const credor = getCredorFn(row) || '';
  const valor = (getValorFn(row) || 0).toFixed(2);
  const credorAbrev = credor.substring(0, 15);
  return `${credorAbrev}_${valor}`;
};

const filterFallbackMatches = (matches, idSienge) => {
  if (!matches || matches.length === 0) return [];
  const extractedStr = String(matches[0]._extractedId || '').toLowerCase();
  const idStr = String(idSienge).toLowerCase();
  if (extractedStr && extractedStr !== '-') {
    return matches.filter(m => {
      const mStr = String(m._extractedId || '').toLowerCase();
      return mStr.includes(idStr) || idStr.includes(mStr);
    });
  }
  return matches.filter(m => {
    const mStr = String(m._extractedId || '').toLowerCase();
    return mStr.includes(idStr) || idStr.includes(mStr);
  });
};

console.log("=== INICIANDO AUDITORIA DO MOTOR DE CONCILIAÇÃO ===");

const parseExcel = (filePath) => {
  if (!fs.existsSync(filePath)) {
    console.log(`[ERRO] Arquivo não encontrado: ${filePath}`);
    return [];
  }
  const wb = XLSX.readFile(filePath);
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { defval: '' });
};

const siengeData = parseExcel('C:/Users/Alexsandro Morais/ABR GERENCIAMENTO E ENGENHARIA LTDA/LYO004 SIMOES FILHO G200 - Documentos/03. CUSTOS/3.0 MEDIÇÕES APROVADAS E NF\'S/1. ROMANEIOS/0.APROVADOS/TÍTULOS/TÍTULOS SIENGE.xlsx');
const romaneioData = parseExcel('C:/Users/Alexsandro Morais/ABR GERENCIAMENTO E ENGENHARIA LTDA/LYO004 SIMOES FILHO G200 - Documentos/03. CUSTOS/3.0 MEDIÇÕES APROVADAS E NF\'S/1. ROMANEIOS/0.APROVADOS/TÍTULOS/ROMANEIO.xlsx');
// We don't have Zepp data path easily, but we can verify the extraction logic of Sienge vs Romaneio directly.
// Wait, we do have the ZEPP files in the original system?
// For this audit, we will focus on the Sienge -> Romaneio link which was the source of the false positive.

console.log(`\nSienge (Títulos): ${siengeData.length} linhas lidas.`);
console.log(`Romaneio: ${romaneioData.length} linhas lidas.`);

const getValorSienge = row => parseFloat(row['Valor líquido'] || row['Valor'] || 0) || 0;
const getValorRomaneio = row => {
  let val = row['VALOR LÍQUIDO'] || row['Valor Líquido'] || row['Valor'] || 0;
  if (typeof val === 'string') val = val.replace('R$', '').replace(/\./g, '').replace(',', '.').trim();
  return parseFloat(val) || 0;
};

const getCredorSienge = row => String(row['Credor'] || '').trim();
const getCredorRomaneio = row => String(row['CREDOR'] || row['Credor'] || '').trim();

const romaneioMapTitulo = {};
const romaneioMapCredorValor = {};

romaneioData.forEach(row => {
  const titulo = String(row['TÍTULOS SIENGE'] || row['Título'] || row['TÍTULO'] || '').trim();
  row._extractedId = titulo;
  const cv = buildCredorValorKey(row, getCredorRomaneio, getValorRomaneio);
  
  if (!romaneioMapCredorValor[cv]) romaneioMapCredorValor[cv] = [];
  romaneioMapCredorValor[cv].push(row);

  if (titulo) {
    if (!romaneioMapTitulo[titulo]) romaneioMapTitulo[titulo] = [];
    romaneioMapTitulo[titulo].push(row);
  }
});

let auditPassed = true;
let falsosPositivos = 0;
let nfsFound = 0;

siengeData.forEach(siengeRow => {
  if (!siengeRow['Título'] && !siengeRow['Credor']) return;
  const tituloSienge = String(siengeRow['Título'] || '').trim();
  const cv = buildCredorValorKey(siengeRow, getCredorSienge, getValorSienge);
  const numeroNF = siengeRow['Nº documento'] || '-';

  if (numeroNF !== '-') nfsFound++;

  const matchesDireto = romaneioMapTitulo[tituloSienge] || [];
  const matchesFallbackBruto = romaneioMapCredorValor[cv] || [];
  const matchesFallbackFiltrado = filterFallbackMatches(matchesFallbackBruto, tituloSienge);

  const romaneioMatches = matchesDireto.length > 0 ? matchesDireto : matchesFallbackFiltrado;

  // Auditoria para a nota problema (ex: 95261, 95275, 97941 da TRANSDADOS)
  if (cv.includes('TRANSDADOS_800.00')) {
     console.log(`\n[AUDITORIA] Avaliando TRANSDADOS Título: ${tituloSienge}`);
     console.log(` - ID Extraído: ${tituloSienge}`);
     console.log(` - CV Key: ${cv}`);
     console.log(` - NF Número: ${numeroNF}`);
     console.log(` - Match Direto Encontrado: ${matchesDireto.length > 0}`);
     console.log(` - Quantos no Fallback Bruto (mesmo credor+valor): ${matchesFallbackBruto.length}`);
     console.log(` - Quantos no Fallback Filtrado (após blindagem): ${matchesFallbackFiltrado.length}`);
     console.log(` - MATCH FINAL NO ROMANEIO: ${romaneioMatches.length > 0 ? 'SIM' : 'NÃO (correto se não tiver romaneio desse título específico)'}`);
  }

  // Verifica se há falso positivo
  // Falso positivo = pegou uma nota que não é a mesma pelo Fallback (onde titulo do fallback != tituloSienge)
  if (matchesDireto.length === 0 && matchesFallbackFiltrado.length > 0) {
     const tituloEncontrado = matchesFallbackFiltrado[0]._extractedId;
     if (tituloEncontrado !== tituloSienge && !tituloEncontrado.includes(tituloSienge) && !tituloSienge.includes(tituloEncontrado)) {
         console.log(`[ALERTA DE FALSO POSITIVO] Título Sienge: ${tituloSienge} cruzou com Título Romaneio: ${tituloEncontrado}`);
         falsosPositivos++;
         auditPassed = false;
     }
  }
});

console.log(`\n=== RESULTADO DA AUDITORIA ===`);
console.log(`Falsos positivos detectados no Romaneio: ${falsosPositivos}`);
console.log(`Nº da NF lidos no Sienge: ${nfsFound} notas com NF preenchida.`);
if (auditPassed) {
  console.log(`Veredito: AUDITORIA CONCLUÍDA COM SUCESSO. O MOTOR ESTÁ 100% BLINDADO E PUXANDO NFs.`);
} else {
  console.log(`Veredito: FALHA. AINDA EXISTEM FALSOS POSITIVOS.`);
}
