const XLSX = require('xlsx');
const fs = require('fs');

const exportFile = "C:\\Users\\Alexsandro Morais\\ABR GERENCIAMENTO E ENGENHARIA LTDA\\LYO004 SIMOES FILHO G200 - Documentos\\03. CUSTOS\\3.0 MEDIÇÕES APROVADAS E NF'S\\1. ROMANEIOS\\0.APROVADOS\\Conciliação Saldos Contábeis\\Conciliacao_Saldos_Contabeis.xlsx";

try {
  const wb = XLSX.readFile(exportFile);
  const sheets = wb.SheetNames;
  console.log("Sheets in exported file:", sheets);

  sheets.forEach(sheetName => {
    if (!sheetName.toLowerCase().includes('fornecedor') && !sheetName.toLowerCase().includes('adiantamento') && !sheetName.toLowerCase().includes('reten')) return;
    console.log(`\n--- Sheet: ${sheetName} ---`);
    const ws = wb.Sheets[sheetName];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    
    // find header row
    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(20, aoa.length); i++) {
      const rowStr = aoa[i].map(c => String(c).toLowerCase()).join(' ');
      if (rowStr.includes('valor de adiantamento') || rowStr.includes('valor líquido') || rowStr.includes('valor liquido') || rowStr.includes('valor de retenção')) {
          headerRowIdx = i;
          break;
      }
    }
    console.log(`Header row idx: ${headerRowIdx}`);
    if (headerRowIdx === -1) return;
    
    let filled = 0;
    let empty = 0;
    const headerRow = aoa[headerRowIdx];
    
    const targetIdxs = [];
    headerRow.forEach((col, idx) => {
      const colStr = String(col).toLowerCase();
      if (colStr.includes('valor de adiantamento') || colStr.includes('valor líquido') || colStr.includes('valor liquido') || colStr.includes('valor de retenção') || colStr === 'valor descontado') {
        targetIdxs.push({ name: colStr, idx });
      }
    });

    for (let i = headerRowIdx + 1; i < aoa.length; i++) {
        const row = aoa[i];
        // simple heuristic: if it has any description/nf text
        if (row.join('').trim().length > 0 && !row.join('').includes('TOTAL')) {
            let hasVal = false;
            let vals = [];
            targetIdxs.forEach(t => {
               const v = row[t.idx];
               vals.push(`${t.name}=${v}`);
               if (v !== '' && v !== undefined && v !== null && v !== 0) hasVal = true;
            });
            if (hasVal) filled++;
            else {
               empty++;
               console.log(`Row ${i} empty ->`, row.slice(0, 10).join(' | '));
            }
        }
    }
    console.log(`Filled: ${filled}, Empty: ${empty}`);
  });
} catch (e) {
  console.error("Error reading file:", e);
}
