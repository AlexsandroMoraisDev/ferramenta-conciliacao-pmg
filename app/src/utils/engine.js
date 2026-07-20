import * as XLSX from 'xlsx';

export const readExcelFile = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: true });
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
};

export const readExcelFileLetterHeaders = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: "A", defval: '', raw: true });
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
};

const normalizeStr = (str) => {
  if (str === undefined || str === null) return '';
  return String(str).toLowerCase().trim();
};

const normalizeNum = (val) => {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  const numStr = String(val).replace(/R\$/g, '').replace(/\./g, '').replace(',', '.').trim();
  const num = parseFloat(numStr);
  return isNaN(num) ? 0 : num;
};

const getCredor = (row) => {
  return normalizeStr(
    row['Credor'] || 
    row['RAZÃO SOCIAL'] || 
    row['Razão Social'] || 
    row['Fornecedor'] || 
    row['Fornecedor*'] ||
    row['Cliente/Fornecedor'] ||
    ''
  );
};

const getValor = (row) => {
  return normalizeNum(
    row['Valor líquido'] || 
    row['Valor Origem'] || 
    row[' VALOR LIQUIDO+JUROS E MULTAS '] || 
    row['VALOR LIQUIDO+JUROS E MULTAS'] ||
    row[' VALOR BRUTO '] || 
    row['VALOR BRUTO'] ||
    row['Valor original'] || 
    row['Total'] || 
    row['VALOR TOTAL DO CONTRATO'] ||
    row['VALOR DO FORNECEDOR'] ||
    row['Total do Pedido'] ||
    row['VALOR'] ||
    row['Total líquido'] ||
    row['VALOR LÍQUIDO/NF A EMITIR'] ||
    row['VALOR TOTAL'] ||
    0
  );
};

const buildCredorValorKey = (row) => {
  const credor = getCredor(row);
  const valor = getValor(row).toFixed(2);
  const credorAbrev = credor.substring(0, 15);
  return `${credorAbrev}_${valor}`;
};

const filterFallbackMatches = (matches, idSienge) => {
  if (!idSienge) return matches;
  return matches.filter(r => {
    if (!r._extractedId) return true;
    const extractedStr = String(r._extractedId).toLowerCase();
    const idStr = String(idSienge).toLowerCase();
    return extractedStr.includes(idStr) || idStr.includes(extractedStr);
  });
};

// ---------------------------------------------------------
// LÓGICA DE TÍTULOS (BLINDADA 100%)
// ---------------------------------------------------------
const processTitulos = (siengeData, zeppData, romaneioData) => {
  const zeppMapTitulo = {};
  const zeppMapCredorValor = {};
  zeppData.forEach(row => {
    const titulo = String(row['Código Origem'] || '').split('/')[0].trim();
    row._extractedId = titulo;
    const cv = buildCredorValorKey(row);
    if (!zeppMapCredorValor[cv]) zeppMapCredorValor[cv] = [];
    zeppMapCredorValor[cv].push(row);

    if (titulo) {
      if (!zeppMapTitulo[titulo]) zeppMapTitulo[titulo] = [];
      zeppMapTitulo[titulo].push(row);
    }
  });

  const romaneioMapTitulo = {};
  const romaneioMapCredorValor = {};
  romaneioData.forEach(row => {
    const titulo = String(row['TÍTULOS SIENGE'] || row['Título'] || row['TÍTULO'] || '').trim();
    row._extractedId = titulo;
    const cv = buildCredorValorKey(row);
    if (!romaneioMapCredorValor[cv]) romaneioMapCredorValor[cv] = [];
    romaneioMapCredorValor[cv].push(row);

    if (titulo) {
      if (!romaneioMapTitulo[titulo]) romaneioMapTitulo[titulo] = [];
      romaneioMapTitulo[titulo].push(row);
    }
  });

  const results = [];
  let kpi = { total: 0, pronto: 0, aprovacao: 0, acao: 0 };

  siengeData.forEach(siengeRow => {
    if (!siengeRow['Título'] && !siengeRow['Credor']) return;
    const tituloSienge = String(siengeRow['Título'] || '').trim();
    const cv = buildCredorValorKey(siengeRow);

    const zeppMatches = (tituloSienge && zeppMapTitulo[tituloSienge]) 
      ? zeppMapTitulo[tituloSienge] 
      : filterFallbackMatches(zeppMapCredorValor[cv] || [], tituloSienge);
      
    const romaneioMatches = (tituloSienge && romaneioMapTitulo[tituloSienge]) 
      ? romaneioMapTitulo[tituloSienge] 
      : filterFallbackMatches(romaneioMapCredorValor[cv] || [], tituloSienge);

    const inZepp = zeppMatches.length > 0;
    const inRomaneio = romaneioMatches.length > 0;

    let acao = '';
    let statusZepp = inZepp ? (zeppMatches[0]['Status'] || 'Aprovado') : 'Não encontrado';
    let vencimentoZepp = inZepp ? (zeppMatches[0]['Dt. Vencto'] || zeppMatches[0]['Dt. vencto'] || '') : '';
    let noRomaneio = inRomaneio ? (romaneioMatches[0]['Nº ROMANEIO'] || romaneioMatches[0]['Nº Romaneio'] || romaneioMatches[0]['Romaneio'] || 'Encontrado') : 'Sem Romaneio';
    let dataEmissao = inZepp ? (zeppMatches[0]['Emissão'] || zeppMatches[0]['emissão'] || '') : '';
    let tipoDocumento = siengeRow['Documento'] || '';

    const zeppStatusLower = statusZepp.toLowerCase();
    
    if (inRomaneio && inZepp && (zeppStatusLower.includes('aprovado') || zeppStatusLower.includes('concluído'))) {
      acao = 'OK: Lançado no Romaneio e Enviado'; kpi.pronto++;
    } else if (inRomaneio && inZepp) {
      acao = 'ALERTA: Em Aprovação no Zepp'; kpi.aprovacao++;
    } else if (inZepp && !inRomaneio) {
      acao = 'ALERTA: Falta Romaneio'; zeppStatusLower.includes('aprov') ? kpi.acao++ : kpi.aprovacao++;
    } else if (inRomaneio && !inZepp) {
      acao = 'ALERTA: Falta Zepp'; kpi.acao++;
    } else {
      acao = 'ALERTA: Falta Romaneio e Falta Zepp'; kpi.acao++;
    }

    kpi.total++;
    results.push({
      id: tituloSienge || '-',
      tipoDocumento: tipoDocumento || '-',
      credor: siengeRow['Credor'] || '-',
      dataEmissao: dataEmissao || '-',
      vencimento: vencimentoZepp || siengeRow['Data competência'] || siengeRow['Data contábil'] || '-',
      numeroNF: siengeRow['Nº documento'] || '-',
      valor: getValor(siengeRow),
      statusZepp,
      noRomaneio,
      observacao: siengeRow['Observação'] || '-',
      acao,
      originalZepp: inZepp ? zeppMatches[0] : null,
      originalSienge: siengeRow
    });
  });

  return { results, kpi };
};

// ---------------------------------------------------------
// LÓGICA DE CONTRATOS
// ---------------------------------------------------------
const processContratos = (siengeData, zeppData, romaneioData) => {
  const zeppMapID = {};
  const zeppMapCV = {};
  zeppData.forEach(row => {
    const id = String(row['Código Origem'] || '').split('/')[0].split('-').pop().trim(); // Ex: "CLSF / 200-3-039.1" -> "039.1" - heurística de ID
    row._extractedId = id;
    const cv = buildCredorValorKey(row);
    if (!zeppMapCV[cv]) zeppMapCV[cv] = [];
    zeppMapCV[cv].push(row);

    if (id) {
      if (!zeppMapID[id]) zeppMapID[id] = [];
      zeppMapID[id].push(row);
    }
  });

  const romMapID = {};
  const romMapCV = {};
  romaneioData.forEach(row => {
    const id = String(row['ID CONTRATO.'] || '').trim();
    row._extractedId = id;
    const cv = buildCredorValorKey(row);
    if (!romMapCV[cv]) romMapCV[cv] = [];
    romMapCV[cv].push(row);

    if (id) {
      if (!romMapID[id]) romMapID[id] = [];
      romMapID[id].push(row);
    }
  });

  const results = [];
  let kpi = { total: 0, pronto: 0, aprovacao: 0, acao: 0 };

  siengeData.forEach(siengeRow => {
    if (!siengeRow['Contrato'] && !siengeRow['Número acordo de preços']) return;
    const idSienge = String(siengeRow['Contrato'] || siengeRow['Número acordo de preços'] || '').trim();
    const cv = buildCredorValorKey(siengeRow);

    const zeppMatches = (idSienge && zeppMapID[idSienge]) 
      ? zeppMapID[idSienge] 
      : filterFallbackMatches(zeppMapCV[cv] || [], idSienge);
      
    const romaneioMatches = (idSienge && romMapID[idSienge]) 
      ? romMapID[idSienge] 
      : filterFallbackMatches(romMapCV[cv] || [], idSienge);

    const inZepp = zeppMatches.length > 0;
    const inRomaneio = romaneioMatches.length > 0;

    let acao = '';
    let statusZepp = inZepp ? (zeppMatches[0]['Status'] || 'Aprovado') : 'Não encontrado';
    let noRomaneio = inRomaneio ? 'Encontrado' : 'Sem Romaneio';

    const zeppStatusLower = statusZepp.toLowerCase();
    
    if (inRomaneio && inZepp && (zeppStatusLower.includes('aprovado') || zeppStatusLower.includes('concluído'))) {
      acao = 'OK: Base Atualizada e Aprovado'; kpi.pronto++;
    } else if (inRomaneio && inZepp) {
      acao = 'ALERTA: Em Aprovação no Zepp'; kpi.aprovacao++;
    } else if (inZepp && !inRomaneio) {
      acao = 'ALERTA: Falta na Base'; zeppStatusLower.includes('aprov') ? kpi.acao++ : kpi.aprovacao++;
    } else if (inRomaneio && !inZepp) {
      acao = 'ALERTA: Falta Zepp'; kpi.acao++;
    } else {
      acao = 'ALERTA: Falta Base e Zepp'; kpi.acao++;
    }

    kpi.total++;
    results.push({
      id: idSienge || '-',
      credor: siengeRow['Fornecedor*'] || siengeRow['Cliente/Fornecedor'] || '-',
      vencimento: siengeRow['Data de Término'] || siengeRow['Data do Contrato'] || '-',
      valor: getValor(siengeRow),
      statusZepp,
      noRomaneio,
      observacao: siengeRow['Obra'] || '-',
      acao,
      originalSienge: siengeRow
    });
  });

  return { results, kpi };
};

// ---------------------------------------------------------
// LÓGICA DE PEDIDOS
// ---------------------------------------------------------
const processPedidos = (siengeData, zeppData, romaneioData) => {
  const zeppMapID = {};
  const zeppMapCV = {};
  zeppData.forEach(row => {
    const id = String(row['Código Origem'] || '').trim();
    row._extractedId = id;
    const cv = buildCredorValorKey(row);
    if (!zeppMapCV[cv]) zeppMapCV[cv] = [];
    zeppMapCV[cv].push(row);

    if (id) {
      if (!zeppMapID[id]) zeppMapID[id] = [];
      zeppMapID[id].push(row);
    }
  });

  const romMapID = {};
  const romMapCV = {};
  romaneioData.forEach(row => {
    const id = String(row['PEDIDO'] || '').trim();
    row._extractedId = id;
    const cv = buildCredorValorKey(row);
    if (!romMapCV[cv]) romMapCV[cv] = [];
    romMapCV[cv].push(row);

    if (id) {
      if (!romMapID[id]) romMapID[id] = [];
      romMapID[id].push(row);
    }
  });

  const results = [];
  let kpi = { total: 0, pronto: 0, aprovacao: 0, acao: 0 };

  siengeData.forEach(siengeRow => {
    if (!siengeRow['N. do Pedido']) return;
    const idSienge = String(siengeRow['N. do Pedido'] || '').trim();
    const cv = buildCredorValorKey(siengeRow);

    // Na base Pedido (Romaneio), o ID vem como "PPC/25660". O Sienge talvez venha "25660".
    const idSiengeNum = idSienge.replace(/\D/g, ''); // Apenas números
    
    // Tenta encontrar ID exato ou contendo o número
    // Tenta encontrar ID exato ou contendo o número
    const zeppMatches = filterFallbackMatches(zeppMapCV[cv] || [], idSiengeNum);
    const romaneioMatches = (idSienge && romMapID[idSienge]) ? romMapID[idSienge] : (
       Object.keys(romMapID).find(k => k.includes(idSiengeNum)) 
        ? romMapID[Object.keys(romMapID).find(k => k.includes(idSiengeNum))] 
        : filterFallbackMatches(romMapCV[cv] || [], idSiengeNum)
    );

    const inZepp = zeppMatches.length > 0;
    const inRomaneio = romaneioMatches.length > 0;

    let acao = '';
    let statusZepp = inZepp ? (zeppMatches[0]['Status'] || 'Aprovado') : 'Não encontrado';
    let noRomaneio = inRomaneio ? 'Encontrado' : 'Sem Romaneio';

    const zeppStatusLower = statusZepp.toLowerCase();
    
    if (inRomaneio && inZepp && (zeppStatusLower.includes('aprovado') || zeppStatusLower.includes('concluído'))) {
      acao = 'OK: CONCILIADO'; kpi.pronto++;
    } else if (inRomaneio && inZepp) {
      acao = 'ALERTA: PENDENTE ZEPP'; kpi.aprovacao++;
    } else if (inZepp && !inRomaneio) {
      acao = 'ALERTA: CADASTRAR ROMANEIO'; kpi.acao++;
    } else if (inRomaneio && !inZepp) {
      acao = 'ALERTA: LANÇAR NO ZEPP'; kpi.acao++;
    } else {
      acao = 'ALERTA: VERIFICAR DADOS'; kpi.acao++;
    }

    kpi.total++;
    results.push({
      id: idSienge || '-',
      credor: siengeRow['Fornecedor'] || '-',
      cnpj: siengeRow['CNPJ/CPF'] || siengeRow['CNPJ'] || (inRomaneio ? (romaneioMatches[0]['CNPJ/CPF'] || romaneioMatches[0]['CNPJ']) : '') || '-',
      apropriacao: inRomaneio ? (romaneioMatches[0]['APROPRIAÇÃO'] || '-') : '-',
      vencimento: siengeRow['Data do Pedido'] || '-', // kept for compatibility with excel export if needed, but won't render
      valor: getValor(siengeRow),
      vinculoFaturamentoDireto: inRomaneio ? (romaneioMatches[0]['VÍNCULO DE FATURAMENTO DIRETO'] || '-') : '-',
      statusZepp,
      noRomaneio, // kept for compatibility
      observacao: siengeRow['Situação dos Pedidos'] || '-',
      acao,
      originalSienge: siengeRow
    });
  });

  return { results, kpi };
};

// ---------------------------------------------------------
// LÓGICA DE MEDIÇÕES
// ---------------------------------------------------------
const processMedicoes = (siengeData, zeppData, romaneioData) => {
  const zeppMapCV = {};
  zeppData.forEach(row => {
    const cv = buildCredorValorKey(row);
    if (!zeppMapCV[cv]) zeppMapCV[cv] = [];
    zeppMapCV[cv].push(row);
  });

  const romMapValor = {};
  romaneioData.forEach(row => {
    // A base de Medições não tem ID e nem Fornecedor claramente, então cruzamos pelo Valor!
    const valorStr = getValor(row).toFixed(2);
    if (!romMapValor[valorStr]) romMapValor[valorStr] = [];
    romMapValor[valorStr].push(row);
  });

  const results = [];
  let kpi = { total: 0, pronto: 0, aprovacao: 0, acao: 0 };

  siengeData.forEach(siengeRow => {
    if (!siengeRow['Contrato'] && !siengeRow['Medição']) return;
    const idSienge = String(siengeRow['Contrato'] || '') + (siengeRow['Medição'] ? ` / ${siengeRow['Medição']}` : '');
    const cv = buildCredorValorKey(siengeRow);
    const valorStr = getValor(siengeRow).toFixed(2);

    const zeppMatches = zeppMapCV[cv] || [];
    const romaneioMatches = romMapValor[valorStr] || [];

    const inZepp = zeppMatches.length > 0;
    const inRomaneio = romaneioMatches.length > 0;

    let acao = '';
    let statusZepp = inZepp ? (zeppMatches[0]['Status'] || 'Aprovado') : 'Não encontrado';
    let noRomaneio = inRomaneio ? 'Encontrado' : 'Sem Romaneio';

    const zeppStatusLower = statusZepp.toLowerCase();
    
    if (inRomaneio && inZepp && (zeppStatusLower.includes('aprovado') || zeppStatusLower.includes('concluído'))) {
      acao = 'OK: Base Atualizada e Aprovado'; kpi.pronto++;
    } else if (inRomaneio && inZepp) {
      acao = 'ALERTA: Em Aprovação no Zepp'; kpi.aprovacao++;
    } else if (inZepp && !inRomaneio) {
      acao = 'ALERTA: Falta na Base'; zeppStatusLower.includes('aprov') ? kpi.acao++ : kpi.aprovacao++;
    } else if (inRomaneio && !inZepp) {
      acao = 'ALERTA: Falta Zepp'; kpi.acao++;
    } else {
      acao = 'ALERTA: Falta Base e Zepp'; kpi.acao++;
    }

    kpi.total++;
    results.push({
      id: idSienge || '-',
      credor: siengeRow['Fornecedor*'] || siengeRow['Cliente/Fornecedor'] || '-',
      vencimento: siengeRow['Data de vencimento'] || siengeRow['Data da medição'] || '-',
      valor: getValor(siengeRow),
      statusZepp,
      noRomaneio,
      observacao: siengeRow['Observação da medição'] || '-',
      acao,
      originalSienge: siengeRow
    });
  });

  return { results, kpi };
};

// ---------------------------------------------------------
// ROUTER PRINCIPAL
// ---------------------------------------------------------
export const processConciliacao = async (files, categoryName = 'Títulos') => {
  const siengeData = files.sienge ? await readExcelFile(files.sienge) : [];
  const zeppData = files.zepp ? await readExcelFile(files.zepp) : [];
  const romaneioData = files.romaneio ? await readExcelFile(files.romaneio) : [];

  switch (categoryName) {
    case 'Títulos':
      return processTitulos(siengeData, zeppData, romaneioData);
    case 'Contratos':
      return processContratos(siengeData, zeppData, romaneioData);
    case 'Pedidos':
      return processPedidos(siengeData, zeppData, romaneioData);
    case 'Medições':
      return processMedicoes(siengeData, zeppData, romaneioData);
    case 'Conciliação Saldos Contábeis':
      return processConciliacaoSaldos(files);
    default:
      return processTitulos(siengeData, zeppData, romaneioData);
  }
};

// ---------------------------------------------------------
// LÓGICA DE CONCILIAÇÃO SALDOS CONTÁBEIS
// ---------------------------------------------------------
export const processConciliacaoSaldos = async (files) => {
  if (!files.planilha1 || !files.planilha2) {
     return { results: [], kpi: { total: 0, pronto: 0, aprovacao: 0, acao: 0 } };
  }

  const plan1Data = await readExcelFileLetterHeaders(files.planilha1);
  
  const plan2DataBuffer = await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(new Uint8Array(e.target.result));
    reader.onerror = (e) => reject(e);
    reader.readAsArrayBuffer(files.planilha2);
  });
  const wb2 = XLSX.read(plan2DataBuffer, { type: 'array', cellDates: true, cellStyles: true });

  const plan1Rows = plan1Data.map(r => ({
     nf: String(r['K'] || '').trim(),
     fornecedor: String(r['I'] || '').trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' '),
     cnpj: String(r['J'] || '').replace(/\D/g, ''),
     tipoAdiantamento: String(r['P'] || '').trim().toUpperCase(),
     valorAdiantamentoSinal: normalizeNum(r['AC']),
     valorAdiantamentoFuturo: normalizeNum(r['AD']),
     valorDescontado: normalizeNum(r['U']),
     contrato: String(r['AI'] || '').trim(),
     valorLiquido: normalizeNum(r['X']),
     valorRetencao: normalizeNum(r['T'])
  })).filter(r => r.nf !== '');

  const groupedPlan1 = {};
  plan1Rows.forEach(row => {
     const suppKey = row.cnpj ? row.cnpj : row.fornecedor;
     const key = `${row.nf}||${suppKey}`;
     if (!groupedPlan1[key]) {
        groupedPlan1[key] = { ...row };
     } else {
        groupedPlan1[key].valorAdiantamentoSinal += row.valorAdiantamentoSinal;
        groupedPlan1[key].valorAdiantamentoFuturo += row.valorAdiantamentoFuturo;
        groupedPlan1[key].valorDescontado += row.valorDescontado;
        groupedPlan1[key].valorLiquido += row.valorLiquido;
        groupedPlan1[key].valorRetencao += row.valorRetencao;
     }
  });

  const normalizeText = (text) => String(text || '').trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');

  const findMatch = (nf, fornecedorName, cnpjVal) => {
     const nfStr = String(nf || '').trim();
     const cnpjStr = String(cnpjVal || '').replace(/\D/g, '');
     const fornStr = normalizeText(fornecedorName);
     
     if (!nfStr && !fornStr && !cnpjStr) return null;

     if (nfStr) {
         if (cnpjStr) {
            const k1 = `${nfStr}||${cnpjStr}`;
            if (groupedPlan1[k1]) return groupedPlan1[k1];
         }
         
         if (fornStr) {
            const k2 = `${nfStr}||${fornStr}`;
            if (groupedPlan1[k2]) return groupedPlan1[k2];
            
            for (const [key, val] of Object.entries(groupedPlan1)) {
               if (val.nf === nfStr && (fornStr.includes(val.fornecedor) || val.fornecedor.includes(fornStr))) {
                   return val;
               }
            }
         }
     } else if (fornStr) {
         let totalSinal = 0;
         let totalFuturo = 0;
         let totalDescontado = 0;
         let contrato = '';
         let found = false;

         for (const [key, val] of Object.entries(groupedPlan1)) {
             if (fornStr.includes(val.fornecedor) || val.fornecedor.includes(fornStr)) {
                 totalSinal += val.valorAdiantamentoSinal || 0;
                 totalFuturo += val.valorAdiantamentoFuturo || 0;
                 totalDescontado += val.valorDescontado || 0;
                 if (val.contrato && !contrato) contrato = val.contrato;
                 found = true;
             }
         }
         if (found) {
             return {
                 isAdiantamentoSummary: true,
                 valorAdiantamentoSinal: totalSinal,
                 valorAdiantamentoFuturo: totalFuturo,
                 valorDescontado: totalDescontado,
                 contrato: contrato
             };
         }
     }
     return null;
  };

  const results = [];
  let kpi = { total: 0, pronto: 0, aprovacao: 0, acao: 0 };

  const sheets = wb2.SheetNames;
  const getSheetByKeyword = (keyword) => sheets.find(s => s.toLowerCase().includes(keyword.toLowerCase()));

  const sheetsNames = wb2.SheetNames;
  const getSheetByKeywordExact = (keyword) => sheetsNames.find(s => s.toLowerCase() === keyword.toLowerCase());
  const adiantamentoSheet = getSheetByKeyword('adiantamento');
  const fornecedoresSheet = getSheetByKeywordExact('fornecedores') || sheetsNames.find(s => s.toLowerCase().includes('fornecedor') && !s.toLowerCase().includes('adiantamento'));
  const retencaoSheet = getSheetByKeyword('reten');

  const processSheet = (sheetName, targetHeadersMap, sheetType) => {
     if (!sheetName) return;
     const ws = wb2.Sheets[sheetName];
     const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
     
     let headerRowIdx = -1;
     let headerMap = {};
     for (let i = 0; i < Math.min(20, aoa.length); i++) {
         const row = aoa[i];
         const rowStr = row.map(c => String(c).toLowerCase()).join(' ');
         if (rowStr.includes('valor de adiantamento') || rowStr.includes('valor líquido') || rowStr.includes('valor liquido') || rowStr.includes('valor de retenção')) {
             headerRowIdx = i;
             row.forEach((col, idx) => {
                 if (col) headerMap[String(col).trim().toLowerCase()] = idx;
             });
             break;
         }
     }

     if (headerRowIdx === -1) return;

     let nfIdx = -1;
     let fornIdx = -1;
     const descIdx = Object.keys(headerMap).find(k => k.includes('descri')) ? headerMap[Object.keys(headerMap).find(k => k.includes('descri'))] : -1;
     const obsIdx = Object.keys(headerMap).find(k => k === 'obs') ? headerMap[Object.keys(headerMap).find(k => k === 'obs')] : -1;
     const directNfIdx = Object.keys(headerMap).find(k => k === 'nf') ? headerMap[Object.keys(headerMap).find(k => k === 'nf')] : -1;

     if (sheetType === 'adiantamento') {
         fornIdx = descIdx;
     } else if (sheetType === 'fornecedores') {
         nfIdx = directNfIdx;
         fornIdx = descIdx;
     } else if (sheetType === 'retencao') {
         nfIdx = obsIdx;
         fornIdx = descIdx;
     }

     const targetIndices = {};
     for (const [key, propName] of Object.entries(targetHeadersMap)) {
         const matchedHeader = Object.keys(headerMap).find(k => k === key.toLowerCase() || k.includes(key.toLowerCase()));
         if (matchedHeader) {
             targetIndices[propName] = headerMap[matchedHeader];
         }
     }

     for (let i = headerRowIdx + 1; i < aoa.length; i++) {
         const row = aoa[i];
         if (!row || row.length === 0) continue;
         
         let nfVal = '';
         let fornVal = '';

         if (sheetType === 'adiantamento') {
             fornVal = fornIdx !== -1 ? String(row[fornIdx]).trim() : '';
             if (!fornVal) continue;
         } else if (sheetType === 'fornecedores') {
             nfVal = nfIdx !== -1 ? String(row[nfIdx]).trim() : '';
             let rawDesc = fornIdx !== -1 ? String(row[fornIdx]).trim() : '';
             if (rawDesc.includes('-')) {
                 fornVal = rawDesc.split('-').slice(1).join('-').trim();
             } else {
                 fornVal = rawDesc;
             }
             if (!nfVal) continue;
         } else if (sheetType === 'retencao') {
             nfVal = nfIdx !== -1 ? String(row[nfIdx]).trim() : '';
             fornVal = fornIdx !== -1 ? String(row[fornIdx]).trim() : '';
             if (!nfVal) continue;
         }

         const match = findMatch(nfVal, fornVal, '');
         kpi.total++;
         
         let acao = 'Pendente de validação';
         let statusZepp = 'Falta na Base Original';

         if (match) {
             kpi.pronto++;
             acao = 'OK: Linha Preenchida com Romaneio';
             statusZepp = 'Aprovado';
             for (const [propName, colIdx] of Object.entries(targetIndices)) {
                 if (propName === 'valorAdiantamento') {
                     row[colIdx] = match.isAdiantamentoSummary ? (match.valorAdiantamentoSinal + match.valorAdiantamentoFuturo) : (match.tipoAdiantamento && match.tipoAdiantamento.includes('FUTURO') ? match.valorAdiantamentoFuturo : match.valorAdiantamentoSinal);
                 } else if (propName === 'valorDescontado') {
                     row[colIdx] = match.valorDescontado;
                 } else if (propName === 'contrato') {
                     row[colIdx] = match.contrato;
                 } else if (propName === 'valorLiquido') {
                     row[colIdx] = match.valorLiquido;
                 } else if (propName === 'valorRetencao') {
                     row[colIdx] = match.valorRetencao;
                 }
             }
         } else {
             kpi.acao++;
         }

         results.push({
             id: nfVal || '-',
             credor: fornVal,
             vencimento: match && match.contrato ? match.contrato : '-', // Contrato
             valor: match ? (match.valorLiquido || 0) : 0,
             valorAdiantamento: match ? ((match.valorAdiantamentoSinal || 0) + (match.valorAdiantamentoFuturo || 0)) : 0,
             valorDescontado: match ? (match.valorDescontado || 0) : 0,
             valorRetencao: match ? (match.valorRetencao || 0) : 0,
             statusZepp,
             noRomaneio: sheetName,
             observacao: '-',
             acao,
             originalSienge: {}
         });
     }

     wb2.Sheets[sheetName] = XLSX.utils.aoa_to_sheet(aoa);
  };

  if (adiantamentoSheet) {
      processSheet(adiantamentoSheet, {
          'Valor de Adiantamento': 'valorAdiantamento',
          'Valor Descontado': 'valorDescontado',
          'Contrato/Pedido': 'contrato',
          'Contrato': 'contrato'
      }, 'adiantamento');
  }

  if (fornecedoresSheet) {
      processSheet(fornecedoresSheet, {
          'Valor liquido': 'valorLiquido',
          'Valor líquido': 'valorLiquido'
      }, 'fornecedores');
  }

  if (retencaoSheet) {
      processSheet(retencaoSheet, {
          'Valor de Retenção': 'valorRetencao',
          'Retenção': 'valorRetencao',
          'Contrato': 'contrato',
          'Contrato/Pedido': 'contrato'
      }, 'retencao');
  }

  return { results, kpi, finalWorkbook: wb2 };
};
