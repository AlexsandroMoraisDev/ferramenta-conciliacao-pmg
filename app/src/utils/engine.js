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

// ---------------------------------------------------------
// LÓGICA DE TÍTULOS (BLINDADA 100%)
// ---------------------------------------------------------
const processTitulos = (siengeData, zeppData, romaneioData) => {
  const zeppMapTitulo = {};
  const zeppMapCredorValor = {};
  zeppData.forEach(row => {
    const cv = buildCredorValorKey(row);
    if (!zeppMapCredorValor[cv]) zeppMapCredorValor[cv] = [];
    zeppMapCredorValor[cv].push(row);

    const titulo = String(row['Código Origem'] || '').split('/')[0].trim();
    if (titulo) {
      if (!zeppMapTitulo[titulo]) zeppMapTitulo[titulo] = [];
      zeppMapTitulo[titulo].push(row);
    }
  });

  const romaneioMapTitulo = {};
  const romaneioMapCredorValor = {};
  romaneioData.forEach(row => {
    const cv = buildCredorValorKey(row);
    if (!romaneioMapCredorValor[cv]) romaneioMapCredorValor[cv] = [];
    romaneioMapCredorValor[cv].push(row);

    const titulo = String(row['TÍTULOS SIENGE'] || row['Título'] || row['TÍTULO'] || '').trim();
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

    const zeppMatches = (tituloSienge && zeppMapTitulo[tituloSienge]) ? zeppMapTitulo[tituloSienge] : (zeppMapCredorValor[cv] || []);
    const romaneioMatches = (tituloSienge && romaneioMapTitulo[tituloSienge]) ? romaneioMapTitulo[tituloSienge] : (romaneioMapCredorValor[cv] || []);

    const inZepp = zeppMatches.length > 0;
    const inRomaneio = romaneioMatches.length > 0;

    let acao = '';
    let statusZepp = inZepp ? (zeppMatches[0]['Status'] || 'Aprovado') : 'Não encontrado';
    let vencimentoZepp = inZepp ? (zeppMatches[0]['Dt. Vencto'] || zeppMatches[0]['Dt. vencto'] || '') : '';
    let noRomaneio = inRomaneio ? (romaneioMatches[0]['Nº ROMANEIO'] || romaneioMatches[0]['Nº Romaneio'] || romaneioMatches[0]['Romaneio'] || 'Encontrado') : 'Sem Romaneio';

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
      credor: siengeRow['Credor'] || '-',
      vencimento: vencimentoZepp || siengeRow['Data competência'] || siengeRow['Data contábil'] || '-',
      valor: getValor(siengeRow),
      statusZepp,
      noRomaneio,
      observacao: siengeRow['Observação'] || '-',
      acao,
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
    const cv = buildCredorValorKey(row);
    if (!zeppMapCV[cv]) zeppMapCV[cv] = [];
    zeppMapCV[cv].push(row);

    const id = String(row['Código Origem'] || '').split('/')[0].split('-').pop().trim(); // Ex: "CLSF / 200-3-039.1" -> "039.1" - heurística de ID
    if (id) {
      if (!zeppMapID[id]) zeppMapID[id] = [];
      zeppMapID[id].push(row);
    }
  });

  const romMapID = {};
  const romMapCV = {};
  romaneioData.forEach(row => {
    const cv = buildCredorValorKey(row);
    if (!romMapCV[cv]) romMapCV[cv] = [];
    romMapCV[cv].push(row);

    const id = String(row['ID CONTRATO.'] || '').trim();
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

    const zeppMatches = (idSienge && zeppMapID[idSienge]) ? zeppMapID[idSienge] : (zeppMapCV[cv] || []);
    const romaneioMatches = (idSienge && romMapID[idSienge]) ? romMapID[idSienge] : (romMapCV[cv] || []);

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
    const cv = buildCredorValorKey(row);
    if (!zeppMapCV[cv]) zeppMapCV[cv] = [];
    zeppMapCV[cv].push(row);

    const id = String(row['Código Origem'] || '').trim();
    if (id) {
      if (!zeppMapID[id]) zeppMapID[id] = [];
      zeppMapID[id].push(row);
    }
  });

  const romMapID = {};
  const romMapCV = {};
  romaneioData.forEach(row => {
    const cv = buildCredorValorKey(row);
    if (!romMapCV[cv]) romMapCV[cv] = [];
    romMapCV[cv].push(row);

    const id = String(row['PEDIDO'] || '').trim();
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
    const zeppMatches = zeppMapCV[cv] || []; // Para Zepp, CV é mais seguro, a menos q Code Origem bata
    const romaneioMatches = (idSienge && romMapID[idSienge]) ? romMapID[idSienge] : (
       Object.keys(romMapID).find(k => k.includes(idSiengeNum)) ? romMapID[Object.keys(romMapID).find(k => k.includes(idSiengeNum))] : (romMapCV[cv] || [])
    );

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
      credor: siengeRow['Fornecedor'] || '-',
      vencimento: siengeRow['Data do Pedido'] || '-',
      valor: getValor(siengeRow),
      statusZepp,
      noRomaneio,
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
    default:
      return processTitulos(siengeData, zeppData, romaneioData);
  }
};
