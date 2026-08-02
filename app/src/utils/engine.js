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

export const readExcelFileRawRows = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: true });
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
// LÓGICA DE TÍTULOS — AUDITADA E CORRIGIDA
// Correções aplicadas:
//   #1 Varredura reversa do Zepp para títulos ausentes no Sienge
//   #2 Guard clause com log ao invés de descarte silencioso
//   #3 Melhor match no Zepp para duplicatas (prioriza Aprovado)
//   #4 Concatena todos os romaneios quando título está em múltiplos
//   #5 Normaliza acento em 'concluído' para comparação segura
//   #6 Data de emissão com fallback no Sienge
// ---------------------------------------------------------

// Normaliza string removendo acentos para comparações seguras
const removeAccents = (str) =>
  String(str || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();

// Seleciona o melhor match do Zepp quando há duplicatas:
// Prioridade: Aprovado/Concluido > outros > primeiro disponível
const getBestZeppMatch = (matches) => {
  if (!matches || matches.length === 0) return null;
  if (matches.length === 1) return matches[0];
  // Prioridade 1: Aprovado ou Concluído
  const aprovado = matches.find(r => {
    const s = removeAccents(r['Status'] || '');
    return s.includes('aprovado') || s.includes('concluido');
  });
  if (aprovado) return aprovado;
  // Prioridade 2: primeiro disponível (evita reprovações antigas no topo)
  return matches[matches.length - 1];
};

// Concatena todos os números de romaneio quando há duplicatas
const getAllRomaneioNums = (matches) => {
  if (!matches || matches.length === 0) return 'Sem Romaneio';
  const nums = matches
    .map(r => r['Nº ROMANEIO'] || r['Nº Romaneio'] || r['Romaneio'] || '')
    .filter(Boolean);
  const uniqNums = [...new Set(nums)];
  return uniqNums.length > 0 ? uniqNums.join(' | ') : 'Encontrado';
};

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

  // Conjunto dos IDs processados (para detectar órfãos no Zepp depois)
  const processedSiengeIds = new Set();

  siengeData.forEach(siengeRow => {
    // CORREÇÃO #2: Guard clause com registro ao invés de descarte silencioso
    if (!siengeRow['Título'] && !siengeRow['Credor']) {
      console.warn('[processTitulos] Linha ignorada por falta de Título e Credor:', siengeRow);
      return;
    }

    const tituloSienge = String(siengeRow['Título'] || '').trim();
    if (tituloSienge) processedSiengeIds.add(tituloSienge);
    const cv = buildCredorValorKey(siengeRow);

    const zeppMatchesAll = (tituloSienge && zeppMapTitulo[tituloSienge])
      ? zeppMapTitulo[tituloSienge]
      : filterFallbackMatches(zeppMapCredorValor[cv] || [], tituloSienge);

    const romaneioMatchesAll = (tituloSienge && romaneioMapTitulo[tituloSienge])
      ? romaneioMapTitulo[tituloSienge]
      : filterFallbackMatches(romaneioMapCredorValor[cv] || [], tituloSienge);

    // CORREÇÃO #3: Usa melhor match no Zepp para duplicatas
    const bestZepp = getBestZeppMatch(zeppMatchesAll);
    const inZepp = !!bestZepp;
    const inRomaneio = romaneioMatchesAll.length > 0;

    let acao = '';
    let statusZepp = inZepp ? (bestZepp['Status'] || 'Aprovado') : 'Não encontrado';
    let vencimentoZepp = inZepp ? (bestZepp['Dt. Vencto'] || bestZepp['Dt. vencto'] || '') : '';

    // CORREÇÃO #4: Concatena todos os romaneios quando há duplicatas
    let noRomaneio = inRomaneio ? getAllRomaneioNums(romaneioMatchesAll) : 'Sem Romaneio';

    // CORREÇÃO #6: Fallback de data de emissão no Sienge quando não encontrado no Zepp
    let dataEmissao = inZepp
      ? (bestZepp['Emissão'] || bestZepp['emissão'] || '')
      : (siengeRow['Data emissão'] || siengeRow['Data de emissão'] || siengeRow['Data Emissão'] || '');

    let tipoDocumento = siengeRow['Documento'] || '';

    // CORREÇÃO #5: Normaliza acento em 'concluído' para comparação segura
    const zeppStatusNorm = removeAccents(statusZepp);

    if (inRomaneio && inZepp && (zeppStatusNorm.includes('aprovado') || zeppStatusNorm.includes('concluido'))) {
      acao = 'OK: Lançado no Romaneio e Enviado'; kpi.pronto++;
    } else if (inRomaneio && inZepp) {
      acao = 'ALERTA: Em Aprovação no Zepp'; kpi.aprovacao++;
    } else if (inZepp && !inRomaneio) {
      acao = 'ALERTA: Falta Romaneio'; zeppStatusNorm.includes('aprov') ? kpi.acao++ : kpi.aprovacao++;
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
      originalZepp: inZepp ? bestZepp : null,
      originalSienge: siengeRow
    });
  });

  // CORREÇÃO #1: Varredura reversa — encontrar títulos APROVADOS no Zepp
  // que não existem no Sienge (ex: título 97588 removido da extração do Sienge).
  // REGRAS ESTRITAS para evitar poluição do resultado:
  //   1. ID deve ser numérico puro (ex: "97588") — exclui "CLSF/XXX", "PPC/XXX"
  //   2. Status deve ser Aprovado ou Concluído (exclui Reprovado — esses saíram por motivo)
  //   3. ID não pode já ter sido processado via Sienge
  const zeppOrphans = new Set();
  zeppData.forEach(row => {
    const titulo = String(row['Código Origem'] || '').split('/')[0].trim();
    if (!titulo || titulo === 'nan') return;

    // Regra 1: ID deve ser estritamente numérico (título do Sienge é sempre número)
    if (!/^\d+$/.test(titulo)) return;

    if (processedSiengeIds.has(titulo)) return; // já processado via Sienge
    if (zeppOrphans.has(titulo)) return; // já adicionado como órfão

    const statusNorm = removeAccents(row['Status'] || '');

    // Regra 2: Só inclui Aprovado ou Concluído — Reprovados saíram do Sienge por motivo legítimo
    const isAprovado = statusNorm.includes('aprovado') || statusNorm.includes('concluido');
    if (!isAprovado) return;

    zeppOrphans.add(titulo);

    const romaneioMatchesAll = romaneioMapTitulo[titulo] || [];
    const inRomaneio = romaneioMatchesAll.length > 0;
    const noRomaneio = inRomaneio ? getAllRomaneioNums(romaneioMatchesAll) : 'Sem Romaneio';
    const statusZepp = row['Status'] || 'Aprovado';
    const zeppStatusNorm = removeAccents(statusZepp);

    let acao = '';
    if (inRomaneio && (zeppStatusNorm.includes('aprovado') || zeppStatusNorm.includes('concluido'))) {
      acao = 'OK: Lançado no Romaneio e Enviado';
      kpi.pronto++;
    } else if (inRomaneio) {
      acao = 'ALERTA: Em Aprovação no Zepp';
      kpi.aprovacao++;
    } else if (zeppStatusNorm.includes('aprovado') || zeppStatusNorm.includes('concluido')) {
      acao = 'ALERTA: Aprovado no Zepp — Não encontrado no Sienge';
      kpi.acao++;
    } else {
      acao = 'ALERTA: No Zepp — Não encontrado no Sienge';
      kpi.acao++;
    }

    kpi.total++;
    results.push({
      id: titulo,
      tipoDocumento: '-',
      credor: row['Credor'] || '-',
      dataEmissao: row['Emissão'] || row['emissão'] || '-',
      vencimento: row['Dt. Vencto'] || row['Dt. vencto'] || '-',
      numeroNF: '-',
      valor: getValor(row),
      statusZepp,
      noRomaneio,
      observacao: '⚠️ Título não encontrado na extração do Sienge',
      acao,
      originalZepp: row,
      originalSienge: null
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
export const parseSiengeBoletimMedicoes = (rawRows) => {
  const medicoes = [];
  let currentObra = '';
  let i = 0;

  while (i < rawRows.length) {
    const row = rawRows[i] || [];
    const rowStr = row.map(c => c !== null && c !== undefined ? String(c).trim() : '');

    if (rowStr[0] === 'Obra') {
      currentObra = rowStr[2] || '';
      i++;
      continue;
    }

    if (rowStr[0] === 'Contrato' && rowStr[1] === 'Fornecedor') {
      const dataRow = rawRows[i + 1] || [];
      const dStr = dataRow.map(c => c !== null && c !== undefined ? String(c).trim() : '');

      const contrato = dStr[0] || '';
      const fornecedor = dStr[1] || '';
      const numMedicao = dStr[4] || '';
      const dtMedicao = dStr[6] || '';
      const dtVencto = dStr[7] || '';

      let totalBruto = 0;
      let totalLiquido = 0;
      let totalImpostos = 0;
      let caucao = 0;
      let descontos = 0;
      let obsText = '';

      let j = i + 1;
      while (j < rawRows.length) {
        const rScan = rawRows[j] || [];
        const rScanStr = rScan.map(c => c !== null && c !== undefined ? String(c).trim() : '');

        if (j > i + 1 && rScanStr[0] === 'Contrato' && rScanStr[1] === 'Fornecedor') break;
        if (j > i + 1 && rScanStr[0] === 'Obra') break;

        // Observação
        if (rScanStr[3] && (rScanStr[3].includes('Medição') || rScanStr[3].includes('DESCONTOS') || rScanStr[3].includes('Abatimento') || rScanStr[3].includes('referente') || rScanStr[3].includes('Medicao'))) {
          obsText += (' ' + rScanStr[3]).trim();
        }

        for (let cIdx = 0; cIdx < rScanStr.length; cIdx++) {
          const cellVal = rScanStr[cIdx];
          if (cellVal === 'Total bruto' && cIdx + 1 < rScan.length) {
            totalBruto = normalizeNum(rScan[cIdx + 1]);
          } else if (cellVal === 'Total líquido' && cIdx + 1 < rScan.length) {
            totalLiquido = normalizeNum(rScan[cIdx + 1]);
          } else if (cellVal === 'Total de impostos retido' && cIdx + 1 < rScan.length) {
            totalImpostos = normalizeNum(rScan[cIdx + 1]);
          } else if (cellVal === 'Caução' && cIdx + 1 < rScan.length) {
            caucao = normalizeNum(rScan[cIdx + 1]);
          } else if (cellVal === 'Descontos' && cIdx + 1 < rScan.length) {
            descontos = normalizeNum(rScan[cIdx + 1]);
          }
        }

        j++;
        if (j - i > 40) break;
      }

      // Regex para extrair "Abatimento de Sinal" da observação (Regra #4)
      let abatimentoSinal = 0;
      const match = obsText.match(/Abatimento\s+de\s+Sinal\s*[=:]\s*R?\$?\s*([\d\.,]+)/i);
      if (match) {
        abatimentoSinal = normalizeNum(match[1]);
      }

      if (contrato || fornecedor) {
        medicoes.push({
          obra: currentObra,
          contrato,
          fornecedor,
          numMedicao,
          dtMedicao,
          dtVencto,
          totalBruto,
          totalImpostos,
          retencao: caucao, // Regra #3: Retenção = Caução
          descontoSinal: abatimentoSinal, // Regra #4: Desconto de Sinal = Abatimento de Sinal
          descontosFD: descontos,
          outrosDescontos: 0,
          totalLiquido,
          calculadoLiquido: totalBruto - totalImpostos - caucao - abatimentoSinal - descontos,
          observacao: obsText.trim()
        });
      }

      i = j;
    } else {
      i++;
    }
  }

  return medicoes;
};

const processMedicoes = (siengeData, zeppData, romaneioData, rawSiengeRows = []) => {
  // 1. Extrair medições do Sienge (Boletim ou Tabela)
  let siengeMedicoes = [];
  if (rawSiengeRows && rawSiengeRows.length > 0) {
    siengeMedicoes = parseSiengeBoletimMedicoes(rawSiengeRows);
  }
  
  if (siengeMedicoes.length === 0 && siengeData && siengeData.length > 0) {
    siengeData.forEach(row => {
      const contrato = row['Contrato'] || row['ID CONTRATO'] || '';
      const numMedicao = row['Medição'] || row['MEDIÇÃO'] || '';
      const fornecedor = row['Fornecedor*'] || row['Fornecedor'] || row['RAZÃO SOCIAL'] || row['Razão Social'] || row['Credor'] || '';
      if (!contrato && !numMedicao && !fornecedor) return;

      const valorBruto = normalizeNum(row['Total bruto'] || row['VALOR TOTAL'] || row['Valor'] || 0);
      const imposto = normalizeNum(row['Total de impostos retido'] || row['IMPOSTO'] || 0);
      const retencao = normalizeNum(row['Caução'] || row['RETENÇÃO'] || 0);
      const obs = String(row['Observação'] || row['Observação da medição'] || row['observacao'] || '');
      let sinal = normalizeNum(row['DESCONTO DE SINAL'] || row['Desconto de Sinal'] || 0);
      if (sinal === 0) {
        const m = obs.match(/Abatimento\s+de\s+Sinal\s*[=:]\s*R?\$?\s*([\d\.,]+)/i);
        if (m) sinal = normalizeNum(m[1]);
      }
      const descontosFD = normalizeNum(row['DESCONTOS FD'] || row['Descontos FD'] || row['Descontos'] || 0);
      const outrosDescontos = normalizeNum(row['OUTROS DESCONTOS'] || row['Outros Descontos'] || 0);
      const valorLiquido = valorBruto - imposto - retencao - sinal - descontosFD - outrosDescontos;

      siengeMedicoes.push({
        contrato,
        numMedicao,
        fornecedor,
        totalBruto: valorBruto,
        totalImpostos: imposto,
        retencao,
        descontoSinal: sinal,
        descontosFD,
        outrosDescontos,
        totalLiquido: normalizeNum(row['Total líquido'] || row['VALOR LÍQUIDO'] || valorLiquido),
        calculadoLiquido: valorLiquido,
        dtVencto: row['Data de vencimento'] || row['Vencimento'] || '',
        observacao: obs
      });
    });
  }

  // 2. Normalizar base de Controle (romaneioData / CONTROLE DE MEDIÇÕES.xlsx)
  const cleanControle = (romaneioData || []).map(row => {
    const cleaned = {};
    Object.keys(row).forEach(k => {
      cleaned[k.trim()] = row[k];
    });
    return cleaned;
  });

  const controleMapKey = {};
  const controleMapValorFornec = {};
  const controleMapValor = {};
  const matchedControleIndexes = new Set();

  cleanControle.forEach((c, idx) => {
    const contrato = String(c['ID CONTRATO'] || '').replace(/\s+/g, '').toUpperCase();
    const med = String(c['MEDIÇÃO'] || '').replace(/\s+/g, '');
    if (contrato && med) {
      controleMapKey[`${contrato}_${med}`] = { row: c, idx };
    }
    const valorTotalStr = normalizeNum(c['VALOR TOTAL']).toFixed(2);
    const fornecNorm = normalizeStr(c['RAZÃO SOCIAL'] || '');
    if (fornecNorm && valorTotalStr !== '0.00') {
      controleMapValorFornec[`${fornecNorm}_${valorTotalStr}`] = { row: c, idx };
    }
    if (valorTotalStr !== '0.00' && !controleMapValor[valorTotalStr]) {
      controleMapValor[valorTotalStr] = { row: c, idx };
    }
  });

  // 3. Normalizar base Zepp (ZEPP_MEDIÇÕES.xlsx)
  const zeppMapKey = {};
  const zeppMapValor = {};
  const matchedZeppIndexes = new Set();

  (zeppData || []).forEach((z, idx) => {
    const codOrigem = String(z['Código Origem'] || '').trim();
    const parts = codOrigem.split('/').map(p => p.trim()).filter(Boolean);
    let contract = '';
    let medicao = '';
    if (parts.length >= 4) {
      contract = `${parts[0]}/${parts[1]}`.replace(/\s+/g, '').toUpperCase();
      medicao = parts[3].replace(/\s+/g, '');
    } else if (parts.length === 3) {
      contract = parts[0].replace(/\s+/g, '').toUpperCase();
      medicao = parts[2].replace(/\s+/g, '');
    } else {
      contract = codOrigem.replace(/\s+/g, '').toUpperCase();
    }

    if (contract && medicao) {
      const key = `${contract}_${medicao}`;
      if (!zeppMapKey[key]) zeppMapKey[key] = [];
      zeppMapKey[key].push({ row: z, idx });
    }

    const valorZepp = normalizeNum(z['Valor']);
    const valorStr = valorZepp.toFixed(2);
    if (valorStr !== '0.00') {
      if (!zeppMapValor[valorStr]) zeppMapValor[valorStr] = [];
      zeppMapValor[valorStr].push({ row: z, idx });
    }
  });

  const results = [];
  let valorTotalMedido = 0;
  let valorAprovado = 0;
  let valorEmAprovacao = 0;
  let valorLancado = 0;
  let valorFaltaLancar = 0;

  // Processar itens do Sienge
  siengeMedicoes.forEach(sm => {
    const contractNorm = String(sm.contrato || '').replace(/\s+/g, '').toUpperCase();
    const medNorm = String(sm.numMedicao || '').replace(/\s+/g, '');
    const key = `${contractNorm}_${medNorm}`;
    const valorTotalStr = normalizeNum(sm.totalBruto).toFixed(2);
    const fornecNorm = normalizeStr(sm.fornecedor || '');

    // Buscar no Zepp
    let zMatch = null;
    if (zeppMapKey[key] && zeppMapKey[key].length > 0) {
      zMatch = zeppMapKey[key][0];
    } else if (zeppMapValor[valorTotalStr] && zeppMapValor[valorTotalStr].length > 0) {
      zMatch = zeppMapValor[valorTotalStr][0];
    }

    if (zMatch) {
      matchedZeppIndexes.add(zMatch.idx);
    }

    const zRow = zMatch ? zMatch.row : null;
    const statusZepp = zRow ? (zRow['Status Atual Proc.'] || zRow['Status'] || 'Aprovado') : 'Não encontrado';
    const statusZeppLower = statusZepp.toLowerCase();

    // Buscar no Controle
    let cMatch = null;
    if (controleMapKey[key]) {
      cMatch = controleMapKey[key];
    } else if (controleMapValorFornec[`${fornecNorm}_${valorTotalStr}`]) {
      cMatch = controleMapValorFornec[`${fornecNorm}_${valorTotalStr}`];
    } else if (controleMapValor[valorTotalStr]) {
      cMatch = controleMapValor[valorTotalStr];
    }

    if (cMatch) {
      matchedControleIndexes.add(cMatch.idx);
    }
    const cRow = cMatch ? cMatch.row : null;

    const valorBruto = sm.totalBruto;
    const valorImposto = sm.totalImpostos;
    const valorRetencao = sm.retencao;
    const valorDescontoSinal = sm.descontoSinal;
    const valorDescontosFD = sm.descontosFD;
    const valorOutrosDescontos = cRow ? normalizeNum(cRow['OUTROS DESCONTOS']) : 0;
    const valorLiquido = sm.calculadoLiquido;

    valorTotalMedido += valorBruto;
    valorLancado += valorBruto;

    let acao = '';
    if (zRow && (statusZeppLower.includes('aprovado') || statusZeppLower.includes('concluido') || statusZeppLower.includes('concluído'))) {
      valorAprovado += valorBruto;
      acao = 'OK: Lançado e Aprovado';
    } else if (zRow && (statusZeppLower.includes('reprovado') || statusZeppLower.includes('cancelado'))) {
      acao = 'ALERTA: Reprovado no Zepp';
    } else if (zRow) {
      valorEmAprovacao += valorBruto;
      acao = 'ALERTA: Em Aprovação no Zepp';
    } else {
      acao = 'ALERTA: Lançado sem Workflow Zepp';
    }

    results.push({
      medicao: sm.numMedicao || (cRow ? cRow['MEDIÇÃO'] : '-'),
      cnpj: cRow ? (cRow['CNPJ'] || '-') : '-',
      credor: sm.fornecedor || (cRow ? cRow['RAZÃO SOCIAL'] : '-'),
      idContrato: sm.contrato || (cRow ? cRow['ID CONTRATO'] : '-'),
      valorContrato: cRow ? normalizeNum(cRow['VALOR DO CONTRATO']) : 0,
      valorTotal: valorBruto,
      imposto: valorImposto,
      retencao: valorRetencao,
      descontoSinal: valorDescontoSinal,
      descontosFD: valorDescontosFD,
      outrosDescontos: valorOutrosDescontos,
      valorLiquido: valorLiquido,
      statusZepp: statusZepp,
      acao: acao,
      // Campos de compatibilidade
      id: `${sm.contrato || (cRow ? cRow['ID CONTRATO'] : '')} / ${sm.numMedicao || (cRow ? cRow['MEDIÇÃO'] : '')}`,
      valor: valorLiquido,
      vencimento: sm.dtVencto || sm.dtMedicao || (cRow ? cRow['Vencimento'] : '-'),
      noRomaneio: cRow ? 'Encontrado' : 'Não Encontrado',
      observacao: sm.observacao || '-'
    });
  });

  // Verificar itens de Controle que não foram encontrados no Sienge
  cleanControle.forEach((c, idx) => {
    if (!matchedControleIndexes.has(idx)) {
      const valorBruto = normalizeNum(c['VALOR TOTAL']);
      const valorLiquido = normalizeNum(c['VALOR LÍQUIDO']);
      valorTotalMedido += valorBruto;
      valorFaltaLancar += valorBruto;

      results.push({
        medicao: c['MEDIÇÃO'] || '-',
        cnpj: c['CNPJ'] || '-',
        credor: c['RAZÃO SOCIAL'] || '-',
        idContrato: c['ID CONTRATO'] || '-',
        valorContrato: normalizeNum(c['VALOR DO CONTRATO']),
        valorTotal: valorBruto,
        imposto: normalizeNum(c['IMPOSTO']),
        retencao: normalizeNum(c['RETENÇÃO']),
        descontoSinal: normalizeNum(c['DESCONTO DE SINAL']),
        descontosFD: normalizeNum(c['DESCONTOS FD']),
        outrosDescontos: normalizeNum(c['OUTROS DESCONTOS']),
        valorLiquido: valorLiquido,
        statusZepp: 'Pendente Lançamento',
        acao: 'ALERTA: Falta Lançar no Sienge',
        id: `${c['ID CONTRATO']} / ${c['MEDIÇÃO']}`,
        valor: valorLiquido,
        vencimento: '-',
        noRomaneio: 'Encontrado no Controle',
        observacao: 'Item presente na planilha de Controle mas não localizado no Boletim do Sienge.'
      });
    }
  });

  const kpi = {
    valorTotalMedido,
    valorAprovado,
    valorEmAprovacao,
    valorLancado,
    valorFaltaLancar,
    total: results.length,
    pronto: results.filter(r => r.acao && r.acao.startsWith('OK')).length,
    aprovacao: results.filter(r => r.acao && r.acao.includes('Em Aprovação')).length,
    acao: results.filter(r => r.acao && r.acao.startsWith('ALERTA')).length
  };

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
    case 'Medições': {
      const rawSiengeRows = files.sienge ? await readExcelFileRawRows(files.sienge) : [];
      return processMedicoes(siengeData, zeppData, romaneioData, rawSiengeRows);
    }
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

  const cleanForInclude = (t) => t.replace(/[^a-z0-9]/g, '');
  const checkSupplierMatch = (s1, s2, isStrict = false) => {
      if (!s1 || !s2) return false;
      const c1 = cleanForInclude(s1);
      const c2 = cleanForInclude(s2);
      if (c1 && c2 && (c1.includes(c2) || c2.includes(c1))) return true;

      if (isStrict) return false;

      const w1 = s1.replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter(w => w.length > 2);
      const w2 = s2.replace(/[^a-z0-9\s]/g, ' ').split(/\s+/).filter(w => w.length > 2);
      let commonCount = 0;
      for (let w of w1) {
          if (w2.includes(w)) commonCount++;
      }
      if (w1.length <= 1 || w2.length <= 1) return commonCount >= 1;
      return commonCount >= 2;
  };

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
               if (val.nf === nfStr && checkSupplierMatch(fornStr, val.fornecedor, false)) {
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
             if (val.fornecedor && checkSupplierMatch(fornStr, val.fornecedor, true)) {
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
