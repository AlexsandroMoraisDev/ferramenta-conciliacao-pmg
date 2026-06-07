import * as XLSX from 'xlsx';

export const readExcelFile = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        resolve(json);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (err) => reject(err);
    reader.readAsArrayBuffer(file);
  });
};

// Normalização para comparar chaves
const normalizeStr = (str) => {
  if (str === undefined || str === null) return '';
  return String(str).toLowerCase().trim().replace(/[\.\-\/]/g, '');
};

const normalizeNum = (val) => {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  const num = parseFloat(String(val).replace(/\./g, '').replace(',', '.'));
  return isNaN(num) ? 0 : num;
};

// Funções para extrair chaves principais de qualquer das planilhas baseado no que existir
const getNota = (row) => {
  return normalizeStr(row['NºNOTA'] || row['Nº documento'] || row['Nº Nota'] || row['Nota Fiscal'] || '');
};

const getCredor = (row) => {
  return normalizeStr(row['Credor'] || row['RAZÃO SOCIAL'] || row['Fornecedor'] || '');
};

const getValor = (row) => {
  return normalizeNum(row['Valor original'] || row['Valor lquido'] || row['VALOR LIQUIDO+JUROS E MULTAS'] || row['Valor Origem'] || row['Valor Bruto'] || row['Valor'] || row['VALOR BRUTO']);
};

export const processConciliacao = async (files) => {
  const siengeData = files.sienge ? await readExcelFile(files.sienge) : [];
  const zeppData = files.zepp ? await readExcelFile(files.zepp) : [];
  const romaneioData = files.romaneio ? await readExcelFile(files.romaneio) : [];

  // Mapear Romaneio e Zepp por uma chave composta: Nota_Credor_Valor
  // Nota: Como podem haver rateios, se não houver Nota, usamos Credor e Valor.
  
  const buildKey = (row) => {
    const nota = getNota(row);
    const credor = getCredor(row);
    const valor = getValor(row).toFixed(2);
    // Se a nota estiver vazia, tentar cruzar pelo menos pelo valor e credor
    return `${nota}_${credor}_${valor}`;
  };

  const zeppMap = {};
  zeppData.forEach(row => {
    const key = buildKey(row);
    if (!zeppMap[key]) zeppMap[key] = [];
    zeppMap[key].push(row);
  });

  const romaneioMap = {};
  romaneioData.forEach(row => {
    const key = buildKey(row);
    if (!romaneioMap[key]) romaneioMap[key] = [];
    romaneioMap[key].push(row);
  });

  const results = [];
  let kpi = { total: 0, pronto: 0, aprovacao: 0, acao: 0 };

  siengeData.forEach((siengeRow) => {
    const key = buildKey(siengeRow);
    const inZepp = zeppMap[key] && zeppMap[key].length > 0;
    const inRomaneio = romaneioMap[key] && romaneioMap[key].length > 0;

    let acao = '';
    let statusZepp = 'Não encontrado';
    let noRomaneio = 'Sem Romaneio';
    
    // Se encontrou no Zepp
    if (inZepp) {
      statusZepp = zeppMap[key][0]['Status'] || zeppMap[key][0]['Status Zepp'] || 'Aprovado';
    }

    // Se encontrou no Romaneio
    if (inRomaneio) {
      noRomaneio = romaneioMap[key][0]['Nº ROMANEIO'] || romaneioMap[key][0]['Nº Romaneio'] || 'Lançado';
    }

    // Regras de Status conforme imagens
    if (inRomaneio && inZepp) {
      acao = 'OK: Lançado no Romaneio e Enviado';
      kpi.pronto++;
    } else if (inZepp && !inRomaneio) {
      acao = 'ALERTA: Falta Romaneio';
      if (statusZepp.toLowerCase().includes('aprov')) {
         kpi.acao++;
      } else {
         kpi.aprovacao++;
      }
    } else if (inRomaneio && !inZepp) {
      acao = 'ALERTA: Falta Zepp';
      kpi.acao++;
    } else {
      acao = 'ALERTA: Falta Romaneio e Falta Zepp';
      kpi.acao++;
    }

    kpi.total++;

    results.push({
      credor: siengeRow['Credor'] || siengeRow['Credor/Fornecedor'] || 'Desconhecido',
      vencimento: siengeRow['Data competncia'] || siengeRow['Data de vencimento'] || siengeRow['Vencimento'] || '-',
      valor: parseFloat(getValor(siengeRow) || 0),
      statusZepp,
      noRomaneio,
      observacao: siengeRow['Observao'] || siengeRow['Observação'] || '-',
      acao,
      originalSienge: siengeRow
    });
  });

  return { results, kpi };
};
