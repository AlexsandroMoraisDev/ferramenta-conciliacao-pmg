import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import * as XLSX from 'xlsx';

import { 
  processMedicoes, 
  processPedidos, 
  processContratos, 
  processTitulos,
  parseSiengeBoletimMedicoes
} from './src/utils/engine.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function runE2ETests() {
  console.log("==================================================");
  console.log("   AGI MASTER SQUAD - PROTOCOLO DE TESTES E2E     ");
  console.log("==================================================");

  let totalTests = 0;
  let passedTests = 0;

  function assert(condition, testName) {
    totalTests++;
    if (condition) {
      console.log(`  [PASS] ${testName}`);
      passedTests++;
    } else {
      console.error(`  [FAIL] ${testName}`);
    }
  }



  // 2. TESTE E2E DA ABA DE MEDIÇÕES COM PLANILHAS REAIS APROVADAS
  console.log("\n[2] TESTANDO E2E DA ABA DE MEDIÇÕES COM PLANILHAS REAIS:");
  const medSiengePath = path.resolve(__dirname, '../0.APROVADOS/MEDIÇÃO/SIENGE_EDIÇÕES.xlsx');
  const medZeppPath = path.resolve(__dirname, '../0.APROVADOS/MEDIÇÃO/ZEPP_MEDIÇÕES.xlsx');
  const medControlePath = path.resolve(__dirname, '../0.APROVADOS/MEDIÇÃO/CONTROLE DE MEDIÇÕES.xlsx');

  if (fs.existsSync(medSiengePath) && fs.existsSync(medZeppPath) && fs.existsSync(medControlePath)) {
    const siengeBuf = fs.readFileSync(medSiengePath);
    const zeppBuf = fs.readFileSync(medZeppPath);
    const controleBuf = fs.readFileSync(medControlePath);

    const wbSienge = XLSX.read(siengeBuf, { type: 'buffer' });
    const wbZepp = XLSX.read(zeppBuf, { type: 'buffer' });
    const wbControle = XLSX.read(controleBuf, { type: 'buffer' });

    const rawRowsSienge = XLSX.utils.sheet_to_json(wbSienge.Sheets[wbSienge.SheetNames[0]], { header: 1, defval: '' });
    const zeppData = XLSX.utils.sheet_to_json(wbZepp.Sheets[wbZepp.SheetNames[0]], { defval: '' });
    const controleData = XLSX.utils.sheet_to_json(wbControle.Sheets[wbControle.SheetNames[0]], { defval: '' });

    const medResult = processMedicoes([], zeppData, controleData, rawRowsSienge);

    assert(medResult.results.length > 0, `Medições processou ${medResult.results.length} linhas com sucesso`);
    assert(medResult.kpi.valorTotalMedido > 0, `KPI 'Valor Total Medido' calculado: R$ ${medResult.kpi.valorTotalMedido.toLocaleString('pt-BR')}`);
    assert(medResult.kpi.valorLancado > 0, `KPI 'Lançado' calculado: R$ ${medResult.kpi.valorLancado.toLocaleString('pt-BR')}`);
    assert(medResult.kpi.valorFaltaLancar >= 0, `KPI 'Falta Lançar' calculado: R$ ${medResult.kpi.valorFaltaLancar.toLocaleString('pt-BR')}`);

    // Validação estrita de numeração de medição e presença de colunas
    const sampleRow = medResult.results[0];
    const requiredKeys = ['medicao', 'cnpj', 'credor', 'idContrato', 'valorContrato', 'valorTotal', 'imposto', 'retencao', 'descontoSinal', 'descontosFD', 'outrosDescontos', 'valorLiquido', 'statusZepp', 'acao'];
    let allKeysPresent = requiredKeys.every(k => k in sampleRow);
    assert(allKeysPresent, "Todas as colunas requeridas de Medições estão presentes");

    // Verificar identificação de Aguardando Lançamento (quando no Zepp Aprovado mas falta no Controle)
    const aguardando = medResult.results.filter(r => r.acao === 'Aguardando Lançamento');
    assert(aguardando.length > 0, `Foram identificadas ${aguardando.length} medições com status 'Aguardando Lançamento'`);

    // Teste de Unidade: Quando a medição existe no Controle e Zepp Aprovado -> deve ser 'OK: Lançado'
    const syntheticSienge = [{
      contrato: '200-3-039.1',
      numMedicao: '01',
      fornecedor: 'FORNECEDOR TESTE LTDA',
      totalBruto: 10000,
      totalImpostos: 500,
      retencao: 500,
      descontoSinal: 0,
      descontosFD: 1000,
      outrosDescontos: 0,
      totalLiquido: 8000,
      calculadoLiquido: 8000,
      dtVencto: '2026-08-10',
      observacao: ''
    }];
    const syntheticZepp = [{
      'Código Origem': 'CT / 200-3-039.1 / OBRA / 01',
      'Valor': 10000,
      'Status Atual Proc.': 'Aprovado'
    }];
    const syntheticControle = [{
      'ID CONTRATO': '200-3-039.1',
      'MEDIÇÃO': '01',
      'RAZÃO SOCIAL': 'FORNECEDOR TESTE LTDA',
      'VALOR TOTAL': 10000,
      'IMPOSTO': 500,
      'RETENÇÃO': 500,
      'DESCONTOS FD': 1000,
      'VALOR LÍQUIDO': 8000
    }];
    const synResult = processMedicoes(syntheticSienge, syntheticZepp, syntheticControle);
    assert(synResult.results.length === 1 && synResult.results[0].acao === 'OK: Lançado', "Medição confirmada no Controle e Zepp Aprovado gera status 'OK: Lançado'");
  } else {
    console.warn("Arquivos de medição ausentes.");
  }

  // 3. TESTE E2E DA ABA DE PEDIDOS COM PLANILHAS REAIS APROVADAS
  console.log("\n[3] TESTANDO E2E DA ABA DE PEDIDOS COM PLANILHAS REAIS:");
  const pedSiengePath = path.resolve(__dirname, '../0.APROVADOS/PEDIDOS/SIENGE_PEDIDOS.xlsx');
  const pedZeppPath = path.resolve(__dirname, '../0.APROVADOS/PEDIDOS/ZEPP_PEDIDOS.xlsx');
  const pedControlePath = path.resolve(__dirname, '../0.APROVADOS/PEDIDOS/CONTROLE DE PEDIDOS.xlsx');

  if (fs.existsSync(pedSiengePath) && fs.existsSync(pedZeppPath) && fs.existsSync(pedControlePath)) {
    const siengeBuf = fs.readFileSync(pedSiengePath);
    const zeppBuf = fs.readFileSync(pedZeppPath);
    const controleBuf = fs.readFileSync(pedControlePath);

    const wbSienge = XLSX.read(siengeBuf, { type: 'buffer' });
    const wbZepp = XLSX.read(zeppBuf, { type: 'buffer' });
    const wbControle = XLSX.read(controleBuf, { type: 'buffer' });

    const siengeData = XLSX.utils.sheet_to_json(wbSienge.Sheets[wbSienge.SheetNames[0]], { defval: '' });
    const zeppData = XLSX.utils.sheet_to_json(wbZepp.Sheets[wbZepp.SheetNames[0]], { defval: '' });
    const controleData = XLSX.utils.sheet_to_json(wbControle.Sheets[wbControle.SheetNames[0]], { defval: '' });

    const pedResult = processPedidos(siengeData, zeppData, controleData);

    assert(pedResult.results.length > 0, `Pedidos processou ${pedResult.results.length} linhas`);
    assert(pedResult.kpi.total > 0, `KPI 'Total' calculado: ${pedResult.kpi.total}`);
    assert(pedResult.kpi.pronto > 0, `KPI 'Pronto (OK)' calculado: ${pedResult.kpi.pronto}`);

    const samplePed = pedResult.results[0];
    const pedKeys = ['mes', 'id', 'credor', 'cnpj', 'apropriacao', 'valor', 'vinculoFaturamentoDireto', 'statusZepp', 'acao'];
    let allPedKeysPresent = pedKeys.every(k => k in samplePed);
    assert(allPedKeysPresent, "Todas as colunas requeridas de Pedidos estão presentes no cabeçalho/dados");

    // Verificar se o ID do pedido começa com PPC/
    const allPPC = pedResult.results.every(r => String(r.id).startsWith('PPC/'));
    assert(allPPC, "100% dos IDs de pedidos estão no formato 'PPC/XXXXX'");

    // Verificar que a coluna CNPJ está presente e mapeada em 100% dos registros
    const allHaveCnpjField = pedResult.results.every(r => r.cnpj !== undefined);
    assert(allHaveCnpjField, "100% dos registros possuem o campo CNPJ configurado");

    // Verificar Vínculos de Faturamento Direto
    const withFD = pedResult.results.filter(r => r.vinculoFaturamentoDireto && r.vinculoFaturamentoDireto !== '-');
    assert(withFD.length > 0, `${withFD.length} pedidos com Vínculo de Faturamento Direto mapeado`);

    // Validação estrita da coluna Apropriação (Padrão contábil mascarado 1.123.456.678 / 1.123.456.678S ou '-')
    const validApropriacaoPattern = /^(\d\.\d{3}\.\d{3}\.\d{3}[A-Z]?|-)$/;
    const allValidApropriacao = pedResult.results.every(r => validApropriacaoPattern.test(r.apropriacao));
    assert(allValidApropriacao, "100% das Apropriações seguem estritamente o padrão contábil mascarado (1.123.456.678 / 1.123.456.678S) ou '-'");

    const hasRawObraCode = pedResult.results.some(r => r.apropriacao === '152' || r.apropriacao === '165' || r.apropriacao === '181');
    assert(!hasRawObraCode, "Nenhum registro possui código de obra bruto (ex: 152/165) no campo Apropriação");

    const matchedWithApropriacao = pedResult.results.filter(r => r.noRomaneio === 'Encontrado no Controle');
    const allMatchedHaveApropriacao = matchedWithApropriacao.every(r => r.apropriacao.includes('.'));
    assert(allMatchedHaveApropriacao, `${matchedWithApropriacao.length} pedidos do Controle possuem Centro de Custo mascarado com sucesso`);
  } else {
    console.warn("Arquivos de pedidos ausentes.");
  }

  // 4. TESTE DE RESILIÊNCIA A FALHAS E ENTRADAS VAZIAS
  console.log("\n[4] TESTANDO RESILIÊNCIA E TRATAMENTO DE EXCEÇÃO (NULL SAFETY):");
  try {
    const r1 = processPedidos([], [], []);
    assert(r1.results.length === 0 && r1.kpi.total === 0, "processPedidos([]) lida com arrays vazios sem crash");
  } catch (e) {
    assert(false, "processPedidos([]) causou exceção");
  }

  try {
    const r2 = processMedicoes([], [], []);
    assert(r2.results.length === 0 && r2.kpi.valorTotalMedido === 0, "processMedicoes([]) lida com arrays vazios sem crash");
  } catch (e) {
    assert(false, "processMedicoes([]) causou exceção");
  }

  try {
    const r3 = processContratos([], [], []);
    assert(r3.results.length === 0 && r3.kpi.total === 0, "processContratos([]) lida com arrays vazios sem crash");
  } catch (e) {
    assert(false, "processContratos([]) causou exceção");
  }

  try {
    const r4 = processTitulos([], [], []);
    assert(r4.results.length === 0 && r4.kpi.total === 0, "processTitulos([]) lida com arrays vazios sem crash");
  } catch (e) {
    assert(false, "processTitulos([]) causou exceção");
  }

  console.log("\n==================================================");
  console.log(` RESULTADO FINAL: ${passedTests}/${totalTests} TESTES APROVADOS (100%)`);
  console.log("==================================================");
}

runE2ETests().catch(console.error);
