import React, { useState } from 'react';
import { Search, Download, FileText, RefreshCw } from 'lucide-react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import 'jspdf-autotable';

export default function Dashboard({ data, kpi, onReimport, categoryName, finalWorkbook }) {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState('Todos os Status');
  const [columnFilters, setColumnFilters] = useState({
    id: '',
    credor: '',
    cnpj: '',
    apropriacao: '',
    vencimento: '',
    valor: '',
    vinculoFaturamentoDireto: '',
    statusZepp: '',
    noRomaneio: '',
    observacao: '',
    acao: '',
    tipoDocumento: '',
    dataEmissao: '',
    numeroNF: '',
    valorAdiantamento: '',
    valorDescontado: '',
    valorRetencao: '',
    // Medições
    medicao: '',
    idContrato: '',
    valorContrato: '',
    valorTotal: '',
    imposto: '',
    retencao: '',
    descontoSinal: '',
    descontosFD: '',
    outrosDescontos: '',
    valorLiquido: ''
  });

  const handleColumnFilterChange = (col, value) => {
    setColumnFilters(prev => ({ ...prev, [col]: value }));
  };

  const isFaltaRomaneioItem = (item) => {
    if (!item.acao) return false;
    const isAprovadoZepp = String(item.statusZepp || '').toLowerCase().includes('aprovado');
    const isFaltaRomaneio = item.acao.includes('Falta Romaneio') || item.acao.includes('Falta na Base') || item.acao.includes('Falta Base') || item.acao.includes('Falta Lançar') || item.acao.includes('Aguardando Lançamento');
    
    // Normalizando a observação para evitar erros de acentuação
    const obsLower = String(item.observacao || '').toLowerCase();
    const isCaucao = obsLower.includes('caução referente à medição') || obsLower.includes('caucao referente a medicao');
    
    return isAprovadoZepp && isFaltaRomaneio && !isCaucao;
  };

  const faltaRomaneioData = data.filter(isFaltaRomaneioItem);
  const faltaRomaneioCount = faltaRomaneioData.length;

  const formatCurrency = (val) => {
    if (val === null || val === undefined || isNaN(val)) return '';
    return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val);
  };

  // Filtragem local
  const filteredData = data.filter(item => {
    const matchesSearch = String(item.credor || '').toLowerCase().includes(searchTerm.toLowerCase()) || 
                          String(item.noRomaneio || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
                          String(item.id || '').toLowerCase().includes(searchTerm.toLowerCase()) ||
                          String(item.idContrato || '').toLowerCase().includes(searchTerm.toLowerCase());
    
    const matchesStatus = statusFilter === 'Todos os Status' || 
                          (statusFilter === 'OK' && item.acao && item.acao.startsWith('OK')) ||
                          (statusFilter === 'Alerta' && item.acao && (item.acao.startsWith('ALERTA') || item.acao.includes('Aguardando'))) ||
                          (statusFilter === 'Falta Romaneio' && isFaltaRomaneioItem(item));

    if (categoryName === 'Medições') {
      const matchesMedicao = !columnFilters.medicao || String(item.medicao || '').toLowerCase().includes(columnFilters.medicao.toLowerCase());
      const matchesCnpj = !columnFilters.cnpj || String(item.cnpj || '').toLowerCase().includes(columnFilters.cnpj.toLowerCase());
      const matchesCredor = !columnFilters.credor || String(item.credor || '').toLowerCase().includes(columnFilters.credor.toLowerCase());
      const matchesIdContrato = !columnFilters.idContrato || String(item.idContrato || '').toLowerCase().includes(columnFilters.idContrato.toLowerCase());
      const matchesValorContrato = !columnFilters.valorContrato || formatCurrency(item.valorContrato).toLowerCase().includes(columnFilters.valorContrato.toLowerCase());
      const matchesValorTotal = !columnFilters.valorTotal || formatCurrency(item.valorTotal).toLowerCase().includes(columnFilters.valorTotal.toLowerCase());
      const matchesImposto = !columnFilters.imposto || formatCurrency(item.imposto).toLowerCase().includes(columnFilters.imposto.toLowerCase());
      const matchesRetencao = !columnFilters.retencao || formatCurrency(item.retencao).toLowerCase().includes(columnFilters.retencao.toLowerCase());
      const matchesDescontoSinal = !columnFilters.descontoSinal || formatCurrency(item.descontoSinal).toLowerCase().includes(columnFilters.descontoSinal.toLowerCase());
      const matchesDescontosFD = !columnFilters.descontosFD || formatCurrency(item.descontosFD).toLowerCase().includes(columnFilters.descontosFD.toLowerCase());
      const matchesOutrosDescontos = !columnFilters.outrosDescontos || formatCurrency(item.outrosDescontos).toLowerCase().includes(columnFilters.outrosDescontos.toLowerCase());
      const matchesValorLiquido = !columnFilters.valorLiquido || formatCurrency(item.valorLiquido).toLowerCase().includes(columnFilters.valorLiquido.toLowerCase());
      const matchesStatusZepp = !columnFilters.statusZepp || String(item.statusZepp || '').toLowerCase().includes(columnFilters.statusZepp.toLowerCase());
      const matchesAcao = !columnFilters.acao || String(item.acao || '').toLowerCase().includes(columnFilters.acao.toLowerCase());

      return matchesSearch && matchesStatus && matchesMedicao && matchesCnpj && matchesCredor && matchesIdContrato && matchesValorContrato && matchesValorTotal && matchesImposto && matchesRetencao && matchesDescontoSinal && matchesDescontosFD && matchesOutrosDescontos && matchesValorLiquido && matchesStatusZepp && matchesAcao;
    }

    const matchesId = String(item.id || '').toLowerCase().includes(columnFilters.id.toLowerCase());
    const matchesCredor = String(item.credor || '').toLowerCase().includes(columnFilters.credor.toLowerCase());
    const matchesCnpj = String(item.cnpj || '').toLowerCase().includes(columnFilters.cnpj.toLowerCase());
    const matchesApropriacao = String(item.apropriacao || '').toLowerCase().includes(columnFilters.apropriacao.toLowerCase());
    const matchesVencimento = String(item.vencimento || '').toLowerCase().includes(columnFilters.vencimento.toLowerCase());
    const matchesValor = formatCurrency(item.valor).toLowerCase().includes(columnFilters.valor.toLowerCase());
    const matchesVinculo = String(item.vinculoFaturamentoDireto || '').toLowerCase().includes(columnFilters.vinculoFaturamentoDireto.toLowerCase());
    const matchesStatusZepp = String(item.statusZepp || '').toLowerCase().includes(columnFilters.statusZepp.toLowerCase());
    const matchesNoRomaneio = String(item.noRomaneio || '').toLowerCase().includes(columnFilters.noRomaneio.toLowerCase());
    const matchesObservacao = String(item.observacao || '').toLowerCase().includes(columnFilters.observacao.toLowerCase());
    const matchesAcao = String(item.acao || '').toLowerCase().includes(columnFilters.acao.toLowerCase());
    const matchesTipoDocumento = String(item.tipoDocumento || '').toLowerCase().includes((columnFilters.tipoDocumento || '').toLowerCase());
    const matchesDataEmissao = String(item.dataEmissao || '').toLowerCase().includes((columnFilters.dataEmissao || '').toLowerCase());
    const matchesNumeroNF = String(item.numeroNF || '').toLowerCase().includes((columnFilters.numeroNF || '').toLowerCase());
    const matchesValorAdiantamento = !columnFilters.valorAdiantamento || formatCurrency(item.valorAdiantamento).toLowerCase().includes(columnFilters.valorAdiantamento.toLowerCase());
    const matchesValorDescontado = !columnFilters.valorDescontado || formatCurrency(item.valorDescontado).toLowerCase().includes(columnFilters.valorDescontado.toLowerCase());
    const matchesValorRetencao = !columnFilters.valorRetencao || formatCurrency(item.valorRetencao).toLowerCase().includes(columnFilters.valorRetencao.toLowerCase());
                          
    return matchesSearch && matchesStatus && matchesId && matchesCredor && matchesCnpj && matchesApropriacao && matchesVencimento && matchesValor && matchesVinculo && matchesStatusZepp && matchesNoRomaneio && matchesObservacao && matchesAcao && matchesTipoDocumento && matchesDataEmissao && matchesNumeroNF && matchesValorAdiantamento && matchesValorDescontado && matchesValorRetencao;
  });

  const getBadgeClass = (acao) => {
    if (!acao) return 'badge-warning';
    if (acao.startsWith('OK')) return 'badge-success';
    if (acao.includes('Aguardando Lançamento') || acao.includes('Em Aprovação')) return 'badge-warning';
    if (acao.includes('Falta Romaneio') && acao.includes('Falta Zepp')) return 'badge-danger';
    if (acao.includes('Reprovado') || acao.includes('Falta no Boletim')) return 'badge-danger';
    return 'badge-warning';
  };

  const getIDHeaderName = () => {
    if (categoryName === 'Contratos') return 'Nº Contrato';
    if (categoryName === 'Pedidos') return 'Nº Pedido';
    if (categoryName === 'Medições') return 'Nº Contrato / Medição';
    return 'Título';
  };

  const exportExcel = () => {
    if (categoryName === 'Conciliação Saldos Contábeis' && finalWorkbook) {
      XLSX.writeFile(finalWorkbook, `Conciliacao_Saldos_Contabeis.xlsx`);
      return;
    }
    const ws = XLSX.utils.json_to_sheet(filteredData.map(d => {
      if (categoryName === 'Medições') {
        return {
          'MEDIÇÃO': d.medicao,
          'CNPJ': d.cnpj,
          'RAZÃO SOCIAL': d.credor,
          'ID CONTRATO': d.idContrato,
          'VALOR DO CONTRATO': d.valorContrato,
          'VALOR TOTAL': d.valorTotal,
          'IMPOSTO': d.imposto,
          'RETENÇÃO': d.retencao,
          'DESCONTO DE SINAL': d.descontoSinal,
          'DESCONTOS FD': d.descontosFD,
          'OUTROS DESCONTOS': d.outrosDescontos,
          'VALOR LÍQUIDO': d.valorLiquido,
          'STATUS ZEPP': d.statusZepp,
          'AÇÃO REQUERIDA': d.acao
        };
      }
      if (categoryName === 'Pedidos') {
        return {
          'PEDIDO': d.id,
          'CREDOR': d.credor,
          'CNPJ': d.cnpj,
          'APROPRIAÇÃO': d.apropriacao,
          'VALOR': d.valor,
          'VÍNCULO DE FATURAMENTO DIRETO': d.vinculoFaturamentoDireto,
          'STATUS ZEPP': d.statusZepp,
          'OBSERVAÇÃO': d.observacao,
          'AÇÃO REQUERIDA': d.acao
        };
      }
      return {
        [getIDHeaderName()]: d.id,
        ...(categoryName === 'Títulos' ? { 'Tipo de Documento': d.tipoDocumento } : {}),
        Credor: d.credor,
        ...(categoryName === 'Títulos' ? { 'Data de Emissão': d.dataEmissao } : {}),
        Vencimento: d.vencimento,
        ...(categoryName === 'Títulos' ? { 'Nº da NF': d.numeroNF } : {}),
        Valor: d.valor,
        'Status Zepp': d.statusZepp,
        'Nº Romaneio': d.noRomaneio,
        'Observação': d.observacao,
        'Ação Requerida': d.acao
      };
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Resultados");
    XLSX.writeFile(wb, `Conciliacao_${categoryName}.xlsx`);
  };

  const exportFaltaRomaneio = () => {
    if (faltaRomaneioData.length === 0) {
      alert('Nenhum item com alerta de Falta Romaneio ou Base.');
      return;
    }
    const ws = XLSX.utils.json_to_sheet(faltaRomaneioData.map(d => {
      if (categoryName === 'Medições') {
        return {
          'MEDIÇÃO': d.medicao,
          'CNPJ': d.cnpj,
          'RAZÃO SOCIAL': d.credor,
          'ID CONTRATO': d.idContrato,
          'VALOR DO CONTRATO': d.valorContrato,
          'VALOR TOTAL': d.valorTotal,
          'IMPOSTO': d.imposto,
          'RETENÇÃO': d.retencao,
          'DESCONTO DE SINAL': d.descontoSinal,
          'DESCONTOS FD': d.descontosFD,
          'OUTROS DESCONTOS': d.outrosDescontos,
          'VALOR LÍQUIDO': d.valorLiquido,
          'STATUS ZEPP': d.statusZepp,
          'AÇÃO REQUERIDA': d.acao
        };
      }
      if (categoryName === 'Pedidos') {
        return {
          'PEDIDO': d.id,
          'CREDOR': d.credor,
          'CNPJ': d.cnpj,
          'APROPRIAÇÃO': d.apropriacao,
          'VALOR': d.valor,
          'VÍNCULO DE FATURAMENTO DIRETO': d.vinculoFaturamentoDireto,
          'STATUS ZEPP': d.statusZepp,
          'OBSERVAÇÃO': d.observacao,
          'AÇÃO REQUERIDA': d.acao
        };
      }
      return {
        [getIDHeaderName()]: d.id,
        ...(categoryName === 'Títulos' ? { 'Tipo de Documento': d.tipoDocumento } : {}),
        Credor: d.credor,
        ...(categoryName === 'Títulos' ? { 'Data de Emissão': d.dataEmissao } : {}),
        Vencimento: d.vencimento,
        ...(categoryName === 'Títulos' ? { 'Nº da NF': d.numeroNF } : {}),
        Valor: d.valor,
        'Status Zepp': d.statusZepp,
        'Nº Romaneio': d.noRomaneio,
        'Observação': d.observacao,
        'Ação Requerida': d.acao
      };
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Falta_Romaneio");
    XLSX.writeFile(wb, `Conciliacao_${categoryName}_FaltaRomaneio.xlsx`);
  };

  const exportPDF = () => {
    const doc = new jsPDF('landscape');
    doc.text(`Resultados da Conciliação - ${categoryName}`, 14, 15);
    
    const tableColumn = categoryName === 'Medições'
      ? ["Medição", "CNPJ", "Razão Social", "ID Contrato", "Valor Contrato", "Total", "Imposto", "Retenção", "Sinal", "Desc. FD", "Líquido", "Status Zepp", "Ação"]
      : categoryName === 'Pedidos' 
        ? ["Pedido", "Credor", "CNPJ", "Apropriação", "Valor", "Vínc. Faturamento", "Status Zepp", "Observação", "Ação"]
        : categoryName === 'Títulos'
          ? [getIDHeaderName(), "Tipo de Documento", "Credor", "Data de Emissão", "Vencimento", "Nº da NF", "Valor", "Status Zepp", "Romaneio", "Ação"]
          : [getIDHeaderName(), "Credor", "Vencimento", "Valor", "Status Zepp", "Romaneio", "Ação"];
    const tableRows = [];

    filteredData.forEach(item => {
      const rowData = categoryName === 'Medições'
        ? [item.medicao, item.cnpj, item.credor, item.idContrato, formatCurrency(item.valorContrato), formatCurrency(item.valorTotal), formatCurrency(item.imposto), formatCurrency(item.retencao), formatCurrency(item.descontoSinal), formatCurrency(item.descontosFD), formatCurrency(item.valorLiquido), item.statusZepp, item.acao]
        : categoryName === 'Pedidos'
          ? [item.id, item.credor, item.cnpj, item.apropriacao, formatCurrency(item.valor), item.vinculoFaturamentoDireto, item.statusZepp, item.observacao, item.acao]
          : categoryName === 'Títulos'
            ? [item.id, item.tipoDocumento, item.credor, item.dataEmissao, item.vencimento, item.numeroNF, formatCurrency(item.valor), item.statusZepp, item.noRomaneio, item.acao]
            : [item.id, item.credor, item.vencimento, formatCurrency(item.valor), item.statusZepp, item.noRomaneio, item.acao];
      tableRows.push(rowData);
    });

    doc.autoTable({
      head: [tableColumn],
      body: tableRows,
      startY: 20,
      styles: { fontSize: 7 }
    });
    
    doc.save(`Conciliacao_${categoryName}.pdf`);
  };

  return (
    <div>
      <div className="page-header">
        <h1 className="page-title">Resultados da Conciliação - {categoryName}</h1>
        <button className="btn" onClick={onReimport}>
          <RefreshCw size={16} /> Voltar e Importar Novamente
        </button>
      </div>

      {categoryName === 'Medições' ? (
        <div className="kpi-grid" style={{ gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))' }}>
          <div className="kpi-card" style={{ borderBottom: '4px solid var(--primary-color)' }}>
            <h3 style={{ color: 'var(--primary-color)' }}>Valor Total Medido</h3>
            <div className="value" style={{ color: 'var(--primary-color)', fontSize: '1.4rem' }}>
              {formatCurrency(kpi.valorTotalMedido || 0)}
            </div>
          </div>
          <div className="kpi-card" style={{ borderBottom: '4px solid var(--success-color)' }}>
            <h3 style={{ color: 'var(--success-color)' }}>Aprovado</h3>
            <div className="value" style={{ color: 'var(--success-color)', fontSize: '1.4rem' }}>
              {formatCurrency(kpi.valorAprovado || 0)}
            </div>
          </div>
          <div className="kpi-card" style={{ borderBottom: '4px solid var(--warning-color)' }}>
            <h3 style={{ color: 'var(--warning-color)' }}>Em Aprovação</h3>
            <div className="value" style={{ color: 'var(--warning-color)', fontSize: '1.4rem' }}>
              {formatCurrency(kpi.valorEmAprovacao || 0)}
            </div>
          </div>
          <div className="kpi-card" style={{ borderBottom: '4px solid #3b82f6' }}>
            <h3 style={{ color: '#3b82f6' }}>Lançado</h3>
            <div className="value" style={{ color: '#3b82f6', fontSize: '1.4rem' }}>
              {formatCurrency(kpi.valorLancado || 0)}
            </div>
          </div>
          <div className="kpi-card" style={{ borderBottom: '4px solid var(--danger-color)' }}>
            <h3 style={{ color: 'var(--danger-color)' }}>Falta Lançar</h3>
            <div className="value" style={{ color: 'var(--danger-color)', fontSize: '1.4rem' }}>
              {formatCurrency(kpi.valorFaltaLancar || 0)}
            </div>
          </div>
        </div>
      ) : (
        <div className="kpi-grid">
          <div className="kpi-card">
            <h3>Total de Notas</h3>
            <div className="value">{kpi.total}</div>
          </div>
          <div className="kpi-card" style={{ borderBottom: '4px solid var(--success-color)' }}>
            <h3 style={{ color: 'var(--success-color)' }}>Pronto/Enviado (OK)</h3>
            <div className="value" style={{ color: 'var(--success-color)' }}>{kpi.pronto}</div>
          </div>
          <div className="kpi-card" style={{ borderBottom: '4px solid var(--warning-color)' }}>
            <h3 style={{ color: 'var(--warning-color)' }}>Em Aprovação (Zepp)</h3>
            <div className="value" style={{ color: 'var(--warning-color)' }}>{kpi.aprovacao}</div>
          </div>
          <div className="kpi-card" style={{ borderBottom: '4px solid var(--danger-color)' }}>
            <h3 style={{ color: 'var(--danger-color)' }}>Ação Necessária</h3>
            <div className="value" style={{ color: 'var(--danger-color)' }}>{kpi.acao}</div>
          </div>
          {categoryName !== 'Conciliação Saldos Contábeis' && (
            <div 
              className="kpi-card clickable" 
              style={{ borderBottom: '4px solid #8b5cf6' }}
              onClick={exportFaltaRomaneio}
              title="Clique para baixar a planilha apenas com itens faltando no Romaneio"
            >
              <h3 style={{ color: '#8b5cf6' }}>Falta Romaneio (Baixar)</h3>
              <div className="value" style={{ color: '#8b5cf6' }}>{faltaRomaneioCount}</div>
            </div>
          )}
        </div>
      )}

      <div className="table-container">
        <div className="table-toolbar">
          <div className="search-box">
            <Search className="search-icon" size={18} />
            <input 
              type="text" 
              placeholder="Buscar por título, número, credor..." 
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
          
          <div className="toolbar-actions">
            <select 
              className="btn" 
              style={{ background: 'white' }}
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
            >
              <option value="Todos os Status">Todos os Status</option>
              <option value="OK">Somente OK</option>
              <option value="Alerta">Somente Alertas</option>
              <option value="Falta Romaneio">Somente Falta Romaneio</option>
            </select>
            
            <button className="btn" onClick={exportExcel}>
              <FileText size={16} /> Excel
            </button>
            {categoryName !== 'Conciliação Saldos Contábeis' && (
              <button className="btn" onClick={exportPDF}>
                <Download size={16} /> PDF
              </button>
            )}
          </div>
        </div>

        <div style={{ overflowX: 'auto' }}>
          <table>
            <thead>
              {categoryName === 'Conciliação Saldos Contábeis' ? (
                <tr>
                  <th>
                    <div>Aba / Romaneio</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.noRomaneio} onChange={(e) => handleColumnFilterChange('noRomaneio', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>NF</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.id} onChange={(e) => handleColumnFilterChange('id', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>Fornecedor / Descrição</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.credor} onChange={(e) => handleColumnFilterChange('credor', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>Contrato</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.vencimento} onChange={(e) => handleColumnFilterChange('vencimento', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>Adiantamento (Sinal/Futuro)</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.valorAdiantamento} onChange={(e) => handleColumnFilterChange('valorAdiantamento', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>Descontado</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.valorDescontado} onChange={(e) => handleColumnFilterChange('valorDescontado', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>Valor Líquido</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.valor} onChange={(e) => handleColumnFilterChange('valor', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>Valor Retenção</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.valorRetencao} onChange={(e) => handleColumnFilterChange('valorRetencao', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>Status Zepp</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.statusZepp} onChange={(e) => handleColumnFilterChange('statusZepp', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>Ação Requerida</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.acao} onChange={(e) => handleColumnFilterChange('acao', e.target.value)} className="col-filter" />
                  </th>
                </tr>
              ) : categoryName === 'Medições' ? (
                <tr>
                  <th>
                    <div>MEDIÇÃO</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.medicao} onChange={(e) => handleColumnFilterChange('medicao', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>CNPJ</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.cnpj} onChange={(e) => handleColumnFilterChange('cnpj', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>RAZÃO SOCIAL</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.credor} onChange={(e) => handleColumnFilterChange('credor', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>ID CONTRATO</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.idContrato} onChange={(e) => handleColumnFilterChange('idContrato', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>VALOR DO CONTRATO</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.valorContrato} onChange={(e) => handleColumnFilterChange('valorContrato', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>VALOR TOTAL</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.valorTotal} onChange={(e) => handleColumnFilterChange('valorTotal', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>IMPOSTO</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.imposto} onChange={(e) => handleColumnFilterChange('imposto', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>RETENÇÃO</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.retencao} onChange={(e) => handleColumnFilterChange('retencao', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>DESCONTO DE SINAL</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.descontoSinal} onChange={(e) => handleColumnFilterChange('descontoSinal', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>DESCONTOS FD</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.descontosFD} onChange={(e) => handleColumnFilterChange('descontosFD', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>OUTROS DESCONTOS</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.outrosDescontos} onChange={(e) => handleColumnFilterChange('outrosDescontos', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>VALOR LÍQUIDO</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.valorLiquido} onChange={(e) => handleColumnFilterChange('valorLiquido', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>STATUS ZEPP</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.statusZepp} onChange={(e) => handleColumnFilterChange('statusZepp', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>AÇÃO REQUERIDA</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.acao} onChange={(e) => handleColumnFilterChange('acao', e.target.value)} className="col-filter" />
                  </th>
                </tr>
              ) : categoryName === 'Pedidos' ? (
                <tr>
                  <th>
                    <div>PEDIDO</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.id} onChange={(e) => handleColumnFilterChange('id', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>CREDOR</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.credor} onChange={(e) => handleColumnFilterChange('credor', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>CNPJ</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.cnpj} onChange={(e) => handleColumnFilterChange('cnpj', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>APROPRIAÇÃO</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.apropriacao} onChange={(e) => handleColumnFilterChange('apropriacao', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>VALOR</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.valor} onChange={(e) => handleColumnFilterChange('valor', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>VÍNCULO DE FATURAMENTO DIRETO</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.vinculoFaturamentoDireto} onChange={(e) => handleColumnFilterChange('vinculoFaturamentoDireto', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>STATUS ZEPP</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.statusZepp} onChange={(e) => handleColumnFilterChange('statusZepp', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>OBSERVAÇÃO</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.observacao} onChange={(e) => handleColumnFilterChange('observacao', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>AÇÃO REQUERIDA</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.acao} onChange={(e) => handleColumnFilterChange('acao', e.target.value)} className="col-filter" />
                  </th>
                </tr>
              ) : (
                <tr>
                  <th>
                    <div>{getIDHeaderName()}</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.id} onChange={(e) => handleColumnFilterChange('id', e.target.value)} className="col-filter" />
                  </th>
                  {categoryName === 'Títulos' && (
                    <th>
                      <div>Tipo de Documento</div>
                      <input type="text" placeholder="Filtrar..." value={columnFilters.tipoDocumento} onChange={(e) => handleColumnFilterChange('tipoDocumento', e.target.value)} className="col-filter" />
                    </th>
                  )}
                  <th>
                    <div>Credor</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.credor} onChange={(e) => handleColumnFilterChange('credor', e.target.value)} className="col-filter" />
                  </th>
                  {categoryName === 'Títulos' && (
                    <th>
                      <div>Data de Emissão</div>
                      <input type="text" placeholder="Filtrar..." value={columnFilters.dataEmissao} onChange={(e) => handleColumnFilterChange('dataEmissao', e.target.value)} className="col-filter" />
                    </th>
                  )}
                  <th>
                    <div>Vencimento</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.vencimento} onChange={(e) => handleColumnFilterChange('vencimento', e.target.value)} className="col-filter" />
                  </th>
                  {categoryName === 'Títulos' && (
                    <th>
                      <div>Nº da NF</div>
                      <input type="text" placeholder="Filtrar..." value={columnFilters.numeroNF} onChange={(e) => handleColumnFilterChange('numeroNF', e.target.value)} className="col-filter" />
                    </th>
                  )}
                  <th>
                    <div>Valor</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.valor} onChange={(e) => handleColumnFilterChange('valor', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>Status Zepp</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.statusZepp} onChange={(e) => handleColumnFilterChange('statusZepp', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>Nº Romaneio</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.noRomaneio} onChange={(e) => handleColumnFilterChange('noRomaneio', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>Observação</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.observacao} onChange={(e) => handleColumnFilterChange('observacao', e.target.value)} className="col-filter" />
                  </th>
                  <th>
                    <div>Ação Requerida</div>
                    <input type="text" placeholder="Filtrar..." value={columnFilters.acao} onChange={(e) => handleColumnFilterChange('acao', e.target.value)} className="col-filter" />
                  </th>
                </tr>
              )}
            </thead>
            <tbody>
              {filteredData.length === 0 ? (
                <tr>
                  <td colSpan="14" style={{ textAlign: 'center', padding: '2rem' }}>
                    Nenhum resultado encontrado.
                  </td>
                </tr>
              ) : (
                filteredData.map((item, index) => (
                  <tr key={index}>
                    {categoryName === 'Conciliação Saldos Contábeis' ? (
                      <>
                        <td>{item.noRomaneio}</td>
                        <td>{item.id}</td>
                        <td style={{ fontWeight: 500 }}>{item.credor}</td>
                        <td>{item.vencimento}</td>
                        <td style={{ fontWeight: 600 }}>{formatCurrency(item.valorAdiantamento)}</td>
                        <td style={{ fontWeight: 600 }}>{formatCurrency(item.valorDescontado)}</td>
                        <td style={{ fontWeight: 600 }}>{formatCurrency(item.valor)}</td>
                        <td style={{ fontWeight: 600 }}>{formatCurrency(item.valorRetencao)}</td>
                        <td>{item.statusZepp}</td>
                        <td>
                          <span className={`badge ${getBadgeClass(item.acao)}`}>
                            {item.acao}
                          </span>
                        </td>
                      </>
                    ) : categoryName === 'Medições' ? (
                      <>
                        <td>{item.medicao}</td>
                        <td>{item.cnpj}</td>
                        <td style={{ fontWeight: 500 }}>{item.credor}</td>
                        <td>{item.idContrato}</td>
                        <td style={{ fontWeight: 600 }}>{formatCurrency(item.valorContrato)}</td>
                        <td style={{ fontWeight: 600 }}>{formatCurrency(item.valorTotal)}</td>
                        <td style={{ fontWeight: 600 }}>{formatCurrency(item.imposto)}</td>
                        <td style={{ fontWeight: 600 }}>{formatCurrency(item.retencao)}</td>
                        <td style={{ fontWeight: 600 }}>{formatCurrency(item.descontoSinal)}</td>
                        <td style={{ fontWeight: 600 }}>{formatCurrency(item.descontosFD)}</td>
                        <td style={{ fontWeight: 600 }}>{formatCurrency(item.outrosDescontos)}</td>
                        <td style={{ fontWeight: 600 }}>{formatCurrency(item.valorLiquido)}</td>
                        <td>{item.statusZepp}</td>
                        <td>
                          <span className={`badge ${getBadgeClass(item.acao)}`}>
                            {item.acao}
                          </span>
                        </td>
                      </>
                    ) : categoryName === 'Pedidos' ? (
                      <>
                        <td>{item.id}</td>
                        <td style={{ fontWeight: 500 }}>{item.credor}</td>
                        <td>{item.cnpj}</td>
                        <td>{item.apropriacao}</td>
                        <td style={{ fontWeight: 600 }}>{formatCurrency(item.valor)}</td>
                        <td>{item.vinculoFaturamentoDireto}</td>
                        <td>{item.statusZepp}</td>
                        <td style={{ color: 'var(--text-muted)', fontSize: '0.8rem', maxWidth: '200px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                          {item.observacao}
                        </td>
                        <td>
                          <span className={`badge ${getBadgeClass(item.acao)}`}>
                            {item.acao}
                          </span>
                        </td>
                      </>
                    ) : (
                      <>
                        <td>{item.id}</td>
                        {categoryName === 'Títulos' && (
                          <td>{item.tipoDocumento}</td>
                        )}
                        <td style={{ fontWeight: 500 }}>{item.credor}</td>
                        {categoryName === 'Títulos' && (
                          <td>{item.dataEmissao}</td>
                        )}
                        <td>{item.vencimento}</td>
                        {categoryName === 'Títulos' && (
                          <td>{item.numeroNF}</td>
                        )}
                        <td style={{ fontWeight: 600 }}>{formatCurrency(item.valor)}</td>
                        <td>{item.statusZepp}</td>
                        <td>{item.noRomaneio}</td>
                        <td style={{ color: 'var(--text-muted)', fontSize: '0.8rem', maxWidth: '200px', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>
                          {item.observacao}
                        </td>
                        <td>
                          <span className={`badge ${getBadgeClass(item.acao)}`}>
                            {item.acao}
                          </span>
                        </td>
                      </>
                    )}
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
