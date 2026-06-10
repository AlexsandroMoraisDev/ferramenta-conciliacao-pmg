import React, { useState } from 'react';
import { Search, Download, FileText, RefreshCw } from 'lucide-react';
import * as XLSX from 'xlsx';
import jsPDF from 'jspdf';
import 'jspdf-autotable';

export default function Dashboard({ data, kpi, onReimport, categoryName }) {
  const [searchTerm, setSearchTerm] = useState('');
  const [statusFilter, setStatusFilter] = useState('Todos os Status');

  const faltaRomaneioData = data.filter(item => {
    const isAprovadoZepp = String(item.statusZepp || '').toLowerCase().includes('aprovado');
    const isFaltaRomaneio = item.acao.includes('Falta Romaneio') || item.acao.includes('Falta na Base') || item.acao.includes('Falta Base');
    
    // Normalizando a observação para evitar erros de acentuação
    const obsLower = String(item.observacao || '').toLowerCase();
    const isCaucao = obsLower.includes('caução referente à medição') || obsLower.includes('caucao referente a medicao');
    
    return isAprovadoZepp && isFaltaRomaneio && !isCaucao;
  });
  const faltaRomaneioCount = faltaRomaneioData.length;

  // Filtragem local
  const filteredData = data.filter(item => {
    const matchesSearch = item.credor.toLowerCase().includes(searchTerm.toLowerCase()) || 
                          item.noRomaneio.toLowerCase().includes(searchTerm.toLowerCase()) ||
                          String(item.id || '').toLowerCase().includes(searchTerm.toLowerCase());
    
    const matchesStatus = statusFilter === 'Todos os Status' || 
                          (statusFilter === 'OK' && item.acao.startsWith('OK')) ||
                          (statusFilter === 'Alerta' && item.acao.startsWith('ALERTA'));
                          
    return matchesSearch && matchesStatus;
  });

  const getBadgeClass = (acao) => {
    if (acao.startsWith('OK')) return 'badge-success';
    if (acao.includes('Falta Romaneio') && acao.includes('Falta Zepp')) return 'badge-danger';
    return 'badge-warning';
  };

  const formatCurrency = (val) => {
    return new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(val);
  };

  const getIDHeaderName = () => {
    if (categoryName === 'Contratos') return 'Nº Contrato';
    if (categoryName === 'Pedidos') return 'Nº Pedido';
    if (categoryName === 'Medições') return 'Nº Contrato / Medição';
    return 'Título';
  };

  const exportExcel = () => {
    const ws = XLSX.utils.json_to_sheet(filteredData.map(d => ({
      [getIDHeaderName()]: d.id,
      Credor: d.credor,
      Vencimento: d.vencimento,
      Valor: d.valor,
      'Status Zepp': d.statusZepp,
      'Nº Romaneio': d.noRomaneio,
      'Observação': d.observacao,
      'Ação Requerida': d.acao
    })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Resultados");
    XLSX.writeFile(wb, `Conciliacao_${categoryName}.xlsx`);
  };

  const exportFaltaRomaneio = () => {
    if (faltaRomaneioData.length === 0) {
      alert('Nenhum item com alerta de Falta Romaneio ou Base.');
      return;
    }
    const ws = XLSX.utils.json_to_sheet(faltaRomaneioData.map(d => ({
      [getIDHeaderName()]: d.id,
      Credor: d.credor,
      Vencimento: d.vencimento,
      Valor: d.valor,
      'Status Zepp': d.statusZepp,
      'Nº Romaneio': d.noRomaneio,
      'Observação': d.observacao,
      'Ação Requerida': d.acao
    })));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Falta_Romaneio");
    XLSX.writeFile(wb, `Conciliacao_${categoryName}_FaltaRomaneio.xlsx`);
  };

  const exportPDF = () => {
    const doc = new jsPDF('landscape');
    doc.text(`Resultados da Conciliação - ${categoryName}`, 14, 15);
    
    const tableColumn = [getIDHeaderName(), "Credor", "Vencimento", "Valor", "Status Zepp", "Romaneio", "Ação"];
    const tableRows = [];

    filteredData.forEach(item => {
      const rowData = [
        item.id,
        item.credor,
        item.vencimento,
        formatCurrency(item.valor),
        item.statusZepp,
        item.noRomaneio,
        item.acao
      ];
      tableRows.push(rowData);
    });

    doc.autoTable({
      head: [tableColumn],
      body: tableRows,
      startY: 20,
      styles: { fontSize: 8 }
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
        <div 
          className="kpi-card clickable" 
          style={{ borderBottom: '4px solid #8b5cf6' }}
          onClick={exportFaltaRomaneio}
          title="Clique para baixar a planilha apenas com itens faltando no Romaneio"
        >
          <h3 style={{ color: '#8b5cf6' }}>Falta Romaneio (Baixar)</h3>
          <div className="value" style={{ color: '#8b5cf6' }}>{faltaRomaneioCount}</div>
        </div>
      </div>

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
            </select>
            
            <button className="btn" onClick={exportExcel}>
              <FileText size={16} /> Excel
            </button>
            <button className="btn" onClick={exportPDF}>
              <Download size={16} /> PDF
            </button>
          </div>
        </div>

        <div style={{ overflowX: 'auto' }}>
          <table>
            <thead>
              <tr>
                <th>{getIDHeaderName()}</th>
                <th>Credor</th>
                <th>Vencimento</th>
                <th>Valor</th>
                <th>Status Zepp</th>
                <th>Nº Romaneio</th>
                <th>Observação</th>
                <th>Ação Requerida</th>
              </tr>
            </thead>
            <tbody>
              {filteredData.length === 0 ? (
                <tr>
                  <td colSpan="8" style={{ textAlign: 'center', padding: '2rem' }}>
                    Nenhum resultado encontrado.
                  </td>
                </tr>
              ) : (
                filteredData.map((item, index) => (
                  <tr key={index}>
                    <td>{item.id}</td>
                    <td style={{ fontWeight: 500 }}>{item.credor}</td>
                    <td>{item.vencimento}</td>
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
