import React, { useState } from 'react';
import { ArrowRight } from 'lucide-react';
import ImportModal from './components/ImportModal';
import Dashboard from './components/Dashboard';
import { processConciliacao } from './utils/engine';
import './index.css';

const TABS = ['Títulos', 'Contratos', 'Pedidos', 'Medições', 'Consolidação CRIVO'];

class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null, errorInfo: null };
  }

  static getDerivedStateFromError(error) {
    return { hasError: true };
  }

  componentDidCatch(error, errorInfo) {
    this.setState({ error, errorInfo });
    console.error("ErrorBoundary caught an error", error, errorInfo);
  }

  render() {
    if (this.state.hasError) {
      return (
        <div style={{ padding: '2rem', color: 'red', background: '#fee2e2', borderRadius: '8px', margin: '2rem' }}>
          <h2>Oops! Algo deu errado no Dashboard.</h2>
          <details style={{ whiteSpace: 'pre-wrap' }}>
            {this.state.error && this.state.error.toString()}
            <br />
            {this.state.errorInfo && this.state.errorInfo.componentStack}
          </details>
          <button className="btn btn-primary" onClick={() => window.location.reload()} style={{marginTop: '1rem'}}>Recarregar</button>
        </div>
      );
    }
    return this.props.children;
  }
}

function App() {
  const [activeTab, setActiveTab] = useState('Títulos');
  const [appState, setAppState] = useState({
    'Títulos': { imported: false, data: [], kpi: null },
    'Contratos': { imported: false, data: [], kpi: null },
    'Pedidos': { imported: false, data: [], kpi: null },
    'Medições': { imported: false, data: [], kpi: null },
    'Consolidação CRIVO': { imported: false, data: [], kpi: null }
  });
  const [isProcessing, setIsProcessing] = useState(false);

  const handleImport = async (files) => {
    setIsProcessing(true);
    try {
      const { results, kpi } = await processConciliacao(files, activeTab);
      setAppState(prev => ({
        ...prev,
        [activeTab]: { imported: true, data: results, kpi }
      }));
    } catch (error) {
      alert("Erro ao processar planilhas. Verifique os formatos.");
      console.error(error);
    } finally {
      setIsProcessing(false);
    }
  };

  const currentTabState = appState[activeTab];

  return (
    <>
      <header className="header">
        <div className="header-title">
          <div className="header-icon">
            <ArrowRight size={20} />
          </div>
          Suíte de Conciliação PMG
        </div>
        
        <div className="tabs">
          {TABS.map(tab => (
            <button 
              key={tab} 
              className={`tab ${activeTab === tab ? 'active' : ''}`}
              onClick={() => setActiveTab(tab)}
            >
              {tab}
            </button>
          ))}
        </div>
      </header>

      <main className="main-content">
        {isProcessing ? (
          <div style={{ display: 'flex', justifyContent: 'center', padding: '4rem' }}>
            <h2>Processando Conciliação...</h2>
          </div>
        ) : !currentTabState.imported ? (
          <ImportModal 
            categoryName={activeTab} 
            onImport={handleImport} 
          />
        ) : (
          <ErrorBoundary>
            <Dashboard 
              data={currentTabState.data} 
              kpi={currentTabState.kpi} 
              categoryName={activeTab}
              onReimport={() => {
                setAppState(prev => ({
                  ...prev,
                  [activeTab]: { imported: false, data: [], kpi: null }
                }));
              }}
            />
          </ErrorBoundary>
        )}
      </main>
    </>
  );
}

export default App;
