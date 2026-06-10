import React, { useState } from 'react';
import { UploadCloud, CheckCircle, X } from 'lucide-react';

export default function ImportModal({ onImport, categoryName, onClose }) {
  const [files, setFiles] = useState({
    romaneio: null,
    sienge: null,
    zepp: null
  });

  const handleFileChange = (type, e) => {
    const file = e.target.files[0];
    if (file) {
      setFiles(prev => ({ ...prev, [type]: file }));
    }
  };

  const handleImport = () => {
    onImport(files);
  };

  const allFilesSelected = files.romaneio && files.sienge && files.zepp;
  const anyFileSelected = files.romaneio || files.sienge || files.zepp;

  return (
    <div className="import-container">
      <div className="import-card">
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <h2>Importar Dados - {categoryName}</h2>
          {onClose && (
            <button onClick={onClose} style={{ background: 'none', border: 'none', cursor: 'pointer', color: 'var(--text-muted)' }}>
              <X size={24} />
            </button>
          )}
        </div>
        
        <p style={{ color: 'var(--text-muted)', marginBottom: '1.5rem' }}>
          Para conciliar a aba de {categoryName}, importe as planilhas extraídas dos 3 sistemas.
        </p>

        <div className="import-grid">
          <FileDropZone 
            title="Sienge" 
            file={files.sienge} 
            onChange={(e) => handleFileChange('sienge', e)} 
          />
          <FileDropZone 
            title="Zepp" 
            file={files.zepp} 
            onChange={(e) => handleFileChange('zepp', e)} 
          />
          <FileDropZone 
            title="Romaneio" 
            file={files.romaneio} 
            onChange={(e) => handleFileChange('romaneio', e)} 
          />
        </div>

        <div style={{ marginTop: '2rem', display: 'flex', justifyContent: 'flex-end', gap: '1rem' }}>
          {onClose && (
            <button className="btn" onClick={onClose}>Cancelar</button>
          )}
          <button 
            className="btn btn-primary" 
            onClick={handleImport}
            disabled={!anyFileSelected}
          >
            Processar Conciliação
          </button>
        </div>
      </div>
    </div>
  );
}

function FileDropZone({ title, file, onChange }) {
  return (
    <div className="file-drop" style={{ borderColor: file ? 'var(--success-color)' : '' }}>
      <input 
        type="file" 
        accept=".xlsx,.xls,.csv" 
        onChange={onChange} 
        style={{ display: 'none' }} 
        id={`file-${title}`}
      />
      <label htmlFor={`file-${title}`} style={{ cursor: 'pointer', display: 'block', width: '100%', height: '100%' }}>
        {file ? (
          <>
            <CheckCircle size={32} color="var(--success-color)" style={{ margin: '0 auto' }} />
            <p style={{ color: 'var(--success-color)', fontWeight: '600' }}>{file.name}</p>
          </>
        ) : (
          <>
            <UploadCloud size={32} color="var(--text-muted)" style={{ margin: '0 auto' }} />
            <p>Selecione a planilha do <strong>{title}</strong></p>
          </>
        )}
      </label>
    </div>
  );
}
