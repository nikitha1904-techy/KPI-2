import React, { useRef } from 'react';
import { Upload, FileSpreadsheet, FileJson, FileText, Check, AlertCircle, Download } from 'lucide-react';

export default function FileUploadPanel({ 
  reportFiles, 
  ontologyFile, 
  onReportUpload, 
  onOntologyUpload, 
  onRemoveReport, 
  onRemoveOntology, 
  onRunMapping, 
  isProcessing, 
  isCompleted,
  onLoadSampleData,
  onDownloadExcel
}) {
  const reportInputRef = useRef(null);
  const ontologyInputRef = useRef(null);

  const handleReportChange = (e) => {
    if (e.target.files.length > 0) {
      onReportUpload(Array.from(e.target.files));
    }
  };

  const handleOntologyChange = (e) => {
    if (e.target.files.length > 0) {
      onOntologyUpload(e.target.files[0]);
    }
  };

  const isRunEnabled = reportFiles.length > 0 && ontologyFile !== null && !isProcessing;

  const getFileIcon = (fileName) => {
    const ext = fileName.split('.').pop().toLowerCase();
    if (ext === 'xlsx' || ext === 'xls' || ext === 'csv') {
      return <FileSpreadsheet className="file-icon text-excel" size={18} />;
    } else if (ext === 'json') {
      return <FileJson className="file-icon text-json" size={18} />;
    }
    return <FileText className="file-icon text-text" size={18} />;
  };

  return (
    <div className="upload-panel card">
      <div className="upload-grid">
        
        {/* Reports Upload */}
        <div className="upload-box-wrapper">
          <label className="upload-label">Actuarial / Insurance Reports</label>
          <div 
            className="upload-dropzone"
            onClick={() => reportInputRef.current?.click()}
          >
            <input 
              type="file" 
              ref={reportInputRef} 
              onChange={handleReportChange} 
              multiple 
              accept=".csv,.xlsx,.xls,.json"
              style={{ display: 'none' }}
            />
            <Upload size={24} className="text-secondary mb-2" />
            <p className="dropzone-text">Upload reports (Excel, CSV, JSON)</p>
            <span className="dropzone-sub">Drag files here or click to browse</span>
          </div>
          
          {reportFiles.length > 0 && (
            <div className="file-list">
              {reportFiles.map((file, idx) => (
                <div key={idx} className="file-item">
                  <div className="file-details">
                    {getFileIcon(file.name)}
                    <span className="file-name" title={file.name}>{file.name}</span>
                  </div>
                  <button 
                    type="button" 
                    className="file-remove-btn" 
                    onClick={(e) => { e.stopPropagation(); onRemoveReport(idx); }}
                    disabled={isProcessing}
                  >
                    &times;
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Ontology Upload */}
        <div className="upload-box-wrapper">
          <label className="upload-label">Enterprise Ontology KPI List</label>
          {ontologyFile ? (
            <div className="ontology-uploaded-card">
              <div className="ontology-file-info">
                <Check className="text-success mr-2" size={20} />
                <div className="ontology-details">
                  <span className="ontology-title">Ontology Loaded</span>
                  <span className="ontology-filename">{ontologyFile.name}</span>
                </div>
              </div>
              <button 
                type="button" 
                className="btn btn-secondary btn-sm"
                onClick={onRemoveOntology}
                disabled={isProcessing}
              >
                Change File
              </button>
            </div>
          ) : (
            <div 
              className="upload-dropzone"
              onClick={() => ontologyInputRef.current?.click()}
            >
              <input 
                type="file" 
                ref={ontologyInputRef} 
                onChange={handleOntologyChange} 
                accept=".csv,.xlsx,.xls,.json"
                style={{ display: 'none' }}
              />
              <Upload size={24} className="text-secondary mb-2" />
              <p className="dropzone-text">Upload KPI ontology (Excel, CSV, JSON)</p>
              <span className="dropzone-sub">Single master file required</span>
            </div>
          )}
        </div>

        {/* Run Panel */}
        <div className="run-panel">
          <div className="run-status-info">
            {reportFiles.length === 0 || !ontologyFile ? (
              <div className="status-badge status-pending">
                <AlertCircle size={14} />
                <span>Requires Files</span>
              </div>
            ) : isProcessing ? (
              <div className="status-badge status-running">
                <span className="pulse-indicator"></span>
                <span>Processing...</span>
              </div>
            ) : isCompleted ? (
              <div className="status-badge status-completed">
                <Check size={14} />
                <span>Mapping Completed</span>
              </div>
            ) : (
              <div className="status-badge status-ready">
                <span>Ready to Run</span>
              </div>
            )}
            
            <p className="run-instruction">
              {reportFiles.length === 0 && ontologyFile === null
                ? "Upload reports and the master ontology KPI list to activate mapping."
                : reportFiles.length === 0
                ? "Please upload at least one insurance report."
                : !ontologyFile
                ? "Please upload the ontology KPI list file."
                : isCompleted 
                ? "Mapping complete. Review outputs in sidebar tabs or re-run."
                : "All files loaded. Execute alignment and formula verification."}
            </p>
            {reportFiles.length === 0 && !ontologyFile && (
              <button
                type="button"
                className="btn-demo-link"
                onClick={onLoadSampleData}
                style={{
                  background: 'none',
                  border: 'none',
                  color: '#818cf8',
                  textDecoration: 'underline',
                  fontSize: '12px',
                  cursor: 'pointer',
                  padding: 0,
                  textAlign: 'left',
                  marginTop: '8px'
                }}
              >
                Or load actuarial sample demo files
              </button>
            )}
          </div>

          <button 
            type="button" 
            className={`btn btn-run ${isRunEnabled ? 'active' : ''}`}
            disabled={!isRunEnabled}
            onClick={onRunMapping}
          >
            {isCompleted ? "Re-Run Mapping" : "Run Mapping"}
          </button>
          {isCompleted && (
            <button
              type="button"
              className="btn mt-2"
              onClick={onDownloadExcel}
              style={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                gap: '8px',
                width: '100%',
                backgroundColor: '#10b981',
                color: 'white',
                border: 'none',
                padding: '10px 16px',
                borderRadius: '6px',
                cursor: 'pointer',
                fontWeight: '600',
                marginTop: '8px'
              }}
            >
              <Download size={16} />
              <span>Download Excel</span>
            </button>
          )}
        </div>

      </div>
    </div>
  );
}
