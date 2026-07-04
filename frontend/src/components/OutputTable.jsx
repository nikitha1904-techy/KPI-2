import React, { useState } from 'react';
import { Search, Download, Database, Check } from 'lucide-react';

export default function OutputTable({ title, description, data = [], columns = [] }) {
  const [search, setSearch] = useState('');
  const [downloaded, setDownloaded] = useState(false);

  const filteredData = data.filter(row => {
    if (!search) return true;
    const lowerSearch = search.toLowerCase();
    return Object.values(row).some(val => 
      String(val).toLowerCase().includes(lowerSearch)
    );
  });

  const handleDownload = () => {
    if (!filteredData || filteredData.length === 0) return;
    
    // Create CSV content with proper escaping
    const csvHeaders = columns.map(col => `"${col.header.replace(/"/g, '""')}"`).join(',');
    const csvRows = filteredData.map(row => {
      return columns.map(col => {
        const val = row[col.accessor];
        const stringVal = val === null || val === undefined ? '' : String(val);
        return `"${stringVal.replace(/"/g, '""')}"`;
      }).join(',');
    });
    const csvContent = [csvHeaders, ...csvRows].join('\n');
    
    // Trigger browser download via dynamic <a> click
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `${title.toLowerCase().replace(/[^a-z0-9]+/g, '_')}_export.csv`);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);

    setDownloaded(true);
    setTimeout(() => setDownloaded(false), 3000);
  };

  const renderCellContent = (row, col) => {
    const value = row[col.accessor];
    
    // Highlight match scores/confidence scores
    if (col.accessor.toLowerCase().includes('score') && typeof value === 'number') {
      let colorClass = 'text-danger';
      if (value >= 90) colorClass = 'badge-score badge-high';
      else if (value >= 75) colorClass = 'badge-score badge-medium';
      
      return <span className={colorClass}>{value}%</span>;
    }

    // Highlight status or priority levels
    if (col.accessor === 'commonalityStatus') {
      let badgeClass = 'status-tag tag-neutral';
      if (value === 'Consistent') badgeClass = 'status-tag tag-success';
      if (value === 'Slight Gap') badgeClass = 'status-tag tag-warning';
      if (value === 'Inconsistent') badgeClass = 'status-tag tag-danger';
      return <span className={badgeClass}>{value}</span>;
    }

    if (col.accessor === 'priority') {
      let badgeClass = 'status-tag tag-neutral';
      if (value === 'High') badgeClass = 'status-tag tag-danger';
      if (value === 'Medium') badgeClass = 'status-tag tag-warning';
      return <span className={badgeClass}>{value}</span>;
    }

    if (col.accessor === 'matchType') {
      let badgeClass = 'status-tag tag-neutral';
      if (value === 'Exact Match') badgeClass = 'status-tag tag-success';
      if (value === 'Close Match') badgeClass = 'status-tag tag-warning';
      return <span className={badgeClass}>{value}</span>;
    }

    return value || <span className="text-muted italic">-</span>;
  };

  return (
    <div className="output-table-container">
      <div className="table-meta-header">
        <div className="title-desc-block">
          <h2 className="workspace-tab-title">{title}</h2>
          <p className="workspace-tab-desc">{description}</p>
        </div>
        
        <div className="table-actions-row">
          <div className="search-box-react">
            <Search className="search-icon-react" size={16} />
            <input 
              type="text" 
              placeholder="Search table rows..." 
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </div>

          <button 
            type="button" 
            className={`btn btn-secondary btn-sm ${downloaded ? 'btn-success' : ''}`}
            onClick={handleDownload}
          >
            {downloaded ? (
              <>
                <Check size={14} />
                <span>CSV Downloaded</span>
              </>
            ) : (
              <>
                <Download size={14} />
                <span>Export CSV</span>
              </>
            )}
          </button>
        </div>
      </div>

      <div className="table-responsive card">
        {filteredData.length === 0 ? (
          <div className="table-empty-state">
            <Database size={36} className="text-muted mb-2" />
            <h4>No Records Match Filter</h4>
            <p>Refine your search term or upload fresh data files.</p>
          </div>
        ) : (
          <table className="enterprise-table">
            <thead>
              <tr>
                {columns.map((col, idx) => (
                  <th key={idx} style={{ width: col.width || 'auto' }}>
                    {col.header}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredData.map((row, rowIdx) => (
                <tr key={rowIdx}>
                  {columns.map((col, colIdx) => (
                    <td key={colIdx}>
                      {renderCellContent(row, col)}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
      
      <div className="table-footer-status">
        Showing {filteredData.length} of {data.length} records.
        <span className="api-placeholder-note">
          [API Endpoint Integration point here for fetching backend mapping data]
        </span>
      </div>
    </div>
  );
}
