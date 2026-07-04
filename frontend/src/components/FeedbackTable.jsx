import React, { useState } from 'react';
import { Search, Save, Check } from 'lucide-react';

export default function FeedbackTable({ feedbackRows = [], onRowUpdate }) {
  const [search, setSearch] = useState('');
  const [savedRowId, setSavedRowId] = useState(null);

  const handleStatusChange = (id, newStatus) => {
    onRowUpdate(id, 'status', newStatus);
  };

  const handleCommentChange = (id, newComment) => {
    onRowUpdate(id, 'comments', newComment);
  };

  const handleSave = (id) => {
    setSavedRowId(id);
    setTimeout(() => setSavedRowId(null), 3000);
  };

  const filteredRows = feedbackRows.filter(row => {
    if (!search) return true;
    const lowerSearch = search.toLowerCase();
    return (
      row.reportName.toLowerCase().includes(lowerSearch) ||
      row.reportKpi.toLowerCase().includes(lowerSearch) ||
      row.ontologyKpi.toLowerCase().includes(lowerSearch) ||
      (row.comments && row.comments.toLowerCase().includes(lowerSearch))
    );
  });

  return (
    <div className="output-table-container">
      <div className="table-meta-header">
        <div className="title-desc-block">
          <h2 className="workspace-tab-title">Feedback Review</h2>
          <p className="workspace-tab-desc">
            Review the mapping results, rate their accuracy, add comments, and save changes to update the ontology pipeline rules.
          </p>
        </div>
        
        <div className="table-actions-row">
          <div className="search-box-react">
            <Search className="search-icon-react" size={16} />
            <input 
              type="text" 
              placeholder="Search feedback..." 
              value={search}
              onChange={(e) => setSearch(e.target.value)}
            />
          </div>
        </div>
      </div>

      <div className="table-responsive card">
        <table className="enterprise-table">
          <thead>
            <tr>
              <th>Report Name</th>
              <th>Report KPI</th>
              <th>Ontology KPI</th>
              <th className="text-center">Match Score</th>
              <th style={{ width: '170px' }}>Validation Status</th>
              <th>Reviewer Comments</th>
              <th className="text-center" style={{ width: '130px' }}>Actions</th>
            </tr>
          </thead>
          <tbody>
            {filteredRows.map((row) => {
              const isSaved = savedRowId === row.id;
              
              return (
                <tr key={row.id}>
                  <td className="font-medium">{row.reportName}</td>
                  <td>{row.reportKpi}</td>
                  <td>{row.ontologyKpi}</td>
                  <td className="text-center">
                    <span className="badge-score badge-high">{row.matchScore}%</span>
                  </td>
                  <td>
                    <select 
                      className="custom-select-react"
                      value={row.status}
                      onChange={(e) => handleStatusChange(row.id, e.target.value)}
                    >
                      <option value="Correct">Correct</option>
                      <option value="Partially Correct">Partially Correct</option>
                      <option value="Incorrect">Incorrect</option>
                    </select>
                  </td>
                  <td>
                    <input 
                      type="text"
                      className="table-input-text"
                      placeholder="Add reviewer notes..."
                      value={row.comments || ''}
                      onChange={(e) => handleCommentChange(row.id, e.target.value)}
                    />
                  </td>
                  <td className="text-center">
                    <button 
                      type="button" 
                      className={`btn btn-sm btn-full ${isSaved ? 'btn-success' : 'btn-primary'}`}
                      onClick={() => handleSave(row.id)}
                    >
                      {isSaved ? (
                        <>
                          <Check size={13} />
                          <span>Saved</span>
                        </>
                      ) : (
                        <>
                          <Save size={13} />
                          <span>Save</span>
                        </>
                      )}
                    </button>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      
      <div className="table-footer-status">
        Manage business validation inputs. 
        <span className="api-placeholder-note">
          [API Endpoint: POST /api/validate-feedback to submit review arrays to Python/LLM pipeline]
        </span>
      </div>
    </div>
  );
}
