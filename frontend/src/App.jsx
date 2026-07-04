import React, { useState, useEffect } from 'react';
import { 
  mockAllKpiMapping, 
  mockBestMatch, 
  mockCommonOntologyMapping, 
  mockFormulaComparison, 
  mockMissingOntologyRecommendations, 
  mockReportWiseMissingKpis, 
  mockFeedbackReview 
} from './data/mockData';

// Modular Component Imports
import FileUploadPanel from './components/FileUploadPanel';
import SidebarTabs from './components/SidebarTabs';
import StatusProgress from './components/StatusProgress';
import OutputTable from './components/OutputTable';
import FeedbackTable from './components/FeedbackTable';
import MethodologyView from './components/MethodologyView';

// Icons
import { Layers, CheckCircle2, ShieldAlert, Sparkles, FileText, CheckSquare, BarChart, RefreshCw, AlertOctagon, HelpCircle, MessageSquare } from 'lucide-react';

export default function App() {
  const [reportFiles, setReportFiles] = useState([]);
  const [ontologyFile, setOntologyFile] = useState(null);
  
  // Pipeline processing states
  const [isProcessing, setIsProcessing] = useState(false);
  const [isCompleted, setIsCompleted] = useState(false);
  const [apiFinished, setApiFinished] = useState(false);
  
  // API results and active states
  const [runId, setRunId] = useState(null);
  const [apiResults, setApiResults] = useState(null);
  const [activeTabId, setActiveTabId] = useState('overview');
  const [errorMessage, setErrorMessage] = useState(null);

  // Sub-toggle state for Pairwise Formula Check tab
  const [pairwiseSubTab, setPairwiseSubTab] = useState('step5_check');

  // Manage feedback entries dynamically
  const [feedbackRows, setFeedbackRows] = useState(mockFeedbackReview);

  // Synchronize feedback rows when API results load
  useEffect(() => {
    if (apiResults && apiResults["3. Best Mapping"]) {
      const formatted = apiResults["3. Best Mapping"].map((row, idx) => ({
        id: idx + 1,
        reportName: row["Report A"],
        reportKpi: row["Report Col Name"],
        ontologyKpi: row["Ontology KPI Name"],
        matchScore: row["similarity_score"] || 0,
        status: row["needs_human_review"] === "Yes" ? "Partially Correct" : "Correct",
        comments: row["actuarial_rationale"] || ""
      }));
      setFeedbackRows(formatted);
    }
  }, [apiResults]);

  // Upload handlers
  const handleReportUpload = (files) => {
    setReportFiles(prev => [...prev, ...files]);
  };

  const handleOntologyUpload = (file) => {
    setOntologyFile(file);
  };

  const handleRemoveReport = (index) => {
    setReportFiles(prev => prev.filter((_, idx) => idx !== index));
  };

  const handleRemoveOntology = () => {
    setOntologyFile(null);
  };

  const handleLoadSampleData = () => {
    setReportFiles([
      { name: "LNBAR 2024 EB Reserve Deatils.xlsx (Sample)" },
      { name: "LNBAR Worksite A Reserve Details.xlsx (Sample)" }
    ]);
    setOntologyFile({ name: "life_annuity_actuarial_reserving_kpi_ontology.xlsx (Sample)" });
  };

  const handleRunMapping = async () => {
    setIsProcessing(true);
    setIsCompleted(false);
    setApiFinished(false);
    setErrorMessage(null);

    try {
      // Determine if we are using uploaded files or sample data fallback
      const hasRealFiles = reportFiles.some(f => f instanceof File) && ontologyFile instanceof File;

      if (hasRealFiles) {
        const formData = new FormData();
        reportFiles.forEach((file) => {
          formData.append("reports", file);
        });
        formData.append("ontology", ontologyFile);

        const uploadRes = await fetch("http://127.0.0.1:8000/api/upload", {
          method: "POST",
          body: formData,
        });

        if (!uploadRes.ok) {
          throw new Error("File upload failed on the server.");
        }
      }

      // Execute pipeline
      const runRes = await fetch("http://127.0.0.1:8000/api/run-mapping", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ run_llm: true }),
      });

      if (!runRes.ok) {
        throw new Error("Pipeline run failed on backend.");
      }

      const runData = await runRes.json();
      setRunId(runData.run_id);

      // Fetch results
      const resultsRes = await fetch(`http://127.0.0.1:8000/api/results/${runData.run_id}`);
      if (!resultsRes.ok) {
        throw new Error("Failed to retrieve pipeline results from server.");
      }

      const resultsData = await resultsRes.json();
      setApiResults(resultsData);
      setApiFinished(true); // Signal to StatusProgress that API call finished successfully
    } catch (err) {
      console.error(err);
      setErrorMessage(err.message);
      setIsProcessing(false);
    }
  };

  const handlePipelineComplete = () => {
    setIsProcessing(false);
    setIsCompleted(true);
    setActiveTabId('all_kpi'); // Focus to results
  };

  const handleDownloadExcel = async () => {
    if (!runId) return;
    try {
      const response = await fetch(`http://127.0.0.1:8000/api/download/${runId}`);
      if (!response.ok) {
        throw new Error("Failed to download Excel file from the server.");
      }
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'final_output_with_all_steps.xlsx');
      link.style.display = 'none';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error("Excel download error:", err);
      alert("Error downloading Excel file: " + err.message);
    }
  };

  const handleRowFeedbackUpdate = (id, field, value) => {
    setFeedbackRows(prev => 
      prev.map(row => row.id === id ? { ...row, [field]: value } : row)
    );
  };

  // Dynamic tabs configuration based on backend results
  const getTabs = () => {
    const list = [
      { id: 'overview', name: 'Overview', icon: <FileText size={16} />, requiresMapping: false },
      { id: 'all_kpi', name: 'All KPI Mapping', icon: <Layers size={16} />, requiresMapping: true },
      { id: 'best_match', name: 'Best Match', icon: <CheckSquare size={16} />, requiresMapping: true },
      { id: 'common', name: 'Common Ontology Mapping', icon: <BarChart size={16} />, requiresMapping: true },
      { id: 'formula_comp', name: 'Formula Comparison', icon: <RefreshCw size={16} />, requiresMapping: true },
      { id: 'pairwise', name: 'Pairwise Formula Check', icon: <Layers size={16} />, requiresMapping: true },
    ];

    if (isCompleted && apiResults) {
      const keys = Object.keys(apiResults);
      
      // Best Mapping report-wise
      keys.forEach(k => {
        if (k.startsWith("Best Mapping - ")) {
          const reportName = k.replace("Best Mapping - ", "");
          list.push({
            id: `best_mapping_${reportName}`,
            name: `Best Mapping - ${reportName}`,
            icon: <CheckSquare size={16} />,
            requiresMapping: true
          });
        }
      });

      // Mapped report-wise
      keys.forEach(k => {
        if (k.endsWith(" Mapped Ontology KPIs")) {
          const reportName = k.replace(" Mapped Ontology KPIs", "");
          list.push({
            id: `mapped_${reportName}`,
            name: `${reportName} Mapped KPIs`,
            icon: <BarChart size={16} />,
            requiresMapping: true
          });
        }
      });

      // Missing report-wise
      keys.forEach(k => {
        if (k.endsWith(" Missing Ontology KPIs")) {
          const reportName = k.replace(" Missing Ontology KPIs", "");
          list.push({
            id: `missing_${reportName}`,
            name: `${reportName} Missing KPIs`,
            icon: <AlertOctagon size={16} />,
            requiresMapping: true
          });
        }
      });

      // Recommendations report-wise
      keys.forEach(k => {
        if (k.endsWith(" LLM Missing KPI Recs")) {
          const reportName = k.replace(" LLM Missing KPI Recs", "");
          list.push({
            id: `recs_${reportName}`,
            name: `${reportName} LLM Missing KPI Recs`,
            icon: <ShieldAlert size={16} />,
            requiresMapping: true
          });
        }
      });
    } else {
      // Default placeholder tabs before mapping run
      list.push({ id: 'report_missing', name: 'Report-wise Missing KPIs', icon: <AlertOctagon size={16} />, requiresMapping: true });
      list.push({ id: 'missing_recs', name: 'Missing Ontology Recommendations', icon: <ShieldAlert size={16} />, requiresMapping: true });
    }

    list.push({ id: 'methodology', name: 'Methodology', icon: <HelpCircle size={16} />, requiresMapping: false });
    list.push({ id: 'feedback', name: 'Feedback Review', icon: <MessageSquare size={16} />, requiresMapping: true });

    return list;
  };

  // Define table structures
  const allKpiColumns = [
    { header: "Report Name", accessor: "reportName", width: "16%" },
    { header: "Report KPI", accessor: "reportKpi", width: "16%" },
    { header: "Report Formula", accessor: "reportFormula", width: "20%" },
    { header: "Ontology KPI", accessor: "ontologyKpi", width: "16%" },
    { header: "Ontology Formula", accessor: "ontologyFormula", width: "18%" },
    { header: "Match Score", accessor: "matchScore", width: "8%" },
    { header: "Rationale", accessor: "rationale", width: "16%" }
  ];

  const bestMatchColumns = [
    { header: "Report Name", accessor: "reportName", width: "12%" },
    { header: "Report KPI", accessor: "reportKpi", width: "14%" },
    { header: "Report Formula", accessor: "reportFormula", width: "15%" },
    { header: "Best Ontology KPI", accessor: "bestOntologyKpi", width: "14%" },
    { header: "Ontology Formula", accessor: "ontologyFormula", width: "15%" },
    { header: "Definition", accessor: "definition", width: "16%" },
    { header: "Confidence", accessor: "confidenceScore", width: "7%" },
    { header: "Match Type", accessor: "matchType", width: "10%" },
    { header: "Rationale", accessor: "rationale", width: "15%" }
  ];

  const commonOntologyColumns = [
    { header: "Ontology KPI", accessor: "ontologyKpi", width: "13%" },
    { header: "Ontology Definition", accessor: "ontologyDefinition", width: "14%" },
    { header: "Ontology Formula", accessor: "ontologyFormula", width: "13%" },
    { header: "Report A KPI", accessor: "reportAKpi", width: "12%" },
    { header: "Report A Formula", accessor: "reportAFormula", width: "12%" },
    { header: "Report B KPI", accessor: "reportBKpi", width: "12%" },
    { header: "Report B Formula", accessor: "reportBFormula", width: "12%" },
    { header: "Report A Score", accessor: "reportAScore", width: "7%" },
    { header: "Report B Score", accessor: "reportBScore", width: "7%" },
    { header: "Status", accessor: "commonalityStatus", width: "8%" }
  ];

  const formulaComparisonColumns = [
    { header: "Report Name", accessor: "reportName", width: "14%" },
    { header: "Report KPI", accessor: "reportKpi", width: "14%" },
    { header: "Ontology KPI", accessor: "ontologyKpi", width: "14%" },
    { header: "Report Formula", accessor: "reportFormula", width: "20%" },
    { header: "Ontology Formula", accessor: "ontologyFormula", width: "20%" },
    { header: "Match Score", accessor: "formulaMatchScore", width: "8%" },
    { header: "Formula Gap", accessor: "formulaGap", width: "22%" },
    { header: "Rationale", accessor: "rationale", width: "18%" }
  ];

  const reportWiseMissingColumns = [
    { header: "Report Name", accessor: "reportName", width: "15%" },
    { header: "Missing KPI", accessor: "missingKpi", width: "15%" },
    { header: "Business Purpose", accessor: "businessPurpose", width: "30%" },
    { header: "Priority", accessor: "priority", width: "10%" },
    { header: "Rationale", accessor: "rationale", width: "30%" } // Fixed Bug accessor "30%" to "rationale"
  ];

  const missingOntologyColumns = [
    { header: "Report Name", accessor: "reportName", width: "15%" },
    { header: "Report Section", accessor: "reportSection", width: "15%" },
    { header: "Recommended KPI", accessor: "recommendedKpi", width: "15%" },
    { header: "Industry Relevance", accessor: "industryRelevance", width: "15%" },
    { header: "Recommended To", accessor: "recommendedTo", width: "12%" },
    { header: "Rationale", accessor: "rationale", width: "28%" }
  ];

  // Dynamic tables schemas
  const mappedKpiColumns = [
    { header: "Report Scope", accessor: "Report Scope", width: "12%" },
    { header: "Ontology KPI", accessor: "Ontology KPI", width: "15%" },
    { header: "Ontology Formula", accessor: "Ontology Formula", width: "20%" },
    { header: "Ontology Definition", accessor: "Ontology Definition", width: "20%" },
    { header: "Mapped Report KPI(s)", accessor: "Mapped Report KPI(s)", width: "15%" },
    { header: "Mapped Report Formula(s)", accessor: "Mapped Report Formula(s)", width: "15%" },
    { header: "Best Mapping Score(s)", accessor: "Best Mapping Score(s)", width: "10%" },
    { header: "Status", accessor: "Status", width: "10%" }
  ];

  const missingKpiColumns = [
    { header: "Report Scope", accessor: "Report Scope", width: "15%" },
    { header: "Ontology KPI", accessor: "Ontology KPI", width: "20%" },
    { header: "Ontology Formula", accessor: "Ontology Formula", width: "20%" },
    { header: "Ontology Definition", accessor: "Ontology Definition", width: "20%" },
    { header: "Missing Reason", accessor: "Missing Reason", width: "25%" }
  ];

  const recommendationColumns = [
    { header: "Report Scope", accessor: "Report Scope", width: "8%" },
    { header: "Ontology KPI", accessor: "Ontology KPI", width: "10%" },
    { header: "Should Add to Report?", accessor: "Should Add to Report?", width: "8%" },
    { header: "Priority", accessor: "Recommendation Priority", width: "8%" },
    { header: "Rationale", accessor: "Priority Rationale", width: "18%" },
    { header: "Report Section", accessor: "Recommended Report Section", width: "10%" },
    { header: "Existing KPIs", accessor: "Related Existing Report KPIs", width: "12%" },
    { header: "Calculation Guidance", accessor: "Potential Formula / Calculation Guidance", width: "14%" }
  ];

  const step5PairwiseColumns = [
    { header: "Ontology KPI", accessor: "Ontology KPI Name", width: "12%" },
    { header: "Report 1", accessor: "Report 1", width: "10%" },
    { header: "Column 1", accessor: "Report Column 1", width: "12%" },
    { header: "Formula 1", accessor: "Report Formula 1", width: "14%" },
    { header: "Score 1", accessor: "Ontology Score 1", width: "6%" },
    { header: "Report 2", accessor: "Report 2", width: "10%" },
    { header: "Column 2", accessor: "Report Column 2", width: "12%" },
    { header: "Formula 2", accessor: "Report Formula 2", width: "14%" },
    { header: "Score 2", accessor: "Ontology Score 2", width: "6%" },
    { header: "Pairwise Score", accessor: "Pairwise Formula Similarity Score", width: "7%" },
    { header: "Verdict", accessor: "Pairwise Verdict", width: "8%" },
    { header: "Rationale", accessor: "Pairwise Rationale", width: "14%" }
  ];

  const llmPairwiseColumns = [
    { header: "Ontology KPI", accessor: "Ontology KPI Name", width: "10%" },
    { header: "Report 1", accessor: "Report 1 File", width: "8%" },
    { header: "Column 1", accessor: "Report 1 Column", width: "10%" },
    { header: "Formula 1", accessor: "Report 1 Formula", width: "12%" },
    { header: "Report 2", accessor: "Report 2 File", width: "8%" },
    { header: "Column 2", accessor: "Report 2 Column", width: "10%" },
    { header: "Formula 2", accessor: "Report 2 Formula", width: "12%" },
    { header: "Pairwise Score", accessor: "Formula Similarity Score", width: "6%" },
    { header: "Band", accessor: "Similarity Band", width: "8%" },
    { header: "Verdict", accessor: "LLM Verdict", width: "8%" },
    { header: "Equivalence Type", accessor: "Formula Equivalence Type", width: "10%" },
    { header: "Rationale", accessor: "Rationale", width: "15%" },
    { header: "Human Review", accessor: "Needs Human Review", width: "6%" }
  ];

  // Helper mapping functions
  const mapAllKpiData = () => {
    if (apiResults && apiResults["2. All Comparisons"]) {
      return apiResults["2. All Comparisons"].map((row, idx) => ({
        id: idx + 1,
        reportName: row["Report A"],
        reportKpi: row["Report Col Name"],
        reportFormula: row["Report Formula"],
        ontologyKpi: row["Ontology KPI Name"],
        ontologyFormula: row["Ontology Formula"],
        matchScore: row["similarity_score"] || 0,
        rationale: row["actuarial_rationale"] || ""
      }));
    }
    return mockAllKpiMapping;
  };

  const mapBestMatchData = () => {
    if (apiResults && apiResults["3. Best Mapping"]) {
      return apiResults["3. Best Mapping"].map((row, idx) => ({
        id: idx + 1,
        reportName: row["Report A"],
        reportKpi: row["Report Col Name"],
        reportFormula: row["Report Formula"],
        bestOntologyKpi: row["Ontology KPI Name"],
        ontologyFormula: row["Ontology Formula"],
        definition: row["Ontology Definition"],
        confidenceScore: row["similarity_score"] || 0,
        matchType: row["similarity_band"] || "Aligned",
        rationale: row["actuarial_rationale"] || ""
      }));
    }
    return mockBestMatch;
  };

  const mapCommonOntologyData = () => {
    if (apiResults && apiResults["4. Common Ontology"]) {
      const rows = apiResults["4. Common Ontology"];
      const groups = {};
      rows.forEach((row) => {
        const kpi = row["Ontology KPI Name"];
        if (!groups[kpi]) {
          groups[kpi] = {
            ontologyKpi: kpi,
            ontologyDefinition: row["Ontology Definition"],
            ontologyFormula: row["Ontology Formula"],
            reports: []
          };
        }
        groups[kpi].reports.push(row);
      });

      return Object.values(groups).map((group, idx) => {
        const r1 = group.reports[0] || {};
        const r2 = group.reports[1] || {};
        return {
          id: idx + 1,
          ontologyKpi: group.ontologyKpi,
          ontologyDefinition: group.ontologyDefinition,
          ontologyFormula: group.ontologyFormula,
          reportAKpi: r1["Report Col Name"] || "-",
          reportAFormula: r1["Report Formula"] || "-",
          reportBKpi: r2["Report Col Name"] || "-",
          reportBFormula: r2["Report Formula"] || "-",
          reportAScore: r1["similarity_score"] || 0,
          reportBScore: r2["similarity_score"] || 0,
          commonalityStatus: r1["similarity_score"] >= 85 && r2["similarity_score"] >= 85 ? "Consistent" : "Slight Gap"
        };
      });
    }
    return mockCommonOntologyMapping;
  };

  const mapFormulaComparisonData = () => {
    if (apiResults && apiResults["5. Step4 Definition Validation"]) {
      return apiResults["5. Step4 Definition Validation"].map((row, idx) => ({
        id: idx + 1,
        reportName: row["Report A"],
        reportKpi: row["Report Col Name"],
        ontologyKpi: row["Ontology KPI Name"],
        reportFormula: row["Report Formula"],
        ontologyFormula: row["Ontology Formula"],
        formulaMatchScore: row["Formula Similarity Score"] || 0,
        formulaGap: row["Step 4 Verdict"] || "Analyzed",
        rationale: row["Step 4 Rationale"] || ""
      }));
    }
    return mockFormulaComparison;
  };

  const mapReportWiseMissingData = () => {
    if (apiResults) {
      const list = [];
      let id = 1;
      Object.keys(apiResults).forEach((sheetName) => {
        if (sheetName.endsWith(" Missing Ontology KPIs")) {
          const reportScope = sheetName.replace(" Missing Ontology KPIs", "");
          apiResults[sheetName].forEach((row) => {
            list.push({
              id: id++,
              reportName: row["Report Scope"] || reportScope,
              missingKpi: row["Ontology KPI"],
              businessPurpose: row["Ontology Definition"],
              priority: "Medium",
              rationale: row["Missing Reason"]
            });
          });
        }
      });
      return list;
    }
    return mockReportWiseMissingKpis;
  };

  const mapRecommendationsData = () => {
    if (apiResults) {
      const list = [];
      let id = 1;
      Object.keys(apiResults).forEach((sheetName) => {
        if (sheetName.endsWith(" LLM Missing KPI Recs")) {
          const reportScope = sheetName.replace(" LLM Missing KPI Recs", "");
          apiResults[sheetName].forEach((row) => {
            if (row["Recommendation Priority"] !== "Not Recommended" && row["Should Add to Report?"] !== "No") {
              list.push({
                id: id++,
                reportName: row["Report Scope"] || reportScope,
                reportSection: row["Recommended Report Section"] || "N/A",
                recommendedKpi: row["Ontology KPI"],
                industryRelevance: row["Industry Relevance"] || "Medium",
                recommendedTo: row["Should Add to Report?"] || "Yes",
                rationale: row["Priority Rationale"] || row["Reason for Recommendation"]
              });
            }
          });
        }
      });
      return list;
    }
    return mockMissingOntologyRecommendations;
  };

  // Render workspace content
  const renderWorkspaceContent = () => {
    // 1. Static standard tabs
    if (activeTabId === 'overview') {
      return (
        <div className="overview-container">
          <div className="workspace-header-block mb-4">
            <h2 className="workspace-tab-title">Standardization & Validation Dashboard</h2>
            <p className="workspace-tab-desc">
              Welcome to the Insurance KPI Ontology Alignment Tool. This application allows insurance firms and auditors to reconcile disparate actuarial spreadsheets and underwriting reports against centralized enterprise business definitions.
            </p>
          </div>

          <div className="overview-cards-grid">
            <div className="card overview-card">
              <Layers className="text-primary mb-2" size={24} />
              <h3>Ontology Alignment</h3>
              <p className="text-secondary">Verify that premium, reserve, and expense terms used across different insurance departments align with regulatory standards.</p>
            </div>
            <div className="card overview-card">
              <CheckCircle2 className="text-success mb-2" size={24} />
              <h3>Formula Auditing</h3>
              <p className="text-secondary">Track mathematical inconsistencies in claims, gross profit, or retrocession equations compared to internal bookkeeping standards.</p>
            </div>
            <div className="card overview-card">
              <ShieldAlert className="text-warning mb-2" size={24} />
              <h3>Gap Detection</h3>
              <p className="text-secondary">Identify KPIs that are completely missing from report files or recommend novel report terms to extend the enterprise ontology.</p>
            </div>
          </div>

          {errorMessage && (
            <div className="card instructions-panel mt-4" style={{ borderColor: '#ef4444', backgroundColor: '#fef2f2' }}>
              <h4 style={{ color: '#ef4444' }}>Pipeline Execution Error</h4>
              <p className="text-secondary mt-1">{errorMessage}</p>
            </div>
          )}

          <div className="card instructions-panel mt-4">
            <h4>Getting Started</h4>
            <ol className="instructions-list mt-2">
              <li>Upload one or more actuarial/claims spreadsheets in the top panel (Excel, CSV, or JSON).</li>
              <li>Upload the enterprise master KPI ontology file.</li>
              <li>Click <strong>Run Mapping</strong> to initiate the alignment process.</li>
              <li>Once complete, the output tabs in the left sidebar will unlock. Select any tab to audit mapping values, evaluate formula differences, and submit review validation checks.</li>
            </ol>
          </div>
        </div>
      );
    }

    if (activeTabId === 'all_kpi') {
      return (
        <OutputTable 
          title="All KPI Mapping"
          description="Comprehensive listing of all extracted report KPIs matched against the enterprise ontology KPI catalog."
          data={mapAllKpiData()}
          columns={allKpiColumns}
        />
      );
    }

    if (activeTabId === 'best_match') {
      return (
        <OutputTable 
          title="Best Match Mapping"
          description="Filters out multiple matches to showcase only the highest-confidence matched ontology target per report KPI."
          data={mapBestMatchData()}
          columns={bestMatchColumns}
        />
      );
    }

    if (activeTabId === 'common') {
      return (
        <OutputTable 
          title="Common Ontology Mapping"
          description="Comparison matrix mapping ontology metrics across multiple uploaded reports to check definition consistency."
          data={mapCommonOntologyData()}
          columns={commonOntologyColumns}
        />
      );
    }

    if (activeTabId === 'formula_comp') {
      return (
        <OutputTable 
          title="Formula Gap Comparison"
          description="Performs mathematical audits between report equations and ontology definitions, flagging structural calculation gaps."
          data={mapFormulaComparisonData()}
          columns={formulaComparisonColumns}
        />
      );
    }

    if (activeTabId === 'pairwise') {
      const step5Data = apiResults ? apiResults["6. Step5 Pairwise Check"] : [];
      const llmPairwiseData = apiResults ? apiResults["LLM Pairwise Formula Check"] : [];
      
      return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '20px' }}>
          <div className="card" style={{ padding: '12px', display: 'flex', gap: '12px', alignItems: 'center', backgroundColor: '#f8fafc' }}>
            <span className="font-semibold text-secondary" style={{ fontSize: '14px' }}>Pairwise View:</span>
            <button
              type="button"
              className={`btn btn-sm ${pairwiseSubTab === 'step5_check' ? 'btn-primary' : 'btn-secondary'}`}
              onClick={() => setPairwiseSubTab('step5_check')}
            >
              Step 5 Pairwise Check
            </button>
            <button
              type="button"
              className={`btn btn-sm ${pairwiseSubTab === 'llm_check' ? 'btn-primary' : 'btn-secondary'}`}
              onClick={() => setPairwiseSubTab('llm_check')}
            >
              LLM Pairwise Formula Check
            </button>
          </div>
          
          {pairwiseSubTab === 'step5_check' ? (
            <OutputTable 
              title="Step 5 Pairwise Check"
              description="Displays formula comparison verdicts and scores for every unique pair of report columns mapped to the same ontology KPI."
              data={step5Data}
              columns={step5PairwiseColumns}
            />
          ) : (
            <OutputTable 
              title="LLM Pairwise Formula Check"
              description="Detailed actuarial discrepancy analysis generated by the LLM comparing each report formula pair."
              data={llmPairwiseData}
              columns={llmPairwiseColumns}
            />
          )}
        </div>
      );
    }

    // 2. Dynamic dynamic report tabs
    if (activeTabId.startsWith("best_mapping_")) {
      const reportName = activeTabId.replace("best_mapping_", "");
      const data = apiResults ? apiResults[`Best Mapping - ${reportName}`] : [];
      const formatted = data.map((row, idx) => ({
        id: idx + 1,
        reportName: row["Report A"],
        reportKpi: row["Report Col Name"],
        reportFormula: row["Report Formula"],
        bestOntologyKpi: row["Ontology KPI Name"],
        ontologyFormula: row["Ontology Formula"],
        definition: row["Ontology Definition"],
        confidenceScore: row["similarity_score"] || 0,
        matchType: row["similarity_band"] || "Aligned",
        rationale: row["actuarial_rationale"] || ""
      }));

      return (
        <OutputTable 
          title={`Best Mapping - ${reportName}`}
          description={`Report-specific best ontology target mappings for ${reportName}.`}
          data={formatted}
          columns={bestMatchColumns}
        />
      );
    }

    if (activeTabId.startsWith("mapped_")) {
      const reportName = activeTabId.replace("mapped_", "");
      const data = apiResults ? apiResults[`${reportName} Mapped Ontology KPIs`] : [];
      return (
        <OutputTable 
          title={`${reportName} Mapped Ontology KPIs`}
          description={`Actuarial ontology KPIs present in the best mapping of ${reportName}.`}
          data={data}
          columns={mappedKpiColumns}
        />
      );
    }

    if (activeTabId.startsWith("missing_")) {
      const reportName = activeTabId.replace("missing_", "");
      const data = apiResults ? apiResults[`${reportName} Missing Ontology KPIs`] : [];
      return (
        <OutputTable 
          title={`${reportName} Missing Ontology KPIs`}
          description={`Standard ontology KPIs absent from the best mapping of ${reportName}.`}
          data={data}
          columns={missingKpiColumns}
        />
      );
    }

    if (activeTabId.startsWith("recs_")) {
      const reportName = activeTabId.replace("recs_", "");
      const data = apiResults ? apiResults[`${reportName} LLM Missing KPI Recs`] : [];
      return (
        <OutputTable 
          title={`${reportName} LLM Missing KPI Recommendations`}
          description={`LLM-suggested report-specific KPIs and calculation guidance to bridge ontology coverage gaps in ${reportName}.`}
          data={data}
          columns={recommendationColumns}
        />
      );
    }

    // 3. Fallback placeholder tabs before run
    if (activeTabId === 'report_missing') {
      return (
        <OutputTable 
          title="Report-wise Missing KPIs"
          description="Lists standard enterprise KPIs that were expected but completely missing from the scanned reports."
          data={mapReportWiseMissingData()}
          columns={reportWiseMissingColumns}
        />
      );
    }

    if (activeTabId === 'missing_recs') {
      return (
        <OutputTable 
          title="Missing Ontology Recommendations"
          description="Recommends novel report-specific KPIs that are highly relevant to add to the enterprise master ontology list."
          data={mapRecommendationsData()}
          columns={missingOntologyColumns}
        />
      );
    }

    if (activeTabId === 'methodology') {
      return <MethodologyView />;
    }

    if (activeTabId === 'feedback') {
      return (
        <FeedbackTable 
          feedbackRows={feedbackRows}
          onRowUpdate={handleRowFeedbackUpdate}
        />
      );
    }

    return <div>Tab not found</div>;
  };

  return (
    <div className="app-container">
      {/* Header section */}
      <header className="app-header-react card">
        <div className="header-brand-block">
          <div className="logo-spark-wrapper">
            <Sparkles className="logo-spark text-primary" size={20} />
          </div>
          <div className="header-text-block">
            <h1>Insurance KPI Ontology Mapping Tool</h1>
            <p>Standardize, validate, and rationalize report KPIs against enterprise ontology</p>
          </div>
        </div>
        <div className="header-pill-tag">
          <span>Enterprise MVP Shell</span>
        </div>
      </header>

      {/* Top File Upload Segment */}
      <FileUploadPanel 
        reportFiles={reportFiles}
        ontologyFile={ontologyFile}
        onReportUpload={handleReportUpload}
        onOntologyUpload={handleOntologyUpload}
        onRemoveReport={handleRemoveReport}
        onRemoveOntology={handleRemoveOntology}
        onRunMapping={handleRunMapping}
        isProcessing={isProcessing}
        isCompleted={isCompleted}
        onLoadSampleData={handleLoadSampleData}
        onDownloadExcel={handleDownloadExcel}
      />

      {/* Conditional Pipeline Progress View */}
      {isProcessing && (
        <StatusProgress 
          onComplete={handlePipelineComplete} 
          apiFinished={apiFinished}
        />
      )}

      {/* Bottom Main Content Panel Split */}
      <div className="main-layout-split">
        {/* Sidebar Navigation */}
        <aside className="sidebar-nav-aside">
          <SidebarTabs 
            tabs={getTabs()}
            activeTabId={activeTabId} 
            onTabSelect={setActiveTabId}
            isCompleted={isCompleted}
          />
        </aside>

        {/* Workspace Display Area */}
        <main className="workspace-main card">
          {renderWorkspaceContent()}
        </main>
      </div>
    </div>
  );
}
