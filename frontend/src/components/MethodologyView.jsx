import React from 'react';
import { Network, Database, Compass, Eye, Shuffle, RefreshCw, Layers, CheckCircle } from 'lucide-react';

const steps = [
  {
    num: "01",
    title: "KPI Extraction",
    icon: <Database className="step-icon text-primary" size={24} />,
    desc: "Scan uploaded Excel sheet tables, CSV summaries, and JSON outputs to extract raw KPI labels, formula definitions, unit specifications, and contextual metadata."
  },
  {
    num: "02",
    title: "Semantic Cross-Comparison",
    icon: <Shuffle className="step-icon text-primary" size={24} />,
    desc: "Pass extracted terms through insurance-trained embedding models to evaluate similarity metrics against the enterprise standard Ontology KPI catalog."
  },
  {
    num: "03",
    title: "Score & Rationale Generation",
    icon: <Layers className="step-icon text-primary" size={24} />,
    desc: "Calculate confidence percentages (0-100%) and assemble natural language justifications outlining structural overlaps or key terminology divergences."
  },
  {
    num: "04",
    title: "Best Match Selection",
    icon: <Compass className="step-icon text-primary" size={24} />,
    desc: "Deduplicate candidate matches, selecting the single highest-probability ontology match for each report KPI based on joint math and definition scores."
  },
  {
    num: "05",
    title: "Commonality Analysis",
    icon: <Network className="step-icon text-primary" size={24} />,
    desc: "Inspect cross-report alignments to highlight consistency levels (e.g. standard mappings across multiple reports) and identify systemic naming differences."
  },
  {
    num: "06",
    title: "Formula Gap Audit",
    icon: <RefreshCw className="step-icon text-primary" size={24} />,
    desc: "Compare mathematical formulations in reports (e.g. gross claim inclusions) against standard ontology formulas, flagging mathematical omissions."
  },
  {
    num: "07",
    title: "Ontology Extension Recommendations",
    icon: <Eye className="step-icon text-primary" size={24} />,
    desc: "Analyze report clauses for novel metrics that represent critical industry standards, prompting additions to expand the master ontology list."
  },
  {
    num: "08",
    title: "Business Validation Compilation",
    icon: <CheckCircle className="step-icon text-primary" size={24} />,
    desc: "Synthesize mapping datasets, reserve mismatch audits, and gap logs into structured business tables to prepare for analyst sign-off and approval."
  }
];

export default function MethodologyView() {
  return (
    <div className="methodology-container">
      <div className="workspace-header-block mb-4">
        <h2 className="workspace-tab-title">Methodology & Pipeline Pipeline Architecture</h2>
        <p className="workspace-tab-desc">
          Structured overview of the 8-phase ontology alignment process used to standardize, compare, and validate insurance KPIs.
        </p>
      </div>

      <div className="methodology-grid">
        {steps.map((step, idx) => (
          <div key={idx} className="methodology-card card">
            <div className="step-badge-row">
              <div className="icon-wrapper bg-glass-surface">
                {step.icon}
              </div>
              <span className="step-number">{step.num}</span>
            </div>
            
            <h3 className="step-card-title">{step.title}</h3>
            <p className="step-card-desc">{step.desc}</p>
          </div>
        ))}
      </div>

      <div className="methodology-pipeline-footer card mt-4">
        <h4 className="footer-pipeline-title">Modifying LLM Prompting & Logic Templates</h4>
        <p className="text-secondary font-body font-normal">
          This system is built as an extensible pipeline shell. Later, you can hook the Python/LangChain backend to this UI.
          The pipeline steps correspond to prompt modules executing sequential operations: extraction, evaluation, formula gap checking, and recommendation prompting.
        </p>
      </div>
    </div>
  );
}
