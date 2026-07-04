import React, { useEffect, useState } from 'react';
import { Play, CheckCircle2 } from 'lucide-react';

const phases = [
  { percent: 12, label: "Extracting report KPIs, formulas, definitions, and metadata..." },
  { percent: 25, label: "Comparing each report KPI against enterprise ontology list..." },
  { percent: 38, label: "Generating semantic match scores and matching rationales..." },
  { percent: 50, label: "Selecting best ontology match per report KPI..." },
  { percent: 63, label: "Identifying common ontology mappings across multiple reports..." },
  { percent: 75, label: "Executing detailed formula syntax and reserve factor comparisons..." },
  { percent: 88, label: "Formulating missing ontology recommendations from report contexts..." },
  { percent: 100, label: "Preparing validation reports and mapping output schemas..." }
];

export default function StatusProgress({ onComplete, apiFinished }) {
  const [progress, setProgress] = useState(0);
  const [activeLogs, setActiveLogs] = useState([]);
  const [currentLabel, setCurrentLabel] = useState("Initializing pipeline...");

  useEffect(() => {
    let currentIdx = 0;
    const intervalTime = 500; // Total duration ~4 seconds

    const timer = setInterval(() => {
      setProgress((prev) => {
        if (prev >= 98 && !apiFinished) {
          setCurrentLabel("Executing LLM scoring and validation pipeline on backend...");
          return 98;
        }
        if (prev >= 100) {
          clearInterval(timer);
          setTimeout(() => {
            onComplete();
          }, 400);
          return 100;
        }
        
        const nextProgress = prev + 3.5;
        const boundedProgress = Math.min(nextProgress, 100);

        // Check if we entered a new phase
        const phase = phases[currentIdx];
        if (phase && boundedProgress >= phase.percent) {
          setCurrentLabel(phase.label);
          setActiveLogs((logs) => [...logs, {
            time: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' }),
            text: phase.label
          }]);
          currentIdx++;
        }

        return boundedProgress;
      });
    }, intervalTime / 4);

    return () => clearInterval(timer);
  }, [onComplete, apiFinished]);

  return (
    <div className="status-progress card">
      <div className="status-progress-header">
        <div className="header-status-title">
          <Play className="spinner-icon text-primary mr-2" size={18} />
          <h3>Aligning Report KPIs with Enterprise Ontology</h3>
        </div>
        <span className="progress-percentage">{Math.round(progress)}%</span>
      </div>

      {/* Progress Bar */}
      <div className="progress-track-large">
        <div 
          className="progress-fill-large" 
          style={{ width: `${progress}%` }}
        ></div>
      </div>

      <div className="current-action-label">
        <span className="text-secondary font-medium">Active Phase: </span>
        <span className="text-primary font-semibold">{currentLabel}</span>
      </div>

      {/* Console Logs Box */}
      <div className="console-logs-wrapper">
        <div className="console-header">Pipeline Execution Logs</div>
        <div className="console-scroll">
          <div className="log-line text-muted">
            [SYSTEM] Starting ontology alignment run...
          </div>
          {activeLogs.map((log, idx) => (
            <div key={idx} className="log-line">
              <span className="log-time">[{log.time}]</span>{' '}
              <span className="log-text">{log.text}</span>
            </div>
          ))}
          {progress >= 100 && (
            <div className="log-line text-success font-semibold flex-row-align mt-1">
              <CheckCircle2 size={13} className="mr-1" />
              [SUCCESS] Pipeline successfully completed. Output tables generated.
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
