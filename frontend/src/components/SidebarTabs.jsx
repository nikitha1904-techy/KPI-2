import React from 'react';
import { Lock } from 'lucide-react';

export default function SidebarTabs({ tabs = [], activeTabId, onTabSelect, isCompleted }) {
  return (
    <div className="sidebar-tabs card">
      <div className="sidebar-header">
        <span className="sidebar-section-title">Analysis Tabs</span>
      </div>
      <nav className="tabs-nav">
        {tabs.map((tab) => {
          const isDisabled = tab.requiresMapping && !isCompleted;
          const isActive = activeTabId === tab.id;

          return (
            <button
              key={tab.id}
              type="button"
              className={`tab-btn ${isActive ? 'active' : ''} ${isDisabled ? 'disabled' : ''}`}
              onClick={() => !isDisabled && onTabSelect(tab.id)}
              disabled={isDisabled}
              title={isDisabled ? "Run mapping to unlock this output tab" : tab.name}
            >
              <span className="tab-btn-content">
                {tab.icon}
                <span className="tab-name-text">{tab.name}</span>
              </span>
              {isDisabled && <Lock className="lock-icon text-muted" size={13} />}
            </button>
          );
        })}
      </nav>
    </div>
  );
}
