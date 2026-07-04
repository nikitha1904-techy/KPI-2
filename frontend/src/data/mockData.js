export const mockAllKpiMapping = [
  {
    id: 1,
    reportName: "Q4 Actuarial Valuation.xlsx",
    reportKpi: "Gross Written Premium (GWP)",
    reportFormula: "Sum of all direct and reinsurance premiums written",
    ontologyKpi: "Gross Written Premium",
    ontologyFormula: "Direct Premium Written + Assumed Reinsurance Premium",
    matchScore: 98,
    rationale: "Exact semantic match. Formula definitions are mathematically equivalent."
  },
  {
    id: 2,
    reportName: "Q4 Actuarial Valuation.xlsx",
    reportKpi: "Loss Ratio (Gross)",
    reportFormula: "Gross Incurred Claims / Gross Earned Premium",
    ontologyKpi: "Gross Loss Ratio",
    ontologyFormula: "Gross Claims Paid + Change in Gross Outstanding Claims / Gross Earned Premium",
    matchScore: 92,
    rationale: "Highly aligned. Report formula uses high-level 'Incurred Claims' which corresponds to Claims Paid + Outstanding Claims in ontology."
  },
  {
    id: 3,
    reportName: "Financial Reserve Report.csv",
    reportKpi: "Net Incurred Claims",
    reportFormula: "Claims Paid - Salvage & Subrogation + Outstanding Claim Reserves",
    ontologyKpi: "Net Claims Incurred",
    ontologyFormula: "Gross Incurred Claims - Reinsurers Share of Incurred Claims",
    matchScore: 85,
    rationale: "Partial formula divergence. Report formula defines gross claims net of recoveries but before reinsurer retrocession share."
  },
  {
    id: 4,
    reportName: "Underwriting Performance.json",
    reportKpi: "Acquisition Cost Ratio",
    reportFormula: "Underwriting Commissions / Direct Written Premium",
    ontologyKpi: "Underwriting Expense Ratio",
    ontologyFormula: "(Acquisition Costs + General Expenses) / Gross Earned Premium",
    matchScore: 78,
    rationale: "Different base definitions. Report uses written premium in denominator and excludes general overhead costs."
  },
  {
    id: 5,
    reportName: "Q4 Actuarial Valuation.xlsx",
    reportKpi: "IBNR Reserve Value",
    reportFormula: "Ultimate Loss Estimate - Paid Claims - Case Reserves",
    ontologyKpi: "Incurred But Not Reported Reserve (IBNR)",
    ontologyFormula: "Ultimate Loss Reserve - Outstanding Case Reserves - Cumulative Paid Claims",
    matchScore: 99,
    rationale: "Complete equivalence. Terminology is standard actuarial definition for IBNR estimation."
  }
];

export const mockBestMatch = [
  {
    id: 1,
    reportName: "Q4 Actuarial Valuation.xlsx",
    reportKpi: "Gross Written Premium (GWP)",
    reportFormula: "Sum of all direct and reinsurance premiums written",
    bestOntologyKpi: "Gross Written Premium",
    ontologyFormula: "Direct Premium Written + Assumed Reinsurance Premium",
    definition: "Total premium charged to policyholders before reinsurance deductions.",
    confidenceScore: 98,
    matchType: "Exact Match",
    rationale: "Direct spelling and math equivalence. No conceptual gaps detected."
  },
  {
    id: 2,
    reportName: "Q4 Actuarial Valuation.xlsx",
    reportKpi: "Loss Ratio (Gross)",
    reportFormula: "Gross Incurred Claims / Gross Earned Premium",
    bestOntologyKpi: "Gross Loss Ratio",
    ontologyFormula: "Gross Claims Paid + Change in Gross Outstanding Claims / Gross Earned Premium",
    definition: "Ratio of claims incurred relative to total earned premiums.",
    confidenceScore: 92,
    matchType: "Close Match",
    rationale: "Actuarial 'Incurred Claims' refers to the paid claims plus change in reserves, mapping closely to the detailed ontology expansion."
  },
  {
    id: 3,
    reportName: "Financial Reserve Report.csv",
    reportKpi: "Net Incurred Claims",
    reportFormula: "Claims Paid - Salvage & Subrogation + Outstanding Claim Reserves",
    bestOntologyKpi: "Net Claims Incurred",
    ontologyFormula: "Gross Incurred Claims - Reinsurers Share of Incurred Claims",
    definition: "Total incurred claims after adjusting for reinsurance recoveries and salvages.",
    confidenceScore: 85,
    matchType: "Partial Match",
    rationale: "Formulas overlap on salvage deductibles but report is missing reinsurer retrocession calculations."
  },
  {
    id: 4,
    reportName: "Underwriting Performance.json",
    reportKpi: "Acquisition Cost Ratio",
    bestOntologyKpi: "Acquisition Expense Ratio",
    reportFormula: "Underwriting Commissions / Direct Written Premium",
    ontologyFormula: "Commissions and Direct Acquisition Expenses / Gross Written Premium",
    definition: "Commissions and direct acquisition costs expressed as a percentage of gross written premium.",
    confidenceScore: 91,
    matchType: "Close Match",
    rationale: "Matches conceptual definition of acquisition expenses; ontology formula includes additional non-commission acquisition charges."
  }
];

export const mockCommonOntologyMapping = [
  {
    id: 1,
    ontologyKpi: "Gross Written Premium",
    ontologyDefinition: "Total premium charged to policyholders before reinsurance deductions.",
    ontologyFormula: "Direct Premium Written + Assumed Reinsurance Premium",
    reportAKpi: "Gross Written Premium (GWP)",
    reportAFormula: "Sum of all direct and reinsurance premiums written",
    reportBKpi: "GWP Total",
    reportBFormula: "Direct premium + Assumed treaty/facultative premium",
    reportAScore: 98,
    reportBScore: 97,
    commonalityStatus: "Consistent"
  },
  {
    id: 2,
    ontologyKpi: "Net Claims Incurred",
    ontologyDefinition: "Total incurred claims after adjusting for reinsurance recoveries and salvages.",
    ontologyFormula: "Gross Incurred Claims - Reinsurers Share of Incurred Claims",
    reportAKpi: "Net Incurred Claims",
    reportAFormula: "Claims Paid - Salvage & Subrogation + Outstanding Claim Reserves",
    reportBKpi: "Net Losses Incurred",
    reportBFormula: "Incurred Losses - Net Reinsurance Recoverable",
    reportAScore: 85,
    reportBScore: 94,
    commonalityStatus: "Slight Gap"
  },
  {
    id: 3,
    ontologyKpi: "Unearned Premium Reserve (UPR)",
    ontologyDefinition: "Portion of premium written that corresponds to the unexpired risk period.",
    ontologyFormula: "Pro-rata allocation of premium based on days remaining on policy",
    reportAKpi: "UPR Balance",
    reportAFormula: "Unearned premium calculated on 1/365 daily pro-rata method",
    reportBKpi: "Deferred Premium Asset",
    reportBFormula: "Premium receivables due post balance sheet date",
    reportAScore: 96,
    reportBScore: 42,
    commonalityStatus: "Inconsistent"
  }
];

export const mockFormulaComparison = [
  {
    id: 1,
    reportName: "Q4 Actuarial Valuation.xlsx",
    reportKpi: "Gross Written Premium (GWP)",
    ontologyKpi: "Gross Written Premium",
    reportFormula: "Sum of all direct and reinsurance premiums written",
    ontologyFormula: "Direct Premium Written + Assumed Reinsurance Premium",
    formulaMatchScore: 98,
    formulaGap: "None",
    rationale: "Identical semantic calculation. 'Direct and reinsurance premiums' maps to 'Direct Premium + Assumed Reinsurance'."
  },
  {
    id: 2,
    reportName: "Financial Reserve Report.csv",
    reportKpi: "Net Incurred Claims",
    ontologyKpi: "Net Claims Incurred",
    reportFormula: "Claims Paid - Salvage & Subrogation + Outstanding Claim Reserves",
    ontologyFormula: "Gross Incurred Claims - Reinsurers Share of Incurred Claims",
    formulaMatchScore: 78,
    formulaGap: "Reinsurance retrocession recovery factors are missing from report formula.",
    rationale: "The report formula focuses purely on claims net of salvage, ignoring reinsurance outward mitigation."
  },
  {
    id: 3,
    reportName: "Underwriting Performance.json",
    reportKpi: "Acquisition Cost Ratio",
    ontologyKpi: "Acquisition Expense Ratio",
    reportFormula: "Underwriting Commissions / Direct Written Premium",
    ontologyFormula: "Commissions and Direct Acquisition Expenses / Gross Written Premium",
    formulaMatchScore: 84,
    formulaGap: "Denominator uses Direct Written Premium instead of Gross Written Premium.",
    rationale: "Direct Written Premium excludes assumed reinsurance premiums, which are included in the ontology's Gross Written Premium."
  }
];

export const mockMissingOntologyRecommendations = [
  {
    id: 1,
    reportName: "Q4 Actuarial Valuation.xlsx",
    reportSection: "Schedule 3: Catastrophe Modeling",
    recommendedKpi: "Probable Maximum Loss (PML)",
    industryRelevance: "High (Solvency II & Rating Agency monitoring)",
    recommendedTo: "Ontology KPI List",
    rationale: "The actuarial report has a dedicated section measuring catastrophe exposure using PML metrics, which is currently missing from the enterprise ontology library."
  },
  {
    id: 2,
    reportName: "Financial Reserve Report.csv",
    reportSection: "Note 4: Claims Development Triangles",
    recommendedKpi: "Tail Factor (Ultimate-to-60 Months)",
    industryRelevance: "Medium (Reserve volatility benchmarking)",
    recommendedTo: "Ontology KPI List",
    rationale: "Identified developmental tail factors in reserving triangles. Standardizing tail factors allows consistent projection comparison across business lines."
  }
];

export const mockReportWiseMissingKpis = [
  {
    id: 1,
    reportName: "Financial Reserve Report.csv",
    missingKpi: "Unearned Premium Reserve (UPR)",
    businessPurpose: "Required for liability estimation and revenue recognition mapping",
    priority: "High",
    rationale: "Ontology standard defines UPR for all underwriting reports, but it is missing from this financial reserve report, which may impair earned-premium match validation."
  },
  {
    id: 2,
    reportName: "Underwriting Performance.json",
    missingKpi: "Net Leverage Ratio",
    businessPurpose: "Assess underwriting exposure relative to capital/surplus",
    priority: "Medium",
    rationale: "Essential for financial strength auditing. The report has premiums and surplus values but fails to output the resulting leverage metric defined in ontology."
  }
];

export const mockFeedbackReview = [
  {
    id: 1,
    reportName: "Q4 Actuarial Valuation.xlsx",
    reportKpi: "Gross Written Premium (GWP)",
    ontologyKpi: "Gross Written Premium",
    matchScore: 98,
    status: "Correct",
    comments: "Verified by Actuarial Lead. Perfect match."
  },
  {
    id: 2,
    reportName: "Financial Reserve Report.csv",
    reportKpi: "Net Incurred Claims",
    ontologyKpi: "Net Claims Incurred",
    matchScore: 85,
    status: "Partially Correct",
    comments: "Formulas differ slightly; need to flag this to the valuation team."
  },
  {
    id: 3,
    reportName: "Underwriting Performance.json",
    reportKpi: "Acquisition Cost Ratio",
    ontologyKpi: "Acquisition Expense Ratio",
    matchScore: 91,
    status: "Correct",
    comments: "Close enough match for high-level expense group mapping."
  }
];
