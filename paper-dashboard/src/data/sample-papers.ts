import type { Paper } from '../types/paper';

export const samplePapers: Paper[] = [
  {
    id: 'pmea-zero-cal',
    title: 'Evaluation of zero-calibration continuous non-invasive arterial pressure monitoring using peripheral mean arterial pressure equivalence algorithm',
    shortTitle: 'Zero-cal PMEA',
    authors: [
      { name: 'T. Onishi', role: 'corresponding', affiliation: 'Dept. of Anesthesiology', email: 'onishi@example.ac.jp' },
      { name: 'K. Yamada', role: 'co-author', affiliation: 'Dept. of Biomedical Engineering' },
      { name: 'S. Tanaka', role: 'co-author', affiliation: 'Dept. of Surgery' },
      { name: 'H. Watanabe', role: 'co-author', affiliation: 'Dept. of Anesthesiology' },
      { name: 'M. Suzuki', role: 'co-author', affiliation: 'Dept. of Anesthesiology' },
    ],
    progress: {
      status: 'revision',
      journal: 'Br J Anaesth',
      impactFactor: 9.1,
      submissionDate: '2026-02-12',
      reviewStartDate: '2026-02-18',
      reviewEndDate: '2026-04-20',
      revisionDueDate: '2026-05-16',
      timeline: [
        { label: 'Drafting', startDate: '2026-01-05', endDate: '2026-02-10', type: 'drafting' },
        { label: 'Submission', startDate: '2026-02-12', endDate: '2026-02-12', type: 'submission' },
        { label: 'Peer Review', startDate: '2026-02-18', endDate: '2026-04-20', type: 'review' },
        { label: 'Revision', startDate: '2026-05-01', type: 'revision' },
      ],
    },
    submissionHistory: [
      { journal: 'Anesthesiology', impactFactor: 8.8, submissionDate: '2025-11-10', decisionDate: '2026-01-03', decision: 'rejected', editorComment: 'Interesting concept but insufficient sample size for primary endpoint.' },
    ],
    links: [
      { label: 'BJA ScholarOne Manuscripts', url: 'https://mc.manuscriptcentral.com/bja', category: 'submission' },
      { label: 'Session: PMEA manuscript draft', url: 'https://app.devin.ai/sessions/example1', category: 'devin' },
      { label: 'Session: Revision figures', url: 'https://app.devin.ai/sessions/example2', category: 'devin' },
      { label: 'bougtoir/bja-initial-compartment-pk', url: 'https://github.com/bougtoir/bja-initial-compartment-pk', category: 'repository' },
    ],
    deliverables: [
      { name: 'Main manuscript (.docx)', type: 'manuscript', version: 'v3.2', lastUpdated: '2026-05-05', status: 'complete' },
      { name: 'Figures (Fig 1-5)', type: 'figure', version: 'v2.0', lastUpdated: '2026-04-28', status: 'in-progress' },
      { name: 'Supplementary tables', type: 'table', version: 'v1.1', lastUpdated: '2026-04-15', status: 'complete' },
      { name: 'Cover letter (revision)', type: 'cover-letter', status: 'pending' },
      { name: 'Response to reviewers', type: 'response', status: 'pending' },
    ],
    deadlines: [
      { label: 'Revision submission deadline', date: '2026-05-16', type: 'revision' },
      { label: 'Proof review (estimated)', date: '2026-05-30', type: 'proof' },
    ],
    notes: [
      { id: 'n1', content: 'Reviewer #2 requested additional Bland-Altman analysis for the subgroup with MAP < 65 mmHg. Need to re-run the pipeline with filtered dataset. Reviewer #1 comments are mostly editorial.', createdAt: '2026-05-01' },
    ],
    todos: [
      { id: 't1', text: 'Re-run BA analysis for MAP < 65 subgroup', done: false, priority: 'high' },
      { id: 't2', text: 'Draft response to Reviewer #2', done: false, priority: 'high' },
      { id: 't3', text: 'Update Figure 3 color scheme', done: true, priority: 'medium' },
      { id: 't4', text: 'Fix reference formatting (APA 7th)', done: false, priority: 'low' },
    ],
    statistics: { wordCount: 4280, wordLimit: 5000, figureCount: 5, figureLimit: 6, tableCount: 3, tableLimit: 4, referenceCount: 42, referenceLimit: 50 },
    costs: {
      apcEstimate: 2850,
      apcCurrency: 'GBP',
      apcPaid: 0,
      fundingSources: [
        { name: 'KAKENHI 24K12345', budget: 500000, allocated: 420000, currency: 'JPY', status: 'approved' },
        { name: 'Departmental fund', budget: 200000, allocated: 0, currency: 'JPY', status: 'pending' },
      ],
    },
    coauthorTasks: [
      { authorName: 'T. Onishi', initials: 'TO', task: 'BA analysis rerun + Response to reviewers', status: 'in-progress', color: '#2a7a8a' },
      { authorName: 'K. Yamada', initials: 'KY', task: 'Statistical methods section review', status: 'waiting', waitingDays: 5, color: '#6a5a8a' },
      { authorName: 'S. Tanaka', initials: 'ST', task: 'IRB documentation update', status: 'done', color: '#5a8a6a' },
      { authorName: 'H. Watanabe', initials: 'HW', task: 'Device calibration data verification', status: 'waiting', waitingDays: 3, color: '#8a6a5a' },
      { authorName: 'M. Suzuki', initials: 'MS', task: 'Final approval', status: 'in-progress', color: '#5a7a8a' },
    ],
  },
  {
    id: 'gdp-tempo',
    title: 'GDP tempo-effect PoC: investment-to-output time-to-build lag',
    shortTitle: 'GDP tempo-effect',
    authors: [
      { name: 'T. Onishi', role: 'first', affiliation: 'Economics Research' },
      { name: 'A. Nakamura', role: 'co-author', affiliation: 'Statistics Dept.' },
    ],
    progress: {
      status: 'submitted',
      journal: 'Rev Income Wealth',
      impactFactor: 2.3,
      submissionDate: '2026-05-08',
      timeline: [
        { label: 'Drafting', startDate: '2026-03-01', endDate: '2026-05-05', type: 'drafting' },
        { label: 'Submission', startDate: '2026-05-08', type: 'submission' },
      ],
    },
    submissionHistory: [],
    links: [
      { label: 'RIW Editorial Manager', url: 'https://www.editorialmanager.com/roiw', category: 'submission' },
      { label: 'bougtoir/wip (GDP branch)', url: 'https://github.com/bougtoir/wip/tree/devin/1776760815-gdp-tempo-poc', category: 'repository' },
    ],
    deliverables: [
      { name: 'Main manuscript (.docx)', type: 'manuscript', version: 'v1.0', lastUpdated: '2026-05-08', status: 'complete' },
      { name: 'Figures & analysis', type: 'figure', version: 'v1.0', lastUpdated: '2026-05-07', status: 'complete' },
    ],
    deadlines: [],
    notes: [{ id: 'n2', content: 'Submitted to RIW. Joint identification approach with CWON data.', createdAt: '2026-05-08' }],
    todos: [{ id: 't5', text: 'Prepare rebuttal outline', done: false, priority: 'low' }],
    statistics: { wordCount: 6200, wordLimit: 8000, figureCount: 4, figureLimit: 8, tableCount: 2, tableLimit: 6, referenceCount: 35, referenceLimit: 60 },
    costs: { apcEstimate: 0, apcCurrency: 'USD', apcPaid: 0, fundingSources: [] },
    coauthorTasks: [
      { authorName: 'T. Onishi', initials: 'TO', task: 'Main analysis & writing', status: 'done', color: '#2a7a8a' },
      { authorName: 'A. Nakamura', initials: 'AN', task: 'Data curation review', status: 'done', color: '#8a5a6a' },
    ],
  },
  {
    id: 'weibull-dropout',
    title: 'Weibull-based clinical trial dropout modelling: a unified framework across therapeutic domains',
    shortTitle: 'Weibull dropout',
    authors: [
      { name: 'T. Onishi', role: 'corresponding', affiliation: 'Clinical Pharmacology' },
      { name: 'R. Ito', role: 'co-author', affiliation: 'Biostatistics' },
      { name: 'Y. Kato', role: 'co-author', affiliation: 'Clinical Trials Unit' },
    ],
    progress: {
      status: 'under-review',
      journal: 'Pharm Res',
      impactFactor: 4.6,
      submissionDate: '2026-04-10',
      reviewStartDate: '2026-04-15',
      timeline: [
        { label: 'Drafting', startDate: '2026-02-01', endDate: '2026-04-05', type: 'drafting' },
        { label: 'Submission', startDate: '2026-04-10', endDate: '2026-04-10', type: 'submission' },
        { label: 'Peer Review', startDate: '2026-04-15', type: 'review' },
      ],
    },
    submissionHistory: [],
    links: [
      { label: 'Pharm Res Editorial Manager', url: 'https://www.editorialmanager.com/pham', category: 'submission' },
      { label: 'bougtoir/wip (Weibull branch)', url: 'https://github.com/bougtoir/wip/tree/weibull-clinical-trial-dropout', category: 'repository' },
    ],
    deliverables: [
      { name: 'Main manuscript (.docx)', type: 'manuscript', version: 'v2.1', lastUpdated: '2026-04-09', status: 'complete' },
      { name: 'Supplementary material', type: 'supplement', version: 'v1.0', lastUpdated: '2026-04-09', status: 'complete' },
    ],
    deadlines: [],
    notes: [],
    todos: [],
    statistics: { wordCount: 5500, wordLimit: 6000, figureCount: 6, figureLimit: 8, tableCount: 4, tableLimit: 6, referenceCount: 48, referenceLimit: 50 },
    costs: { apcEstimate: 3000, apcCurrency: 'USD', apcPaid: 0, fundingSources: [{ name: 'AMED Grant', budget: 800000, allocated: 300000, currency: 'JPY', status: 'approved' }] },
    coauthorTasks: [
      { authorName: 'T. Onishi', initials: 'TO', task: 'Framework design & writing', status: 'done', color: '#2a7a8a' },
      { authorName: 'R. Ito', initials: 'RI', task: 'Statistical validation', status: 'done', color: '#5a8a6a' },
      { authorName: 'Y. Kato', initials: 'YK', task: 'Clinical data curation', status: 'done', color: '#8a6a5a' },
    ],
  },
  {
    id: 'spectral-causality',
    title: 'Spectral causality: frequency-domain Granger analysis with topological constraints',
    shortTitle: 'Spectral causality',
    authors: [
      { name: 'T. Onishi', role: 'first', affiliation: 'Computational Methods' },
    ],
    progress: {
      status: 'drafting',
      journal: 'JMLR',
      impactFactor: 6.0,
      timeline: [
        { label: 'Drafting', startDate: '2026-04-20', type: 'drafting' },
      ],
    },
    submissionHistory: [],
    links: [
      { label: 'bougtoir/wip (spectral branch)', url: 'https://github.com/bougtoir/wip/tree/devin/1777462980-spectral-causality-brainstorm', category: 'repository' },
    ],
    deliverables: [
      { name: 'Draft manuscript', type: 'manuscript', version: 'v0.3', lastUpdated: '2026-05-07', status: 'in-progress' },
    ],
    deadlines: [
      { label: 'JMLR abstract deadline', date: '2026-06-23', type: 'submission' },
    ],
    notes: [{ id: 'n3', content: 'Need to finalize the topological constraint formulation. Consider ECD approach.', createdAt: '2026-05-07' }],
    todos: [
      { id: 't6', text: 'Formalize spectral decomposition theorem', done: false, priority: 'high' },
      { id: 't7', text: 'Run simulations on synthetic data', done: false, priority: 'medium' },
    ],
    statistics: { wordCount: 2100, wordLimit: 10000, figureCount: 2, figureLimit: 10, tableCount: 0, tableLimit: 5, referenceCount: 18, referenceLimit: 80 },
    costs: { apcEstimate: 0, apcCurrency: 'USD', apcPaid: 0, fundingSources: [] },
    coauthorTasks: [
      { authorName: 'T. Onishi', initials: 'TO', task: 'All sections', status: 'in-progress', color: '#2a7a8a' },
    ],
  },
  {
    id: 'medical-sapir-whorf',
    title: 'Medical Sapir-Whorf hypothesis: linguistic relativity in clinical decision-making across cultures',
    shortTitle: 'Medical Sapir-Whorf',
    authors: [
      { name: 'T. Onishi', role: 'corresponding', affiliation: 'Dept. of Anesthesiology' },
      { name: 'J. Park', role: 'co-author', affiliation: 'Linguistics Dept.' },
    ],
    progress: {
      status: 'submitted',
      journal: 'Lancet',
      impactFactor: 168.9,
      submissionDate: '2026-05-05',
      timeline: [
        { label: 'Drafting', startDate: '2026-03-15', endDate: '2026-05-03', type: 'drafting' },
        { label: 'Submission', startDate: '2026-05-05', type: 'submission' },
      ],
    },
    submissionHistory: [],
    links: [
      { label: 'Lancet EES', url: 'https://ees.elsevier.com/thelancet', category: 'submission' },
      { label: 'bougtoir/wip (sapir branch)', url: 'https://github.com/bougtoir/wip/tree/devin/1777545460-medical-sapir-whorf-paper', category: 'repository' },
    ],
    deliverables: [
      { name: 'Main manuscript', type: 'manuscript', version: 'v1.0', lastUpdated: '2026-05-05', status: 'complete' },
    ],
    deadlines: [],
    notes: [],
    todos: [],
    statistics: { wordCount: 3800, wordLimit: 4500, figureCount: 3, figureLimit: 5, tableCount: 2, tableLimit: 3, referenceCount: 30, referenceLimit: 40 },
    costs: { apcEstimate: 5000, apcCurrency: 'USD', apcPaid: 0, fundingSources: [] },
    coauthorTasks: [
      { authorName: 'T. Onishi', initials: 'TO', task: 'Clinical framework & writing', status: 'done', color: '#2a7a8a' },
      { authorName: 'J. Park', initials: 'JP', task: 'Linguistic analysis', status: 'done', color: '#6a8a5a' },
    ],
  },
  {
    id: 'cryo-scoping',
    title: 'Cryoanesthesia in competitive sport: a scoping review of WADA TUE considerations',
    shortTitle: 'Cryoanesthesia review',
    authors: [
      { name: 'T. Onishi', role: 'first', affiliation: 'Sports Medicine' },
      { name: 'L. Chen', role: 'co-author', affiliation: 'Anti-Doping Research' },
    ],
    progress: {
      status: 'accepted',
      journal: 'WADA-TUE Review',
      submissionDate: '2026-03-20',
      reviewStartDate: '2026-03-25',
      reviewEndDate: '2026-04-28',
      acceptanceDate: '2026-04-30',
      timeline: [
        { label: 'Drafting', startDate: '2026-01-15', endDate: '2026-03-18', type: 'drafting' },
        { label: 'Submission', startDate: '2026-03-20', endDate: '2026-03-20', type: 'submission' },
        { label: 'Review', startDate: '2026-03-25', endDate: '2026-04-28', type: 'review' },
        { label: 'Accepted', startDate: '2026-04-30', endDate: '2026-04-30', type: 'decision' },
      ],
    },
    submissionHistory: [],
    links: [
      { label: 'bougtoir/wip (cryo branch)', url: 'https://github.com/bougtoir/wip/tree/devin/1774686640-cryoanesthesia-scoping-review', category: 'repository' },
    ],
    deliverables: [
      { name: 'Final manuscript', type: 'manuscript', version: 'v2.0', lastUpdated: '2026-04-30', status: 'complete' },
    ],
    deadlines: [
      { label: 'Proof return', date: '2026-05-20', type: 'proof' },
    ],
    notes: [],
    todos: [{ id: 't8', text: 'Return proof corrections', done: false, priority: 'medium' }],
    statistics: { wordCount: 4000, wordLimit: 5000, figureCount: 2, figureLimit: 4, tableCount: 3, tableLimit: 4, referenceCount: 55, referenceLimit: 60 },
    costs: { apcEstimate: 0, apcCurrency: 'USD', apcPaid: 0, fundingSources: [] },
    coauthorTasks: [
      { authorName: 'T. Onishi', initials: 'TO', task: 'Lead author', status: 'done', color: '#2a7a8a' },
      { authorName: 'L. Chen', initials: 'LC', task: 'WADA policy review', status: 'done', color: '#8a5a7a' },
    ],
  },
  {
    id: 'anesthesia-regional',
    title: 'Regional variation in anesthesia practice patterns across Japan: a national database study',
    shortTitle: 'Anesthesia variation JP',
    authors: [
      { name: 'T. Onishi', role: 'corresponding', affiliation: 'Dept. of Anesthesiology' },
    ],
    progress: {
      status: 'drafting',
      journal: 'Target TBD',
      timeline: [
        { label: 'Drafting', startDate: '2026-04-01', type: 'drafting' },
      ],
    },
    submissionHistory: [],
    links: [
      { label: 'bougtoir/wip (regional branch)', url: 'https://github.com/bougtoir/wip/tree/devin/1774690359-anesthesia-regional-variation-japan', category: 'repository' },
    ],
    deliverables: [
      { name: 'Draft manuscript', type: 'manuscript', version: 'v0.1', lastUpdated: '2026-05-01', status: 'in-progress' },
    ],
    deadlines: [],
    notes: [{ id: 'n4', content: 'Need to decide target journal. Consider JA or Anesthesiology.', createdAt: '2026-05-01' }],
    todos: [
      { id: 't9', text: 'Complete NDB data extraction', done: false, priority: 'high' },
      { id: 't10', text: 'Select target journal', done: false, priority: 'medium' },
    ],
    statistics: { wordCount: 1200, wordLimit: 5000, figureCount: 1, figureLimit: 6, tableCount: 1, tableLimit: 4, referenceCount: 12, referenceLimit: 40 },
    costs: { apcEstimate: 0, apcCurrency: 'JPY', apcPaid: 0, fundingSources: [] },
    coauthorTasks: [
      { authorName: 'T. Onishi', initials: 'TO', task: 'All sections', status: 'in-progress', color: '#2a7a8a' },
    ],
  },
];
