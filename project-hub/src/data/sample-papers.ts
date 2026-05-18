import type { Paper } from '../types/paper';

export const samplePapers: Paper[] = [
  {
    id: 'webapp-redesign',
    title: 'Corporate Website Redesign — Next.js Migration & UI Overhaul',
    shortTitle: 'Web Redesign',
    authors: [
      { name: 'A. Tanaka', role: 'lead', affiliation: 'Engineering', email: 'tanaka@example.com' },
      { name: 'M. Suzuki', role: 'member', affiliation: 'Design' },
      { name: 'K. Yamada', role: 'member', affiliation: 'Engineering' },
      { name: 'R. Sato', role: 'member', affiliation: 'QA' },
    ],
    progress: {
      status: 'in-progress',
      journal: 'Client: Acme Corp',
      impactFactor: 8,
      submissionDate: '2026-03-01',
      reviewStartDate: '2026-04-15',
      revisionDueDate: '2026-06-30',
      timeline: [
        { label: 'Planning', startDate: '2026-01-10', endDate: '2026-02-28', type: 'planning' },
        { label: 'Development Sprint 1', startDate: '2026-03-01', endDate: '2026-03-28', type: 'development' },
        { label: 'Development Sprint 2', startDate: '2026-04-01', endDate: '2026-04-25', type: 'development' },
        { label: 'QA & UAT', startDate: '2026-05-01', type: 'testing' },
      ],
    },
    submissionHistory: [],
    links: [
      { label: 'Jira Board', url: 'https://acme.atlassian.net/jira/software/projects/WEB', category: 'board' },
      { label: 'Session: Homepage prototype', url: 'https://app.devin.ai/sessions/example1', category: 'devin' },
      { label: 'Session: API integration', url: 'https://app.devin.ai/sessions/example2', category: 'devin' },
      { label: 'acme-corp/website-v2', url: 'https://github.com/acme-corp/website-v2', category: 'repository' },
    ],
    deliverables: [
      { name: 'Homepage (Next.js)', type: 'feature', version: 'v2.1', lastUpdated: '2026-05-05', status: 'complete' },
      { name: 'Product catalog pages', type: 'feature', version: 'v1.3', lastUpdated: '2026-04-28', status: 'in-progress' },
      { name: 'CMS integration', type: 'api', version: 'v1.0', lastUpdated: '2026-04-15', status: 'complete' },
      { name: 'Mobile responsive styles', type: 'design', status: 'in-progress' },
      { name: 'Performance audit report', type: 'docs', status: 'pending' },
    ],
    deadlines: [
      { label: 'Sprint 3 demo', date: '2026-05-23', type: 'sprint' },
      { label: 'UAT sign-off', date: '2026-06-15', type: 'milestone' },
      { label: 'Production launch', date: '2026-06-30', type: 'release' },
    ],
    notes: [
      { id: 'n1', content: 'Client requested dark mode support — added to Sprint 3 backlog. Need to confirm with design team on color palette. Lighthouse score target: 90+ across all pages.', createdAt: '2026-05-01' },
    ],
    todos: [
      { id: 't1', text: 'Implement dark mode toggle', done: false, priority: 'high' },
      { id: 't2', text: 'Fix mobile nav hamburger menu on Safari', done: false, priority: 'high' },
      { id: 't3', text: 'Add structured data (JSON-LD) for SEO', done: true, priority: 'medium' },
      { id: 't4', text: 'Update README with deployment guide', done: false, priority: 'low' },
    ],
    statistics: { wordCount: 48, wordLimit: 60, figureCount: 12, figureLimit: 15, tableCount: 3, tableLimit: 5, referenceCount: 8, referenceLimit: 10 },
    costs: {
      apcEstimate: 45000,
      apcCurrency: 'USD',
      apcPaid: 22500,
      fundingSources: [
        { name: 'Acme Corp contract', budget: 50000, allocated: 45000, currency: 'USD', status: 'approved' },
      ],
    },
    coauthorTasks: [
      { authorName: 'A. Tanaka', initials: 'AT', task: 'Sprint 3 feature implementation', status: 'in-progress', color: '#2a7a8a' },
      { authorName: 'M. Suzuki', initials: 'MS', task: 'Dark mode design mockups', status: 'in-progress', color: '#6a5a8a' },
      { authorName: 'K. Yamada', initials: 'KY', task: 'API endpoint optimization', status: 'done', color: '#5a8a6a' },
      { authorName: 'R. Sato', initials: 'RS', task: 'E2E test suite for Sprint 2', status: 'waiting', waitingDays: 2, color: '#8a6a5a' },
    ],
  },
  {
    id: 'mobile-app',
    title: 'Cross-Platform Mobile App — React Native + Expo',
    shortTitle: 'Mobile App',
    authors: [
      { name: 'Y. Kimura', role: 'lead', affiliation: 'Mobile Team' },
      { name: 'H. Ito', role: 'member', affiliation: 'Backend' },
    ],
    progress: {
      status: 'testing',
      journal: 'Internal Product',
      impactFactor: 9,
      submissionDate: '2026-02-01',
      timeline: [
        { label: 'Planning & Prototyping', startDate: '2026-01-15', endDate: '2026-02-10', type: 'planning' },
        { label: 'Core Development', startDate: '2026-02-15', endDate: '2026-04-10', type: 'development' },
        { label: 'Beta Testing', startDate: '2026-04-15', type: 'testing' },
      ],
    },
    submissionHistory: [],
    links: [
      { label: 'Linear Board', url: 'https://linear.app/team/mobile', category: 'board' },
      { label: 'org/mobile-app', url: 'https://github.com/org/mobile-app', category: 'repository' },
    ],
    deliverables: [
      { name: 'iOS build (TestFlight)', type: 'feature', version: 'v0.9.2', lastUpdated: '2026-05-02', status: 'complete' },
      { name: 'Android build (Play Console)', type: 'feature', version: 'v0.9.2', lastUpdated: '2026-05-02', status: 'complete' },
      { name: 'Push notification service', type: 'api', version: 'v1.0', status: 'in-progress' },
    ],
    deadlines: [
      { label: 'Beta feedback deadline', date: '2026-05-20', type: 'milestone' },
      { label: 'App Store submission', date: '2026-06-01', type: 'release' },
    ],
    notes: [],
    todos: [
      { id: 't1', text: 'Fix crash on Android 12 deep linking', done: false, priority: 'high' },
      { id: 't2', text: 'Implement biometric auth', done: false, priority: 'medium' },
    ],
    statistics: { wordCount: 32, wordLimit: 40, figureCount: 8, figureLimit: 10, tableCount: 2, tableLimit: 5, referenceCount: 5, referenceLimit: 8 },
    costs: { apcEstimate: 15000, apcCurrency: 'USD', apcPaid: 12000, fundingSources: [{ name: 'Product budget Q2', budget: 20000, allocated: 15000, currency: 'USD', status: 'approved' }] },
    coauthorTasks: [
      { authorName: 'Y. Kimura', initials: 'YK', task: 'Push notification integration', status: 'in-progress', color: '#2a7a8a' },
      { authorName: 'H. Ito', initials: 'HI', task: 'Backend API for notifications', status: 'done', color: '#5a8a6a' },
    ],
  },
  {
    id: 'data-pipeline',
    title: 'Real-Time Data Pipeline — Kafka + Spark Streaming',
    shortTitle: 'Data Pipeline',
    authors: [
      { name: 'S. Nakamura', role: 'lead', affiliation: 'Data Engineering' },
      { name: 'T. Watanabe', role: 'member', affiliation: 'DevOps' },
      { name: 'E. Kobayashi', role: 'member', affiliation: 'Data Science' },
    ],
    progress: {
      status: 'in-review',
      journal: 'Infrastructure',
      timeline: [
        { label: 'Architecture Design', startDate: '2026-02-01', endDate: '2026-02-28', type: 'planning' },
        { label: 'Implementation', startDate: '2026-03-01', endDate: '2026-04-30', type: 'development' },
        { label: 'Code Review', startDate: '2026-05-01', type: 'review' },
      ],
    },
    submissionHistory: [],
    links: [
      { label: 'Confluence: Architecture Doc', url: 'https://wiki.example.com/data-pipeline', category: 'other' },
      { label: 'org/data-pipeline', url: 'https://github.com/org/data-pipeline', category: 'repository' },
    ],
    deliverables: [
      { name: 'Kafka consumer service', type: 'feature', version: 'v1.0', lastUpdated: '2026-04-25', status: 'complete' },
      { name: 'Spark Streaming jobs', type: 'feature', version: 'v0.8', lastUpdated: '2026-04-30', status: 'in-progress' },
      { name: 'Monitoring dashboard (Grafana)', type: 'infrastructure', status: 'pending' },
      { name: 'Runbook documentation', type: 'docs', status: 'pending' },
    ],
    deadlines: [
      { label: 'Code review completion', date: '2026-05-15', type: 'milestone' },
      { label: 'Staging deployment', date: '2026-05-25', type: 'release' },
    ],
    notes: [
      { id: 'n1', content: 'Need to benchmark throughput with production-scale data. Current test: 50K events/sec. Target: 200K events/sec.', createdAt: '2026-04-28' },
    ],
    todos: [
      { id: 't1', text: 'Load test with 200K events/sec', done: false, priority: 'high' },
      { id: 't2', text: 'Set up dead letter queue', done: false, priority: 'medium' },
    ],
    statistics: { wordCount: 15, wordLimit: 25, figureCount: 4, figureLimit: 8, tableCount: 1, tableLimit: 3, referenceCount: 3, referenceLimit: 5 },
    costs: { apcPaid: 0, fundingSources: [{ name: 'Infrastructure budget', budget: 30000, allocated: 18000, currency: 'USD', status: 'approved' }] },
    coauthorTasks: [
      { authorName: 'S. Nakamura', initials: 'SN', task: 'Spark job optimization', status: 'in-progress', color: '#2a7a8a' },
      { authorName: 'T. Watanabe', initials: 'TW', task: 'K8s deployment manifests', status: 'done', color: '#5a8a6a' },
      { authorName: 'E. Kobayashi', initials: 'EK', task: 'Data quality validation rules', status: 'waiting', waitingDays: 4, color: '#8a6a5a' },
    ],
  },
  {
    id: 'ml-recommendation',
    title: 'Product Recommendation Engine — Collaborative Filtering + LLM',
    shortTitle: 'ML Recommender',
    authors: [
      { name: 'L. Chen', role: 'lead', affiliation: 'ML Team' },
      { name: 'P. Kim', role: 'member', affiliation: 'ML Team' },
    ],
    progress: {
      status: 'planning',
      journal: 'R&D Initiative',
      timeline: [
        { label: 'Research & Benchmarking', startDate: '2026-05-01', endDate: '2026-05-31', type: 'planning' },
      ],
    },
    submissionHistory: [],
    links: [
      { label: 'Notion: Research Notes', url: 'https://notion.so/ml-recommender', category: 'other' },
      { label: 'org/ml-recommender', url: 'https://github.com/org/ml-recommender', category: 'repository' },
    ],
    deliverables: [
      { name: 'Benchmark report', type: 'docs', status: 'in-progress' },
      { name: 'Prototype API', type: 'api', status: 'pending' },
    ],
    deadlines: [
      { label: 'Research phase complete', date: '2026-05-31', type: 'milestone' },
      { label: 'MVP demo to stakeholders', date: '2026-07-15', type: 'demo' },
    ],
    notes: [],
    todos: [
      { id: 't1', text: 'Evaluate ALS vs. neural CF baselines', done: false, priority: 'high' },
      { id: 't2', text: 'Set up MLflow experiment tracking', done: true, priority: 'medium' },
    ],
    statistics: { wordCount: 5, wordLimit: 30, figureCount: 0, figureLimit: 10, tableCount: 0, tableLimit: 5, referenceCount: 0, referenceLimit: 8 },
    costs: { apcPaid: 0, fundingSources: [{ name: 'R&D Innovation Fund', budget: 25000, allocated: 5000, currency: 'USD', status: 'approved' }] },
    coauthorTasks: [
      { authorName: 'L. Chen', initials: 'LC', task: 'Baseline model implementation', status: 'in-progress', color: '#2a7a8a' },
      { authorName: 'P. Kim', initials: 'PK', task: 'Dataset preparation', status: 'done', color: '#6a5a8a' },
    ],
  },
  {
    id: 'security-audit',
    title: 'Annual Security Audit & Compliance (SOC 2 Type II)',
    shortTitle: 'Security Audit',
    authors: [
      { name: 'J. Park', role: 'lead', affiliation: 'Security' },
      { name: 'D. Lee', role: 'member', affiliation: 'Compliance' },
    ],
    progress: {
      status: 'approved',
      journal: 'Compliance',
      submissionDate: '2026-01-15',
      acceptanceDate: '2026-04-30',
      timeline: [
        { label: 'Scope Definition', startDate: '2026-01-15', endDate: '2026-02-01', type: 'planning' },
        { label: 'Evidence Collection', startDate: '2026-02-01', endDate: '2026-03-31', type: 'development' },
        { label: 'External Audit', startDate: '2026-04-01', endDate: '2026-04-25', type: 'review' },
        { label: 'Report Finalized', startDate: '2026-04-30', endDate: '2026-04-30', type: 'milestone' },
      ],
    },
    submissionHistory: [],
    links: [
      { label: 'Vanta Dashboard', url: 'https://app.vanta.com', category: 'board' },
      { label: 'org/security-policies', url: 'https://github.com/org/security-policies', category: 'repository' },
    ],
    deliverables: [
      { name: 'SOC 2 Type II Report', type: 'docs', version: 'Final', lastUpdated: '2026-04-30', status: 'complete' },
      { name: 'Remediation tracker', type: 'docs', version: 'v1.2', lastUpdated: '2026-04-20', status: 'complete' },
    ],
    deadlines: [
      { label: 'Certificate renewal', date: '2027-04-30', type: 'milestone' },
    ],
    notes: [
      { id: 'n1', content: 'All critical findings remediated. 2 low-severity items have accepted risk waivers.', createdAt: '2026-04-30' },
    ],
    todos: [],
    statistics: { wordCount: 8, wordLimit: 10, figureCount: 2, figureLimit: 5, tableCount: 1, tableLimit: 3, referenceCount: 4, referenceLimit: 5 },
    costs: { apcEstimate: 35000, apcCurrency: 'USD', apcPaid: 35000, fundingSources: [{ name: 'Security budget FY2026', budget: 40000, allocated: 35000, currency: 'USD', status: 'approved' }] },
    coauthorTasks: [
      { authorName: 'J. Park', initials: 'JP', task: 'Remediation verification', status: 'done', color: '#2a7a8a' },
      { authorName: 'D. Lee', initials: 'DL', task: 'Policy documentation update', status: 'done', color: '#5a8a6a' },
    ],
  },
  {
    id: 'devops-migration',
    title: 'CI/CD Pipeline Migration — Jenkins to GitHub Actions',
    shortTitle: 'CI/CD Migration',
    authors: [
      { name: 'N. Fujita', role: 'lead', affiliation: 'DevOps' },
    ],
    progress: {
      status: 'completed',
      journal: 'DevOps',
      submissionDate: '2026-01-05',
      acceptanceDate: '2026-03-20',
      publicationDate: '2026-03-25',
      timeline: [
        { label: 'Audit existing pipelines', startDate: '2026-01-05', endDate: '2026-01-20', type: 'planning' },
        { label: 'Migration', startDate: '2026-01-21', endDate: '2026-03-10', type: 'development' },
        { label: 'Validation & Rollout', startDate: '2026-03-11', endDate: '2026-03-25', type: 'deployment' },
      ],
    },
    submissionHistory: [],
    links: [
      { label: 'org/.github', url: 'https://github.com/org/.github', category: 'repository' },
    ],
    deliverables: [
      { name: 'Reusable workflow templates', type: 'infrastructure', version: 'v1.0', lastUpdated: '2026-03-20', status: 'complete' },
      { name: 'Migration runbook', type: 'docs', version: 'v2.0', lastUpdated: '2026-03-25', status: 'complete' },
    ],
    deadlines: [],
    notes: [
      { id: 'n1', content: '23 repos migrated. Jenkins decommissioned. Average build time reduced by 40%.', createdAt: '2026-03-25' },
    ],
    todos: [],
    statistics: { wordCount: 23, wordLimit: 25, figureCount: 5, figureLimit: 5, tableCount: 2, tableLimit: 3, referenceCount: 3, referenceLimit: 5 },
    costs: { apcPaid: 0, fundingSources: [{ name: 'Platform budget', budget: 10000, allocated: 8500, currency: 'USD', status: 'approved' }] },
    coauthorTasks: [
      { authorName: 'N. Fujita', initials: 'NF', task: 'Post-migration monitoring', status: 'done', color: '#2a7a8a' },
    ],
  },
  {
    id: 'design-system',
    title: 'Design System v2 — Component Library & Documentation',
    shortTitle: 'Design System v2',
    authors: [
      { name: 'S. Hayashi', role: 'lead', affiliation: 'Design' },
      { name: 'T. Morita', role: 'member', affiliation: 'Frontend' },
      { name: 'A. Ogawa', role: 'manager', affiliation: 'Product' },
    ],
    progress: {
      status: 'on-hold',
      journal: 'Design',
      timeline: [
        { label: 'Audit & Token Definition', startDate: '2026-02-01', endDate: '2026-03-15', type: 'planning' },
        { label: 'Component Build', startDate: '2026-03-16', endDate: '2026-04-20', type: 'development' },
      ],
    },
    submissionHistory: [
      { journal: 'Q1 Roadmap', submissionDate: '2025-12-01', decisionDate: '2026-01-15', decision: 'withdrawn', editorComment: 'Deprioritized due to Q1 client work. Rescheduled to Q2.' },
    ],
    links: [
      { label: 'Figma: Design System v2', url: 'https://figma.com/design-system-v2', category: 'other' },
      { label: 'Storybook', url: 'https://storybook.example.com', category: 'other' },
      { label: 'org/design-system', url: 'https://github.com/org/design-system', category: 'repository' },
    ],
    deliverables: [
      { name: 'Design tokens (JSON)', type: 'design', version: 'v2.0', lastUpdated: '2026-03-15', status: 'complete' },
      { name: 'React component library', type: 'feature', version: 'v0.6', lastUpdated: '2026-04-20', status: 'in-progress' },
      { name: 'Storybook documentation', type: 'docs', status: 'in-progress' },
    ],
    deadlines: [
      { label: 'Resume development', date: '2026-06-01', type: 'sprint' },
      { label: 'v2.0 release', date: '2026-08-01', type: 'release' },
    ],
    notes: [
      { id: 'n1', content: 'On hold — frontend team pulled to client project. Resume after Sprint 8 (June). 60% of components completed.', createdAt: '2026-04-22' },
    ],
    todos: [
      { id: 't1', text: 'Complete form components (Select, Checkbox, Radio)', done: false, priority: 'high' },
      { id: 't2', text: 'Add accessibility tests (axe-core)', done: false, priority: 'medium' },
    ],
    statistics: { wordCount: 18, wordLimit: 30, figureCount: 6, figureLimit: 10, tableCount: 1, tableLimit: 3, referenceCount: 2, referenceLimit: 5 },
    costs: { apcPaid: 0, fundingSources: [{ name: 'Design Ops budget', budget: 15000, allocated: 9000, currency: 'USD', status: 'approved' }] },
    coauthorTasks: [
      { authorName: 'S. Hayashi', initials: 'SH', task: 'Token finalization', status: 'done', color: '#2a7a8a' },
      { authorName: 'T. Morita', initials: 'TM', task: 'React component build (paused)', status: 'waiting', waitingDays: 18, color: '#6a5a8a' },
      { authorName: 'A. Ogawa', initials: 'AO', task: 'Stakeholder approval for v2 scope', status: 'done', color: '#5a8a6a' },
    ],
  },
];
