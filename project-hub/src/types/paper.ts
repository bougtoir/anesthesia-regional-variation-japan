export type PaperStatus =
  | 'planning'
  | 'in-progress'
  | 'in-review'
  | 'on-hold'
  | 'testing'
  | 'approved'
  | 'completed'
  | 'cancelled';

export interface Author {
  name: string;
  affiliation?: string;
  email?: string;
  role: 'lead' | 'manager' | 'member';
}

export interface TimelineEvent {
  label: string;
  startDate: string;
  endDate?: string;
  type: 'planning' | 'development' | 'review' | 'testing' | 'deployment' | 'milestone';
}

export interface Progress {
  status: PaperStatus;
  journal: string;
  impactFactor?: number;
  submissionDate?: string;
  reviewStartDate?: string;
  reviewEndDate?: string;
  revisionDueDate?: string;
  acceptanceDate?: string;
  publicationDate?: string;
  timeline: TimelineEvent[];
}

export interface Link {
  label: string;
  url: string;
  category: 'board' | 'devin' | 'repository' | 'other';
}

export interface Deliverable {
  name: string;
  type: 'feature' | 'bugfix' | 'docs' | 'design' | 'api' | 'infrastructure' | 'other';
  version?: string;
  lastUpdated?: string;
  status: 'complete' | 'in-progress' | 'pending';
}

export interface Deadline {
  label: string;
  date: string;
  type: 'sprint' | 'milestone' | 'release' | 'demo' | 'other';
}

export interface Note {
  id: string;
  content: string;
  createdAt: string;
}

export interface TodoItem {
  id: string;
  text: string;
  done: boolean;
  priority: 'high' | 'medium' | 'low';
}

export interface Statistics {
  wordCount: number;
  wordLimit?: number;
  figureCount: number;
  figureLimit?: number;
  tableCount: number;
  tableLimit?: number;
  referenceCount: number;
  referenceLimit?: number;
  pageCount?: number;
}

export interface FundingSource {
  name: string;
  budget: number;
  allocated: number;
  currency: string;
  status: 'approved' | 'pending' | 'rejected';
}

export interface CostInfo {
  apcEstimate?: number;
  apcCurrency?: string;
  apcPaid: number;
  fundingSources: FundingSource[];
}

export interface CoauthorTask {
  authorName: string;
  initials: string;
  task: string;
  status: 'done' | 'in-progress' | 'waiting';
  waitingDays?: number;
  color?: string;
}

export interface SubmissionRecord {
  journal: string;
  impactFactor?: number;
  submissionDate: string;
  decisionDate?: string;
  decision: 'rejected' | 'withdrawn' | 'desk-reject';
  editorComment?: string;
}

export interface Paper {
  id: string;
  title: string;
  shortTitle: string;
  authors: Author[];
  progress: Progress;
  submissionHistory: SubmissionRecord[];
  links: Link[];
  deliverables: Deliverable[];
  deadlines: Deadline[];
  notes: Note[];
  todos: TodoItem[];
  statistics: Statistics;
  costs: CostInfo;
  coauthorTasks: CoauthorTask[];
}

export type PaneId =
  | 'title-authors'
  | 'progress'
  | 'links'
  | 'deliverables'
  | 'deadlines'
  | 'notes-todo'
  | 'statistics'
  | 'costs-funding'
  | 'coauthor-tasks';

export interface PaneConfig {
  id: PaneId;
  label: string;
  number: number;
  docked: boolean;
  floatingPosition?: { x: number; y: number };
  floatingSize?: { width: number; height: number };
}

export interface LayoutConfig {
  panes: PaneConfig[];
}
