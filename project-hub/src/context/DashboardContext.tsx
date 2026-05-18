import { createContext, useContext, useCallback, useState, useMemo, type ReactNode } from 'react';
import type { Paper, PaneConfig, PaneId } from '../types/paper';
import { useLocalStorage } from '../hooks/useLocalStorage';
import { useServerEvents } from '../hooks/useServerEvents';
import { samplePapers } from '../data/sample-papers';

const DEFAULT_PANES: PaneConfig[] = [
  { id: 'title-authors', label: 'Project & Team', number: 1, docked: true },
  { id: 'progress', label: 'Progress Timeline', number: 2, docked: true },
  { id: 'links', label: 'Links', number: 3, docked: true },
  { id: 'deliverables', label: 'Deliverables', number: 4, docked: true },
  { id: 'deadlines', label: 'Deadlines', number: 5, docked: true },
  { id: 'notes-todo', label: 'Notes & TODO', number: 6, docked: true },
  { id: 'statistics', label: 'Metrics', number: 7, docked: true },
  { id: 'costs-funding', label: 'Budget & Funding', number: 8, docked: false, floatingPosition: { x: 100, y: 100 }, floatingSize: { width: 480, height: 360 } },
  { id: 'coauthor-tasks', label: 'Team Tasks', number: 9, docked: false, floatingPosition: { x: 200, y: 150 }, floatingSize: { width: 500, height: 400 } },
];

interface DashboardContextType {
  papers: Paper[];
  setPapers: (papers: Paper[] | ((prev: Paper[]) => Paper[])) => void;
  selectedPaperId: string;
  setSelectedPaperId: (id: string) => void;
  selectedPaper: Paper | undefined;
  panes: PaneConfig[];
  setPanes: (panes: PaneConfig[] | ((prev: PaneConfig[]) => PaneConfig[])) => void;
  togglePaneDocked: (id: PaneId) => void;
  settingsOpen: boolean;
  setSettingsOpen: (open: boolean) => void;
  addPaperDialogOpen: boolean;
  setAddPaperDialogOpen: (open: boolean) => void;
  editingPaper: Paper | null;
  setEditingPaper: (paper: Paper | null) => void;
  updatePaper: (paper: Paper) => void;
  addPaper: (paper: Paper) => void;
  deletePaper: (id: string) => void;
  exportJson: () => string;
  exportYaml: () => string;
  importData: (data: string, format: 'json' | 'yaml') => void;
}

const DashboardContext = createContext<DashboardContextType | null>(null);

export function DashboardProvider({ children }: { children: ReactNode }) {
  const [papers, setPapers] = useLocalStorage<Paper[]>('projecthub-projects', samplePapers);
  const [selectedPaperId, setSelectedPaperId] = useLocalStorage<string>('projecthub-selected', samplePapers[0]?.id ?? '');
  const [panes, setPanes] = useLocalStorage<PaneConfig[]>('projecthub-panes', DEFAULT_PANES);
  const [settingsOpen, setSettingsOpen] = useLocalStorage<boolean>('projecthub-settings', false);
  const [addPaperDialogOpen, setAddPaperDialogOpen] = useState(false);
  const [editingPaper, setEditingPaper] = useState<Paper | null>(null);

  const selectedPaper = papers.find((p) => p.id === selectedPaperId);

  // SSE: receive real-time updates from server plugins (e.g. folder-watcher)
  const sseHandlers = useMemo(
    () => ({
      'project-update': (data: unknown) => {
        const { project } = data as { project: Paper };
        if (project?.id) setPapers((prev) => prev.map((p) => (p.id === project.id ? project : p)));
      },
      'project-add': (data: unknown) => {
        const { project } = data as { project: Paper };
        if (project?.id) setPapers((prev) => (prev.some((p) => p.id === project.id) ? prev : [...prev, project]));
      },
      'project-delete': (data: unknown) => {
        const { id } = data as { id: string };
        if (id) setPapers((prev) => prev.filter((p) => p.id !== id));
      },
      'project-replace': (data: unknown) => {
        const { projects } = data as { projects: Paper[] };
        if (Array.isArray(projects)) setPapers(projects);
      },
    }),
    [setPapers],
  );
  useServerEvents(sseHandlers);

  const togglePaneDocked = useCallback(
    (id: PaneId) => {
      setPanes((prev) => prev.map((p) => (p.id === id ? { ...p, docked: !p.docked } : p)));
    },
    [setPanes],
  );

  const updatePaper = useCallback(
    (paper: Paper) => {
      setPapers((prev) => prev.map((p) => (p.id === paper.id ? paper : p)));
    },
    [setPapers],
  );

  const addPaper = useCallback(
    (paper: Paper) => {
      setPapers((prev) => [...prev, paper]);
      setSelectedPaperId(paper.id);
    },
    [setPapers, setSelectedPaperId],
  );

  const deletePaper = useCallback(
    (id: string) => {
      setPapers((prev) => {
        const next = prev.filter((p) => p.id !== id);
        if (selectedPaperId === id && next.length > 0) {
          setSelectedPaperId(next[0].id);
        }
        return next;
      });
    },
    [setPapers, selectedPaperId, setSelectedPaperId],
  );

  const exportJson = useCallback(() => JSON.stringify(papers, null, 2), [papers]);

  const exportYaml = useCallback(() => {
    // Simple YAML serializer for project data
    const yamlLines: string[] = ['projects:'];
    for (const paper of papers) {
      yamlLines.push(`  - id: "${paper.id}"`);
      yamlLines.push(`    title: "${paper.title}"`);
      yamlLines.push(`    shortTitle: "${paper.shortTitle}"`);
      yamlLines.push(`    authors:`);
      for (const a of paper.authors) {
        yamlLines.push(`      - name: "${a.name}"`);
        yamlLines.push(`        role: "${a.role}"`);
        if (a.affiliation) yamlLines.push(`        affiliation: "${a.affiliation}"`);
      }
      yamlLines.push(`    progress:`);
      yamlLines.push(`      status: "${paper.progress.status}"`);
      yamlLines.push(`      journal: "${paper.progress.journal}"`);
      if (paper.progress.impactFactor) yamlLines.push(`      impactFactor: ${paper.progress.impactFactor}`);
      if (paper.progress.submissionDate) yamlLines.push(`      submissionDate: "${paper.progress.submissionDate}"`);
      if (paper.progress.revisionDueDate) yamlLines.push(`      revisionDueDate: "${paper.progress.revisionDueDate}"`);
      if (paper.submissionHistory?.length) {
        yamlLines.push(`    submissionHistory:`);
        for (const sh of paper.submissionHistory) {
          yamlLines.push(`      - journal: "${sh.journal}"`);
          if (sh.impactFactor) yamlLines.push(`        impactFactor: ${sh.impactFactor}`);
          yamlLines.push(`        submissionDate: "${sh.submissionDate}"`);
          if (sh.decisionDate) yamlLines.push(`        decisionDate: "${sh.decisionDate}"`);
          yamlLines.push(`        decision: "${sh.decision}"`);
          if (sh.editorComment) yamlLines.push(`        editorComment: "${sh.editorComment}"`);
        }
      }
      yamlLines.push(`    links:`);
      for (const l of paper.links) {
        yamlLines.push(`      - label: "${l.label}"`);
        yamlLines.push(`        url: "${l.url}"`);
        yamlLines.push(`        category: "${l.category}"`);
      }
      yamlLines.push(`    deliverables:`);
      for (const d of paper.deliverables) {
        yamlLines.push(`      - name: "${d.name}"`);
        yamlLines.push(`        type: "${d.type}"`);
        yamlLines.push(`        status: "${d.status}"`);
      }
      yamlLines.push(`    statistics:`);
      yamlLines.push(`      wordCount: ${paper.statistics.wordCount}`);
      if (paper.statistics.wordLimit) yamlLines.push(`      wordLimit: ${paper.statistics.wordLimit}`);
      yamlLines.push(`      figureCount: ${paper.statistics.figureCount}`);
      yamlLines.push(`      tableCount: ${paper.statistics.tableCount}`);
      yamlLines.push(`      referenceCount: ${paper.statistics.referenceCount}`);
    }
    return yamlLines.join('\n');
  }, [papers]);

  const importData = useCallback(
    (data: string, format: 'json' | 'yaml') => {
      try {
        if (format === 'json') {
          const parsed = JSON.parse(data) as Paper[];
          if (Array.isArray(parsed)) {
            setPapers(parsed);
            if (parsed.length > 0) setSelectedPaperId(parsed[0].id);
          }
        } else {
          // For YAML, we use js-yaml
          import('js-yaml').then((yaml) => {
            try {
              const parsed = yaml.load(data) as { projects?: Paper[]; papers?: Paper[] };
              const items = parsed?.projects ?? parsed?.papers;
              if (items && Array.isArray(items)) {
                setPapers(items);
                if (items.length > 0) setSelectedPaperId(items[0].id);
              }
            } catch (e) {
              console.error('YAML import failed:', e);
            }
          }).catch((e) => console.error('Failed to load js-yaml:', e));
        }
      } catch (e) {
        console.error('Import failed:', e);
      }
    },
    [setPapers, setSelectedPaperId],
  );

  return (
    <DashboardContext.Provider
      value={{
        papers,
        setPapers,
        selectedPaperId,
        setSelectedPaperId,
        selectedPaper,
        panes,
        setPanes,
        togglePaneDocked,
        settingsOpen,
        setSettingsOpen,
        addPaperDialogOpen,
        setAddPaperDialogOpen,
        editingPaper,
        setEditingPaper,
        updatePaper,
        addPaper,
        deletePaper,
        exportJson,
        exportYaml,
        importData,
      }}
    >
      {children}
    </DashboardContext.Provider>
  );
}

// eslint-disable-next-line react-refresh/only-export-components
export function useDashboard() {
  const ctx = useContext(DashboardContext);
  if (!ctx) throw new Error('useDashboard must be used within DashboardProvider');
  return ctx;
}
