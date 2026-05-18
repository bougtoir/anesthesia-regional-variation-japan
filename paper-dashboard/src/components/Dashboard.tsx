import { useDashboard } from '../context/DashboardContext';
import { Sidebar } from './Sidebar';
import { PaneWrapper } from './PaneWrapper';
import { FloatingPane } from './FloatingPane';
import { SettingsPanel } from './SettingsPanel';
import { AddPaperDialog } from './AddPaperDialog';
import { TitleAuthorsPane } from './panes/TitleAuthorsPane';
import { ProgressPane } from './panes/ProgressPane';
import { LinksPane } from './panes/LinksPane';
import { DeliverablesPane } from './panes/DeliverablesPane';
import { DeadlinesPane } from './panes/DeadlinesPane';
import { NotesTodoPane } from './panes/NotesTodoPane';
import { StatisticsPane } from './panes/StatisticsPane';
import { CostsFundingPane } from './panes/CostsFundingPane';
import { CoauthorTasksPane } from './panes/CoauthorTasksPane';
import type { PaneId, Paper } from '../types/paper';

function getPaneContent(paneId: PaneId, paper: Paper) {
  switch (paneId) {
    case 'title-authors':
      return <TitleAuthorsPane paper={paper} />;
    case 'progress':
      return <ProgressPane paper={paper} />;
    case 'links':
      return <LinksPane paper={paper} />;
    case 'deliverables':
      return <DeliverablesPane paper={paper} />;
    case 'deadlines':
      return <DeadlinesPane paper={paper} />;
    case 'notes-todo':
      return <NotesTodoPane paper={paper} />;
    case 'statistics':
      return <StatisticsPane paper={paper} />;
    case 'costs-funding':
      return <CostsFundingPane paper={paper} />;
    case 'coauthor-tasks':
      return <CoauthorTasksPane paper={paper} />;
  }
}

const PANE_FLEX: Record<PaneId, number> = {
  'title-authors': 1.2,
  progress: 2.5,
  links: 1,
  deliverables: 1,
  deadlines: 0.9,
  'notes-todo': 1.5,
  statistics: 1,
  'costs-funding': 1,
  'coauthor-tasks': 1.3,
};

const ROW_GROUPS: PaneId[][] = [
  ['title-authors', 'progress'],
  ['links', 'deliverables', 'deadlines'],
  ['notes-todo', 'statistics'],
  ['costs-funding', 'coauthor-tasks'],
];

export function Dashboard() {
  const {
    selectedPaper,
    panes,
    settingsOpen,
    setSettingsOpen,
    addPaperDialogOpen,
    setAddPaperDialogOpen,
    editingPaper,
    setEditingPaper,
    exportJson,
    exportYaml,
    deletePaper,
  } = useDashboard();

  const dockedPanes = panes.filter((p) => p.docked);
  const floatingPanes = panes.filter((p) => !p.docked);
  const dockedIds = new Set(dockedPanes.map((p) => p.id));

  const dockedRows: { paneId: PaneId; flex: number }[][] = [];
  for (const group of ROW_GROUPS) {
    const row = group.filter((id) => dockedIds.has(id)).map((id) => ({ paneId: id, flex: PANE_FLEX[id] }));
    if (row.length > 0) dockedRows.push(row);
  }

  const orphanDocked = dockedPanes.filter((p) => !ROW_GROUPS.flat().includes(p.id));
  if (orphanDocked.length > 0) {
    dockedRows.push(orphanDocked.map((p) => ({ paneId: p.id, flex: 1 })));
  }

  const handleExport = (format: 'json' | 'yaml') => {
    const data = format === 'json' ? exportJson() : exportYaml();
    const blob = new Blob([data], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `papers.${format === 'json' ? 'json' : 'yaml'}`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column', background: '#f0ede6' }}>
      {/* Top Bar */}
      <div
        style={{
          height: 44,
          background: 'linear-gradient(135deg, #1a3a4a 0%, #2a5a6a 100%)',
          display: 'flex',
          alignItems: 'center',
          padding: '0 20px',
          color: '#e8e4da',
          boxShadow: '0 2px 8px rgba(0,0,0,0.15)',
          zIndex: 100,
          flexShrink: 0,
        }}
      >
        <div style={{ fontSize: 16, fontWeight: 700, letterSpacing: 1, marginRight: 32 }}>
          Paper<span style={{ color: '#8cc' }}>Hub</span>
        </div>
        <div style={{ marginLeft: 'auto', display: 'flex', gap: 10, alignItems: 'center' }}>
          <button
            onClick={() => {
              setEditingPaper(null);
              setAddPaperDialogOpen(true);
            }}
            style={{
              background: '#4a9',
              color: 'white',
              border: 'none',
              padding: '5px 14px',
              borderRadius: 4,
              fontSize: 12,
              cursor: 'pointer',
              fontWeight: 600,
            }}
          >
            + New Paper
          </button>
          {selectedPaper && (
            <button
              onClick={() => {
                if (confirm(`Delete "${selectedPaper.shortTitle}"?`)) {
                  deletePaper(selectedPaper.id);
                }
              }}
              style={{
                background: 'none',
                border: '1px solid rgba(255,255,255,0.3)',
                color: '#faa',
                padding: '4px 10px',
                borderRadius: 4,
                fontSize: 12,
                cursor: 'pointer',
              }}
            >
              Delete
            </button>
          )}
          <button
            onClick={() => handleExport('json')}
            style={{
              background: 'none',
              border: '1px solid rgba(255,255,255,0.3)',
              color: '#ccc',
              padding: '4px 10px',
              borderRadius: 4,
              fontSize: 12,
              cursor: 'pointer',
            }}
          >
            Export JSON
          </button>
          <button
            onClick={() => handleExport('yaml')}
            style={{
              background: 'none',
              border: '1px solid rgba(255,255,255,0.3)',
              color: '#ccc',
              padding: '4px 10px',
              borderRadius: 4,
              fontSize: 12,
              cursor: 'pointer',
            }}
          >
            Export YAML
          </button>
          <button
            onClick={() => setSettingsOpen(!settingsOpen)}
            style={{
              background: 'none',
              border: '1px solid rgba(255,255,255,0.3)',
              color: '#ccc',
              padding: '4px 10px',
              borderRadius: 4,
              fontSize: 12,
              cursor: 'pointer',
            }}
          >
            Settings
          </button>
        </div>
      </div>

      {/* Main Layout */}
      <div style={{ display: 'flex', flex: 1, overflow: 'hidden' }}>
        <Sidebar />

        {/* Content Area */}
        <div
          style={{
            flex: 1,
            overflowY: 'auto',
            padding: 16,
            display: 'flex',
            flexDirection: 'column',
            gap: 12,
          }}
        >
          {selectedPaper ? (
            <>
              {dockedRows.map((row, ri) => (
                <div key={ri} style={{ display: 'flex', gap: 12, flexShrink: 0 }}>
                  {row.map(({ paneId, flex }) => {
                    const paneConfig = panes.find((p) => p.id === paneId);
                    if (!paneConfig) return null;
                    return (
                      <div key={paneId} style={{ flex }}>
                        <PaneWrapper paneId={paneId} number={paneConfig.number} title={paneConfig.label}>
                          {getPaneContent(paneId, selectedPaper)}
                        </PaneWrapper>
                      </div>
                    );
                  })}
                </div>
              ))}
            </>
          ) : (
            <div
              style={{
                flex: 1,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                color: '#aaa',
                fontSize: 14,
              }}
            >
              Select a paper from the sidebar or add a new one
            </div>
          )}
        </div>
      </div>

      {/* Floating Panes */}
      {selectedPaper &&
        floatingPanes.map((pane) => (
          <FloatingPane
            key={pane.id}
            paneId={pane.id}
            number={pane.number}
            title={pane.label}
            initialPosition={pane.floatingPosition}
            initialSize={pane.floatingSize}
          >
            {getPaneContent(pane.id, selectedPaper)}
          </FloatingPane>
        ))}

      {/* Settings Panel */}
      {settingsOpen && <SettingsPanel />}

      {/* Add/Edit Paper Dialog */}
      {(addPaperDialogOpen || editingPaper) && <AddPaperDialog />}
    </div>
  );
}
