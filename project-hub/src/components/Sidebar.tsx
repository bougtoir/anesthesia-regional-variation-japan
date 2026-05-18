import { useDashboard } from '../context/DashboardContext';
import { STATUS_DOT_COLORS } from './common/statusColors';

export function Sidebar() {
  const { papers, selectedPaperId, setSelectedPaperId, setAddPaperDialogOpen, setEditingPaper } = useDashboard();

  return (
    <div
      style={{
        width: 240,
        background: '#e8e4da',
        borderRight: '1px solid #d0ccc0',
        overflowY: 'auto',
        flexShrink: 0,
        display: 'flex',
        flexDirection: 'column',
      }}
    >
      <div
        style={{
          padding: '14px 16px 8px',
          fontSize: 11,
          fontWeight: 700,
          textTransform: 'uppercase',
          letterSpacing: 1,
          color: '#888',
        }}
      >
        Projects ({papers.length})
      </div>

      <div style={{ flex: 1, overflowY: 'auto' }}>
        {papers.map((paper) => {
          const active = paper.id === selectedPaperId;
          return (
            <div
              key={paper.id}
              onClick={() => setSelectedPaperId(paper.id)}
              onDoubleClick={() => setEditingPaper(paper)}
              style={{
                padding: '10px 16px',
                cursor: 'pointer',
                borderLeft: `3px solid ${active ? '#2a7a8a' : 'transparent'}`,
                background: active ? '#fff' : undefined,
                fontWeight: active ? 600 : 400,
                display: 'flex',
                alignItems: 'center',
                gap: 10,
                transition: 'all 0.15s',
              }}
            >
              <div
                style={{
                  width: 8,
                  height: 8,
                  borderRadius: '50%',
                  background: STATUS_DOT_COLORS[paper.progress.status] ?? '#ccc',
                  flexShrink: 0,
                }}
              />
              <div style={{ minWidth: 0 }}>
                <div
                  style={{
                    fontSize: 12,
                    lineHeight: 1.4,
                    overflow: 'hidden',
                    textOverflow: 'ellipsis',
                    whiteSpace: 'nowrap',
                  }}
                >
                  {paper.shortTitle}
                </div>
                <div style={{ fontSize: 10, color: '#999', marginTop: 2 }}>
                  {paper.progress.journal} ({paper.progress.status})
                </div>
              </div>
            </div>
          );
        })}
      </div>

      <div style={{ padding: '12px 16px', borderTop: '1px solid #d0ccc0' }}>
        <button
          onClick={() => {
            setEditingPaper(null);
            setAddPaperDialogOpen(true);
          }}
          style={{
            width: '100%',
            padding: 8,
            background: 'none',
            border: '1px dashed #aaa',
            borderRadius: 4,
            color: '#777',
            cursor: 'pointer',
            fontSize: 12,
          }}
        >
          + Add New Project
        </button>
      </div>
    </div>
  );
}
