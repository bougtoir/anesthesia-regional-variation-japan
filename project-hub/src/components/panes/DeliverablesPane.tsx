import { useState } from 'react';
import type { Paper, Deliverable } from '../../types/paper';
import { useDashboard } from '../../context/DashboardContext';

const TYPE_ICONS: Record<string, string> = {
  feature: '\u2B50',
  bugfix: '\uD83D\uDC1B',
  docs: '\uD83D\uDCC4',
  design: '\uD83C\uDFA8',
  api: '\u2699',
  infrastructure: '\uD83D\uDEE0',
  other: '\uD83D\uDCC1',
};

const STATUS_STYLES: Record<string, { color: string; label: string }> = {
  complete: { color: '#40a060', label: 'Complete' },
  'in-progress': { color: '#e0a040', label: 'In Progress' },
  pending: { color: '#bbb', label: 'Pending' },
};

const STATUS_CYCLE: Deliverable['status'][] = ['pending', 'in-progress', 'complete'];

const inputStyle: React.CSSProperties = {
  padding: '4px 8px',
  border: '1px solid #d0ccc0',
  borderRadius: 4,
  fontSize: 11,
  background: '#faf8f3',
  fontFamily: 'inherit',
};

interface Props {
  paper: Paper;
}

export function DeliverablesPane({ paper }: Props) {
  const { updatePaper } = useDashboard();
  const [showAdd, setShowAdd] = useState(false);
  const [newName, setNewName] = useState('');
  const [newType, setNewType] = useState<Deliverable['type']>('feature');
  const [newVersion, setNewVersion] = useState('');
  const [newStatus, setNewStatus] = useState<Deliverable['status']>('pending');

  const cycleStatus = (index: number) => {
    const d = paper.deliverables[index];
    const curIdx = STATUS_CYCLE.indexOf(d.status);
    const nextStatus = STATUS_CYCLE[(curIdx + 1) % STATUS_CYCLE.length];
    const updated = paper.deliverables.map((item, i) =>
      i === index ? { ...item, status: nextStatus, lastUpdated: new Date().toISOString().slice(0, 10) } : item,
    );
    updatePaper({ ...paper, deliverables: updated });
  };

  const addDeliverable = () => {
    if (!newName.trim()) return;
    updatePaper({
      ...paper,
      deliverables: [
        ...paper.deliverables,
        {
          name: newName.trim(),
          type: newType,
          version: newVersion || undefined,
          lastUpdated: new Date().toISOString().slice(0, 10),
          status: newStatus,
        },
      ],
    });
    setNewName('');
    setNewType('feature');
    setNewVersion('');
    setNewStatus('pending');
    setShowAdd(false);
  };

  const removeDeliverable = (index: number) => {
    updatePaper({ ...paper, deliverables: paper.deliverables.filter((_, i) => i !== index) });
  };

  if (paper.deliverables.length === 0 && !showAdd) {
    return (
      <div>
        <div style={{ color: '#aaa', fontSize: 12 }}>No deliverables</div>
        <button
          onClick={() => setShowAdd(true)}
          style={{ marginTop: 6, background: 'none', border: '1px dashed #ccc', borderRadius: 4, color: '#888', cursor: 'pointer', fontSize: 10, padding: '3px 8px' }}
        >
          + Add Deliverable
        </button>
      </div>
    );
  }

  return (
    <div>
      {paper.deliverables.map((d, i) => {
        const st = STATUS_STYLES[d.status] ?? STATUS_STYLES.pending;
        return (
          <div
            key={i}
            style={{
              display: 'flex',
              alignItems: 'center',
              padding: '6px 8px',
              borderBottom: '1px solid #f0ede6',
              gap: 8,
              background: i % 2 === 1 ? '#faf8f3' : undefined,
            }}
          >
            <span style={{ fontSize: 16 }}>{TYPE_ICONS[d.type] ?? TYPE_ICONS.other}</span>
            <div style={{ flex: 1 }}>
              <div style={{ fontSize: 12, fontWeight: 600 }}>{d.name}</div>
              <div style={{ fontSize: 10, color: '#999' }}>
                {d.version && `${d.version} \u2022 `}
                {d.lastUpdated ?? 'Not started'}
              </div>
            </div>
            <div
              onClick={() => cycleStatus(i)}
              style={{ fontSize: 10, color: st.color, fontWeight: 600, cursor: 'pointer' }}
              title="Click to cycle status"
            >
              {st.label}
            </div>
            <button
              onClick={() => removeDeliverable(i)}
              style={{ background: 'none', border: 'none', color: '#ccc', cursor: 'pointer', fontSize: 12 }}
              title="Remove"
            >&times;</button>
          </div>
        );
      })}

      {showAdd ? (
        <div style={{ marginTop: 6, padding: 8, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da' }}>
          <div style={{ display: 'flex', gap: 4, marginBottom: 4 }}>
            <input value={newName} onChange={(e) => setNewName(e.target.value)} placeholder="Name" style={{ ...inputStyle, flex: 1 }} />
            <select value={newType} onChange={(e) => setNewType(e.target.value as Deliverable['type'])} style={inputStyle}>
              <option value="feature">Feature</option>
              <option value="bugfix">Bugfix</option>
              <option value="docs">Documentation</option>
              <option value="design">Design</option>
              <option value="api">API</option>
              <option value="infrastructure">Infrastructure</option>
              <option value="other">Other</option>
            </select>
          </div>
          <div style={{ display: 'flex', gap: 4 }}>
            <input value={newVersion} onChange={(e) => setNewVersion(e.target.value)} placeholder="Version (e.g. v1.0)" style={{ ...inputStyle, flex: 1 }} />
            <select value={newStatus} onChange={(e) => setNewStatus(e.target.value as Deliverable['status'])} style={inputStyle}>
              <option value="pending">Pending</option>
              <option value="in-progress">In Progress</option>
              <option value="complete">Complete</option>
            </select>
            <button onClick={addDeliverable} style={{ ...inputStyle, background: '#2a7a8a', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 600 }}>Add</button>
            <button onClick={() => setShowAdd(false)} style={{ ...inputStyle, cursor: 'pointer' }}>Cancel</button>
          </div>
        </div>
      ) : (
        <button
          onClick={() => setShowAdd(true)}
          style={{ marginTop: 6, background: 'none', border: '1px dashed #ccc', borderRadius: 4, color: '#888', cursor: 'pointer', fontSize: 10, padding: '3px 8px' }}
        >
          + Add Deliverable
        </button>
      )}
    </div>
  );
}
