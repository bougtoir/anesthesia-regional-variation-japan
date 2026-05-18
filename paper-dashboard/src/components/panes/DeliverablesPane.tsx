import { useState } from 'react';
import type { Paper, Deliverable } from '../../types/paper';
import { useDashboard } from '../../context/DashboardContext';

const TYPE_ICONS: Record<string, string> = {
  manuscript: '\uD83D\uDCC4',
  figure: '\uD83D\uDCCA',
  table: '\uD83D\uDCC8',
  supplement: '\uD83D\uDCCE',
  'cover-letter': '\u2709',
  response: '\uD83D\uDCC4',
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
  const [newType, setNewType] = useState<Deliverable['type']>('manuscript');
  const [newVersion, setNewVersion] = useState('');
  const [newStatus, setNewStatus] = useState<Deliverable['status']>('pending');

  // Edit state
  const [editingIdx, setEditingIdx] = useState<number | null>(null);
  const [editName, setEditName] = useState('');
  const [editType, setEditType] = useState<Deliverable['type']>('manuscript');
  const [editVersion, setEditVersion] = useState('');

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
    setNewType('manuscript');
    setNewVersion('');
    setNewStatus('pending');
    setShowAdd(false);
  };

  const removeDeliverable = (index: number) => {
    updatePaper({ ...paper, deliverables: paper.deliverables.filter((_, i) => i !== index) });
  };

  const startEdit = (i: number) => {
    const d = paper.deliverables[i];
    setEditName(d.name);
    setEditType(d.type);
    setEditVersion(d.version ?? '');
    setEditingIdx(i);
  };

  const saveEdit = () => {
    if (editingIdx === null || !editName.trim()) return;
    const updated = paper.deliverables.map((d, i) =>
      i === editingIdx ? { ...d, name: editName.trim(), type: editType, version: editVersion || undefined, lastUpdated: new Date().toISOString().slice(0, 10) } : d,
    );
    updatePaper({ ...paper, deliverables: updated });
    setEditingIdx(null);
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
        if (editingIdx === i) {
          return (
            <div key={i} style={{ padding: 6, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da', marginBottom: 4 }}>
              <div style={{ display: 'flex', gap: 4, marginBottom: 4 }}>
                <input value={editName} onChange={(e) => setEditName(e.target.value)} placeholder="Name" style={{ ...inputStyle, flex: 1 }} />
                <select value={editType} onChange={(e) => setEditType(e.target.value as Deliverable['type'])} style={inputStyle}>
                  <option value="manuscript">Manuscript</option>
                  <option value="figure">Figure</option>
                  <option value="table">Table</option>
                  <option value="supplement">Supplement</option>
                  <option value="cover-letter">Cover Letter</option>
                  <option value="response">Response</option>
                  <option value="other">Other</option>
                </select>
              </div>
              <div style={{ display: 'flex', gap: 4 }}>
                <input value={editVersion} onChange={(e) => setEditVersion(e.target.value)} placeholder="Version (opt)" style={{ ...inputStyle, flex: 1 }} />
                <button onClick={saveEdit} style={{ ...inputStyle, background: '#2a7a8a', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 600 }}>Save</button>
                <button onClick={() => setEditingIdx(null)} style={{ ...inputStyle, cursor: 'pointer' }}>Cancel</button>
              </div>
            </div>
          );
        }
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
            <div
              style={{ flex: 1, cursor: 'pointer' }}
              onClick={() => startEdit(i)}
              title="Click to edit"
            >
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
              <option value="manuscript">Manuscript</option>
              <option value="figure">Figure</option>
              <option value="table">Table</option>
              <option value="supplement">Supplement</option>
              <option value="cover-letter">Cover Letter</option>
              <option value="response">Response</option>
              <option value="other">Other</option>
            </select>
          </div>
          <div style={{ display: 'flex', gap: 4 }}>
            <input value={newVersion} onChange={(e) => setNewVersion(e.target.value)} placeholder="Version (opt)" style={{ ...inputStyle, flex: 1 }} />
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
