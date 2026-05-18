import { useState } from 'react';
import type { Paper, Deadline } from '../../types/paper';
import { useDashboard } from '../../context/DashboardContext';

const inputStyle: React.CSSProperties = {
  padding: '4px 8px',
  border: '1px solid #d0ccc0',
  borderRadius: 4,
  fontSize: 11,
  background: '#faf8f3',
  fontFamily: 'inherit',
};

function daysUntil(dateStr: string): number {
  const d = new Date(dateStr);
  const now = new Date();
  return Math.ceil((d.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
}

interface Props {
  paper: Paper;
}

export function DeadlinesPane({ paper }: Props) {
  const { updatePaper } = useDashboard();
  const [showAdd, setShowAdd] = useState(false);
  const [newLabel, setNewLabel] = useState('');
  const [newDate, setNewDate] = useState('');
  const [newType, setNewType] = useState<Deadline['type']>('other');

  const allDeadlines = [...paper.deadlines];
  if (paper.progress.revisionDueDate && !allDeadlines.some((d) => d.date === paper.progress.revisionDueDate)) {
    allDeadlines.unshift({
      label: 'Project deadline',
      date: paper.progress.revisionDueDate,
      type: 'milestone',
    });
  }

  const addDeadline = () => {
    if (!newLabel.trim() || !newDate) return;
    updatePaper({
      ...paper,
      deadlines: [...paper.deadlines, { label: newLabel.trim(), date: newDate, type: newType }],
    });
    setNewLabel('');
    setNewDate('');
    setNewType('other');
    setShowAdd(false);
  };

  const removeDeadline = (index: number) => {
    updatePaper({ ...paper, deadlines: paper.deadlines.filter((_, i) => i !== index) });
  };

  return (
    <div>
      {allDeadlines.length === 0 && !showAdd && (
        <div style={{ color: '#aaa', fontSize: 12 }}>No upcoming deadlines</div>
      )}

      {allDeadlines
        .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime())
        .map((dl, i) => {
          const days = daysUntil(dl.date);
          const countdownColor = days <= 7 ? '#d04040' : days <= 21 ? '#d09030' : '#40a060';
          const isFromProgress = dl.label === 'Project deadline' && dl.date === paper.progress.revisionDueDate && !paper.deadlines.some((d) => d.date === dl.date);
          const realIdx = paper.deadlines.indexOf(dl);

          return (
            <div
              key={i}
              style={{
                display: 'flex',
                alignItems: 'center',
                padding: '7px 8px',
                borderBottom: '1px solid #f0ede6',
                gap: 10,
                background: i % 2 === 1 ? '#faf8f3' : undefined,
              }}
            >
              <div style={{ minWidth: 50, textAlign: 'center' }}>
                <div style={{ fontSize: 18, fontWeight: 700, color: countdownColor }}>
                  {days < 0 ? 'Past' : days}
                </div>
                <div style={{ fontSize: 9, color: '#888' }}>{days < 0 ? 'due' : 'days'}</div>
              </div>
              <div style={{ flex: 1 }}>
                <div style={{ fontSize: 12, fontWeight: 600 }}>{dl.label}</div>
                <div style={{ fontSize: 10, color: '#999' }}>
                  {new Date(dl.date).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' })}
                </div>
              </div>
              {!isFromProgress && realIdx >= 0 && (
                <button
                  onClick={() => removeDeadline(realIdx)}
                  style={{ background: 'none', border: 'none', color: '#ccc', cursor: 'pointer', fontSize: 12 }}
                  title="Remove"
                >&times;</button>
              )}
            </div>
          );
        })}

      {showAdd ? (
        <div style={{ marginTop: 6, padding: 8, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da' }}>
          <div style={{ display: 'flex', gap: 4, marginBottom: 4 }}>
            <input value={newLabel} onChange={(e) => setNewLabel(e.target.value)} placeholder="Deadline label" style={{ ...inputStyle, flex: 1 }} />
            <select value={newType} onChange={(e) => setNewType(e.target.value as Deadline['type'])} style={inputStyle}>
              <option value="sprint">Sprint</option>
              <option value="milestone">Milestone</option>
              <option value="release">Release</option>
              <option value="demo">Demo</option>
              <option value="other">Other</option>
            </select>
          </div>
          <div style={{ display: 'flex', gap: 4 }}>
            <input type="date" value={newDate} onChange={(e) => setNewDate(e.target.value)} style={{ ...inputStyle, flex: 1 }} />
            <button onClick={addDeadline} style={{ ...inputStyle, background: '#2a7a8a', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 600 }}>Add</button>
            <button onClick={() => setShowAdd(false)} style={{ ...inputStyle, cursor: 'pointer' }}>Cancel</button>
          </div>
        </div>
      ) : (
        <button
          onClick={() => setShowAdd(true)}
          style={{ marginTop: 6, background: 'none', border: '1px dashed #ccc', borderRadius: 4, color: '#888', cursor: 'pointer', fontSize: 10, padding: '3px 8px' }}
        >
          + Add Deadline
        </button>
      )}
    </div>
  );
}
