import { useState } from 'react';
import type { Paper, TimelineEvent } from '../../types/paper';
import { useDashboard } from '../../context/DashboardContext';
import { StatusBadge } from '../common/StatusBadge';
import { GanttChart } from '../common/GanttChart';

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

export function ProgressPane({ paper }: Props) {
  const { updatePaper } = useDashboard();
  const p = paper.progress;
  const [showAddEvent, setShowAddEvent] = useState(false);
  const [evLabel, setEvLabel] = useState('');
  const [evStart, setEvStart] = useState('');
  const [evEnd, setEvEnd] = useState('');
  const [evType, setEvType] = useState<TimelineEvent['type']>('milestone');

  const addEvent = () => {
    if (!evLabel.trim() || !evStart) return;
    updatePaper({
      ...paper,
      progress: {
        ...paper.progress,
        timeline: [
          ...paper.progress.timeline,
          { label: evLabel.trim(), startDate: evStart, endDate: evEnd || undefined, type: evType },
        ],
      },
    });
    setEvLabel('');
    setEvStart('');
    setEvEnd('');
    setShowAddEvent(false);
  };

  const removeEvent = (index: number) => {
    updatePaper({
      ...paper,
      progress: {
        ...paper.progress,
        timeline: paper.progress.timeline.filter((_, i) => i !== index),
      },
    });
  };

  const addHistoryRecord = () => {
    const journal = prompt('Previous iteration name:');
    if (!journal) return;
    const subDate = prompt('Submission date (YYYY-MM-DD):');
    if (!subDate) return;
    const decDate = prompt('Decision date (YYYY-MM-DD):');
    const decision = prompt('Outcome (rejected / withdrawn / desk-reject):') as 'rejected' | 'withdrawn' | 'desk-reject';
    if (!decision) return;
    const comment = prompt('Stakeholder comment (optional):') ?? '';
    const ifStr = prompt('Priority (optional, 1-10):');
    updatePaper({
      ...paper,
      submissionHistory: [
        ...paper.submissionHistory,
        {
          journal,
          impactFactor: ifStr ? parseFloat(ifStr) : undefined,
          submissionDate: subDate,
          decisionDate: decDate || undefined,
          decision,
          editorComment: comment || undefined,
        },
      ],
    });
  };

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 10 }}>
        <StatusBadge status={p.status} />
        <span style={{ fontSize: 11, color: '#888' }}>
          {p.journal}
          {p.impactFactor ? ` \u2022 Priority: ${p.impactFactor}/10` : ''}
        </span>
      </div>
      <GanttChart events={p.timeline} />
      <div style={{ marginTop: 10, display: 'flex', gap: 16, fontSize: 10, color: '#aaa' }}>
        {p.submissionDate && <span>Started: {p.submissionDate}</span>}
        {p.revisionDueDate && (
          <span style={{ color: '#d08050', fontWeight: 600 }}>Deadline: {p.revisionDueDate}</span>
        )}
        {p.acceptanceDate && <span style={{ color: '#40a060' }}>Approved: {p.acceptanceDate}</span>}
      </div>

      {/* Timeline events list with delete */}
      <div style={{ marginTop: 8 }}>
        {p.timeline.map((ev, i) => (
          <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 10, color: '#888', padding: '2px 0' }}>
            <span style={{ flex: 1 }}>{ev.label} ({ev.startDate}{ev.endDate ? ` ~ ${ev.endDate}` : ''})</span>
            <button
              onClick={() => removeEvent(i)}
              style={{ background: 'none', border: 'none', color: '#ccc', cursor: 'pointer', fontSize: 10 }}
              title="Remove event"
            >
              &times;
            </button>
          </div>
        ))}
      </div>

      {/* Add timeline event */}
      {showAddEvent ? (
        <div style={{ marginTop: 6, padding: 8, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da' }}>
          <div style={{ display: 'flex', gap: 4, marginBottom: 4 }}>
            <input value={evLabel} onChange={(e) => setEvLabel(e.target.value)} placeholder="Event label" style={{ ...inputStyle, flex: 1 }} />
            <select value={evType} onChange={(e) => setEvType(e.target.value as TimelineEvent['type'])} style={inputStyle}>
              <option value="planning">Planning</option>
              <option value="development">Development</option>
              <option value="review">Review</option>
              <option value="testing">Testing</option>
              <option value="deployment">Deployment</option>
              <option value="milestone">Milestone</option>
            </select>
          </div>
          <div style={{ display: 'flex', gap: 4 }}>
            <input type="date" value={evStart} onChange={(e) => setEvStart(e.target.value)} style={{ ...inputStyle, flex: 1 }} />
            <input type="date" value={evEnd} onChange={(e) => setEvEnd(e.target.value)} placeholder="End (opt)" style={{ ...inputStyle, flex: 1 }} />
            <button onClick={addEvent} style={{ ...inputStyle, background: '#2a7a8a', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 600 }}>Add</button>
            <button onClick={() => setShowAddEvent(false)} style={{ ...inputStyle, cursor: 'pointer' }}>Cancel</button>
          </div>
        </div>
      ) : (
        <div style={{ marginTop: 6, display: 'flex', gap: 6 }}>
          <button
            onClick={() => setShowAddEvent(true)}
            style={{ background: 'none', border: '1px dashed #ccc', borderRadius: 4, color: '#888', cursor: 'pointer', fontSize: 10, padding: '3px 8px' }}
          >
            + Add Timeline Event
          </button>
        </div>
      )}

      {/* Submission History */}
      {paper.submissionHistory.length > 0 && (
        <div style={{ marginTop: 10, borderTop: '1px solid #e8e4da', paddingTop: 8 }}>
          <div style={{ fontSize: 10, fontWeight: 700, textTransform: 'uppercase', color: '#888', letterSpacing: 0.5, marginBottom: 4 }}>
            Previous Iterations ({paper.submissionHistory.length})
          </div>
          {paper.submissionHistory.map((h, i) => (
            <div key={i} style={{ fontSize: 11, padding: '4px 8px', background: i % 2 === 0 ? '#faf8f3' : undefined, borderRadius: 3, marginBottom: 2 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                <span style={{ fontWeight: 600 }}>{h.journal}</span>
                {h.impactFactor && <span style={{ color: '#aaa', fontSize: 10 }}>P: {h.impactFactor}</span>}
                <span style={{
                  fontSize: 9, padding: '1px 6px', borderRadius: 8, fontWeight: 600,
                  background: h.decision === 'rejected' ? '#fde0e0' : '#fdf0d0',
                  color: h.decision === 'rejected' ? '#c04040' : '#a07020',
                }}>
                  {h.decision}
                </span>
                <button
                  onClick={() => updatePaper({ ...paper, submissionHistory: paper.submissionHistory.filter((_, j) => j !== i) })}
                  style={{ marginLeft: 'auto', background: 'none', border: 'none', color: '#ccc', cursor: 'pointer', fontSize: 10 }}
                >&times;</button>
              </div>
              <div style={{ fontSize: 10, color: '#999' }}>
                {h.submissionDate}{h.decisionDate ? ` → ${h.decisionDate}` : ''}
              </div>
              {h.editorComment && <div style={{ fontSize: 10, color: '#888', fontStyle: 'italic', marginTop: 2 }}>{h.editorComment}</div>}
            </div>
          ))}
        </div>
      )}
      <div style={{ marginTop: 4 }}>
        <button
          onClick={addHistoryRecord}
          style={{ background: 'none', border: '1px dashed #ccc', borderRadius: 4, color: '#888', cursor: 'pointer', fontSize: 10, padding: '3px 8px' }}
        >
          + Add Previous Iteration
        </button>
      </div>
    </div>
  );
}
