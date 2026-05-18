import { useState } from 'react';
import type { Paper, TimelineEvent, PaperStatus } from '../../types/paper';
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

const clickableStyle: React.CSSProperties = {
  cursor: 'pointer',
  borderBottom: '1px dashed transparent',
};

const ALL_STATUSES: PaperStatus[] = ['drafting', 'internal-review', 'submitted', 'under-review', 'revision', 'accepted', 'published', 'rejected'];

interface Props {
  paper: Paper;
}

export function ProgressPane({ paper }: Props) {
  const { updatePaper } = useDashboard();
  const p = paper.progress;

  // Inline editing state for header fields
  const [editingField, setEditingField] = useState<string | null>(null);
  const [draft, setDraft] = useState('');

  // Add event state
  const [showAddEvent, setShowAddEvent] = useState(false);
  const [evLabel, setEvLabel] = useState('');
  const [evStart, setEvStart] = useState('');
  const [evEnd, setEvEnd] = useState('');
  const [evType, setEvType] = useState<TimelineEvent['type']>('milestone');

  // Edit event state
  const [editingEvIdx, setEditingEvIdx] = useState<number | null>(null);
  const [editEvLabel, setEditEvLabel] = useState('');
  const [editEvStart, setEditEvStart] = useState('');
  const [editEvEnd, setEditEvEnd] = useState('');
  const [editEvType, setEditEvType] = useState<TimelineEvent['type']>('milestone');

  const updateProgress = (partial: Partial<Paper['progress']>) => {
    updatePaper({ ...paper, progress: { ...paper.progress, ...partial } });
  };

  const startEditField = (field: string, value: string) => {
    setEditingField(field);
    setDraft(value);
  };

  const saveField = (field: string) => {
    if (field === 'journal') updateProgress({ journal: draft });
    if (field === 'impactFactor') updateProgress({ impactFactor: draft ? parseFloat(draft) : undefined });
    if (field === 'submissionDate') updateProgress({ submissionDate: draft || undefined });
    if (field === 'revisionDueDate') updateProgress({ revisionDueDate: draft || undefined });
    if (field === 'acceptanceDate') updateProgress({ acceptanceDate: draft || undefined });
    setEditingField(null);
  };

  const addEvent = () => {
    if (!evLabel.trim() || !evStart) return;
    updateProgress({
      timeline: [
        ...paper.progress.timeline,
        { label: evLabel.trim(), startDate: evStart, endDate: evEnd || undefined, type: evType },
      ],
    });
    setEvLabel('');
    setEvStart('');
    setEvEnd('');
    setShowAddEvent(false);
  };

  const removeEvent = (index: number) => {
    updateProgress({ timeline: paper.progress.timeline.filter((_, i) => i !== index) });
  };

  const startEditEvent = (i: number) => {
    const ev = p.timeline[i];
    setEditEvLabel(ev.label);
    setEditEvStart(ev.startDate);
    setEditEvEnd(ev.endDate ?? '');
    setEditEvType(ev.type);
    setEditingEvIdx(i);
  };

  const saveEditEvent = () => {
    if (editingEvIdx === null || !editEvLabel.trim() || !editEvStart) return;
    const updated = p.timeline.map((ev, i) =>
      i === editingEvIdx
        ? { ...ev, label: editEvLabel.trim(), startDate: editEvStart, endDate: editEvEnd || undefined, type: editEvType }
        : ev,
    );
    updateProgress({ timeline: updated });
    setEditingEvIdx(null);
  };

  const addHistoryRecord = () => {
    const journal = prompt('Previous journal name:');
    if (!journal) return;
    const subDate = prompt('Submission date (YYYY-MM-DD):');
    if (!subDate) return;
    const decDate = prompt('Decision date (YYYY-MM-DD):');
    const decision = prompt('Decision (rejected / withdrawn / desk-reject):') as 'rejected' | 'withdrawn' | 'desk-reject';
    if (!decision) return;
    const comment = prompt('Editor comment (optional):') ?? '';
    const ifStr = prompt('Impact factor (optional):');
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

  const renderInlineField = (field: string, value: string, placeholder: string, extraStyle?: React.CSSProperties) => {
    if (editingField === field) {
      return (
        <input
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          onKeyDown={(e) => { if (e.key === 'Enter') saveField(field); if (e.key === 'Escape') setEditingField(null); }}
          onBlur={() => saveField(field)}
          autoFocus
          type={field.includes('Date') ? 'date' : 'text'}
          style={{ ...inputStyle, ...extraStyle }}
        />
      );
    }
    return (
      <span
        onClick={() => startEditField(field, value)}
        style={{ ...clickableStyle, ...extraStyle }}
        onMouseEnter={(e) => (e.currentTarget.style.borderBottomColor = '#ccc')}
        onMouseLeave={(e) => (e.currentTarget.style.borderBottomColor = 'transparent')}
        title={`Click to edit ${placeholder}`}
      >
        {value || <span style={{ color: '#ccc', fontStyle: 'italic' }}>{placeholder}</span>}
      </span>
    );
  };

  return (
    <div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 10, flexWrap: 'wrap' }}>
        {editingField === 'status' ? (
          <select
            value={draft}
            onChange={(e) => { updateProgress({ status: e.target.value as PaperStatus }); setEditingField(null); }}
            onBlur={() => setEditingField(null)}
            autoFocus
            style={inputStyle}
          >
            {ALL_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
          </select>
        ) : (
          <span onClick={() => startEditField('status', p.status)} style={{ cursor: 'pointer' }} title="Click to change status">
            <StatusBadge status={p.status} />
          </span>
        )}
        <span style={{ fontSize: 11, color: '#888', display: 'flex', alignItems: 'center', gap: 4 }}>
          {renderInlineField('journal', p.journal, 'Set journal')}
          {p.impactFactor !== undefined ? (
            <span> &bull; IF: {renderInlineField('impactFactor', String(p.impactFactor ?? ''), 'IF')}</span>
          ) : (
            <span
              onClick={() => startEditField('impactFactor', '')}
              style={{ color: '#ccc', cursor: 'pointer', fontSize: 10 }}
              title="Click to add Impact Factor"
            > + IF</span>
          )}
        </span>
      </div>
      <GanttChart events={p.timeline} />
      <div style={{ marginTop: 10, display: 'flex', gap: 16, fontSize: 10, color: '#aaa', flexWrap: 'wrap' }}>
        <span>Submitted: {renderInlineField('submissionDate', p.submissionDate ?? '', 'Set date')}</span>
        <span style={{ color: '#d08050', fontWeight: 600 }}>
          Revision due: {renderInlineField('revisionDueDate', p.revisionDueDate ?? '', 'Set date')}
        </span>
        <span style={{ color: '#40a060' }}>
          Accepted: {renderInlineField('acceptanceDate', p.acceptanceDate ?? '', 'Set date')}
        </span>
      </div>

      {/* Timeline events list with edit/delete */}
      <div style={{ marginTop: 8 }}>
        {p.timeline.map((ev, i) => (
          <div key={i}>
            {editingEvIdx === i ? (
              <div style={{ padding: 6, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da', marginBottom: 4 }}>
                <div style={{ display: 'flex', gap: 4, marginBottom: 4 }}>
                  <input value={editEvLabel} onChange={(e) => setEditEvLabel(e.target.value)} placeholder="Label" style={{ ...inputStyle, flex: 1 }} />
                  <select value={editEvType} onChange={(e) => setEditEvType(e.target.value as TimelineEvent['type'])} style={inputStyle}>
                    <option value="drafting">Drafting</option>
                    <option value="submission">Submission</option>
                    <option value="review">Review</option>
                    <option value="revision">Revision</option>
                    <option value="decision">Decision</option>
                    <option value="milestone">Milestone</option>
                  </select>
                </div>
                <div style={{ display: 'flex', gap: 4 }}>
                  <input type="date" value={editEvStart} onChange={(e) => setEditEvStart(e.target.value)} style={{ ...inputStyle, flex: 1 }} />
                  <input type="date" value={editEvEnd} onChange={(e) => setEditEvEnd(e.target.value)} placeholder="End (opt)" style={{ ...inputStyle, flex: 1 }} />
                  <button onClick={saveEditEvent} style={{ ...inputStyle, background: '#2a7a8a', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 600 }}>Save</button>
                  <button onClick={() => setEditingEvIdx(null)} style={{ ...inputStyle, cursor: 'pointer' }}>Cancel</button>
                </div>
              </div>
            ) : (
              <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 10, color: '#888', padding: '2px 0' }}>
                <span
                  onClick={() => startEditEvent(i)}
                  style={{ flex: 1, cursor: 'pointer', borderBottom: '1px dashed transparent' }}
                  onMouseEnter={(e) => (e.currentTarget.style.borderBottomColor = '#ccc')}
                  onMouseLeave={(e) => (e.currentTarget.style.borderBottomColor = 'transparent')}
                  title="Click to edit"
                >
                  {ev.label} ({ev.startDate}{ev.endDate ? ` ~ ${ev.endDate}` : ''})
                </span>
                <button
                  onClick={() => removeEvent(i)}
                  style={{ background: 'none', border: 'none', color: '#ccc', cursor: 'pointer', fontSize: 10 }}
                  title="Remove event"
                >
                  &times;
                </button>
              </div>
            )}
          </div>
        ))}
      </div>

      {/* Add timeline event */}
      {showAddEvent ? (
        <div style={{ marginTop: 6, padding: 8, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da' }}>
          <div style={{ display: 'flex', gap: 4, marginBottom: 4 }}>
            <input value={evLabel} onChange={(e) => setEvLabel(e.target.value)} placeholder="Event label" style={{ ...inputStyle, flex: 1 }} />
            <select value={evType} onChange={(e) => setEvType(e.target.value as TimelineEvent['type'])} style={inputStyle}>
              <option value="drafting">Drafting</option>
              <option value="submission">Submission</option>
              <option value="review">Review</option>
              <option value="revision">Revision</option>
              <option value="decision">Decision</option>
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
            Previous Submissions ({paper.submissionHistory.length})
          </div>
          {paper.submissionHistory.map((h, i) => (
            <div key={i} style={{ fontSize: 11, padding: '4px 8px', background: i % 2 === 0 ? '#faf8f3' : undefined, borderRadius: 3, marginBottom: 2 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                <span style={{ fontWeight: 600 }}>{h.journal}</span>
                {h.impactFactor && <span style={{ color: '#aaa', fontSize: 10 }}>IF: {h.impactFactor}</span>}
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
                {h.submissionDate}{h.decisionDate ? ` \u2192 ${h.decisionDate}` : ''}
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
          + Add Previous Submission
        </button>
      </div>
    </div>
  );
}
