import { useState } from 'react';
import type { Paper } from '../../types/paper';
import { useDashboard } from '../../context/DashboardContext';

const inputStyle: React.CSSProperties = {
  padding: '4px 8px',
  border: '1px solid #d0ccc0',
  borderRadius: 4,
  fontSize: 11,
  background: '#faf8f3',
  fontFamily: 'inherit',
  width: '100%',
  textAlign: 'center' as const,
  boxSizing: 'border-box' as const,
};

interface Props {
  paper: Paper;
}

type StatField = 'wordCount' | 'wordLimit' | 'figureCount' | 'figureLimit' | 'tableCount' | 'tableLimit' | 'referenceCount' | 'referenceLimit';

interface StatCardProps {
  value: number;
  valueField: StatField;
  label: string;
  limit?: number;
  limitField: StatField;
  editingField: StatField | null;
  draft: string;
  setDraft: (v: string) => void;
  saveEdit: () => void;
  setEditingField: (f: StatField | null) => void;
  startEdit: (field: StatField, value: number | undefined) => void;
}

function StatCard({ value, valueField, label, limit, limitField, editingField, draft, setDraft, saveEdit, setEditingField, startEdit }: StatCardProps) {
  const pct = limit ? (value / limit) * 100 : 0;
  const fillColor = pct > 95 ? '#d04040' : pct > 75 ? '#e0a040' : '#40a060';

  return (
    <div
      style={{
        padding: 10,
        background: '#faf8f3',
        borderRadius: 4,
        border: '1px solid #e8e4da',
        textAlign: 'center',
      }}
    >
      {editingField === valueField ? (
        <input
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          onKeyDown={(e) => { if (e.key === 'Enter') saveEdit(); if (e.key === 'Escape') setEditingField(null); }}
          onBlur={saveEdit}
          autoFocus
          type="number"
          style={{ ...inputStyle, fontSize: 22, fontWeight: 700, color: '#1a3a4a' }}
        />
      ) : (
        <div
          onClick={() => startEdit(valueField, value)}
          style={{ fontSize: 22, fontWeight: 700, color: '#1a3a4a', cursor: 'pointer' }}
          title="Click to edit count"
        >
          {value.toLocaleString()}
        </div>
      )}
      <div style={{ fontSize: 10, color: '#888', marginTop: 2 }}>{label}</div>
      {editingField === limitField ? (
        <input
          value={draft}
          onChange={(e) => setDraft(e.target.value)}
          onKeyDown={(e) => { if (e.key === 'Enter') saveEdit(); if (e.key === 'Escape') setEditingField(null); }}
          onBlur={saveEdit}
          autoFocus
          type="number"
          placeholder="Limit"
          style={{ ...inputStyle, fontSize: 10, marginTop: 2 }}
        />
      ) : (
        <div
          onClick={() => startEdit(limitField, limit)}
          style={{ fontSize: 10, color: '#bbb', cursor: 'pointer', marginTop: 2 }}
          title="Click to edit limit"
        >
          {limit ? `Limit: ${limit.toLocaleString()}` : '+ Set limit'}
        </div>
      )}
      {limit && editingField !== limitField && (
        <div style={{ height: 4, background: '#e8e4da', borderRadius: 2, marginTop: 6, overflow: 'hidden' }}>
          <div
            style={{
              height: '100%',
              borderRadius: 2,
              background: fillColor,
              width: `${Math.min(pct, 100)}%`,
            }}
          />
        </div>
      )}
    </div>
  );
}

export function StatisticsPane({ paper }: Props) {
  const { updatePaper } = useDashboard();
  const s = paper.statistics;
  const [editingField, setEditingField] = useState<StatField | null>(null);
  const [draft, setDraft] = useState('');

  const startEdit = (field: StatField, value: number | undefined) => {
    setEditingField(field);
    setDraft(String(value ?? ''));
  };

  const saveEdit = () => {
    if (!editingField) return;
    const val = draft ? parseInt(draft) : undefined;
    updatePaper({
      ...paper,
      statistics: { ...paper.statistics, [editingField]: editingField.includes('Limit') ? val : (val ?? 0) },
    });
    setEditingField(null);
  };

  const cardProps = { editingField, draft, setDraft, saveEdit, setEditingField, startEdit };

  return (
    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
      <StatCard value={s.wordCount} valueField="wordCount" label="Word Count" limit={s.wordLimit} limitField="wordLimit" {...cardProps} />
      <StatCard value={s.figureCount} valueField="figureCount" label="Figures" limit={s.figureLimit} limitField="figureLimit" {...cardProps} />
      <StatCard value={s.tableCount} valueField="tableCount" label="Tables" limit={s.tableLimit} limitField="tableLimit" {...cardProps} />
      <StatCard value={s.referenceCount} valueField="referenceCount" label="References" limit={s.referenceLimit} limitField="referenceLimit" {...cardProps} />
    </div>
  );
}
