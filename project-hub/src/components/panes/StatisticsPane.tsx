import type { Paper } from '../../types/paper';

interface Props {
  paper: Paper;
}

function StatCard({ value, label, limit }: { value: number; label: string; limit?: number }) {
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
      <div style={{ fontSize: 22, fontWeight: 700, color: '#1a3a4a' }}>{value.toLocaleString()}</div>
      <div style={{ fontSize: 10, color: '#888', marginTop: 2 }}>{label}</div>
      {limit && <div style={{ fontSize: 10, color: '#bbb' }}>Limit: {limit.toLocaleString()}</div>}
      {limit && (
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
  const s = paper.statistics;
  return (
    <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
      <StatCard value={s.wordCount} label="Tasks" limit={s.wordLimit} />
      <StatCard value={s.figureCount} label="Issues" limit={s.figureLimit} />
      <StatCard value={s.tableCount} label="Milestones" limit={s.tableLimit} />
      <StatCard value={s.referenceCount} label="Sprints" limit={s.referenceLimit} />
    </div>
  );
}
