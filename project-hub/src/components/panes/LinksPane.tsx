import { useState } from 'react';
import type { Paper, Link } from '../../types/paper';
import { useDashboard } from '../../context/DashboardContext';

const CATEGORY_LABELS: Record<string, string> = {
  board: 'Project Board',
  devin: 'Devin Sessions',
  repository: 'Repository',
  other: 'Other',
};

const CATEGORY_ICONS: Record<string, string> = {
  board: '\uD83D\uDCCB',
  devin: '\u2699',
  repository: '\uD83D\uDCC1',
  other: '\uD83D\uDD17',
};

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

export function LinksPane({ paper }: Props) {
  const { updatePaper } = useDashboard();
  const [showAdd, setShowAdd] = useState(false);
  const [newLabel, setNewLabel] = useState('');
  const [newUrl, setNewUrl] = useState('');
  const [newCategory, setNewCategory] = useState<Link['category']>('other');

  const grouped = paper.links.reduce(
    (acc, link) => {
      const cat = link.category;
      if (!acc[cat]) acc[cat] = [];
      acc[cat].push(link);
      return acc;
    },
    {} as Record<string, typeof paper.links>,
  );

  const addLink = () => {
    if (!newLabel.trim() || !newUrl.trim()) return;
    updatePaper({
      ...paper,
      links: [...paper.links, { label: newLabel.trim(), url: newUrl.trim(), category: newCategory }],
    });
    setNewLabel('');
    setNewUrl('');
    setNewCategory('other');
    setShowAdd(false);
  };

  const removeLink = (index: number) => {
    updatePaper({ ...paper, links: paper.links.filter((_, i) => i !== index) });
  };

  return (
    <div>
      {paper.links.length === 0 && !showAdd && (
        <div style={{ color: '#aaa', fontSize: 12 }}>No links added</div>
      )}

      {Object.entries(grouped).map(([cat, links]) => (
        <div key={cat} style={{ marginBottom: 10 }}>
          <div
            style={{
              fontSize: 10,
              fontWeight: 700,
              textTransform: 'uppercase',
              color: '#888',
              marginBottom: 4,
              letterSpacing: 0.5,
            }}
          >
            {CATEGORY_LABELS[cat] ?? cat}
          </div>
          {links.map((link, i) => {
            const globalIdx = paper.links.indexOf(link);
            return (
              <div key={i} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: '4px 0', fontSize: 12 }}>
                <span style={{ fontSize: 14 }}>{CATEGORY_ICONS[cat] ?? '\uD83D\uDD17'}</span>
                <a
                  href={link.url}
                  target="_blank"
                  rel="noopener noreferrer"
                  style={{ color: '#2a7a8a', textDecoration: 'none', flex: 1 }}
                >
                  {link.label}
                </a>
                <button
                  onClick={() => removeLink(globalIdx)}
                  style={{ background: 'none', border: 'none', color: '#ccc', cursor: 'pointer', fontSize: 12 }}
                  title="Remove link"
                >&times;</button>
              </div>
            );
          })}
        </div>
      ))}

      {showAdd ? (
        <div style={{ marginTop: 6, padding: 8, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da' }}>
          <div style={{ display: 'flex', gap: 4, marginBottom: 4 }}>
            <input value={newLabel} onChange={(e) => setNewLabel(e.target.value)} placeholder="Label" style={{ ...inputStyle, flex: 1 }} />
            <select value={newCategory} onChange={(e) => setNewCategory(e.target.value as Link['category'])} style={inputStyle}>
              <option value="board">Project Board</option>
              <option value="devin">Devin Session</option>
              <option value="repository">Repository</option>
              <option value="other">Other</option>
            </select>
          </div>
          <div style={{ display: 'flex', gap: 4 }}>
            <input value={newUrl} onChange={(e) => setNewUrl(e.target.value)} placeholder="URL" style={{ ...inputStyle, flex: 1 }} />
            <button onClick={addLink} style={{ ...inputStyle, background: '#2a7a8a', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 600 }}>Add</button>
            <button onClick={() => setShowAdd(false)} style={{ ...inputStyle, cursor: 'pointer' }}>Cancel</button>
          </div>
        </div>
      ) : (
        <button
          onClick={() => setShowAdd(true)}
          style={{ marginTop: 6, background: 'none', border: '1px dashed #ccc', borderRadius: 4, color: '#888', cursor: 'pointer', fontSize: 10, padding: '3px 8px' }}
        >
          + Add Link
        </button>
      )}
    </div>
  );
}
