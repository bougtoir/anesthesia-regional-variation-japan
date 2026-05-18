import { useState } from 'react';
import type { Paper, Author } from '../../types/paper';
import { useDashboard } from '../../context/DashboardContext';

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

export function TitleAuthorsPane({ paper }: Props) {
  const { updatePaper } = useDashboard();
  const [editingTitle, setEditingTitle] = useState(false);
  const [titleDraft, setTitleDraft] = useState('');
  const [editingIdx, setEditingIdx] = useState<number | null>(null);
  const [showAddAuthor, setShowAddAuthor] = useState(false);
  const [newName, setNewName] = useState('');
  const [newRole, setNewRole] = useState<Author['role']>('co-author');
  const [newAffiliation, setNewAffiliation] = useState('');
  const [newEmail, setNewEmail] = useState('');

  const [editName, setEditName] = useState('');
  const [editRole, setEditRole] = useState<Author['role']>('co-author');
  const [editAffiliation, setEditAffiliation] = useState('');
  const [editEmail, setEditEmail] = useState('');

  const startEdit = (i: number) => {
    const a = paper.authors[i];
    setEditName(a.name);
    setEditRole(a.role);
    setEditAffiliation(a.affiliation ?? '');
    setEditEmail(a.email ?? '');
    setEditingIdx(i);
  };

  const saveEdit = () => {
    if (editingIdx === null || !editName.trim()) return;
    const updated = paper.authors.map((a, i) =>
      i === editingIdx
        ? { ...a, name: editName.trim(), role: editRole, affiliation: editAffiliation || undefined, email: editEmail || undefined }
        : a,
    );
    updatePaper({ ...paper, authors: updated });
    setEditingIdx(null);
  };

  const addAuthor = () => {
    if (!newName.trim()) return;
    updatePaper({
      ...paper,
      authors: [
        ...paper.authors,
        { name: newName.trim(), role: newRole, affiliation: newAffiliation || undefined, email: newEmail || undefined },
      ],
    });
    setNewName('');
    setNewRole('co-author');
    setNewAffiliation('');
    setNewEmail('');
    setShowAddAuthor(false);
  };

  const removeAuthor = (index: number) => {
    updatePaper({ ...paper, authors: paper.authors.filter((_, i) => i !== index) });
  };

  const startEditTitle = () => {
    setTitleDraft(paper.title);
    setEditingTitle(true);
  };

  const saveTitle = () => {
    if (titleDraft.trim()) {
      updatePaper({ ...paper, title: titleDraft.trim(), shortTitle: paper.shortTitle || titleDraft.trim().slice(0, 30) });
    }
    setEditingTitle(false);
  };

  return (
    <div>
      {editingTitle ? (
        <textarea
          value={titleDraft}
          onChange={(e) => setTitleDraft(e.target.value)}
          onKeyDown={(e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); saveTitle(); } if (e.key === 'Escape') setEditingTitle(false); }}
          onBlur={saveTitle}
          autoFocus
          style={{ ...inputStyle, width: '100%', fontSize: 16, fontWeight: 700, color: '#1a3a4a', lineHeight: 1.4, resize: 'vertical', minHeight: 48, boxSizing: 'border-box' }}
        />
      ) : (
        <div
          onClick={startEditTitle}
          style={{ fontSize: 16, fontWeight: 700, color: '#1a3a4a', marginBottom: 8, lineHeight: 1.4, cursor: 'pointer', borderBottom: '1px dashed transparent' }}
          onMouseEnter={(e) => (e.currentTarget.style.borderBottomColor = '#ccc')}
          onMouseLeave={(e) => (e.currentTarget.style.borderBottomColor = 'transparent')}
          title="Click to edit title"
        >
          {paper.title}
        </div>
      )}
      <div style={{ margin: '8px 0 6px', fontSize: 11, color: '#888' }}>Authors</div>
      <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
        {paper.authors.map((a, i) => (
          <div key={i}>
            {editingIdx === i ? (
              <div style={{ padding: 6, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da' }}>
                <div style={{ display: 'flex', gap: 4, marginBottom: 4 }}>
                  <input value={editName} onChange={(e) => setEditName(e.target.value)} placeholder="Name" style={{ ...inputStyle, flex: 1 }} />
                  <select value={editRole} onChange={(e) => setEditRole(e.target.value as Author['role'])} style={inputStyle}>
                    <option value="corresponding">Corresponding</option>
                    <option value="first">First Author</option>
                    <option value="co-author">Co-author</option>
                  </select>
                </div>
                <div style={{ display: 'flex', gap: 4 }}>
                  <input value={editAffiliation} onChange={(e) => setEditAffiliation(e.target.value)} placeholder="Affiliation" style={{ ...inputStyle, flex: 1 }} />
                  <input value={editEmail} onChange={(e) => setEditEmail(e.target.value)} placeholder="Email" style={{ ...inputStyle, flex: 1 }} />
                </div>
                <div style={{ display: 'flex', gap: 4, marginTop: 4, justifyContent: 'flex-end' }}>
                  <button onClick={saveEdit} style={{ ...inputStyle, background: '#2a7a8a', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 600 }}>Save</button>
                  <button onClick={() => setEditingIdx(null)} style={{ ...inputStyle, cursor: 'pointer' }}>Cancel</button>
                </div>
              </div>
            ) : (
              <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                <div
                  onClick={() => startEdit(i)}
                  style={{
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: 4,
                    padding: '3px 10px',
                    background: a.role === 'corresponding' ? '#e0f0e8' : '#f0ede6',
                    border: a.role === 'corresponding' ? '1px solid #b0d8c0' : 'none',
                    borderRadius: 12,
                    fontSize: 11,
                    color: '#555',
                    fontWeight: a.role === 'corresponding' ? 600 : 400,
                    cursor: 'pointer',
                  }}
                  title="Click to edit"
                >
                  {a.role === 'corresponding' && (
                    <span style={{ fontSize: 9, background: '#2a7a8a', color: 'white', padding: '1px 5px', borderRadius: 8, fontWeight: 600 }}>CA</span>
                  )}
                  {a.role === 'first' && (
                    <span style={{ fontSize: 9, background: '#6a8a4a', color: 'white', padding: '1px 5px', borderRadius: 8, fontWeight: 600 }}>1st</span>
                  )}
                  {a.name}
                </div>
                <button
                  onClick={() => removeAuthor(i)}
                  style={{ background: 'none', border: 'none', color: '#ccc', cursor: 'pointer', fontSize: 12, padding: 0 }}
                  title="Remove author"
                >&times;</button>
              </div>
            )}
          </div>
        ))}
      </div>
      {paper.authors.some((a) => a.affiliation) && (
        <div style={{ marginTop: 8, fontSize: 10, color: '#999', lineHeight: 1.6 }}>
          {paper.authors
            .filter((a) => a.affiliation)
            .map((a) => `${a.name}: ${a.affiliation}${a.email ? ` <${a.email}>` : ''}`)
            .join(' | ')}
        </div>
      )}

      {/* Add author */}
      {showAddAuthor ? (
        <div style={{ marginTop: 6, padding: 8, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da' }}>
          <div style={{ display: 'flex', gap: 4, marginBottom: 4 }}>
            <input value={newName} onChange={(e) => setNewName(e.target.value)} placeholder="Name" style={{ ...inputStyle, flex: 1 }} />
            <select value={newRole} onChange={(e) => setNewRole(e.target.value as Author['role'])} style={inputStyle}>
              <option value="corresponding">Corresponding</option>
              <option value="first">First Author</option>
              <option value="co-author">Co-author</option>
            </select>
          </div>
          <div style={{ display: 'flex', gap: 4 }}>
            <input value={newAffiliation} onChange={(e) => setNewAffiliation(e.target.value)} placeholder="Affiliation" style={{ ...inputStyle, flex: 1 }} />
            <input value={newEmail} onChange={(e) => setNewEmail(e.target.value)} placeholder="Email" style={{ ...inputStyle, flex: 1 }} />
          </div>
          <div style={{ display: 'flex', gap: 4, marginTop: 4, justifyContent: 'flex-end' }}>
            <button onClick={addAuthor} style={{ ...inputStyle, background: '#2a7a8a', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 600 }}>Add</button>
            <button onClick={() => setShowAddAuthor(false)} style={{ ...inputStyle, cursor: 'pointer' }}>Cancel</button>
          </div>
        </div>
      ) : (
        <button
          onClick={() => setShowAddAuthor(true)}
          style={{ marginTop: 6, background: 'none', border: '1px dashed #ccc', borderRadius: 4, color: '#888', cursor: 'pointer', fontSize: 10, padding: '3px 8px' }}
        >
          + Add Author
        </button>
      )}
    </div>
  );
}
