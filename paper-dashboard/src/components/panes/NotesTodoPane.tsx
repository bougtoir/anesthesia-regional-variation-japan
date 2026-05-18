import { useState } from 'react';
import type { Paper } from '../../types/paper';
import { useDashboard } from '../../context/DashboardContext';

const PRIORITY_STYLES: Record<string, { bg: string; color: string; label: string }> = {
  high: { bg: '#fde0e0', color: '#c04040', label: 'HIGH' },
  medium: { bg: '#fdf0d0', color: '#a07020', label: 'MED' },
  low: { bg: '#e0f0e8', color: '#408060', label: 'LOW' },
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

export function NotesTodoPane({ paper }: Props) {
  const { updatePaper } = useDashboard();
  const [newNote, setNewNote] = useState('');
  const [newTodo, setNewTodo] = useState('');
  const [newPriority, setNewPriority] = useState<'high' | 'medium' | 'low'>('medium');

  // Edit note state
  const [editingNoteId, setEditingNoteId] = useState<string | null>(null);
  const [editNoteDraft, setEditNoteDraft] = useState('');

  // Edit todo state
  const [editingTodoId, setEditingTodoId] = useState<string | null>(null);
  const [editTodoDraft, setEditTodoDraft] = useState('');

  const toggleTodo = (id: string) => {
    updatePaper({
      ...paper,
      todos: paper.todos.map((t) => (t.id === id ? { ...t, done: !t.done } : t)),
    });
  };

  const addNote = () => {
    if (!newNote.trim()) return;
    updatePaper({
      ...paper,
      notes: [
        ...paper.notes,
        { id: `n-${Date.now()}`, content: newNote.trim(), createdAt: new Date().toISOString().slice(0, 10) },
      ],
    });
    setNewNote('');
  };

  const removeNote = (id: string) => {
    updatePaper({ ...paper, notes: paper.notes.filter((n) => n.id !== id) });
  };

  const startEditNote = (id: string, content: string) => {
    setEditingNoteId(id);
    setEditNoteDraft(content);
  };

  const saveNote = () => {
    if (!editingNoteId || !editNoteDraft.trim()) return;
    updatePaper({
      ...paper,
      notes: paper.notes.map((n) => (n.id === editingNoteId ? { ...n, content: editNoteDraft.trim() } : n)),
    });
    setEditingNoteId(null);
  };

  const addTodo = () => {
    if (!newTodo.trim()) return;
    updatePaper({
      ...paper,
      todos: [...paper.todos, { id: `t-${Date.now()}`, text: newTodo.trim(), done: false, priority: newPriority }],
    });
    setNewTodo('');
  };

  const removeTodo = (id: string) => {
    updatePaper({ ...paper, todos: paper.todos.filter((t) => t.id !== id) });
  };

  const startEditTodo = (id: string, text: string) => {
    setEditingTodoId(id);
    setEditTodoDraft(text);
  };

  const saveTodo = () => {
    if (!editingTodoId || !editTodoDraft.trim()) return;
    updatePaper({
      ...paper,
      todos: paper.todos.map((t) => (t.id === editingTodoId ? { ...t, text: editTodoDraft.trim() } : t)),
    });
    setEditingTodoId(null);
  };

  const cyclePriority = (id: string) => {
    const cycle: Array<'high' | 'medium' | 'low'> = ['low', 'medium', 'high'];
    updatePaper({
      ...paper,
      todos: paper.todos.map((t) => {
        if (t.id !== id) return t;
        const curIdx = cycle.indexOf(t.priority);
        return { ...t, priority: cycle[(curIdx + 1) % cycle.length] };
      }),
    });
  };

  return (
    <div>
      <div style={{ marginBottom: 12 }}>
        <div
          style={{
            fontSize: 10,
            fontWeight: 700,
            textTransform: 'uppercase',
            color: '#888',
            letterSpacing: 0.5,
            marginBottom: 6,
            paddingBottom: 4,
            borderBottom: '1px solid #e8e4da',
          }}
        >
          Notes
        </div>
        {paper.notes.map((note) => (
          <div key={note.id}>
            {editingNoteId === note.id ? (
              <div style={{ marginBottom: 6 }}>
                <textarea
                  value={editNoteDraft}
                  onChange={(e) => setEditNoteDraft(e.target.value)}
                  onKeyDown={(e) => { if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); saveNote(); } if (e.key === 'Escape') setEditingNoteId(null); }}
                  autoFocus
                  style={{ ...inputStyle, width: '100%', minHeight: 48, resize: 'vertical', boxSizing: 'border-box' }}
                />
                <div style={{ display: 'flex', gap: 4, marginTop: 2, justifyContent: 'flex-end' }}>
                  <button onClick={saveNote} style={{ ...inputStyle, background: '#2a7a8a', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 600, fontSize: 10, padding: '2px 8px' }}>Save</button>
                  <button onClick={() => setEditingNoteId(null)} style={{ ...inputStyle, cursor: 'pointer', fontSize: 10, padding: '2px 8px' }}>Cancel</button>
                </div>
              </div>
            ) : (
              <div
                style={{
                  fontSize: 12,
                  lineHeight: 1.6,
                  color: '#555',
                  padding: '6px 8px',
                  background: '#faf8f3',
                  borderRadius: 4,
                  borderLeft: '3px solid #d0ccc0',
                  marginBottom: 6,
                  cursor: 'pointer',
                  position: 'relative',
                }}
                onClick={() => startEditNote(note.id, note.content)}
                title="Click to edit"
              >
                {note.content}
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: 2 }}>
                  <span style={{ fontSize: 9, color: '#bbb' }}>{note.createdAt}</span>
                  <button
                    onClick={(e) => { e.stopPropagation(); removeNote(note.id); }}
                    style={{ background: 'none', border: 'none', color: '#ccc', cursor: 'pointer', fontSize: 10 }}
                    title="Remove note"
                  >&times;</button>
                </div>
              </div>
            )}
          </div>
        ))}
        <div style={{ display: 'flex', gap: 4, marginTop: 4 }}>
          <input
            value={newNote}
            onChange={(e) => setNewNote(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && addNote()}
            placeholder="Add note..."
            style={{
              flex: 1,
              padding: '4px 8px',
              border: '1px solid #d0ccc0',
              borderRadius: 4,
              fontSize: 11,
              background: '#faf8f3',
            }}
          />
          <button
            onClick={addNote}
            style={{
              padding: '4px 10px',
              border: 'none',
              borderRadius: 4,
              background: '#2a7a8a',
              color: 'white',
              fontSize: 11,
              cursor: 'pointer',
            }}
          >
            +
          </button>
        </div>
      </div>

      <div>
        <div
          style={{
            fontSize: 10,
            fontWeight: 700,
            textTransform: 'uppercase',
            color: '#888',
            letterSpacing: 0.5,
            marginBottom: 6,
            paddingBottom: 4,
            borderBottom: '1px solid #e8e4da',
          }}
        >
          Action Items
        </div>
        {paper.todos.map((todo) => {
          const ps = PRIORITY_STYLES[todo.priority];
          if (editingTodoId === todo.id) {
            return (
              <div key={todo.id} style={{ display: 'flex', gap: 4, padding: '4px 0' }}>
                <input
                  value={editTodoDraft}
                  onChange={(e) => setEditTodoDraft(e.target.value)}
                  onKeyDown={(e) => { if (e.key === 'Enter') saveTodo(); if (e.key === 'Escape') setEditingTodoId(null); }}
                  autoFocus
                  style={{ ...inputStyle, flex: 1 }}
                />
                <button onClick={saveTodo} style={{ ...inputStyle, background: '#2a7a8a', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 600, fontSize: 10 }}>Save</button>
                <button onClick={() => setEditingTodoId(null)} style={{ ...inputStyle, cursor: 'pointer', fontSize: 10 }}>Cancel</button>
              </div>
            );
          }
          return (
            <div
              key={todo.id}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 8,
                padding: '5px 0',
                fontSize: 12,
              }}
            >
              <div
                onClick={() => toggleTodo(todo.id)}
                style={{
                  width: 16,
                  height: 16,
                  border: todo.done ? 'none' : '2px solid #ccc',
                  borderRadius: 3,
                  background: todo.done ? '#40a060' : undefined,
                  color: 'white',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  fontSize: 10,
                  cursor: 'pointer',
                  flexShrink: 0,
                }}
              >
                {todo.done ? '\u2713' : ''}
              </div>
              <span
                onClick={() => cyclePriority(todo.id)}
                style={{
                  fontSize: 9,
                  padding: '1px 6px',
                  borderRadius: 8,
                  fontWeight: 700,
                  background: ps.bg,
                  color: ps.color,
                  cursor: 'pointer',
                }}
                title="Click to cycle priority"
              >
                {ps.label}
              </span>
              <span
                onClick={() => startEditTodo(todo.id, todo.text)}
                style={{
                  textDecoration: todo.done ? 'line-through' : undefined,
                  color: todo.done ? '#999' : undefined,
                  cursor: 'pointer',
                  flex: 1,
                  borderBottom: '1px dashed transparent',
                }}
                onMouseEnter={(e) => (e.currentTarget.style.borderBottomColor = '#ccc')}
                onMouseLeave={(e) => (e.currentTarget.style.borderBottomColor = 'transparent')}
                title="Click to edit"
              >
                {todo.text}
              </span>
              <button
                onClick={() => removeTodo(todo.id)}
                style={{ background: 'none', border: 'none', color: '#ccc', cursor: 'pointer', fontSize: 10 }}
                title="Remove"
              >&times;</button>
            </div>
          );
        })}
        <div style={{ display: 'flex', gap: 4, marginTop: 6 }}>
          <select
            value={newPriority}
            onChange={(e) => setNewPriority(e.target.value as 'high' | 'medium' | 'low')}
            style={{ padding: '4px', border: '1px solid #d0ccc0', borderRadius: 4, fontSize: 10, background: '#faf8f3' }}
          >
            <option value="high">HIGH</option>
            <option value="medium">MED</option>
            <option value="low">LOW</option>
          </select>
          <input
            value={newTodo}
            onChange={(e) => setNewTodo(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && addTodo()}
            placeholder="Add action item..."
            style={{
              flex: 1,
              padding: '4px 8px',
              border: '1px solid #d0ccc0',
              borderRadius: 4,
              fontSize: 11,
              background: '#faf8f3',
            }}
          />
          <button
            onClick={addTodo}
            style={{
              padding: '4px 10px',
              border: 'none',
              borderRadius: 4,
              background: '#2a7a8a',
              color: 'white',
              fontSize: 11,
              cursor: 'pointer',
            }}
          >
            +
          </button>
        </div>
      </div>
    </div>
  );
}
