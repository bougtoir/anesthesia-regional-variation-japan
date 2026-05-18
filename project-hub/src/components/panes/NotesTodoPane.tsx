import { useState } from 'react';
import type { Paper } from '../../types/paper';
import { useDashboard } from '../../context/DashboardContext';

const PRIORITY_STYLES: Record<string, { bg: string; color: string; label: string }> = {
  high: { bg: '#fde0e0', color: '#c04040', label: 'HIGH' },
  medium: { bg: '#fdf0d0', color: '#a07020', label: 'MED' },
  low: { bg: '#e0f0e8', color: '#408060', label: 'LOW' },
};

interface Props {
  paper: Paper;
}

export function NotesTodoPane({ paper }: Props) {
  const { updatePaper } = useDashboard();
  const [newNote, setNewNote] = useState('');
  const [newTodo, setNewTodo] = useState('');
  const [newPriority, setNewPriority] = useState<'high' | 'medium' | 'low'>('medium');

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

  const addTodo = () => {
    if (!newTodo.trim()) return;
    updatePaper({
      ...paper,
      todos: [...paper.todos, { id: `t-${Date.now()}`, text: newTodo.trim(), done: false, priority: newPriority }],
    });
    setNewTodo('');
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
          <div
            key={note.id}
            style={{
              fontSize: 12,
              lineHeight: 1.6,
              color: '#555',
              padding: '6px 8px',
              background: '#faf8f3',
              borderRadius: 4,
              borderLeft: '3px solid #d0ccc0',
              marginBottom: 6,
            }}
          >
            {note.content}
            <div style={{ fontSize: 9, color: '#bbb', marginTop: 2 }}>{note.createdAt}</div>
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
                style={{
                  fontSize: 9,
                  padding: '1px 6px',
                  borderRadius: 8,
                  fontWeight: 700,
                  background: ps.bg,
                  color: ps.color,
                }}
              >
                {ps.label}
              </span>
              <span style={{ textDecoration: todo.done ? 'line-through' : undefined, color: todo.done ? '#999' : undefined }}>
                {todo.text}
              </span>
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
