import { useState } from 'react';
import type { Paper, CoauthorTask } from '../../types/paper';
import { useDashboard } from '../../context/DashboardContext';

const TASK_STATUS: Record<string, { bg: string; color: string; label: string }> = {
  done: { bg: '#e0f0e8', color: '#408060', label: 'Done' },
  'in-progress': { bg: '#fdf0d0', color: '#a07020', label: 'In Progress' },
  waiting: { bg: '#fde0e0', color: '#c04040', label: 'Waiting' },
};

const STATUS_CYCLE: CoauthorTask['status'][] = ['waiting', 'in-progress', 'done'];

const AVATAR_COLORS = ['#2a7a8a', '#6a5a8a', '#5a8a6a', '#8a6a5a', '#5a7a8a', '#8a5a7a', '#6a8a5a', '#8a5a6a'];

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

export function CoauthorTasksPane({ paper }: Props) {
  const { updatePaper } = useDashboard();
  const [showAdd, setShowAdd] = useState(false);
  const [newAuthor, setNewAuthor] = useState('');
  const [newTask, setNewTask] = useState('');
  const [newStatus, setNewStatus] = useState<CoauthorTask['status']>('in-progress');

  // Edit state
  const [editingIdx, setEditingIdx] = useState<number | null>(null);
  const [editAuthor, setEditAuthor] = useState('');
  const [editTask, setEditTask] = useState('');

  const cycleStatus = (index: number) => {
    const ct = paper.coauthorTasks[index];
    const curIdx = STATUS_CYCLE.indexOf(ct.status);
    const nextStatus = STATUS_CYCLE[(curIdx + 1) % STATUS_CYCLE.length];
    const updated = paper.coauthorTasks.map((item, i) =>
      i === index ? { ...item, status: nextStatus } : item,
    );
    updatePaper({ ...paper, coauthorTasks: updated });
  };

  const addTask = () => {
    if (!newAuthor.trim() || !newTask.trim()) return;
    const name = newAuthor.trim();
    const initials = name
      .split(/\s+/)
      .map((w) => w[0])
      .join('')
      .toUpperCase()
      .slice(0, 2);
    const color = AVATAR_COLORS[paper.coauthorTasks.length % AVATAR_COLORS.length];
    updatePaper({
      ...paper,
      coauthorTasks: [
        ...paper.coauthorTasks,
        { authorName: name, initials, task: newTask.trim(), status: newStatus, color },
      ],
    });
    setNewAuthor('');
    setNewTask('');
    setNewStatus('in-progress');
    setShowAdd(false);
  };

  const removeTask = (index: number) => {
    updatePaper({ ...paper, coauthorTasks: paper.coauthorTasks.filter((_, i) => i !== index) });
  };

  const startEdit = (i: number) => {
    const ct = paper.coauthorTasks[i];
    setEditAuthor(ct.authorName);
    setEditTask(ct.task);
    setEditingIdx(i);
  };

  const saveEdit = () => {
    if (editingIdx === null || !editAuthor.trim() || !editTask.trim()) return;
    const name = editAuthor.trim();
    const initials = name.split(/\s+/).map((w) => w[0]).join('').toUpperCase().slice(0, 2);
    const updated = paper.coauthorTasks.map((ct, i) =>
      i === editingIdx ? { ...ct, authorName: name, initials, task: editTask.trim() } : ct,
    );
    updatePaper({ ...paper, coauthorTasks: updated });
    setEditingIdx(null);
  };

  if (paper.coauthorTasks.length === 0 && !showAdd) {
    return (
      <div>
        <div style={{ color: '#aaa', fontSize: 12 }}>No tasks assigned</div>
        <button
          onClick={() => setShowAdd(true)}
          style={{ marginTop: 6, background: 'none', border: '1px dashed #ccc', borderRadius: 4, color: '#888', cursor: 'pointer', fontSize: 10, padding: '3px 8px' }}
        >
          + Add Task
        </button>
      </div>
    );
  }

  return (
    <div>
      {paper.coauthorTasks.map((ct, i) => {
        const ts = TASK_STATUS[ct.status] ?? TASK_STATUS['in-progress'];
        if (editingIdx === i) {
          return (
            <div key={i} style={{ padding: 6, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da', marginBottom: 4 }}>
              <div style={{ display: 'flex', gap: 4, marginBottom: 4 }}>
                <input value={editAuthor} onChange={(e) => setEditAuthor(e.target.value)} placeholder="Author name" style={{ ...inputStyle, flex: 1 }} />
              </div>
              <div style={{ display: 'flex', gap: 4 }}>
                <input value={editTask} onChange={(e) => setEditTask(e.target.value)} placeholder="Task description" style={{ ...inputStyle, flex: 1 }} />
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
              gap: 8,
              padding: '6px 8px',
              borderBottom: '1px solid #f0ede6',
              background: i % 2 === 1 ? '#faf8f3' : undefined,
            }}
          >
            <div
              style={{
                width: 28,
                height: 28,
                borderRadius: '50%',
                background: ct.color ?? '#2a7a8a',
                color: 'white',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                fontSize: 11,
                fontWeight: 700,
                flexShrink: 0,
              }}
            >
              {ct.initials}
            </div>
            <div
              style={{ flex: 1, cursor: 'pointer' }}
              onClick={() => startEdit(i)}
              title="Click to edit"
            >
              <div style={{ fontSize: 12, fontWeight: 600 }}>{ct.authorName}</div>
              <div style={{ fontSize: 10, color: '#999' }}>{ct.task}</div>
            </div>
            <span
              onClick={() => cycleStatus(i)}
              style={{
                fontSize: 10,
                padding: '2px 8px',
                borderRadius: 10,
                fontWeight: 600,
                background: ts.bg,
                color: ts.color,
                cursor: 'pointer',
              }}
              title="Click to cycle status"
            >
              {ts.label}
              {ct.status === 'waiting' && ct.waitingDays ? ` (${ct.waitingDays}d)` : ''}
            </span>
            <button
              onClick={() => removeTask(i)}
              style={{ background: 'none', border: 'none', color: '#ccc', cursor: 'pointer', fontSize: 12 }}
              title="Remove"
            >&times;</button>
          </div>
        );
      })}

      {showAdd ? (
        <div style={{ marginTop: 6, padding: 8, background: '#faf8f3', borderRadius: 4, border: '1px solid #e8e4da' }}>
          <div style={{ display: 'flex', gap: 4, marginBottom: 4 }}>
            <input value={newAuthor} onChange={(e) => setNewAuthor(e.target.value)} placeholder="Author name" style={{ ...inputStyle, flex: 1 }} />
            <select value={newStatus} onChange={(e) => setNewStatus(e.target.value as CoauthorTask['status'])} style={inputStyle}>
              <option value="in-progress">In Progress</option>
              <option value="waiting">Waiting</option>
              <option value="done">Done</option>
            </select>
          </div>
          <div style={{ display: 'flex', gap: 4 }}>
            <input value={newTask} onChange={(e) => setNewTask(e.target.value)} placeholder="Task description" style={{ ...inputStyle, flex: 1 }} />
            <button onClick={addTask} style={{ ...inputStyle, background: '#2a7a8a', color: 'white', border: 'none', cursor: 'pointer', fontWeight: 600 }}>Add</button>
            <button onClick={() => setShowAdd(false)} style={{ ...inputStyle, cursor: 'pointer' }}>Cancel</button>
          </div>
        </div>
      ) : (
        <button
          onClick={() => setShowAdd(true)}
          style={{ marginTop: 6, background: 'none', border: '1px dashed #ccc', borderRadius: 4, color: '#888', cursor: 'pointer', fontSize: 10, padding: '3px 8px' }}
        >
          + Add Task
        </button>
      )}
    </div>
  );
}
