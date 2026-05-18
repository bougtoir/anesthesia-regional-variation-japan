import { useCallback, useRef } from 'react';
import { useDashboard } from '../context/DashboardContext';
import type { PaneConfig } from '../types/paper';

export function SettingsPanel() {
  const { panes, setPanes, setSettingsOpen, exportJson, exportYaml, importData } = useDashboard();
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleToggle = useCallback(
    (id: string) => {
      setPanes((prev: PaneConfig[]) => prev.map((p: PaneConfig) => (p.id === id ? { ...p, docked: !p.docked } : p)));
    },
    [setPanes],
  );

  const handleDragStart = useCallback((e: React.DragEvent, index: number) => {
    e.dataTransfer.setData('text/plain', String(index));
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent, dropIndex: number) => {
      e.preventDefault();
      const dragIndex = parseInt(e.dataTransfer.getData('text/plain'), 10);
      if (isNaN(dragIndex) || dragIndex === dropIndex) return;
      setPanes((prev: PaneConfig[]) => {
        const next = [...prev];
        const [moved] = next.splice(dragIndex, 1);
        next.splice(dropIndex, 0, moved);
        return next;
      });
    },
    [setPanes],
  );

  const handleExport = useCallback(
    (format: 'json' | 'yaml') => {
      const data = format === 'json' ? exportJson() : exportYaml();
      const blob = new Blob([data], { type: 'text/plain' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `papers.${format === 'json' ? 'json' : 'yaml'}`;
      a.click();
      URL.revokeObjectURL(url);
    },
    [exportJson, exportYaml],
  );

  const handleImport = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  const handleFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;
      const format = file.name.endsWith('.yaml') || file.name.endsWith('.yml') ? 'yaml' : 'json';
      const reader = new FileReader();
      reader.onload = () => {
        importData(reader.result as string, format);
      };
      reader.readAsText(file);
      e.target.value = '';
    },
    [importData],
  );

  return (
    <div
      style={{
        position: 'fixed',
        top: 44,
        right: 0,
        bottom: 0,
        width: 320,
        background: '#fff',
        boxShadow: '-4px 0 16px rgba(0,0,0,0.1)',
        zIndex: 300,
        overflowY: 'auto',
        borderLeft: '1px solid #d0ccc0',
      }}
    >
      <div
        style={{
          padding: 16,
          borderBottom: '1px solid #e8e4da',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
        }}
      >
        <span style={{ fontSize: 14, fontWeight: 700, color: '#1a3a4a' }}>Dashboard Settings</span>
        <button
          onClick={() => setSettingsOpen(false)}
          style={{ background: 'none', border: 'none', fontSize: 18, cursor: 'pointer', color: '#999' }}
        >
          &times;
        </button>
      </div>

      <div style={{ padding: '14px 16px', borderBottom: '1px solid #f0ede6' }}>
        <div
          style={{
            fontSize: 11,
            fontWeight: 700,
            textTransform: 'uppercase',
            color: '#888',
            letterSpacing: 0.5,
            marginBottom: 10,
          }}
        >
          Pane Display
        </div>
        <div style={{ fontSize: 10, color: '#aaa', marginBottom: 8 }}>
          Drag to reorder. Toggle to dock/float.
        </div>
        {panes.map((pane, i) => (
          <div
            key={pane.id}
            draggable
            onDragStart={(e) => handleDragStart(e, i)}
            onDragOver={(e) => e.preventDefault()}
            onDrop={(e) => handleDrop(e, i)}
            style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              padding: '6px 0',
              cursor: 'grab',
            }}
          >
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              <span style={{ color: '#ccc', fontSize: 14, cursor: 'grab' }}>{'\u2630'}</span>
              <span style={{ fontSize: 12 }}>
                {pane.number}. {pane.label}
              </span>
            </div>
            <div
              onClick={() => handleToggle(pane.id)}
              style={{
                width: 36,
                height: 20,
                borderRadius: 10,
                background: pane.docked ? '#2a7a8a' : '#d0ccc0',
                position: 'relative',
                cursor: 'pointer',
                transition: 'background 0.2s',
              }}
            >
              <div
                style={{
                  width: 16,
                  height: 16,
                  borderRadius: '50%',
                  background: 'white',
                  position: 'absolute',
                  top: 2,
                  left: pane.docked ? 18 : 2,
                  transition: 'left 0.2s',
                }}
              />
            </div>
          </div>
        ))}
      </div>

      <div style={{ padding: '14px 16px' }}>
        <div
          style={{
            fontSize: 11,
            fontWeight: 700,
            textTransform: 'uppercase',
            color: '#888',
            letterSpacing: 0.5,
            marginBottom: 10,
          }}
        >
          Data
        </div>
        <button
          onClick={handleImport}
          style={{
            width: '100%',
            marginBottom: 6,
            padding: 8,
            border: '1px solid #d0ccc0',
            background: '#faf8f3',
            borderRadius: 4,
            fontSize: 12,
            cursor: 'pointer',
            color: '#666',
          }}
        >
          Import JSON / YAML
        </button>
        <div style={{ display: 'flex', gap: 6 }}>
          <button
            onClick={() => handleExport('json')}
            style={{
              flex: 1,
              padding: 8,
              border: '1px solid #d0ccc0',
              background: '#faf8f3',
              borderRadius: 4,
              fontSize: 12,
              cursor: 'pointer',
              color: '#666',
            }}
          >
            Export JSON
          </button>
          <button
            onClick={() => handleExport('yaml')}
            style={{
              flex: 1,
              padding: 8,
              border: '1px solid #d0ccc0',
              background: '#faf8f3',
              borderRadius: 4,
              fontSize: 12,
              cursor: 'pointer',
              color: '#666',
            }}
          >
            Export YAML
          </button>
        </div>
        <input ref={fileInputRef} type="file" accept=".json,.yaml,.yml" onChange={handleFileChange} style={{ display: 'none' }} />
      </div>
    </div>
  );
}
