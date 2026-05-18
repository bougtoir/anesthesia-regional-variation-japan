import type { ReactNode } from 'react';
import type { PaneId } from '../types/paper';
import { useDashboard } from '../context/DashboardContext';

interface Props {
  paneId: PaneId;
  number: number;
  title: string;
  children: ReactNode;
  onAdd?: () => void;
}

export function PaneWrapper({ paneId, number, title, children, onAdd }: Props) {
  const { togglePaneDocked } = useDashboard();

  return (
    <div
      style={{
        background: '#fff',
        borderRadius: 6,
        boxShadow: '0 1px 4px rgba(0,0,0,0.08)',
        border: '1px solid #ddd8cc',
        overflow: 'hidden',
        display: 'flex',
        flexDirection: 'column',
        minWidth: 0,
      }}
    >
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          padding: '8px 14px',
          background: '#faf8f3',
          borderBottom: '1px solid #e8e4da',
          cursor: 'grab',
          flexShrink: 0,
        }}
      >
        <div
          style={{
            width: 20,
            height: 20,
            marginRight: 8,
            fontSize: 12,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            background: '#e8e4da',
            borderRadius: 4,
            fontWeight: 700,
            color: '#888',
          }}
        >
          {number}
        </div>
        <div style={{ fontSize: 12, fontWeight: 700, color: '#444', flex: 1 }}>{title}</div>
        <div style={{ display: 'flex', gap: 4 }}>
          {onAdd && (
            <button
              onClick={onAdd}
              style={{
                width: 22,
                height: 22,
                border: 'none',
                background: 'none',
                borderRadius: 3,
                color: '#999',
                cursor: 'pointer',
                fontSize: 14,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
              }}
              title="Add"
            >
              +
            </button>
          )}
          <button
            onClick={() => togglePaneDocked(paneId)}
            style={{
              width: 22,
              height: 22,
              border: 'none',
              background: 'none',
              borderRadius: 3,
              color: '#999',
              cursor: 'pointer',
              fontSize: 11,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
            }}
            title="Undock (float)"
          >
            {'\u2197'}
          </button>
        </div>
      </div>
      <div style={{ padding: '12px 14px', flex: 1, overflow: 'auto' }}>{children}</div>
    </div>
  );
}
