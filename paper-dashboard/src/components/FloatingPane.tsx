import { useState, useRef, useCallback, type ReactNode } from 'react';
import type { PaneId } from '../types/paper';
import { useDashboard } from '../context/DashboardContext';

interface Props {
  paneId: PaneId;
  number: number;
  title: string;
  children: ReactNode;
  initialPosition?: { x: number; y: number };
  initialSize?: { width: number; height: number };
}

export function FloatingPane({
  paneId,
  number,
  title,
  children,
  initialPosition = { x: 100, y: 100 },
  initialSize = { width: 480, height: 360 },
}: Props) {
  const { togglePaneDocked, setPanes } = useDashboard();
  const [pos, setPos] = useState(initialPosition);
  const [size, setSize] = useState(initialSize);
  const [minimized, setMinimized] = useState(false);
  const dragRef = useRef<{ startX: number; startY: number; origX: number; origY: number } | null>(null);
  const resizeRef = useRef<{ startX: number; startY: number; origW: number; origH: number } | null>(null);

  const onDragStart = useCallback(
    (e: React.MouseEvent) => {
      e.preventDefault();
      dragRef.current = { startX: e.clientX, startY: e.clientY, origX: pos.x, origY: pos.y };

      const onMove = (ev: MouseEvent) => {
        if (!dragRef.current) return;
        setPos({
          x: dragRef.current.origX + ev.clientX - dragRef.current.startX,
          y: dragRef.current.origY + ev.clientY - dragRef.current.startY,
        });
      };
      const onUp = () => {
        dragRef.current = null;
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
      };
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    },
    [pos],
  );

  const onResizeStart = useCallback(
    (e: React.MouseEvent) => {
      e.preventDefault();
      e.stopPropagation();
      resizeRef.current = { startX: e.clientX, startY: e.clientY, origW: size.width, origH: size.height };

      const onMove = (ev: MouseEvent) => {
        if (!resizeRef.current) return;
        const newW = Math.max(280, resizeRef.current.origW + ev.clientX - resizeRef.current.startX);
        const newH = Math.max(200, resizeRef.current.origH + ev.clientY - resizeRef.current.startY);
        setSize({ width: newW, height: newH });
      };
      const onUp = () => {
        resizeRef.current = null;
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        setPanes((prev) =>
          prev.map((p) =>
            p.id === paneId
              ? { ...p, floatingPosition: pos, floatingSize: size }
              : p,
          ),
        );
      };
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    },
    [size, paneId, pos, setPanes],
  );

  return (
    <div
      style={{
        position: 'fixed',
        left: pos.x,
        top: pos.y,
        width: size.width,
        height: minimized ? 'auto' : size.height,
        zIndex: 200,
        boxShadow: '0 8px 32px rgba(0,0,0,0.18)',
        border: '1px solid #ccc8bc',
        borderRadius: 8,
        background: '#fff',
        overflow: 'hidden',
        display: 'flex',
        flexDirection: 'column',
      }}
    >
      <div
        onMouseDown={onDragStart}
        style={{
          display: 'flex',
          alignItems: 'center',
          padding: '8px 14px',
          background: 'linear-gradient(180deg, #faf8f3, #f0ede6)',
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
        <span style={{ fontSize: 9, color: '#d09030', marginRight: 8 }}>floating</span>
        <div style={{ display: 'flex', gap: 4 }}>
          <button
            onClick={() => setMinimized(!minimized)}
            style={{
              width: 22,
              height: 22,
              border: 'none',
              background: 'none',
              borderRadius: 3,
              color: '#999',
              cursor: 'pointer',
              fontSize: 14,
            }}
            title={minimized ? 'Expand' : 'Minimize'}
          >
            {minimized ? '\u25A1' : '\u2013'}
          </button>
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
            }}
            title="Pin to dock"
          >
            {'\u{1F4CC}'}
          </button>
        </div>
      </div>
      {!minimized && (
        <>
          <div style={{ padding: '12px 14px', flex: 1, overflow: 'auto' }}>{children}</div>
          <div
            onMouseDown={onResizeStart}
            style={{
              position: 'absolute',
              right: 0,
              bottom: 0,
              width: 16,
              height: 16,
              cursor: 'nwse-resize',
              background: 'linear-gradient(135deg, transparent 50%, #ccc8bc 50%)',
              borderRadius: '0 0 8px 0',
            }}
          />
        </>
      )}
    </div>
  );
}
