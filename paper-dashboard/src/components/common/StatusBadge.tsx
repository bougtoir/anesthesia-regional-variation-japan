import type { PaperStatus } from '../../types/paper';

const STATUS_STYLES: Record<PaperStatus, { bg: string; text: string; label: string }> = {
  drafting: { bg: '#f8f0e0', text: '#907020', label: 'DRAFTING' },
  'internal-review': { bg: '#e8e0f0', text: '#605090', label: 'INTERNAL REVIEW' },
  submitted: { bg: '#e0eef8', text: '#3070a0', label: 'SUBMITTED' },
  'under-review': { bg: '#f0e8f8', text: '#7050a0', label: 'UNDER REVIEW' },
  revision: { bg: '#f8eee0', text: '#a06830', label: 'REVISION' },
  accepted: { bg: '#e0f0e8', text: '#308050', label: 'ACCEPTED' },
  published: { bg: '#d8f0d8', text: '#206830', label: 'PUBLISHED' },
  rejected: { bg: '#f8e0e0', text: '#a03030', label: 'REJECTED' },
};

export function StatusBadge({ status }: { status: PaperStatus }) {
  const s = STATUS_STYLES[status];
  return (
    <span
      style={{
        display: 'inline-block',
        padding: '2px 8px',
        borderRadius: 10,
        fontSize: 10,
        fontWeight: 700,
        background: s.bg,
        color: s.text,
      }}
    >
      {s.label}
    </span>
  );
}
