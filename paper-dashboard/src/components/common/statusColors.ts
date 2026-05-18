import type { PaperStatus } from '../../types/paper';

export const STATUS_DOT_COLORS: Record<PaperStatus, string> = {
  drafting: '#e0a040',
  'internal-review': '#8070b0',
  submitted: '#5090d0',
  'under-review': '#9070c0',
  revision: '#d08050',
  accepted: '#40a060',
  published: '#208040',
  rejected: '#c04040',
};
