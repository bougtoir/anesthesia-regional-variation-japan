import type { PaperStatus } from '../../types/paper';

export const STATUS_DOT_COLORS: Record<PaperStatus, string> = {
  planning: '#e0a040',
  'in-progress': '#5090d0',
  'in-review': '#9070c0',
  'on-hold': '#8070b0',
  testing: '#d08050',
  approved: '#40a060',
  completed: '#208040',
  cancelled: '#c04040',
};
