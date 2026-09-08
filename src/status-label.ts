import type { Severity } from './types';

export const statusLabel = (severity: Severity): string => ({
  match: 'Exact match',
  critical: 'Critical',
  warning: 'Warning',
  revision: 'Warning: Revision',
})[severity];
