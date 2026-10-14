import type { HistoryRecord } from '../types';

const STORAGE_KEY = 'cv_historico';

export function getHistory(): HistoryRecord[] {
  if (typeof window === 'undefined') return [];
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    return raw ? JSON.parse(raw) : [];
  } catch (e) {
    console.error('Failed to parse history:', e);
    return [];
  }
}

export function addHistory(record: HistoryRecord): void {
  if (typeof window === 'undefined') return;
  try {
    const current = getHistory();
    const updated = [record, ...current].slice(0, 50);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(updated));
  } catch (e) {
    console.error('Failed to save history:', e);
  }
}

export function clearHistory(): void {
  if (typeof window === 'undefined') return;
  localStorage.removeItem(STORAGE_KEY);
}
