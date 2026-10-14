import type { Workbook } from 'exceljs';

export interface ManifestStore {
  code: string;
  volumes: number;
  occurrences: number;
  ignored: number;
}

export interface ManifestBatch {
  id: string;
  name: string;
  entries: ManifestStore[];
}

export interface SheetStore {
  route: string;
  display: string;
  key: string;
  row: number;
  lojaCol: number;
  volCol: number;
  currentVol: number | null;
}

export interface SheetRoute {
  name: string;
  totalRow: number | null;
  volCol: number;
  stores: SheetStore[];
}

export interface SheetData {
  workbook?: Workbook;
  sheetName: string;
  routes: SheetRoute[];
  fileName: string;
  sourceType: 'local_xlsx' | 'google_sheets';
  spreadsheetId?: string;
  spreadsheetTitle?: string;
  spreadsheetUrl?: string;
}

export interface HistoryRecord {
  id: string;
  date: string;
  manifest: string;
  stores: number;
  volumes: number;
  filled: number;
  missing: number;
  status: 'Concluído' | 'Com pendências';
  targetType?: 'local_xlsx' | 'google_sheets';
  spreadsheetTitle?: string;
  entries: Array<{
    code: string;
    volumes: number;
    occurrences: number;
  }>;
}
