import type { SheetData, SheetRoute, SheetStore } from '../types';
import { normalizeCode } from './manifestParser';

export function extractSpreadsheetId(input: string): string | null {
  const trimmed = input.trim();
  if (!trimmed) return null;

  // Pattern: https://docs.google.com/spreadsheets/d/{spreadsheetId}/...
  const urlMatch = trimmed.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (urlMatch && urlMatch[1]) {
    return urlMatch[1];
  }

  // If it's already an ID (alphanumeric, dashes, underscores, typically 25-50 chars)
  if (/^[a-zA-Z0-9-_]{20,60}$/.test(trimmed)) {
    return trimmed;
  }

  return null;
}

export function colToLetter(col: number): string {
  let temp: number;
  let letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = Math.floor((col - temp - 1) / 26);
  }
  return letter;
}

export interface SpreadsheetMeta {
  title: string;
  sheets: Array<{
    sheetId: number;
    title: string;
  }>;
}

export async function fetchSpreadsheetMeta(
  spreadsheetId: string,
  accessToken: string
): Promise<SpreadsheetMeta> {
  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?fields=properties.title,sheets.properties(sheetId,title)`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!res.ok) {
    const errorBody = await res.json().catch(() => ({}));
    const message =
      errorBody?.error?.message ||
      (res.status === 403
        ? 'Sem permissão para acessar esta planilha. Verifique o compartilhamento.'
        : res.status === 404
        ? 'Planilha não encontrada. Verifique o link ou ID informado.'
        : `Erro ${res.status} ao acessar a API do Google Sheets`);
    throw new Error(message);
  }

  const data = await res.json();
  return {
    title: data.properties?.title || 'Planilha Google Sheets',
    sheets: (data.sheets || []).map((s: any) => ({
      sheetId: s.properties.sheetId,
      title: s.properties.title,
    })),
  };
}

export async function readGoogleSheet(
  spreadsheetId: string,
  sheetName: string,
  accessToken: string
): Promise<SheetData> {
  // Fetch values up to 200 rows and 26 columns (A to Z)
  const range = `'${sheetName}'!A1:Z200`;
  const res = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(range)}`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (!res.ok) {
    const errorBody = await res.json().catch(() => ({}));
    throw new Error(
      errorBody?.error?.message || `Falha ao ler os dados da aba '${sheetName}'`
    );
  }

  const result = await res.json();
  const rows: string[][] = result.values || [];

  const routes: SheetRoute[] = [];
  const maxRow = rows.length;
  let maxCol = 0;
  for (const row of rows) {
    if (row.length > maxCol) maxCol = row.length;
  }

  const getCellStr = (r: number, c: number): string => {
    const row = rows[r - 1];
    if (!row) return '';
    return String(row[c - 1] ?? '').trim();
  };

  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const headerLoja = getCellStr(r, c).toUpperCase();
      if (headerLoja !== 'LOJA') continue;

      const volCol = c + 1;
      const headerVol = getCellStr(r, volCol).toUpperCase();
      if (headerVol !== 'VOL') continue;

      // Find route name 1 to 3 rows above the LOJA header
      let routeName = '';
      for (let prevR = r - 1; prevR >= 1 && prevR >= r - 3; prevR--) {
        const val = getCellStr(prevR, c);
        if (val) {
          routeName = val;
          break;
        }
      }
      if (!routeName) {
        routeName = `ROTA (${colToLetter(c)}${r})`;
      }

      const stores: SheetStore[] = [];
      let totalRow: number | null = null;

      for (let curR = r + 1; curR <= maxRow; curR++) {
        const cellVal = getCellStr(curR, c);
        if (/^TOTAL/i.test(cellVal)) {
          totalRow = curR;
          break;
        }

        if (!cellVal) {
          if (!stores.length) continue;
          break;
        }

        if (!/^\d{3,5}$/.test(cellVal)) {
          break;
        }

        const volStr = getCellStr(curR, volCol);
        const parsedVol = parseInt(volStr, 10);
        const currentVol = Number.isNaN(parsedVol) ? null : parsedVol;

        stores.push({
          route: routeName,
          display: cellVal,
          key: normalizeCode(cellVal),
          row: curR,
          lojaCol: c,
          volCol,
          currentVol,
        });
      }

      if (stores.length > 0) {
        routes.push({
          name: routeName,
          totalRow,
          volCol,
          stores,
        });
      }
    }
  }

  return {
    sheetName,
    routes,
    fileName: `Google Sheets - ${sheetName}`,
    sourceType: 'google_sheets',
    spreadsheetId,
    spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,
  };
}

export interface GoogleSheetsBatchResult {
  updatedCells: number;
  updatedRangeCount: number;
}

export async function batchUpdateGoogleSheet(
  spreadsheetId: string,
  sheetName: string,
  sheetData: SheetData,
  storeMap: Map<string, number>,
  dateIso: string,
  orderNumber: string,
  accessToken: string
): Promise<GoogleSheetsBatchResult> {
  const updates: Array<{ range: string; values: Array<Array<string | number>> }> = [];

  // 1. Update store volumes
  for (const route of sheetData.routes) {
    for (const store of route.stores) {
      const newVol = storeMap.get(store.key);
      if (newVol !== undefined) {
        const cellRef = `'${sheetName}'!${colToLetter(store.volCol)}${store.row}`;
        updates.push({
          range: cellRef,
          values: [[newVol]],
        });
        store.currentVol = newVol;
      }
    }

    // 2. Recalculate and update TOTAL row if present
    if (route.totalRow) {
      const sum = route.stores.reduce(
        (acc, item) => acc + (storeMap.get(item.key) ?? item.currentVol ?? 0),
        0
      );
      const totalCellRef = `'${sheetName}'!${colToLetter(route.volCol)}${route.totalRow}`;
      updates.push({
        range: totalCellRef,
        values: [[sum]],
      });
    }
  }

  // 3. Search and update DATA and PEDIDO cells in first few rows
  const metaRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(
      `'${sheetName}'!A1:Z10`
    )}`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`,
      },
    }
  );

  if (metaRes.ok) {
    const metaData = await metaRes.json();
    const rows: string[][] = metaData.values || [];
    const [year, month, day] = dateIso.split('-');

    for (let r = 1; r <= rows.length; r++) {
      const row = rows[r - 1] || [];
      for (let c = 1; c <= row.length; c++) {
        const val = String(row[c - 1] || '').trim();

        // Update DATA cell
        if (/^DATA\s*:/i.test(val)) {
          const dateCellRef = `'${sheetName}'!${colToLetter(c)}${r}`;
          updates.push({
            range: dateCellRef,
            values: [[`DATA:     ${day} / ${month} / ${year}`]],
          });
        }

        // Update PEDIDO cell
        if (orderNumber && /PEDIDO/i.test(val)) {
          const pedidoCellRef = `'${sheetName}'!${colToLetter(c)}${r}`;
          if (/\d/.test(val)) {
            updates.push({
              range: pedidoCellRef,
              values: [[val.replace(/\d+/, orderNumber)]],
            });
          } else {
            updates.push({
              range: pedidoCellRef,
              values: [[`PEDIDO: ${orderNumber}`]],
            });
          }
        }
      }
    }
  }

  if (updates.length === 0) {
    throw new Error('Nenhuma célula para atualizar.');
  }

  const batchRes = await fetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values:batchUpdate`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        valueInputOption: 'USER_ENTERED',
        data: updates,
      }),
    }
  );

  if (!batchRes.ok) {
    const errData = await batchRes.json().catch(() => ({}));
    throw new Error(
      errData?.error?.message ||
        `Erro ${batchRes.status} ao gravar dados na planilha do Google Sheets`
    );
  }

  const batchResult = await batchRes.json();
  return {
    updatedCells: batchResult.totalUpdatedCells || updates.length,
    updatedRangeCount: updates.length,
  };
}
