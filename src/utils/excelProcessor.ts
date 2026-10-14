import ExcelJS from 'exceljs';
import type { SheetData, SheetRoute, SheetStore } from '../types';
import { normalizeCode } from './manifestParser';

export const cellToString = (val: unknown): string => {
  if (val == null) return '';
  if (typeof val === 'object') {
    if ('result' in (val as Record<string, unknown>)) {
      return String((val as { result?: unknown }).result ?? '').trim();
    }
    if ('richText' in (val as Record<string, unknown>)) {
      const rt = (val as { richText: Array<{ text: string }> }).richText;
      return rt.map((item) => item.text).join('').trim();
    }
  }
  return String(val).trim();
};

export async function loadWorkbookFromBuffer(file: File): Promise<SheetData> {
  const workbook = new ExcelJS.Workbook();
  const buffer = await file.arrayBuffer();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error('A planilha está vazia ou não possui abas.');
  }

  const routes: SheetRoute[] = [];
  const maxRow = worksheet.rowCount;
  const maxCol = worksheet.columnCount;

  for (let r = 1; r <= maxRow; r++) {
    for (let c = 1; c <= maxCol; c++) {
      const headerLoja = cellToString(worksheet.getCell(r, c).value).toUpperCase();
      if (headerLoja !== 'LOJA') continue;

      const volCol = c + 1;
      const headerVol = cellToString(worksheet.getCell(r, volCol).value).toUpperCase();
      if (headerVol !== 'VOL') continue;

      // Find route name 1 to 3 rows above the LOJA header
      let routeName = '';
      for (let prevR = r - 1; prevR >= 1 && prevR >= r - 3; prevR--) {
        const val = cellToString(worksheet.getCell(prevR, c).value);
        if (val) {
          routeName = val;
          break;
        }
      }
      if (!routeName) {
        routeName = `ROTA (${worksheet.getCell(r, c).address})`;
      }

      const stores: SheetStore[] = [];
      let totalRow: number | null = null;

      for (let curR = r + 1; curR <= maxRow; curR++) {
        const cellVal = cellToString(worksheet.getCell(curR, c).value);
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

        const volVal = worksheet.getCell(curR, volCol).value;
        const currentVol = typeof volVal === 'number' ? volVal : null;

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
    workbook,
    sheetName: worksheet.name,
    routes,
    fileName: file.name,
    sourceType: 'local_xlsx',
  };
}

export function fillWorkbookWithManifest(sheetData: SheetData, storeMap: Map<string, number>): void {
  if (!sheetData.workbook) return;
  const ws = sheetData.workbook.getWorksheet(sheetData.sheetName);
  if (!ws) return;

  for (const route of sheetData.routes) {
    for (const store of route.stores) {
      const newVol = storeMap.get(store.key);
      if (newVol !== undefined) {
        ws.getCell(store.row, store.volCol).value = newVol;
        store.currentVol = newVol;
      }
    }

    if (route.totalRow) {
      const sum = route.stores.reduce(
        (acc, item) => acc + (storeMap.get(item.key) ?? item.currentVol ?? 0),
        0
      );
      ws.getCell(route.totalRow, route.volCol).value = sum;
    }
  }
}

export function updateSheetDate(sheetData: SheetData, dateIso: string): void {
  if (!sheetData.workbook) return;
  const ws = sheetData.workbook.getWorksheet(sheetData.sheetName);
  if (!ws) return;

  const [year, month, day] = dateIso.split('-');
  const maxR = Math.min(ws.rowCount, 6);
  const maxC = Math.min(ws.columnCount, 6);

  for (let r = 1; r <= maxR; r++) {
    for (let c = 1; c <= maxC; c++) {
      const cell = ws.getCell(r, c);
      const strVal = cellToString(cell.value);
      if (/^DATA\s*:/i.test(strVal)) {
        cell.value = `DATA:     ${day} / ${month} / ${year}`;
        return;
      }
    }
  }
}

export function updateSheetOrderNumber(sheetData: SheetData, orderNumber: string): void {
  if (!orderNumber || !sheetData.workbook) return;
  const ws = sheetData.workbook.getWorksheet(sheetData.sheetName);
  if (!ws) return;

  const maxR = Math.min(ws.rowCount, 8);
  const maxC = Math.min(ws.columnCount, 20);

  // Check 7-digit numbers first
  for (let r = 1; r <= maxR; r++) {
    for (let c = 1; c <= maxC; c++) {
      const cell = ws.getCell(r, c);
      const str = cellToString(cell.value);
      if (/^\d{7}$/.test(str)) {
        cell.value = typeof cell.value === 'number' ? Number(orderNumber) : orderNumber;
        return;
      }
    }
  }

  // Check cell containing PEDIDO
  for (let r = 1; r <= maxR; r++) {
    for (let c = 1; c <= maxC; c++) {
      const cell = ws.getCell(r, c);
      const str = cellToString(cell.value);
      if (/PEDIDO/i.test(str)) {
        if (/\d/.test(str)) {
          cell.value = str.replace(/\d+/, orderNumber);
          return;
        }
        for (let nextC = c + 1; nextC <= maxC; nextC++) {
          const neighbor = ws.getCell(r, nextC);
          if (!cellToString(neighbor.value)) {
            neighbor.value = orderNumber;
            return;
          }
        }
        cell.value = `${str} ${orderNumber}`;
        return;
      }
    }
  }
}

export async function downloadFilledWorkbook(sheetData: SheetData, dateIso: string): Promise<void> {
  if (!sheetData.workbook) {
    throw new Error('Não há arquivo Excel local para download.');
  }
  const buffer = await sheetData.workbook.xlsx.writeBuffer();
  const [year, month, day] = dateIso.split('-');
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `CONFERENCIA_PM_${day}-${month}-${year}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

export function calculateJulianOrderNumber(dateVal: string | Date): string {
  let y: number, m: number, d: number;
  if (typeof dateVal === 'string') {
    const match = dateVal.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!match) return '';
    y = Number(match[1]);
    m = Number(match[2]);
    d = Number(match[3]);
  } else {
    y = dateVal.getFullYear();
    m = dateVal.getMonth() + 1;
    d = dateVal.getDate();
  }

  if (!y || !m || !d) return '';

  const isLeapYear = (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
  const daysInMonths = [
    31,
    isLeapYear ? 29 : 28,
    31,
    30,
    31,
    30,
    31,
    31,
    30,
    31,
    30,
    31,
  ];

  let dayOfYear = d;
  for (let i = 0; i < m - 1; i++) {
    dayOfYear += daysInMonths[i];
  }

  return `${String(y).padStart(4, '0')}${String(dayOfYear).padStart(3, '0')}`;
}

export async function generateDemoTemplateWorkbook(): Promise<File> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('CONFERENCIA');

  // Title and metadata
  ws.getCell('A1').value = 'CONFERÊNCIA DE VOLUMES POR ROTA';
  ws.getCell('A1').font = { bold: true, size: 14 };

  const today = new Date().toISOString().slice(0, 10);
  const [year, month, day] = today.split('-');
  ws.getCell('A2').value = `DATA:     ${day} / ${month} / ${year}`;
  ws.getCell('A2').font = { bold: true };

  ws.getCell('C2').value = `PEDIDO: ${calculateJulianOrderNumber(today)}`;
  ws.getCell('C2').font = { bold: true };

  // Route 1 (Cols A & B)
  ws.getCell('A4').value = 'ROTA 01 - ZONA SUL';
  ws.getCell('A4').font = { bold: true, color: { argb: 'FF1E3A8A' } };
  ws.getCell('A5').value = 'LOJA';
  ws.getCell('B5').value = 'VOL';
  ws.getCell('A5').font = { bold: true };
  ws.getCell('B5').font = { bold: true };

  const r1Stores = ['0811', '1305', '1322'];
  r1Stores.forEach((store, idx) => {
    ws.getCell(`A${6 + idx}`).value = store;
    ws.getCell(`B${6 + idx}`).value = null;
  });
  ws.getCell(`A${6 + r1Stores.length}`).value = 'TOTAL';
  ws.getCell(`A${6 + r1Stores.length}`).font = { bold: true };
  ws.getCell(`B${6 + r1Stores.length}`).value = 0;
  ws.getCell(`B${6 + r1Stores.length}`).font = { bold: true };

  // Route 2 (Cols D & E)
  ws.getCell('D4').value = 'ROTA 02 - CENTRO';
  ws.getCell('D4').font = { bold: true, color: { argb: 'FF1E3A8A' } };
  ws.getCell('D5').value = 'LOJA';
  ws.getCell('E5').value = 'VOL';
  ws.getCell('D5').font = { bold: true };
  ws.getCell('E5').font = { bold: true };

  const r2Stores = ['1362', '1450', '1488'];
  r2Stores.forEach((store, idx) => {
    ws.getCell(`D${6 + idx}`).value = store;
    ws.getCell(`E${6 + idx}`).value = null;
  });
  ws.getCell(`D${6 + r2Stores.length}`).value = 'TOTAL';
  ws.getCell(`D${6 + r2Stores.length}`).font = { bold: true };
  ws.getCell(`E${6 + r2Stores.length}`).value = 0;
  ws.getCell(`E${6 + r2Stores.length}`).font = { bold: true };

  // Auto-width
  ws.columns.forEach((column) => {
    column.width = 16;
  });

  const buffer = await wb.xlsx.writeBuffer();
  return new File([buffer], 'PLANILHA_MODELO_DEMONSTRACAO.xlsx', {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}
