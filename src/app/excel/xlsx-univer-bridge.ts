import * as XLSX from 'xlsx';

// Convert xlsx binary data into univer snapshot structure
export function xlsxToWorkbookData(buffer: ArrayBuffer, id: string): any {
  const wb = XLSX.read(buffer, { type: 'array' });

  const sheets: Record<string, any> = {};
  const sheetOrder: string[] = [];

  for (let i = 0; i < wb.SheetNames.length; i++) {
    const sheetName = wb.SheetNames[i];
    const sheetId = `sheet-${i}`;
    const ws = wb.Sheets[sheetName];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');

    const cellData: Record<number, Record<number, any>> = {};

    for (let r = range.s.r; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cellRef = XLSX.utils.encode_cell({ r, c });
        const cell = ws[cellRef];
        if (!cell) continue;

        if (!cellData[r]) {
          cellData[r] = {};
        }

        cellData[r][c] = {
          v: cell.v,
          f: cell.f ? `=${cell.f}` : undefined,
        };
      }
    }

    sheets[sheetId] = {
      id: sheetId,
      name: sheetName,
      rowCount: Math.max(range.e.r + 1, 100),
      columnCount: Math.max(range.e.c + 1, 20),
      cellData,
    };
    sheetOrder.push(sheetId);
  }

  return {
    id,
    name: id,
    sheetOrder,
    sheets,
    styles: {},
  };
}

// Convert univer snapshot back into xlsx ArrayBuffer
export function workbookDataToXlsxArrayBuffer(univerData: any): ArrayBuffer {
  const wb = XLSX.utils.book_new();
  const order = univerData?.sheetOrder || [];

  for (const sheetId of order) {
    const sheet = univerData.sheets?.[sheetId];
    if (!sheet) continue;

    const ws: XLSX.WorkSheet = {};
    let maxRow = 0;
    let maxCol = 0;

    const cellData = sheet.cellData || {};
    const rowKeys = Object.keys(cellData);

    for (const rKey of rowKeys) {
      const r = Number(rKey);
      const rowObj = cellData[rKey];
      const colKeys = Object.keys(rowObj);

      for (const cKey of colKeys) {
        const c = Number(cKey);
        const cell = rowObj[cKey];
        const cellRef = XLSX.utils.encode_cell({ r, c });

        if (cell.f) {
          const cleanFormula = cell.f.startsWith('=') ? cell.f.substring(1) : cell.f;
          ws[cellRef] = { f: cleanFormula, v: cell.v };
        } else {
          ws[cellRef] = { v: cell.v };
        }

        if (r > maxRow) maxRow = r;
        if (c > maxCol) maxCol = c;
      }
    }

    ws['!ref'] = XLSX.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: maxRow, c: maxCol },
    });

    XLSX.utils.book_append_sheet(wb, ws, sheet.name);
  }

  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
}
