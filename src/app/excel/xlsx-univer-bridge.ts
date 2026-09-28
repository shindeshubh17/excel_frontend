import * as ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

function parseColor(color: any): string | undefined {
  if (!color || typeof color !== 'object') return undefined;
  if (color.argb && typeof color.argb === 'string') {
    const raw = color.argb.trim();
    return '#' + raw.slice(-6);
  }
  return undefined;
}

// Convert xlsx binary data into univer snapshot structure
export async function xlsxToWorkbookData(buffer: ArrayBuffer, id: string): Promise<any> {
  try {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    const sheets: Record<string, any> = {};
    const sheetOrder: string[] = [];

    wb.eachSheet((ws, sheetIdNum) => {
      const sheetId = `sheet-${sheetIdNum - 1}`;
      const cellData: Record<number, Record<number, any>> = {};

      let maxRow = ws.rowCount || 0;
      let maxCol = ws.columnCount || 0;

      ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const r = rowNumber - 1; // 0-based for Univer

        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const c = colNumber - 1; // 0-based for Univer

          // Extract value and formula
          let val: any = undefined;
          let formula: string | undefined = undefined;

          if (cell.formula) {
            formula = cell.formula.startsWith('=') ? cell.formula : `=${cell.formula}`;
            // If formula has a pre-calculated result, keep it
            val = cell.result !== undefined ? cell.result : (cell.value && typeof cell.value === 'object' && 'result' in cell.value ? (cell.value as any).result : undefined);
          } else if (cell.value !== null && cell.value !== undefined) {
            if (typeof cell.value === 'object') {
              if ('richText' in cell.value && Array.isArray((cell.value as any).richText)) {
                val = (cell.value as any).richText.map((t: any) => t.text || '').join('');
              } else if ('text' in cell.value) {
                val = (cell.value as any).text;
              } else {
                val = cell.text || '';
              }
            } else {
              val = cell.value;
            }
          }

          // Extract style: bold, italic, text color, background color
          const style: any = {};
          if (cell.font) {
            if (cell.font.bold) style.bl = 1;
            if (cell.font.italic) style.it = 1;
            const fontColor = parseColor(cell.font.color);
            if (fontColor) style.cl = { rgb: fontColor };
          }

          if (cell.fill && cell.fill.type === 'pattern') {
            const patternFill = cell.fill as any;
            const bgColor = parseColor(patternFill.fgColor || patternFill.bgColor);
            if (bgColor) style.bg = { rgb: bgColor };
          }

          const hasStyle = Object.keys(style).length > 0;

          // Only set cell if it contains a value, formula, or style
          if (val !== undefined || formula || hasStyle) {
            if (!cellData[r]) {
              cellData[r] = {};
            }
            const cellItem: any = {};
            if (val !== undefined) cellItem.v = val;
            if (formula) cellItem.f = formula;
            if (hasStyle) cellItem.s = style;

            cellData[r][c] = cellItem;
          }

          if (r >= maxRow) maxRow = r + 1;
          if (c >= maxCol) maxCol = c + 1;
        });
      });

      sheets[sheetId] = {
        id: sheetId,
        name: ws.name,
        rowCount: Math.max(maxRow, 100),
        columnCount: Math.max(maxCol, 20),
        cellData,
      };
      sheetOrder.push(sheetId);
    });

    if (sheetOrder.length > 0) {
      return {
        id,
        name: id,
        sheetOrder,
        sheets,
        styles: {},
      };
    }
  } catch (err) {
    console.warn('ExcelJS failed to parse, falling back to XLSX:', err);
  }

  // Fallback to basic SheetJS parsing if ExcelJS fails or produces no sheets
  return fallbackXlsxToWorkbookData(buffer, id);
}

// Convert univer snapshot back into xlsx ArrayBuffer
export async function workbookDataToXlsxArrayBuffer(univerData: any): Promise<ArrayBuffer> {
  try {
    const wb = new ExcelJS.Workbook();
    const order = univerData?.sheetOrder || [];

    for (const sheetId of order) {
      const sheet = univerData.sheets?.[sheetId];
      if (!sheet) continue;

      const ws = wb.addWorksheet(sheet.name || 'Sheet');
      const cellData = sheet.cellData || {};

      for (const rKey of Object.keys(cellData)) {
        const r = Number(rKey);
        const rowObj = cellData[rKey];
        if (!rowObj) continue;

        for (const cKey of Object.keys(rowObj)) {
          const c = Number(cKey);
          const cell = rowObj[cKey];
          if (!cell) continue;

          // ExcelJS uses 1-based indexing for rows and columns
          const excelCell = ws.getCell(r + 1, c + 1);

          // Handle formula and value
          if (cell.f) {
            const cleanFormula = cell.f.startsWith('=') ? cell.f.substring(1) : cell.f;
            excelCell.value = {
              formula: cleanFormula,
              result: cell.v !== undefined ? cell.v : undefined,
            };
          } else if (cell.v !== undefined) {
            excelCell.value = cell.v;
          }

          // Handle styling: bold, italic, text color, background color
          const s = typeof cell.s === 'string' ? univerData.styles?.[cell.s] : cell.s;
          if (s) {
            if (s.bl || s.it || s.cl?.rgb) {
              excelCell.font = {
                bold: !!s.bl,
                italic: !!s.it,
                color: s.cl?.rgb ? { argb: 'FF' + s.cl.rgb.replace('#', '') } : undefined,
              };
            }
            if (s.bg?.rgb) {
              excelCell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FF' + s.bg.rgb.replace('#', '') },
              };
            }
          }
        }
      }
    }

    const buffer = await wb.xlsx.writeBuffer();
    return buffer as ArrayBuffer;
  } catch (err) {
    console.warn('ExcelJS failed to write, falling back to XLSX:', err);
    return fallbackWorkbookDataToXlsxArrayBuffer(univerData);
  }
}

// Convert xlsx binary data into univer snapshot using SheetJS
function fallbackXlsxToWorkbookData(buffer: ArrayBuffer, id: string): any {
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

function fallbackWorkbookDataToXlsxArrayBuffer(univerData: any): ArrayBuffer {
  const wb = XLSX.utils.book_new();
  const order = univerData?.sheetOrder || [];

  for (const sheetId of order) {
    const sheet = univerData.sheets?.[sheetId];
    if (!sheet) continue;

    const ws: XLSX.WorkSheet = {};
    let maxRow = 0;
    let maxCol = 0;
    const cellData = sheet.cellData || {};

    for (const rKey of Object.keys(cellData)) {
      const r = Number(rKey);
      const rowObj = cellData[rKey];
      for (const cKey of Object.keys(rowObj)) {
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
