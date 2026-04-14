import ExcelJS from 'exceljs';

const HDR_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: '01696F' } };
const SUB_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EAF3F2' } };
const BORDER = {
  top: { style: 'thin', color: { argb: 'D4D1CA' } },
  left: { style: 'thin', color: { argb: 'D4D1CA' } },
  bottom: { style: 'thin', color: { argb: 'D4D1CA' } },
  right: { style: 'thin', color: { argb: 'D4D1CA' } }
};

function applyHeader(ws, range, title) {
  ws.mergeCells(range);
  const cell = ws.getCell(range.split(':')[0]);
  cell.value = title;
  cell.fill = HDR_FILL;
  cell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFF' } };
  cell.alignment = { horizontal: 'left', vertical: 'middle' };
}

function styleGrid(ws, startRow, endRow, endCol, pctStartCol = 2) {
  for (let r = startRow; r <= endRow; r++) {
    for (let c = 1; c <= endCol; c++) {
      const cell = ws.getRow(r).getCell(c);
      cell.border = BORDER;
      cell.font = { name: 'Calibri', size: 11, bold: r === startRow };
      if (r === startRow) cell.fill = SUB_FILL;
      if (r > startRow && c >= pctStartCol && typeof cell.value === 'number') cell.numFmt = '0%';
      cell.alignment = { vertical: 'middle', wrapText: true };
    }
  }
}

export async function buildWorkbook(report, fileName) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Perplexity';
  wb.created = new Date();

  const summary = wb.addWorksheet('САММАРИ');
  const full = wb.addWorksheet('Полные таблицы');
  const sig = wb.addWorksheet('Значимости');
  const audience = wb.addWorksheet('Аудитория');

  [summary, full, sig].forEach(ws => {
    ws.columns = [{ width: 42 }, { width: 18 }, { width: 18 }, { width: 18 }, { width: 18 }, { width: 18 }, { width: 18 }];
    ws.views = [{ state: 'frozen', ySplit: 2 }];
  });
  audience.columns = [{ width: 34 }, { width: 42 }, { width: 14 }, { width: 14 }];
  audience.views = [{ state: 'frozen', ySplit: 2 }];

  applyHeader(summary, 'A1:G1', `САММАРИ | ${fileName}`);
  summary.addRow(['Показатель', 'ТОП-2']);
  report.summary.forEach(item => summary.addRow([item.questionLabel, item.values[0]]));
  styleGrid(summary, 2, summary.rowCount, 2, 2);

  applyHeader(full, 'A1:G1', 'ПОЛНЫЕ ТАБЛИЦЫ');
  let fullRow = 2;
  for (const table of report.fullTables) {
    full.getCell(`A${fullRow}`).value = table.questionLabel;
    full.getCell(`A${fullRow}`).font = { name: 'Calibri', bold: true, size: 11 };
    fullRow += 1;
    full.addRow(['Ответ', 'Доля']);
    full.addRow(['ТОП-2', table.top2]);
    table.entries.forEach(entry => full.addRow([entry.label, entry.value]));
    styleGrid(full, fullRow, fullRow + 1 + table.entries.length, 2, 2);
    fullRow = full.rowCount + 2;
  }

  applyHeader(sig, 'A1:G1', 'ЗНАЧИМОСТИ');
  let sigRow = 2;
  for (const section of report.significance) {
    sig.getCell(`A${sigRow}`).value = section.stem || 'Блок';
    sig.getCell(`A${sigRow}`).font = { name: 'Calibri', bold: true, size: 11 };
    sigRow += 1;
    sig.addRow(['Вариант ответа', ...section.concepts.map((_, i) => String.fromCharCode(65 + i))]);
    section.rows.forEach(row => {
      const vals = row.values.map((v, i) => row.significance[i].length ? `${Math.round(v.pct * 100)}% ${row.significance[i].join(',')}` : v.pct);
      sig.addRow([row.rowLabel, ...vals]);
    });
    styleGrid(sig, sigRow, sig.rowCount, 1 + section.concepts.length, 99);
    sigRow = sig.rowCount + 2;
  }

  applyHeader(audience, 'A1:D1', 'АУДИТОРИЯ');
  audience.addRow(['Вопрос', 'Категория', 'Доля', 'База']);
  report.audience.forEach(block => {
    block.categories.forEach(cat => audience.addRow([block.questionLabel, cat.label, cat.value, block.base]));
  });
  styleGrid(audience, 2, audience.rowCount, 4, 3);
  for (let r = 3; r <= audience.rowCount; r++) audience.getRow(r).getCell(4).numFmt = '#,##0';

  return wb.xlsx.writeBuffer();
}
