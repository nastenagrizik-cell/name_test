import XLSX from 'xlsx';
import { normalizeText, unique } from './utils.js';

export function loadWorkbook(buffer) {
  return XLSX.read(buffer, { type: 'buffer' });
}

export function firstSheetRows(workbook) {
  const firstName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstName];
  return XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true });
}

export function inferColumns(rows) {
  const headers = rows.length ? Object.keys(rows[0]) : [];
  return headers.map((header, index) => ({
    header,
    index,
    normalized: normalizeText(header)
  }));
}

export function inferConceptGroups(columns) {
  const conceptMap = new Map();
  columns.forEach(col => {
    const match = col.header.match(/^(Q\d+_\d+|Q\d+|q\d+)(?:_(\d+))?/i);
    if (!match) return;
    const prefix = (match[1] || '').toUpperCase();
    if (!conceptMap.has(prefix)) conceptMap.set(prefix, []);
    conceptMap.get(prefix).push(col.header);
  });
  return [...conceptMap.entries()].map(([prefix, cols]) => ({ prefix, columns: cols }));
}

export function inferQuestionBlocks(columns) {
  const blocks = [];
  const grouped = new Map();
  columns.forEach(col => {
    const key = col.header.split('_')[0];
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(col.header);
  });
  for (const [key, cols] of grouped.entries()) {
    blocks.push({ key, columns: cols });
  }
  return blocks;
}

export function inferAudienceColumns(columns) {
  return columns.filter(col => /пол|возраст|частота|новин|ресторан|посещ|fast food|bk/i.test(col.normalized));
}

export function parseSurvey(buffer) {
  const workbook = loadWorkbook(buffer);
  const rows = firstSheetRows(workbook);
  const columns = inferColumns(rows);
  return {
    workbook,
    rows,
    columns,
    questionBlocks: inferQuestionBlocks(columns),
    conceptGroups: inferConceptGroups(columns),
    audienceColumns: inferAudienceColumns(columns),
    sheetNames: workbook.SheetNames,
    rowCount: rows.length,
    headers: unique(columns.map(c => c.header))
  };
}
