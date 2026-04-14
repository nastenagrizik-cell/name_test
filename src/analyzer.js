import { APP_CONFIG } from './config.js';
import { isScaleQuestion, letterFromIndex, normalizeText, percent, safeNumber, top2FromCounts } from './utils.js';

function countValues(rows, column) {
  const counts = new Map();
  let base = 0;
  for (const row of rows) {
    const value = row[column];
    if (value === '' || value === null || value === undefined) continue;
    base += 1;
    counts.set(value, (counts.get(value) || 0) + 1);
  }
  return { counts, base };
}

function zTestProportions(p1, n1, p2, n2) {
  if (!n1 || !n2) return { significant: false, z: null };
  const pooled = ((p1 * n1) + (p2 * n2)) / (n1 + n2);
  const se = Math.sqrt(pooled * (1 - pooled) * ((1 / n1) + (1 / n2)));
  if (!se) return { significant: false, z: null };
  const z = (p1 - p2) / se;
  return { significant: Math.abs(z) >= APP_CONFIG.zCriticalTwoTailed, z };
}

function inferQuestionLabel(columnName) {
  const cleaned = String(columnName).replace(/_/g, ' ').trim();
  return cleaned;
}

function buildScaleTables(rows, columns) {
  const result = [];
  columns.forEach(column => {
    const values = rows.map(r => r[column.header]).filter(v => v !== '' && v !== null && v !== undefined);
    if (!isScaleQuestion(values)) return;
    const { counts, base } = countValues(rows, column.header);
    const entries = [1, 2, 3, 4, 5].map(v => ({ label: String(v), value: percent(counts.get(v) || 0, base), base }));
    result.push({
      questionKey: column.header,
      questionLabel: inferQuestionLabel(column.header),
      base,
      top2: percent(top2FromCounts(counts, APP_CONFIG.top2ScaleValues), base),
      entries
    });
  });
  return result;
}

function compareAll(tables) {
  const comparisons = [];
  for (const table of tables) {
    const topRow = { rowLabel: `${table.questionLabel} — ТОП-2`, marks: {} };
    comparisons.push(topRow);
  }
  return comparisons;
}

function inferAudience(rows, columns) {
  return columns.map(col => {
    const { counts, base } = countValues(rows, col.header);
    const categories = [...counts.entries()].map(([label, count]) => ({ label: String(label), value: percent(count, base) }));
    return { questionLabel: col.header, base, categories };
  });
}

function groupByStem(scaleTables) {
  const groups = new Map();
  for (const table of scaleTables) {
    const stem = normalizeText(table.questionLabel.replace(/q\d+/i, '').trim()) || table.questionLabel;
    if (!groups.has(stem)) groups.set(stem, []);
    groups.get(stem).push(table);
  }
  return [...groups.entries()].map(([stem, items]) => ({ stem, items }));
}

function significanceByGroup(groups) {
  const sections = [];
  for (const group of groups) {
    if (group.items.length < 2) continue;
    const rowDefs = [
      { label: 'ТОП-2', getter: item => item.top2 },
      { label: '1', getter: item => item.entries.find(e => e.label === '1')?.value || 0 },
      { label: '2', getter: item => item.entries.find(e => e.label === '2')?.value || 0 },
      { label: '3', getter: item => item.entries.find(e => e.label === '3')?.value || 0 },
      { label: '4', getter: item => item.entries.find(e => e.label === '4')?.value || 0 },
      { label: '5', getter: item => item.entries.find(e => e.label === '5')?.value || 0 }
    ];
    const rows = rowDefs.map(def => {
      const values = group.items.map(item => ({
        concept: item.questionLabel,
        pct: def.getter(item),
        base: item.base
      }));
      const significance = values.map((current, idx) => {
        const higherThan = [];
        values.forEach((other, j) => {
          if (idx === j) return;
          const test = zTestProportions(current.pct, current.base, other.pct, other.base);
          if (test.significant && current.pct > other.pct) higherThan.push(letterFromIndex(j));
        });
        return higherThan;
      });
      return { rowLabel: def.label, values, significance };
    });
    sections.push({ stem: group.stem, concepts: group.items.map(i => i.questionLabel), rows });
  }
  return sections;
}

export function analyzeSurvey(parsed) {
  const scaleTables = buildScaleTables(parsed.rows, parsed.columns);
  const grouped = groupByStem(scaleTables);
  return {
    profile: {
      rowCount: parsed.rowCount,
      sheetNames: parsed.sheetNames,
      headers: parsed.headers.slice(0, 50)
    },
    summary: scaleTables.map(item => ({ questionLabel: item.questionLabel, values: [item.top2] })),
    fullTables: scaleTables,
    significance: significanceByGroup(grouped),
    audience: inferAudience(parsed.rows, parsed.audienceColumns),
    meta: {
      conceptGroups: parsed.conceptGroups,
      questionBlocks: parsed.questionBlocks,
      audienceColumns: parsed.audienceColumns.map(c => c.header)
    }
  };
}
