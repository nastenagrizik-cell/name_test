export function normalizeText(value) {
  return String(value ?? '')
    .replace(/\s+/g, ' ')
    .replace(/[ё]/gi, 'е')
    .trim()
    .toLowerCase();
}

export function slugify(value) {
  return String(value ?? 'project')
    .trim()
    .toLowerCase()
    .replace(/[^a-zа-я0-9]+/gi, '-')
    .replace(/^-+|-+$/g, '') || 'project';
}

export function unique(arr) {
  return [...new Set(arr)];
}

export function letterFromIndex(index) {
  return String.fromCharCode(65 + index);
}

export function safeNumber(value) {
  if (value === null || value === undefined || value === '') return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

export function percent(value, base) {
  if (!base) return 0;
  return value / base;
}

export function top2FromCounts(countMap, topValues) {
  return topValues.reduce((sum, key) => sum + (countMap.get(key) || 0), 0);
}

export function isScaleQuestion(values) {
  const nums = unique(values.map(v => safeNumber(v)).filter(v => v !== null)).sort((a, b) => a - b);
  if (!nums.length) return false;
  return nums.every(n => Number.isInteger(n) && n >= 1 && n <= 5);
}
