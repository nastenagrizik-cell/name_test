(() => {
  const state = { rows: [], headers: [], data: [], wb: null, analysis: null };
  const $ = id => document.getElementById(id);
  const statusEl = $('status');
  const fileEl = $('dataFile');
  const projectEl = $('projectName');
  const validateBtn = $('validateBtn');
  const generateBtn = $('generateBtn');
  const themeToggle = $('themeToggle');

  let theme = window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  document.documentElement.setAttribute('data-theme', theme);
  themeToggle?.addEventListener('click', () => {
    theme = theme === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', theme);
  });

  const normalize = v => String(v ?? '').replace(/\s+/g, ' ').trim();
  const lower = v => normalize(v).toLowerCase();
  const pct = (n, d) => d ? n / d : 0;
  const fmtPct = v => `${Math.round(v * 100)}%`;
  const alphaIndex = i => { const letters='ABCDEFGHIJKLMNOPQRSTUVWXYZ'; let n=i,s=''; do { s=letters[n%26]+s; n=Math.floor(n/26)-1; } while(n>=0); return s; };

  function setStatus(text, level = 'ok') {
    statusEl.textContent = text;
    statusEl.className = `status ${level}`;
  }

  function toObjects(rows, headers) {
    return rows.slice(1).filter(r => r.some(v => normalize(v) !== '')).map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = r[i]; });
      return obj;
    });
  }

  async function parseWorkbook(file) {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const firstSheet = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
    const headers = (rows[0] || []).map(normalize);
    return { wb, rows, headers, data: toObjects(rows, headers) };
  }

  function findHeader(headers, variants) {
    return headers.find(h => variants.some(v => lower(h) === lower(v) || lower(h).includes(lower(v))));
  }

  function detectBaseFields(headers) {
    return {
      mono: findHeader(headers, ['QMonadicGroup', 'MonadicGroup']),
      gender: findHeader(headers, ['Qgender', 'укажите ваш пол']),
      age: findHeader(headers, ['Qagerange', 'укажите ваш возраст']),
      citySize: findHeader(headers, ['Qrucitysize', 'размер населенного пункта']),
      city: findHeader(headers, ['Qrucity', 'в каком населенном пункте'])
    };
  }

  function parseMonadicConcepts(data, monoHeader) {
    if (!monoHeader) return [];
    const values = [...new Set(data.map(r => normalize(r[monoHeader])).filter(Boolean))];
    return values.sort((a, b) => Number(a || 0) - Number(b || 0)).map(v => ({ code: v, label: `Вариант ${v}` }));
  }

  function scoreFromRaw(value) {
    const s = lower(value);
    if (!s || s === 'nan') return null;
    if (s === 'checked') return 1;
    if (s === 'unchecked') return 0;
    const m = s.match(/(^|\s|,)([1-5])(\s|$|-|,)/);
    if (m) return Number(m[2]);
    if (/совсем не/i.test(s)) return 1;
    if (/скорее не/i.test(s)) return 2;
    if (/нейтр|затруд/i.test(s)) return 3;
    if (/скорее да|скорее нравится|подходит/i.test(s)) return 4;
    if (/очень|полностью|точно/i.test(s)) return 5;
    return null;
  }

  function classifyHeaders(headers) {
    const blocks = { image: [], direct: [] };
    headers.forEach(h => {
      const id = normalize(h).toLowerCase();
      if (/^q61r\d+$/.test(id)) blocks.image.push(h);
      if (/^(q47|q81|q90|q91)$/.test(id)) blocks.direct.push(h);
    });
    return blocks;
  }

  function groupedScaleBlocks(headers) {
    const ids = headers.map(h => normalize(h).toLowerCase());
    const groups = [
      { title: 'Нравится название', cols: ['q37r1','q37r2','q37r7','q37r3'] },
      { title: 'Подходит для блюда', cols: ['q46r1','q46r2','q46r7','q46r3'] },
      { title: 'Подходит для бренда', cols: ['q50r1','q50r2','q50r7','q50r3'] },
      { title: 'Намерение посетить БК', cols: ['q42r1','q42r2','q42r7','q42r3'] },
      { title: 'Намерение купить', cols: ['q87r1','q87r2','q87r7','q87r3'] }
    ];
    return groups.map(g => ({ title: g.title, cols: g.cols.filter(c => ids.includes(c)) })).filter(g => g.cols.length > 0);
  }

  function ztest(p1, n1, p2, n2) {
    if (!n1 || !n2) return 0;
    const p = (p1 * n1 + p2 * n2) / (n1 + n2);
    const se = Math.sqrt(p * (1 - p) * (1 / n1 + 1 / n2));
    return se ? (p1 - p2) / se : 0;
  }

  function summarizeMonadic(data, monoHeader, concepts, blocks) {
    const result = { concepts, summary: [], full: [], sig: [] };
    blocks.forEach(block => {
      const row = { title: block.title, values: [] };
      const sigRow = { title: `ТОП-2 ${block.title}`, values: [] };
      const fullSection = { title: block.title.toUpperCase(), rows: [['ТОП-2 (4+5)'], ['1'], ['2'], ['3'], ['4'], ['5']] };
      const distStore = [];

      concepts.forEach((concept, idx) => {
        const subset = data.filter(r => normalize(r[monoHeader]) === concept.code);
        const col = block.cols[Math.min(idx, block.cols.length - 1)];
        const counts = { 1:0, 2:0, 3:0, 4:0, 5:0 };
        let base = 0;

        subset.forEach(r => {
          const s = scoreFromRaw(r[col]);
          if (s && counts[s] !== undefined) {
            counts[s] += 1;
            base += 1;
          }
        });

        const top2 = pct(counts[4] + counts[5], base);
        row.values.push(fmtPct(top2));
        distStore.push({ value: top2, base });
        fullSection.rows[0].push(fmtPct(top2));
        [1,2,3,4,5].forEach((v, pos) => fullSection.rows[pos + 1].push(fmtPct(pct(counts[v], base))));
      });

      distStore.forEach((cur, i) => {
        const letters = [];
        distStore.forEach((other, j) => {
          if (i === j) return;
          const z = ztest(cur.value, cur.base, other.value, other.base);
          if (z > 1.96) letters.push(alphaIndex(j));
        });
        sigRow.values.push(`${fmtPct(cur.value)}${letters.length ? ' > ' + letters.join(',') : ''}`);
      });

      result.summary.push(row);
      result.full.push(fullSection);
      result.sig.push(sigRow);
    });
    return result;
  }

  function audienceRows(data, fields) {
    const out = [];
    [['Пол', fields.gender], ['Возраст', fields.age], ['Гео', fields.city], ['Размер города', fields.citySize]].forEach(([title, header]) => {
      if (!header) return;
      const counts = new Map();
      data.forEach(r => {
        const key = normalize(r[header]) || '(пусто)';
        counts.set(key, (counts.get(key) || 0) + 1);
      });
      [...counts.entries()].sort((a, b) => b[1] - a[1]).forEach(([label, count]) => {
        out.push([title, label, fmtPct(pct(count, data.length))]);
      });
    });
    return out;
  }

  async function runValidation() {
    const file = fileEl.files?.[0];
    if (!file) {
      setStatus('Сначала загрузите файл базы.', 'err');
      return;
    }

    try {
      const parsed = await parseWorkbook(file);
      state.rows = parsed.rows;
      state.headers = parsed.headers;
      state.data = parsed.data;
      state.wb = parsed.wb;

      const fields = detectBaseFields(state.headers);
      const concepts = parseMonadicConcepts(state.data, fields.mono);
      const scales = groupedScaleBlocks(state.headers);
      const blocks = classifyHeaders(state.headers);
      const warnings = [];

      if (!fields.mono) warnings.push('Не найдена monadic-переменная.');
      if (concepts.length < 2) warnings.push('Найдено меньше двух вариантов названий.');
      if (scales.length < 3) warnings.push('Распознано мало шкальных блоков.');
      if (!fields.gender || !fields.age) warnings.push('Демографический блок найден не полностью.');

      state.analysis = { fields, concepts, scales, blocks, warnings };

      const text = [
        `Файл: ${file.name}`,
        `Интервью: ${state.data.length}`,
        `Колонки: ${state.headers.length}`,
        `Monadic-переменная: ${fields.mono || 'не найдена'}`,
        `Число вариантов: ${concepts.length}`,
        `Шкальные блоки: ${scales.length}`,
        `Имиджевый блок: ${blocks.image.length ? 'найден' : 'не найден'}`,
        `Прямое сравнение: ${blocks.direct.length ? 'найдено' : 'не найдено'}`,
        `Аудитория: ${fields.gender || fields.age || fields.city || fields.citySize ? 'найдена частично/полностью' : 'не найдена'}`,
        warnings.length ? '' : 'Все ключевые блоки для сборки распознаны.',
        ...warnings.map(w => `Предупреждение: ${w}`)
      ].join('\n');

      setStatus(text, warnings.length ? 'warn' : 'ok');
    } catch (e) {
      console.error(e);
      setStatus(`Ошибка чтения файла: ${e.message}`, 'err');
    }
  }

  const BORDER = { top:{style:'thin',color:{argb:'FFE5E7EB'}}, left:{style:'thin',color:{argb:'FFE5E7EB'}}, bottom:{style:'thin',color:{argb:'FFE5E7EB'}}, right:{style:'thin',color:{argb:'FFE5E7EB'}} };
  function styleHeader(row, fill) {
    row.eachCell(cell => {
      cell.font = { name:'Calibri', size:11, bold:true, color:{argb:'FFFFFFFF'} };
      cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:fill} };
      cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
      cell.border = BORDER;
    });
  }
  function styleCell(cell, bold=false) {
    cell.font = { name:'Calibri', size:11, bold };
    cell.border = BORDER;
    cell.alignment = { vertical:'middle', horizontal:'left', wrapText:true };
  }

  async function generateWorkbook() {
    if (!state.analysis) {
      await runValidation();
      if (!state.analysis) return;
    }

    const { fields, concepts, scales } = state.analysis;
    if (!fields.mono || !concepts.length) {
      setStatus('База не распознана. Сначала выполните проверку.', 'err');
      return;
    }

    const project = (projectEl.value || 'Тест названий').trim();
    const analysis = summarizeMonadic(state.data, fields.mono, concepts, scales);
    analysis.audience = audienceRows(state.data, fields);

    const wb = new ExcelJS.Workbook();
    wb.creator = 'Perplexity';
    const names = concepts.map(c => c.label);
    const endCol = 1 + names.length;

    const s = wb.addWorksheet('САММАРИ');
    s.columns = [{ width: 42 }, ...names.map(() => ({ width: 18 }))];
    s.addRow([project]);
    s.mergeCells(1,1,1,endCol);
    styleCell(s.getCell('A1'), true);
    s.getCell('A1').font = { name:'Calibri', size:14, bold:true };
    s.addRow(['Варианты', ...names]);
    styleHeader(s.getRow(2), '0F766E');
    s.addRow([`База: n=${state.data.length}`, ...names.map(() => state.data.length)]);
    s.addRow([]);
    s.addRow(['ОСНОВНЫЕ ПОКАЗАТЕЛИ']);
    styleCell(s.getCell(`A${s.rowCount}`), true);
    analysis.summary.forEach(r => s.addRow([r.title, ...r.values]));
    s.eachRow((row, i) => row.eachCell(cell => styleCell(cell, i === 1)));

    const p = wb.addWorksheet('полные таблицы');
    p.columns = [{ width: 42 }, ...names.map(() => ({ width: 18 }))];
    p.addRow(['ПОЛНЫЕ ТАБЛИЦЫ']);
    p.mergeCells(1,1,1,endCol);
    styleCell(p.getCell('A1'), true);
    p.getCell('A1').font = { name:'Calibri', size:14, bold:true };
    p.addRow(['Варианты', ...names]);
    styleHeader(p.getRow(2), '0F766E');
    analysis.full.forEach(section => {
      p.addRow([section.title]);
      styleCell(p.getCell(`A${p.rowCount}`), true);
      section.rows.forEach(r => p.addRow(r));
      p.addRow([]);
    });
    p.eachRow((row, i) => row.eachCell(cell => styleCell(cell, i === 1)));

    const z = wb.addWorksheet('значимости');
    z.columns = [{ width: 42 }, ...names.map(() => ({ width: 20 }))];
    z.addRow(['ЗНАЧИМОСТИ']);
    z.mergeCells(1,1,1,endCol);
    styleCell(z.getCell('A1'), true);
    z.getCell('A1').font = { name:'Calibri', size:14, bold:true };
    z.addRow(['Варианты', ...names.map((n, i) => `${n} (${alphaIndex(i)})`)]);
    styleHeader(z.getRow(2), '7C3AED');
    analysis.sig.forEach(r => z.addRow([r.title, ...r.values]));
    z.eachRow((row, i) => row.eachCell(cell => styleCell(cell, i === 1)));

    const a = wb.addWorksheet('Аудитория');
    a.columns = [{ width: 28 }, { width: 34 }, { width: 12 }];
    a.addRow(['АУДИТОРИЯ']);
    a.mergeCells('A1:C1');
    styleCell(a.getCell('A1'), true);
    a.getCell('A1').font = { name:'Calibri', size:14, bold:true };
    a.addRow(['Срез', 'Категория', '%']);
    styleHeader(a.getRow(2), 'B45309');
    analysis.audience.forEach(r => a.addRow(r));
    a.eachRow((row, i) => row.eachCell(cell => styleCell(cell, i === 1)));

    const buffer = await wb.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const link = document.createElement('a');
    const safeName = project.replace(/\s+/g, '_');
    link.href = URL.createObjectURL(blob);
    link.download = `${safeName}_topline_v4.xlsx`;
    document.body.appendChild(link);
    link.click();
    setTimeout(() => {
      URL.revokeObjectURL(link.href);
      link.remove();
    }, 1000);

    setStatus('Итоговый XLSX собран. Если цифры выглядят неполными, сначала нажмите «Проверить базу» и посмотрите, все ли блоки распознаны.', 'ok');
  }

  validateBtn?.addEventListener('click', () => { runValidation(); });
  generateBtn?.addEventListener('click', () => { generateWorkbook().catch(e => { console.error(e); setStatus(`Ошибка сборки: ${e.message}`, 'err'); }); });
})();
