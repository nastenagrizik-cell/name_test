// app.js

const baseInput = document.getElementById('baseFile');
const statusEl = document.getElementById('status');
const mappingSection = document.getElementById('mappingSection');
const extraSection = document.getElementById('extraSection');
const runSection = document.getElementById('runSection');
const standardGroupsEl = document.getElementById('standardGroups');
const extraQuestionsEl = document.getElementById('extraQuestions');
const runBtn = document.getElementById('runBtn');

let baseFile;
let parsed = null;
let autoMapping = null;
let userConfig = null;

if (typeof XLSX !== 'object') {
  if (statusEl) {
    statusEl.textContent = 'Ошибка: библиотека XLSX не загружена. Попробуйте обновить страницу.';
    statusEl.className = 'status error';
  }
} else if (!baseInput) {
  if (statusEl) {
    statusEl.textContent = 'Ошибка: на странице не найден input с id="baseFile".';
    statusEl.className = 'status error';
  }
} else {
  baseInput.addEventListener('change', async e => {
    baseFile = e.target.files[0] || null;
    resetState();

    if (!baseFile) {
      status('Файл не выбран');
      return;
    }

    status('Читаю базу...\nЭто может занять до минуты.');

    try {
      const arrayBuffer = await baseFile.arrayBuffer();
      const wb = XLSX.read(arrayBuffer, { type: 'array' });

      const sheetName = wb.SheetNames[wb.SheetNames.length - 1];
      const sheet = wb.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

      const { header, rows } = splitHeaderRows(data);
      parsed = { header, rows };

      autoMapping = autoDetectMapping(header);
      renderStandardMappingUI(autoMapping, header);
      renderExtraQuestionsUI(autoMapping, header);

      mappingSection.style.display = '';
      extraSection.style.display = '';
      runSection.style.display = '';

      status(
        'База загружена. Проверьте найденные вопросы и доп.метрики, затем нажмите «Посчитать топлайн».',
        true
      );
    } catch (e) {
      console.error(e);
      status('Ошибка при чтении файла: ' + (e && e.message ? e.message : String(e)), false, true);
    }
  });

  runBtn.addEventListener('click', () => {
    if (!parsed || !autoMapping) {
      status('Сначала загрузите файл и дождитесь определения вопросов.', false, true);
      return;
    }

    try {
      userConfig = collectUserConfig(autoMapping);
    } catch (e) {
      status('Нужно завершить настройку вопросов: ' + e.message, false, true);
      return;
    }

    try {
      runBtn.disabled = true;
      status('Считаю топлайн...\nПодождите, формируется Excel.');

      const { header, rows } = parsed;
      const concepts = inferConcepts(header, userConfig);

      const stdResults = calcStandardBlocks(rows, userConfig, concepts, header);
      const extraResults = calcExtraBlocks(rows, userConfig);
      const audienceRes = calcAudience(rows, userConfig);
      const signifRes = calcSignificance(stdResults, concepts, rows.length);

      const outWb = XLSX.utils.book_new();

      XLSX.utils.book_append_sheet(outWb, makeSummarySheetStyled(stdResults, concepts, signifRes), 'САММАРИ');
      XLSX.utils.book_append_sheet(outWb, makeFullSheetStyled(stdResults, concepts), 'полные таблицы');
      XLSX.utils.book_append_sheet(outWb, makeSignifSheetStyled(stdResToSignifInput(stdResults), concepts, signifRes), 'значимости');
      XLSX.utils.book_append_sheet(outWb, makeAudienceSheetStyled(audienceRes), 'Аудитория');

      const outName = 'Topline_' + (baseFile.name.replace(/\.[^.]+$/, '') || 'output') + '.xlsx';
      XLSX.writeFile(outWb, outName);

      status('Готово. Файл ' + outName + ' сохранён.', true);
    } catch (e) {
      console.error(e);
      status('Ошибка при расчете: ' + (e && e.message ? e.message : String(e)), false, true);
    } finally {
      runBtn.disabled = false;
    }
  });
}

function resetState() {
  parsed = null;
  autoMapping = null;
  userConfig = null;

  if (mappingSection) mappingSection.style.display = 'none';
  if (extraSection) extraSection.style.display = 'none';
  if (runSection) runSection.style.display = 'none';
  if (standardGroupsEl) standardGroupsEl.innerHTML = '';
  if (extraQuestionsEl) extraQuestionsEl.innerHTML = '';
}

function status(text, ok = false, isError = false) {
  if (!statusEl) return;
  statusEl.textContent = text;
  statusEl.classList.toggle('ok', ok);
  statusEl.classList.toggle('error', isError);
}

function splitHeaderRows(data) {
  if (!data || data.length < 2) return { header: [], rows: [] };

  const varNames = (data[0] || []).map(v => String(v || '').trim());
  const questionTexts = (data[1] || []).map(v => String(v || '').trim());
  const header = questionTexts.map((txt, i) => txt || varNames[i] || `col_${i}`);
  const rows = data.slice(2).filter(r => r && r.some(v => v !== null && v !== ''));

  return { header, rows };
}

function autoDetectMapping(header) {
  const std = {
    like: [],
    fitDish: [],
    fitBrand: [],
    visitBK: [],
    buyDish: [],
    image: [],
    directLike: [],
    directBuy: [],
    audience: {
      sex: null,
      age: null,
      freqNew: null,
      freqProd: null,
      freqBK: null
    }
  };

  header.forEach((h, idx) => {
    if (h.includes('Оцените, пожалуйста, насколько вам нравится или не нравится каждое из этих названий')) {
      std.like.push(idx);
    }
    if (h.includes('Насколько каждое из этих названий') && h.includes('подходит или не подходит для этого')) {
      std.fitDish.push(idx);
    }
    if (h.includes('А теперь оцените, насколько каждое из этих названий подходит или не подходит для бренда Бургер Кинг')) {
      std.fitBrand.push(idx);
    }
    if (h.includes('Скажите, насколько вероятно, что Вы посетите ресторан Бургер Кинг')) {
      std.visitBK.push(idx);
    }
    if (h.includes('Для каждого названия укажите, насколько вероятно, что Вы купите')) {
      std.buyDish.push(idx);
    }

    if (h.includes('Понятное и простое название - ')) std.image.push({ key: 'Понятное и простое название', idx });
    if (h.includes('Вызывает аппетит')) std.image.push({ key: 'Вызывает аппетит, звучит вкусно', idx });
    if (h.startsWith('Вызывает у меня доверие')) std.image.push({ key: 'Вызывает доверие', idx });
    if (h.startsWith('Добавляет премиальности')) std.image.push({ key: 'Добавляет премиальности', idx });
    if (h.startsWith('Хочется попробовать такой капучино')) std.image.push({ key: 'Хочется попробовать такой капучино', idx });
    if (h.startsWith('Понятно какой будет вкус')) std.image.push({ key: 'Понятно какой будет вкус', idx });
    if (h.startsWith('Звучит как качественный продукт')) std.image.push({ key: 'Звучит как качественный продукт', idx });
    if (h.startsWith('Оригинальное, отличается от других')) std.image.push({ key: 'Оригинальное, отличается от других', idx });
    if (h.startsWith('Хочется рассказать/поделиться в социальных сетях')) std.image.push({ key: 'Хочется рассказать/поделиться в соц.сетях', idx });

    if (h.includes('Какое из перечисленных ниже названий')) std.directLike.push(idx);
    if (h.includes('С каким из этих названий вы бы купили')) std.directBuy.push(idx);

    if (h.includes('Укажите Ваш пол')) std.audience.sex = idx;
    if (h.includes('Укажите Ваш возраст')) std.audience.age = idx;
    if (h.includes('Как часто Вы берете новинки в категории горячих напитков')) std.audience.freqNew = idx;
    if (h.includes('Как часто вы покупаете капучино')) std.audience.freqProd = idx;
    if (h.includes('Как часто вы посещаете Бургер Кинг')) std.audience.freqBK = idx;
  });

  const used = new Set([
    ...std.like, ...std.fitDish, ...std.fitBrand, ...std.visitBK, ...std.buyDish,
    ...std.image.map(x => x.idx), ...std.directLike, ...std.directBuy,
    std.audience.sex, std.audience.age, std.audience.freqNew, std.audience.freqProd, std.audience.freqBK
  ].filter(v => v !== null && v !== undefined));

  const extraCandidates = header.map((h, idx) => {
    if (used.has(idx)) return null;
    if (!h) return null;
    const lower = h.toLowerCase();
    const looksClosed =
      lower.includes('насколько') ||
      lower.includes('оцените') ||
      lower.includes('выберите') ||
      lower.includes('какое из перечисленных');
    if (!looksClosed) return null;
    return { idx, header: h };
  }).filter(Boolean);

  return { std, extraCandidates };
}

function renderStandardMappingUI(mapping, header) {
  const groups = [
    { key: 'like', label: 'Нравится название (шкала 1–5, Top‑2)', indexes: mapping.std.like },
    { key: 'fitDish', label: 'Подходит для блюда (шкала 1–5, Top‑2)', indexes: mapping.std.fitDish },
    { key: 'fitBrand', label: 'Подходит для бренда (шкала 1–5, Top‑2)', indexes: mapping.std.fitBrand },
    { key: 'visitBK', label: 'Намерение посетить БК (шкала 1–5, Top‑2)', indexes: mapping.std.visitBK },
    { key: 'buyDish', label: 'Намерение купить (шкала 1–5, Top‑2)', indexes: mapping.std.buyDish },
    { key: 'directCompare', label: 'Прямое сравнение', indexes: [...mapping.std.directLike, ...mapping.std.directBuy] }
  ];

  standardGroupsEl.innerHTML = '';

  groups.forEach(group => {
    const col = document.createElement('div');
    col.className = 'col-half mapping-group';

    const title = document.createElement('div');
    title.className = 'mapping-group-title';
    title.textContent = group.label;
    col.appendChild(title);

    const list = document.createElement('div');
    list.className = 'mapping-list';

    if (!group.indexes.length) {
      const empty = document.createElement('div');
      empty.className = 'mapping-item';
      empty.innerHTML = '<small>Колонки не найдены по ключевым словам</small>';
      list.appendChild(empty);
    } else {
      group.indexes.forEach(idx => {
        const item = document.createElement('div');
        item.className = 'mapping-item';
        const id = `std-${group.key}-${idx}`;

        let stdKey = group.key;
        if (group.key === 'directCompare') {
          stdKey = mapping.std.directLike.includes(idx) ? 'directLike' : 'directBuy';
        }

        item.innerHTML = `
          <input type="checkbox" id="${id}" data-std-key="${stdKey}" data-col-idx="${idx}" checked>
          <label for="${id}">
            <small>${header[idx]}</small>
          </label>
        `;
        list.appendChild(item);
      });
    }

    col.appendChild(list);
    standardGroupsEl.appendChild(col);
  });
}

function renderExtraQuestionsUI(mapping) {
  extraQuestionsEl.innerHTML = '';

  if (!mapping.extraCandidates.length) {
    extraQuestionsEl.innerHTML = '<div class="status">Дополнительные закрытые вопросы не найдены.</div>';
    return;
  }

  mapping.extraCandidates.forEach(q => {
    const wrap = document.createElement('div');
    wrap.className = 'card';
    wrap.style.marginBottom = '1rem';

    wrap.innerHTML = `
      <div class="field">
        <label>
          <input type="checkbox" data-extra-idx="${q.idx}" checked>
          Использовать этот вопрос как доп.метрику
        </label>
        <div><small>${q.header}</small></div>
      </div>

      <div class="row">
        <div class="col-half">
          <div class="field">
            <label>Название метрики в топлайне</label>
            <input type="text" data-extra-idx="${q.idx}" data-role="title" placeholder="Напр. Осведомленность">
          </div>
          <div class="field">
            <label>Тип вопроса</label>
            <select data-extra-idx="${q.idx}" data-role="type">
              <option value="scale5">Шкала 1–5</option>
              <option value="single">Single choice</option>
            </select>
          </div>
        </div>
        <div class="col-half">
          <div class="field">
            <label>Куда выводить</label>
            <div class="pill-checkboxes" data-extra-idx="${q.idx}" data-role="where">
              <label><input type="checkbox" value="summary" checked> САММАРИ</label>
              <label><input type="checkbox" value="full" checked> полные таблицы</label>
              <label><input type="checkbox" value="signif"> значимости</label>
            </div>
          </div>
        </div>
      </div>
    `;

    extraQuestionsEl.appendChild(wrap);
  });
}

function collectUserConfig(mapping) {
  const stdSelected = {
    like: [],
    fitDish: [],
    fitBrand: [],
    visitBK: [],
    buyDish: [],
    directLike: [],
    directBuy: [],
    image: mapping.std.image.slice(),
    audience: mapping.std.audience
  };

  document.querySelectorAll('input[type="checkbox"][data-std-key]').forEach(cb => {
    if (!cb.checked) return;
    const key = cb.getAttribute('data-std-key');
    const idx = Number(cb.getAttribute('data-col-idx'));
    stdSelected[key].push(idx);
  });

  const extra = [];

  mapping.extraCandidates.forEach(q => {
    const enabledCb = document.querySelector(`input[type="checkbox"][data-extra-idx="${q.idx}"]`);
    if (!enabledCb || !enabledCb.checked) return;

    const titleInput = document.querySelector(`input[data-extra-idx="${q.idx}"][data-role="title"]`);
    const typeSelect = document.querySelector(`select[data-extra-idx="${q.idx}"][data-role="type"]`);
    const whereWrap = document.querySelector(`div[data-extra-idx="${q.idx}"][data-role="where"]`);

    const title = (titleInput?.value || '').trim();
    if (!title) throw new Error('У доп.вопроса "' + q.header + '" не задано название метрики.');

    const qtype = typeSelect?.value || 'scale5';
    const where = [];
    whereWrap.querySelectorAll('input[type="checkbox"]').forEach(cb => {
      if (cb.checked) where.push(cb.value);
    });

    if (!where.length) throw new Error('У доп.метрики "' + title + '" не выбрано, куда выводить.');

    extra.push({ idx: q.idx, header: q.header, title, type: qtype, where });
  });

  return { std: stdSelected, extra };
}

function inferConcepts(header, config) {
  const sourceCols =
    config.std.like.length ? config.std.like :
    config.std.fitDish.length ? config.std.fitDish :
    config.std.fitBrand.length ? config.std.fitBrand :
    config.std.visitBK.length ? config.std.visitBK :
    config.std.buyDish.length ? config.std.buyDish : [];

  if (!sourceCols.length) return [{ code: 'A', label: 'Название A' }];

  const labels = sourceCols.map(colIdx => {
    const text = String(header[colIdx] || '').trim();
    const parts = text.split(' - ');
    return parts.length > 1 ? parts[parts.length - 1].trim() : text;
  });

  return labels.map((label, i) => ({
    code: String.fromCharCode(65 + i),
    label: label || `Название ${String.fromCharCode(65 + i)}`
  }));
}

function getCell(row, idx) {
  if (idx == null || idx < 0) return null;
  return row[idx];
}

function parseScaleValue(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return v;
  const s = String(v).trim();
  const m = s.match(/^([1-5])/);
  return m ? Number(m[1]) : null;
}

function findConceptIndexByHeader(headerText, concepts) {
  const text = String(headerText || '').trim();
  for (let i = 0; i < concepts.length; i++) if (text.endsWith(concepts[i].label)) return i;
  for (let i = 0; i < concepts.length; i++) if (text.includes(concepts[i].label)) return i;
  return -1;
}

function calcStandardBlocks(rows, config, concepts, header) {
  const n = rows.length;

  function top2ByCols(cols) {
    const res = Array(concepts.length).fill(0);
    rows.forEach(r => {
      cols.forEach((col, i) => {
        if (i >= concepts.length) return;
        const v = parseScaleValue(getCell(r, col));
        if (v === 4 || v === 5) res[i]++;
      });
    });
    return res.map(v => v / n);
  }

  function dist5(cols) {
    const arr = Array.from({ length: concepts.length }, () => ({ '1':0, '2':0, '3':0, '4':0, '5':0 }));
    rows.forEach(r => {
      cols.forEach((col, i) => {
        if (i >= concepts.length) return;
        const v = parseScaleValue(getCell(r, col));
        if (v >= 1 && v <= 5) arr[i][String(v)]++;
      });
    });
    return arr.map(d => ({
      '1': d['1'] / n,
      '2': d['2'] / n,
      '3': d['3'] / n,
      '4': d['4'] / n,
      '5': d['5'] / n,
      top2: (d['4'] + d['5']) / n
    }));
  }

  function imageBlock() {
    const names = [
      'Понятное и простое название',
      'Вызывает аппетит, звучит вкусно',
      'Вызывает доверие',
      'Добавляет премиальности',
      'Хочется попробовать такой капучино',
      'Понятно какой будет вкус',
      'Звучит как качественный продукт',
      'Оригинальное, отличается от других',
      'Хочется рассказать/поделиться в соц.сетях'
    ];

    const res = {};
    names.forEach(nm => res[nm] = Array(concepts.length).fill(0));

    rows.forEach(r => {
      config.std.image.forEach(({ key, idx }) => {
        const val = String(getCell(r, idx) || '').trim();
        if (!val) return;
        const conceptIndex = findConceptIndexByHeader(header[idx], concepts);
        if (conceptIndex === -1) return;
        if (res[key]) res[key][conceptIndex]++;
      });
    });

    Object.keys(res).forEach(k => {
      res[k] = res[k].map(v => v / n);
    });

    return res;
  }

  function directSingle(cols) {
    const counts = {};
    cols.forEach(idx => {
      rows.forEach(r => {
        const v = String(getCell(r, idx) || '').trim();
        if (!v) return;
        counts[v] = (counts[v] || 0) + 1;
      });
    });

    const perConcept = Array(concepts.length).fill(0);
    let none = 0;

    concepts.forEach((c, i) => {
      const key = Object.keys(counts).find(k => k.includes(c.label));
      perConcept[i] = key ? counts[key] / n : 0;
    });

    const noneKey = Object.keys(counts).find(k => k.toLowerCase().includes('ни одно'));
    if (noneKey) none = counts[noneKey] / n;

    return { perConcept, none };
  }

  return {
    n,
    scales: {
      like: dist5(config.std.like),
      fitDish: dist5(config.std.fitDish),
      fitBrand: dist5(config.std.fitBrand),
      visitBK: dist5(config.std.visitBK),
      buyDish: dist5(config.std.buyDish)
    },
    top2: {
      like: top2ByCols(config.std.like),
      fitDish: top2ByCols(config.std.fitDish),
      fitBrand: top2ByCols(config.std.fitBrand),
      visitBK: top2ByCols(config.std.visitBK),
      buyDish: top2ByCols(config.std.buyDish)
    },
    image: imageBlock(),
    direct: {
      likeMost: directSingle(config.std.directLike),
      buyFirst: directSingle(config.std.directBuy)
    }
  };
}

function calcExtraBlocks(rows, config) {
  const n = rows.length;
  const result = [];

  config.extra.forEach(q => {
    if (q.type === 'scale5') {
      const counts = { '1':0, '2':0, '3':0, '4':0, '5':0 };
      rows.forEach(r => {
        const v = parseScaleValue(getCell(r, q.idx));
        if (v >= 1 && v <= 5) counts[String(v)]++;
      });

      result.push({
        kind: 'scale5',
        title: q.title,
        where: q.where,
        dist: {
          '1': counts['1'] / n,
          '2': counts['2'] / n,
          '3': counts['3'] / n,
          '4': counts['4'] / n,
          '5': counts['5'] / n,
          top2: (counts['4'] + counts['5']) / n
        }
      });
    } else {
      const counts = {};
      rows.forEach(r => {
        const v = String(getCell(r, q.idx) || '').trim();
        if (!v) return;
        counts[v] = (counts[v] || 0) + 1;
      });

      result.push({
        kind: 'single',
        title: q.title,
        where: q.where,
        dist: Object.entries(counts)
          .map(([cat, c]) => ({ cat, p: c / n }))
          .sort((a, b) => b.p - a.p)
      });
    }
  });

  return result;
}

function calcAudience(rows, config) {
  const n = rows.length;

  function freq(idx) {
    if (idx == null || idx < 0) return [];
    const counts = {};
    rows.forEach(r => {
      const v = String(getCell(r, idx) || '').trim();
      if (!v) return;
      counts[v] = (counts[v] || 0) + 1;
    });
    return Object.entries(counts)
      .map(([label, c]) => ({ label, p: c / n }))
      .sort((a, b) => b.p - a.p);
  }

  return {
    n,
    sex: freq(config.std.audience.sex),
    age: freq(config.std.audience.age),
    freqNew: freq(config.std.audience.freqNew),
    freqProd: freq(config.std.audience.freqProd),
    freqBK: freq(config.std.audience.freqBK)
  };
}

function zTest(p1, p2, n1, n2) {
  const p = (p1 * n1 + p2 * n2) / (n1 + n2);
  const se = Math.sqrt(p * (1 - p) * (1 / n1 + 1 / n2));
  if (!se) return 0;
  return (p1 - p2) / se;
}

function calcSignificance(stdRes, concepts, n) {
  const alphaZ = 1.96;
  const signif = {
    top2: {},
    scales: {},
    image: {},
    directMax: {}
  };

  function labelsFor(arr) {
    return arr.map((p, i) => {
      const greater = [];
      arr.forEach((q, j) => {
        if (i === j) return;
        const z = zTest(p, q, n, n);
        if (z > alphaZ) greater.push(concepts[j].code);
      });
      return greater;
    });
  }

  ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish'].forEach(k => {
    signif.top2[k] = labelsFor(stdRes.top2[k]);
  });

  ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish'].forEach(k => {
    signif.scales[k] = {};
    ['top2', '1', '2', '3', '4', '5'].forEach(level => {
      signif.scales[k][level] = labelsFor(stdRes.scales[k].map(d => d[level]));
    });
  });

  Object.entries(stdRes.image).forEach(([k, vals]) => {
    signif.image[k] = labelsFor(vals);
  });

  function maxMask(arr) {
    const max = Math.max(...arr);
    return arr.map(v => max > 0 && v === max);
  }

  signif.directMax.likeMost = maxMask(stdRes.direct.likeMost.perConcept);
  signif.directMax.buyFirst = maxMask(stdRes.direct.buyFirst.perConcept);

  return signif;
}

function stdResToSignifInput(stdRes) {
  return stdRes;
}

function cellRef(r, c) {
  return XLSX.utils.encode_cell({ r, c });
}

function setCell(ws, r, c, value, style = null) {
  ws[cellRef(r, c)] = { t: typeof value === 'number' ? 'n' : 's', v: value };
  if (style) ws[cellRef(r, c)].s = JSON.parse(JSON.stringify(style));
  return ws[cellRef(r, c)];
}

function ensureCell(ws, r, c, value = '') {
  const ref = cellRef(r, c);
  if (!ws[ref]) ws[ref] = { t: typeof value === 'number' ? 'n' : 's', v: value };
  return ws[ref];
}

function setPercent(ws, r, c, value, style) {
  const cell = setCell(ws, r, c, Number(value || 0), style);
  cell.z = '0%';
  return cell;
}

function mergeRange(ws, sRow, sCol, eRow, eCol) {
  if (!ws['!merges']) ws['!merges'] = [];
  ws['!merges'].push({
    s: { r: sRow, c: sCol },
    e: { r: eRow, c: eCol }
  });
}

function applySheetRangeRef(ws, endRow, endCol) {
  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: endRow, c: endCol }
  });
}

function borderAll() {
  return {
    top: { style: 'thin', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } }
  };
}

function hexFill(rgb) {
  return { patternType: 'solid', fgColor: { rgb } };
}

const STYLES = {
  title: {
    font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 14 },
    fill: hexFill('244C73'),
    alignment: { horizontal: 'left', vertical: 'center' },
    border: borderAll()
  },
  section: {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: hexFill('244C73'),
    alignment: { horizontal: 'left', vertical: 'center' },
    border: borderAll()
  },
  blockTitle: {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: hexFill('5E86B4'),
    alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    border: borderAll()
  },
  headerCenter: {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: hexFill('244C73'),
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: borderAll()
  },
  base: {
    font: { italic: true, color: { rgb: '333333' } },
    fill: hexFill('D9D9D9'),
    alignment: { horizontal: 'left', vertical: 'center' },
    border: borderAll()
  },
  label: {
    alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    border: borderAll()
  },
  top2Label: {
    font: { bold: true },
    fill: hexFill('DCE6F1'),
    alignment: { horizontal: 'left', vertical: 'center' },
    border: borderAll()
  },
  top2Row: {
    font: { bold: true },
    fill: hexFill('DCE6F1'),
    alignment: { horizontal: 'center', vertical: 'center' },
    border: borderAll()
  },
  percent: {
    alignment: { horizontal: 'center', vertical: 'center' },
    border: borderAll()
  },
  percentGreen: {
    alignment: { horizontal: 'center', vertical: 'center' },
    fill: hexFill('70AD47'),
    border: borderAll()
  },
  signifTextGreen: {
    font: { bold: true, color: { rgb: '000000' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    fill: hexFill('70AD47'),
    border: borderAll()
  },
  legendGreen: {
    alignment: { horizontal: 'center', vertical: 'center' },
    fill: hexFill('92D050'),
    border: borderAll()
  },
  legendAccent: {
    font: { bold: true, color: { rgb: 'C55A11' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    fill: hexFill('FFF2CC'),
    border: borderAll()
  },
  legendText: {
    alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    border: borderAll()
  }
};

function isStrong2Plus(arr, index) {
  return Array.isArray(arr) && Array.isArray(arr[index]) && arr[index].length >= 2;
}

function blockTitleForKey(key) {
  return {
    like: 'НАСКОЛЬКО НРАВИТСЯ НАЗВАНИЕ',
    fitDish: 'НАСКОЛЬКО ПОДХОДИТ ДЛЯ ЭТОГО БЛЮДА',
    fitBrand: 'НАСКОЛЬКО ПОДХОДИТ ДЛЯ БРЕНДА БУРГЕР КИНГ В ЦЕЛОМ',
    visitBK: 'НАМЕРЕНИЕ ПОСЕТИТЬ БУРГЕР КИНГ, ЕСЛИ ПОЯВИТСЯ В МЕНЮ',
    buyDish: 'НАМЕРЕНИЕ КУПИТЬ ПО ПРИЕМЛЕМОЙ ЦЕНЕ'
  }[key];
}

function scaleLabelsForBlock(key) {
  return {
    like: ['1 - Совсем не нравится', '2', '3', '4', '5 - Очень нравится'],
    fitDish: ['1 - Точно не подходит', '2', '3', '4', '5 - Полностью подходит'],
    fitBrand: ['1 - Точно не подходит', '2', '3', '4', '5 - Полностью подходит'],
    visitBK: ['1 - Точно не посещу', '2', '3', '4', '5 - Точно посещу'],
    buyDish: ['1 - Точно не куплю', '2', '3', '4', '5 - Точно куплю']
  }[key];
}

function blockValueRows(stdRes, key) {
  const distArr = stdRes.scales[key];
  const labels = scaleLabelsForBlock(key);
  return [
    ['ТОП-2 (сумма 4+5)', distArr.map(d => d.top2), 'top2'],
    [labels[0], distArr.map(d => d['1']), '1'],
    [labels[1], distArr.map(d => d['2']), '2'],
    [labels[2], distArr.map(d => d['3']), '3'],
    [labels[3], distArr.map(d => d['4']), '4'],
    [labels[4], distArr.map(d => d['5']), '5']
  ];
}

function makeSummarySheetStyled(stdRes, concepts, signifRes) {
  const ws = {};
  const lastCol = concepts.length;

  ws['!cols'] = [{ wch: 42 }, ...Array.from({ length: concepts.length }, () => ({ wch: 15 }))];

  let row = 0;
  setCell(ws, row, 0, 'САММАРИ: ТОП-2 (сумма оценок 4 и 5)', STYLES.title);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  setCell(ws, row, 0, 'Тестируемые варианты названий', STYLES.headerCenter);
  concepts.forEach((c, i) => setCell(ws, row, i + 1, c.label, STYLES.headerCenter));
  row++;

  setCell(ws, row, 0, `База: n=${stdRes.n} респондентов | Все значения в %`, STYLES.base);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  setCell(ws, row, 1, 'xx', STYLES.legendGreen);
  setCell(ws, row, 2, 'значимо выше 2 и более других названий', STYLES.legendText);
  if (lastCol >= 2) mergeRange(ws, row, 2, row, lastCol);
  row++;

  setCell(ws, row, 0, 'ОСНОВНЫЕ ПОКАЗАТЕЛИ', STYLES.section);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  setCell(ws, row, 0, 'Показатель', STYLES.headerCenter);
  concepts.forEach((c, i) => setCell(ws, row, i + 1, c.label, STYLES.headerCenter));
  row++;

  [
    ['Нравится название', stdRes.top2.like, 'like'],
    ['Подходит для блюда', stdRes.top2.fitDish, 'fitDish'],
    ['Подходит для бренда', stdRes.top2.fitBrand, 'fitBrand'],
    ['Намерение посетить БК', stdRes.top2.visitBK, 'visitBK'],
    ['Намерение купить', stdRes.top2.buyDish, 'buyDish']
  ].forEach(([label, vals, key]) => {
    setCell(ws, row, 0, label, STYLES.label);
    vals.forEach((v, i) => {
      setPercent(ws, row, i + 1, v, isStrong2Plus(signifRes.top2[key], i) ? STYLES.percentGreen : STYLES.percent);
    });
    row++;
  });

  row++;
  setCell(ws, row, 0, 'Прямое сравнение', STYLES.section);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  [
    ['Нравится больше всего', stdRes.direct.likeMost.perConcept, signifRes.directMax.likeMost],
    ['Куплю в первую очередь', stdRes.direct.buyFirst.perConcept, signifRes.directMax.buyFirst]
  ].forEach(([label, vals, mask]) => {
    setCell(ws, row, 0, label, STYLES.label);
    vals.forEach((v, i) => {
      setPercent(ws, row, i + 1, v, mask[i] ? STYLES.percentGreen : STYLES.percent);
    });
    row++;
  });

  row++;
  setCell(ws, row, 0, 'ИМИДЖЕВЫЙ БЛОК', STYLES.section);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  setCell(ws, row, 0, 'Показатель', STYLES.headerCenter);
  concepts.forEach((c, i) => setCell(ws, row, i + 1, c.label, STYLES.headerCenter));
  row++;

  Object.entries(stdRes.image).forEach(([label, vals]) => {
    setCell(ws, row, 0, label, STYLES.label);
    vals.forEach((v, i) => {
      setPercent(ws, row, i + 1, v, isStrong2Plus(signifRes.image[label], i) ? STYLES.percentGreen : STYLES.percent);
    });
    row++;
  });

  applySheetRangeRef(ws, row, lastCol);
  return ws;
}

function makeFullSheetStyled(stdRes, concepts) {
  const ws = {};
  const lastCol = concepts.length;

  ws['!cols'] = [{ wch: 46 }, ...Array.from({ length: concepts.length }, () => ({ wch: 15 }))];

  let row = 0;
  setCell(ws, row, 0, 'ПОЛНЫЕ ДАННЫЕ ПО ВСЕМ ВОПРОСАМ', STYLES.title);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  setCell(ws, row, 0, 'Тестируемые варианты названий', STYLES.headerCenter);
  concepts.forEach((c, i) => setCell(ws, row, i + 1, c.label, STYLES.headerCenter));
  row++;

  setCell(ws, row, 0, `База: n=${stdRes.n} респондентов | Все значения в %`, STYLES.base);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish'].forEach(key => {
    setCell(ws, row, 0, blockTitleForKey(key), STYLES.blockTitle);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    blockValueRows(stdRes, key).forEach(([label, vals, level]) => {
      setCell(ws, row, 0, label, level === 'top2' ? STYLES.top2Label : STYLES.label);
      vals.forEach((v, i) => setPercent(ws, row, i + 1, v, level === 'top2' ? STYLES.top2Row : STYLES.percent));
      row++;
    });

    row++;
  });

  setCell(ws, row, 0, 'ИМИДЖЕВЫЙ БЛОК', STYLES.section);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  Object.entries(stdRes.image).forEach(([label, vals]) => {
    setCell(ws, row, 0, label, STYLES.label);
    vals.forEach((v, i) => setPercent(ws, row, i + 1, v, STYLES.percent));
    row++;
  });

  row++;
  setCell(ws, row, 0, 'ПРЯМОЕ СРАВНЕНИЕ', STYLES.section);
  mergeRange(ws, row, 0, row, 2);
  row++;

  setCell(ws, row, 0, 'Название', STYLES.headerCenter);
  setCell(ws, row, 1, 'Нравится больше всего', STYLES.headerCenter);
  setCell(ws, row, 2, 'Куплю в первую очередь', STYLES.headerCenter);
  row++;

  concepts.forEach((c, i) => {
    setCell(ws, row, 0, c.label, STYLES.label);
    setPercent(ws, row, 1, stdRes.direct.likeMost.perConcept[i] || 0, STYLES.percent);
    setPercent(ws, row, 2, stdRes.direct.buyFirst.perConcept[i] || 0, STYLES.percent);
    row++;
  });

  setCell(ws, row, 0, 'Ни одно из них', STYLES.label);
  setPercent(ws, row, 1, stdRes.direct.likeMost.none || 0, STYLES.percent);
  setPercent(ws, row, 2, stdRes.direct.buyFirst.none || 0, STYLES.percent);

  applySheetRangeRef(ws, row, Math.max(lastCol, 2));
  return ws;
}

function signifCellText(value, letters) {
  const pct = Math.round((value || 0) * 100) + '%';
  return letters && letters.length ? `${pct} ${letters.join(',')}` : pct;
}

function writeSignifBlock(ws, startRow, startCol, stdRes, concepts, signifRes, mode) {
  let row = startRow;
  const lastCol = startCol + concepts.length;

  setCell(ws, row, startCol, 'ПОЛНЫЕ ДАННЫЕ ПО ВСЕМ ВОПРОСАМ', STYLES.title);
  mergeRange(ws, row, startCol, row, lastCol);
  row++;

  setCell(ws, row, startCol, 'Тестируемые варианты названий', STYLES.headerCenter);
  concepts.forEach((c, i) => setCell(ws, row, startCol + 1 + i, c.label, STYLES.headerCenter));
  row++;

  setCell(ws, row, startCol, `База: n=${stdRes.n} респондентов | Все значения в %`, STYLES.base);
  mergeRange(ws, row, startCol, row, lastCol);
  row++;

  if (mode === 'green') {
    setCell(ws, row, startCol + 1, 'xx', STYLES.legendGreen);
    setCell(ws, row, startCol + 2, 'значимо выше 2 и более других названий', STYLES.legendText);
    if (lastCol >= startCol + 2) mergeRange(ws, row, startCol + 2, row, lastCol);
  } else {
    setCell(ws, row, startCol + 1, 'Х', STYLES.legendAccent);
    setCell(ws, row, startCol + 2, 'значимо выше по сравнению с другими вариантами', STYLES.legendText);
    if (lastCol >= startCol + 2) mergeRange(ws, row, startCol + 2, row, lastCol);
  }
  row++;

  ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish'].forEach(key => {
    setCell(ws, row, startCol, blockTitleForKey(key), STYLES.blockTitle);
    mergeRange(ws, row, startCol, row, lastCol);
    row++;

    blockValueRows(stdRes, key).forEach(([label, vals, level]) => {
      setCell(ws, row, startCol, label, level === 'top2' ? STYLES.top2Label : STYLES.label);

      vals.forEach((v, i) => {
        if (mode === 'green') {
          const style = isStrong2Plus(signifRes.scales[key][level], i)
            ? STYLES.percentGreen
            : (level === 'top2' ? STYLES.top2Row : STYLES.percent);
          setPercent(ws, row, startCol + 1 + i, v, style);
        } else {
          const letters = signifRes.scales[key][level][i];
          const style = letters && letters.length ? STYLES.signifTextGreen : (level === 'top2' ? STYLES.top2Row : STYLES.percent);
          setCell(ws, row, startCol + 1 + i, signifCellText(v, letters), style);
        }
      });

      row++;
    });

    row++;
  });

  setCell(ws, row, startCol, 'ИМИДЖЕВЫЙ БЛОК', STYLES.section);
  mergeRange(ws, row, startCol, row, lastCol);
  row++;

  Object.entries(stdRes.image).forEach(([label, vals]) => {
    setCell(ws, row, startCol, label, STYLES.label);
    vals.forEach((v, i) => {
      if (mode === 'green') {
        setPercent(
          ws,
          row,
          startCol + 1 + i,
          v,
          isStrong2Plus(signifRes.image[label], i) ? STYLES.percentGreen : STYLES.percent
        );
      } else {
        const letters = signifRes.image[label][i];
        setCell(
          ws,
          row,
          startCol + 1 + i,
          signifCellText(v, letters),
          letters && letters.length ? STYLES.signifTextGreen : STYLES.percent
        );
      }
    });
    row++;
  });

  row++;
  setCell(ws, row, startCol, 'ПРЯМОЕ СРАВНЕНИЕ', STYLES.section);
  mergeRange(ws, row, startCol, row, startCol + 2);
  row++;

  setCell(ws, row, startCol, 'Название', STYLES.headerCenter);
  setCell(ws, row, startCol + 1, 'Нравится больше всего', STYLES.headerCenter);
  setCell(ws, row, startCol + 2, 'Куплю в первую очередь', STYLES.headerCenter);
  row++;

  concepts.forEach((c, i) => {
    setCell(ws, row, startCol, c.label, STYLES.label);
    if (mode === 'green') {
      setPercent(ws, row, startCol + 1, stdRes.direct.likeMost.perConcept[i] || 0, signifRes.directMax.likeMost[i] ? STYLES.percentGreen : STYLES.percent);
      setPercent(ws, row, startCol + 2, stdRes.direct.buyFirst.perConcept[i] || 0, signifRes.directMax.buyFirst[i] ? STYLES.percentGreen : STYLES.percent);
    } else {
      setCell(ws, row, startCol + 1, Math.round((stdRes.direct.likeMost.perConcept[i] || 0) * 100) + '%', STYLES.percent);
      setCell(ws, row, startCol + 2, Math.round((stdRes.direct.buyFirst.perConcept[i] || 0) * 100) + '%', STYLES.percent);
    }
    row++;
  });

  setCell(ws, row, startCol, 'Ни одно из них', STYLES.label);
  if (mode === 'green') {
    setPercent(ws, row, startCol + 1, stdRes.direct.likeMost.none || 0, STYLES.percent);
    setPercent(ws, row, startCol + 2, stdRes.direct.buyFirst.none || 0, STYLES.percent);
  } else {
    setCell(ws, row, startCol + 1, Math.round((stdRes.direct.likeMost.none || 0) * 100) + '%', STYLES.percent);
    setCell(ws, row, startCol + 2, Math.round((stdRes.direct.buyFirst.none || 0) * 100) + '%', STYLES.percent);
  }
  row++;

  return { endRow: row, endCol: lastCol };
}

function makeSignifSheetStyled(stdRes, concepts, signifRes) {
  const ws = {};

  const leftCols = [{ wch: 42 }, ...Array.from({ length: concepts.length }, () => ({ wch: 15 }))];
  const spacer = [{ wch: 4 }];
  const rightCols = [{ wch: 42 }, ...Array.from({ length: concepts.length }, () => ({ wch: 18 }))];

  ws['!cols'] = [...leftCols, ...spacer, ...rightCols];

  const left = writeSignifBlock(ws, 0, 0, stdRes, concepts, signifRes, 'green');
  const rightStart = leftCols.length + 1;
  const right = writeSignifBlock(ws, 0, rightStart, stdRes, concepts, signifRes, 'letters');

  applySheetRangeRef(ws, Math.max(left.endRow, right.endRow), Math.max(left.endCol, right.endCol));
  return ws;
}

function makeAudienceSheetStyled(audienceRes) {
  const ws = {};
  ws['!cols'] = [{ wch: 42 }, { wch: 18 }];

  let row = 0;

  setCell(ws, row, 0, 'АУДИТОРИЯ', STYLES.title);
  mergeRange(ws, row, 0, row, 1);
  row++;

  setCell(ws, row, 0, `База: n=${audienceRes.n} респондентов | Все значения в %`, STYLES.base);
  mergeRange(ws, row, 0, row, 1);
  row++;

  row = writeAudienceBlock(ws, row, 'СОЦИАЛЬНО-ДЕМОГРАФИЧЕСКИЙ ПРОФИЛЬ', [
    { title: 'Пол', rows: audienceRes.sex },
    { title: 'Возраст', rows: audienceRes.age }
  ]);

  row = writeAudienceBlock(ws, row, 'АУДИТОРИЯ И ПОВЕДЕНИЕ', [
    { title: 'Частота взятия новинок (горячие напитки)', rows: audienceRes.freqNew },
    { title: 'Частота покупки капучино', rows: audienceRes.freqProd },
    { title: 'Частота посещения Бургер Кинг', rows: audienceRes.freqBK }
  ]);

  applySheetRangeRef(ws, row, 1);
  return ws;
}

function writeAudienceBlock(ws, startRow, sectionTitle, blocks) {
  let row = startRow;

  setCell(ws, row, 0, sectionTitle, STYLES.section);
  mergeRange(ws, row, 0, row, 1);
  row++;

  blocks.forEach(block => {
    setCell(ws, row, 0, block.title, STYLES.blockTitle);
    mergeRange(ws, row, 0, row, 1);
    row++;

    setCell(ws, row, 0, 'Категория', STYLES.headerCenter);
    setCell(ws, row, 1, 'Доля', STYLES.headerCenter);
    row++;

    block.rows.forEach(item => {
      setCell(ws, row, 0, item.label, STYLES.label);
      setPercent(ws, row, 1, item.p, STYLES.percent);
      row++;
    });

    row++;
  });

  return row;
}
