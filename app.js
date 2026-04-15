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
      userConfig = collectUserConfig(autoMapping, parsed.header);
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
      const extraResults = calcExtraBlocks(rows, userConfig, concepts, header);
      const audienceRes = calcAudience(rows, userConfig, header);
      const signifRes = calcSignificance(stdResults, extraResults, concepts, rows.length);

      const outWb = XLSX.utils.book_new();

      const summarySheet = makeSummarySheetStyled(stdResults, extraResults, concepts, signifRes);
      XLSX.utils.book_append_sheet(outWb, summarySheet, 'САММАРИ');

      const fullSheet = makeFullSheetStyled(stdResults, extraResults, concepts);
      XLSX.utils.book_append_sheet(outWb, fullSheet, 'полные таблицы');

      const signifSheet = makeSignifSheetStyled(stdResults, concepts, signifRes);
      XLSX.utils.book_append_sheet(outWb, signifSheet, 'значимости');

      const audienceSheet = makeAudienceSheet(audienceRes);
      applyPercentFormat(audienceSheet, 0);
      XLSX.utils.book_append_sheet(outWb, audienceSheet, 'Аудитория');

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

// ---------- УТИЛИТЫ СОСТОЯНИЯ / СТАТУС ----------

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

// ---------- ПАРСИНГ БАЗЫ ----------

function splitHeaderRows(data) {
  if (!data || data.length < 2) {
    return { header: [], rows: [] };
  }

  const varNames = (data[0] || []).map(v => String(v || '').trim());
  const questionTexts = (data[1] || []).map(v => String(v || '').trim());

  const header = questionTexts.map((txt, i) => txt || varNames[i] || `col_${i}`);
  const rows = data.slice(2).filter(r => r && r.some(v => v !== null && v !== ''));

  return { header, rows };
}

// ---------- АВТОПОИСК ВОПРОСОВ ----------

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
      freqNew: null,
      freqProd: null,
      sex: null,
      age: null,
      freqBK: null
    }
  };

  header.forEach((h, idx) => {
    if (h.includes('Оцените, пожалуйста, насколько вам нравится или не нравится каждое из этих названий')) {
      std.like.push(idx);
    }

    if (h.includes('Насколько каждое из этих названий') &&
        h.includes('подходит или не подходит для этого')) {
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

    if (h.includes('Понятное и простое название - ')) {
      std.image.push({ type: 'image', key: 'Понятное и простое название', idx });
    }
    if (h.includes('Вызывает аппетит')) {
      std.image.push({ type: 'image', key: 'Вызывает аппетит, звучит вкусно', idx });
    }
    if (h.startsWith('Вызывает у меня доверие')) {
      std.image.push({ type: 'image', key: 'Вызывает доверие', idx });
    }
    if (h.startsWith('Добавляет премиальности')) {
      std.image.push({ type: 'image', key: 'Добавляет премиальности', idx });
    }
    if (h.startsWith('Хочется попробовать такой капучино')) {
      std.image.push({ type: 'image', key: 'Хочется попробовать такой капучино', idx });
    }
    if (h.startsWith('Понятно какой будет вкус')) {
      std.image.push({ type: 'image', key: 'Понятно какой будет вкус', idx });
    }
    if (h.startsWith('Звучит как качественный продукт')) {
      std.image.push({ type: 'image', key: 'Звучит как качественный продукт', idx });
    }
    if (h.startsWith('Оригинальное, отличается от других')) {
      std.image.push({ type: 'image', key: 'Оригинальное, отличается от других', idx });
    }
    if (h.startsWith('Хочется рассказать/поделиться в социальных сетях')) {
      std.image.push({ type: 'image', key: 'Хочется рассказать/поделиться в соц.сетях', idx });
    }

    if (h.includes('Какое из перечисленных ниже названий')) {
      std.directLike.push(idx);
    }
    if (h.includes('С каким из этих названий вы бы купили')) {
      std.directBuy.push(idx);
    }

    if (h.includes('Как часто Вы берете новинки в категории горячих напитков')) {
      std.audience.freqNew = idx;
    }
    if (h.includes('Как часто вы покупаете капучино')) {
      std.audience.freqProd = idx;
    }
    if (h.includes('Укажите Ваш пол')) {
      std.audience.sex = idx;
    }
    if (h.includes('Укажите Ваш возраст')) {
      std.audience.age = idx;
    }
    if (h.includes('Как часто вы посещаете Бургер Кинг')) {
      std.audience.freqBK = idx;
    }
  });

  const usedIndexes = new Set([
    ...std.like,
    ...std.fitDish,
    ...std.fitBrand,
    ...std.visitBK,
    ...std.buyDish,
    ...std.image.map(o => o.idx),
    ...std.directLike,
    ...std.directBuy,
    std.audience.freqNew,
    std.audience.freqProd,
    std.audience.sex,
    std.audience.age,
    std.audience.freqBK
  ].filter(v => v !== null && v !== undefined));

  const extraCandidates = header.map((h, idx) => {
    if (usedIndexes.has(idx)) return null;
    if (!h) return null;

    const lower = h.toLowerCase();
    const looksClosed = lower.includes('насколько') ||
                        lower.includes('оцените') ||
                        lower.includes('выберите') ||
                        lower.includes('какое из перечисленных');

    if (!looksClosed) return null;
    return { idx, header: h };
  }).filter(Boolean);

  return { std, extraCandidates };
}

// ---------- UI ----------

function renderStandardMappingUI(mapping, header) {
  const groups = [
    { key: 'like', label: 'Нравится название (шкала 1–5, Top‑2)', indexes: mapping.std.like },
    { key: 'fitDish', label: 'Подходит для блюда (шкала 1–5, Top‑2)', indexes: mapping.std.fitDish },
    { key: 'fitBrand', label: 'Подходит для бренда (шкала 1–5, Top‑2)', indexes: mapping.std.fitBrand },
    { key: 'visitBK', label: 'Намерение посетить БК (шкала 1–5, Top‑2)', indexes: mapping.std.visitBK },
    { key: 'buyDish', label: 'Намерение купить блюдо (шкала 1–5, Top‑2)', indexes: mapping.std.buyDish },
    {
      key: 'directCompare',
      label: 'Прямое сравнение (все single-choice вопросы этого типа)',
      indexes: [...mapping.std.directLike, ...mapping.std.directBuy]
    }
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
            <input type="text" data-extra-idx="${q.idx}" data-role="title" placeholder="Напр. 'Осведомленность о новинке'">
          </div>
          <div class="field">
            <label>Тип вопроса</label>
            <select data-extra-idx="${q.idx}" data-role="type">
              <option value="scale5">Шкала 1–5 (считаем Top‑2)</option>
              <option value="single">Single choice (один вариант ответа)</option>
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

// ---------- КОНФИГ ----------

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
    if (!title) {
      throw new Error('У доп.вопроса "' + q.header + '" не задано название метрики в топлайне.');
    }

    const qtype = typeSelect?.value || 'scale5';
    const where = [];

    whereWrap.querySelectorAll('input[type="checkbox"]').forEach(cb => {
      if (cb.checked) where.push(cb.value);
    });

    if (!where.length) {
      throw new Error('У доп.метрики "' + title + '" не выбрано, куда выводить.');
    }

    extra.push({
      idx: q.idx,
      header: q.header,
      title,
      type: qtype,
      where
    });
  });

  return { std: stdSelected, extra };
}

// ---------- КОНЦЕПТЫ ----------

function inferConcepts(header, config) {
  const sourceCols =
    (config.std.like && config.std.like.length ? config.std.like : null) ||
    (config.std.fitDish && config.std.fitDish.length ? config.std.fitDish : null) ||
    (config.std.fitBrand && config.std.fitBrand.length ? config.std.fitBrand : null) ||
    (config.std.visitBK && config.std.visitBK.length ? config.std.visitBK : null) ||
    (config.std.buyDish && config.std.buyDish.length ? config.std.buyDish : null) ||
    [];

  if (!sourceCols.length) {
    return [{ code: 'A', label: 'Название A' }];
  }

  const labels = sourceCols.map(colIdx => {
    const text = String(header[colIdx] || '').trim();
    const parts = text.split(' - ');
    if (parts.length > 1) return parts[parts.length - 1].trim();
    return text;
  });

  return labels.map((lab, i) => ({
    code: String.fromCharCode(65 + i),
    label: lab || `Название ${String.fromCharCode(65 + i)}`
  }));
}

// ---------- ВСПОМОГАТЕЛЬНОЕ ----------

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

  for (let i = 0; i < concepts.length; i++) {
    if (text.endsWith(concepts[i].label)) return i;
  }
  for (let i = 0; i < concepts.length; i++) {
    if (text.includes(concepts[i].label)) return i;
  }
  return -1;
}

function normalizeText(s) {
  return String(s || '').trim().toLowerCase();
}

// ---------- РАСЧЁТЫ ----------

function calcStandardBlocks(rows, config, concepts, header) {
  const n = rows.length;

  function top2ByCols(colIndexes) {
    const res = Array(concepts.length).fill(0);
    rows.forEach(r => {
      colIndexes.forEach((col, i) => {
        if (i >= concepts.length) return;
        const v = parseScaleValue(getCell(r, col));
        if (v === 4 || v === 5) res[i] += 1;
      });
    });
    return res.map(c => c / n);
  }

  function dist5(colIndexes) {
    const res = Array.from({ length: concepts.length }, () => ({ '1':0,'2':0,'3':0,'4':0,'5':0 }));
    rows.forEach(r => {
      colIndexes.forEach((col, i) => {
        if (i >= concepts.length) return;
        const v = parseScaleValue(getCell(r, col));
        if (v >= 1 && v <= 5) res[i][String(v)] += 1;
      });
    });
    return res.map(d => {
      const o = {};
      ['1','2','3','4','5'].forEach(k => o[k] = d[k] / n);
      o.top2 = (d['4'] + d['5']) / n;
      return o;
    });
  }

  function imageBlock() {
    const attrList = [
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

    const result = {};
    attrList.forEach(name => {
      result[name] = Array(concepts.length).fill(0);
    });

    rows.forEach(r => {
      config.std.image.forEach(({ key, idx }) => {
        const val = String(getCell(r, idx) || '').trim();
        if (!val) return;
        const conceptIndex = findConceptIndexByHeader(header[idx], concepts);
        if (conceptIndex === -1) return;
        if (result[key]) result[key][conceptIndex] += 1;
      });
    });

    Object.keys(result).forEach(k => {
      result[k] = result[k].map(c => c / n);
    });

    return result;
  }

  function directSingle(colIndexes) {
    const counts = {};
    colIndexes.forEach(idx => {
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
    const idx = q.idx;
    const type = q.type;

    if (type === 'scale5') {
      const counts = { '1':0,'2':0,'3':0,'4':0,'5':0 };
      rows.forEach(r => {
        const v = parseScaleValue(getCell(r, idx));
        if (v >= 1 && v <= 5) counts[String(v)] += 1;
      });

      const dist = {};
      ['1','2','3','4','5'].forEach(k => dist[k] = counts[k] / n);
      dist.top2 = (counts['4'] + counts['5']) / n;

      result.push({
        kind: 'scale5',
        title: q.title,
        where: q.where,
        dist
      });
    } else if (type === 'single') {
      const counts = {};
      rows.forEach(r => {
        const v = String(getCell(r, idx) || '').trim();
        if (!v) return;
        counts[v] = (counts[v] || 0) + 1;
      });

      const rowsOut = Object.entries(counts)
        .map(([cat, c]) => ({ cat, p: c / n }))
        .sort((a, b) => b.p - a.p);

      result.push({
        kind: 'single',
        title: q.title,
        where: q.where,
        dist: rowsOut
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
      .map(([label, c]) => [label, c / n])
      .sort((a, b) => b[1] - a[1]);
  }

  return {
    freqNew: freq(config.std.audience.freqNew),
    freqProd: freq(config.std.audience.freqProd),
    sex: freq(config.std.audience.sex),
    age: freq(config.std.audience.age),
    freqBK: freq(config.std.audience.freqBK),
    n
  };
}

// ---------- ЗНАЧИМОСТИ ----------

function zTest(p1, p2, n1, n2) {
  const p = (p1 * n1 + p2 * n2) / (n1 + n2);
  const se = Math.sqrt(p * (1 - p) * (1 / n1 + 1 / n2));
  if (!se) return 0;
  return (p1 - p2) / se;
}

function calcSignificance(stdRes, extraRes, concepts, n) {
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

  const stdTop2Keys = ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish'];
  stdTop2Keys.forEach(k => {
    const arr = stdRes.top2[k];
    if (!arr || !arr.length) return;
    signif.top2[k] = labelsFor(arr);
  });

  const stdScaleKeys = ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish'];
  stdScaleKeys.forEach(k => {
    const distArr = stdRes.scales[k];
    signif.scales[k] = {};
    ['top2','1','2','3','4','5'].forEach(level => {
      const arr = distArr.map(d => d[level]);
      signif.scales[k][level] = labelsFor(arr);
    });
  });

  Object.entries(stdRes.image).forEach(([k, vals]) => {
    signif.image[k] = labelsFor(vals);
  });

  function markMax(arr) {
    const max = Math.max(...arr);
    return arr.map(v => v === max && max > 0);
  }

  signif.directMax.likeMost = markMax(stdRes.direct.likeMost.perConcept);
  signif.directMax.buyFirst = markMax(stdRes.direct.buyFirst.perConcept);

  return signif;
}

// ---------- EXCEL FORMAT ----------

function formatPercentCell(cell, decimals = 0) {
  if (!cell) return;
  if (typeof cell.v !== 'number') return;
  cell.t = 'n';
  cell.z = decimals > 0 ? `0.${'0'.repeat(decimals)}%` : '0%';
}

function applyPercentFormat(ws, decimals = 0) {
  if (!ws || !ws['!ref']) return ws;
  const range = XLSX.utils.decode_range(ws['!ref']);

  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      if (!cell) continue;
      if (typeof cell.v !== 'number') continue;
      if (cell.v >= 0 && cell.v <= 1) formatPercentCell(cell, decimals);
    }
  }

  return ws;
}

function cellRef(r, c) {
  return XLSX.utils.encode_cell({ r, c });
}

function ensureCell(ws, r, c, value = '') {
  const ref = cellRef(r, c);
  if (!ws[ref]) ws[ref] = { t: typeof value === 'number' ? 'n' : 's', v: value };
  return ws[ref];
}

function setCell(ws, r, c, value, style = null) {
  const ref = cellRef(r, c);
  ws[ref] = { t: typeof value === 'number' ? 'n' : 's', v: value };
  if (style) ws[ref].s = style;
  return ws[ref];
}

function mergeRange(ws, sRow, sCol, eRow, eCol) {
  if (!ws['!merges']) ws['!merges'] = [];
  ws['!merges'].push({
    s: { r: sRow, c: sCol },
    e: { r: eRow, c: eCol }
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
    fill: hexFill('1F4E78'),
    alignment: { horizontal: 'left', vertical: 'center' },
    border: borderAll()
  },
  section: {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: hexFill('1F4E78'),
    alignment: { horizontal: 'left', vertical: 'center' },
    border: borderAll()
  },
  blockTitle: {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: hexFill('4F81BD'),
    alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    border: borderAll()
  },
  headerCenter: {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: hexFill('1F4E78'),
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
  labelBold: {
    font: { bold: true },
    alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    border: borderAll()
  },
  percent: {
    alignment: { horizontal: 'center', vertical: 'center' },
    border: borderAll()
  },
  top2Row: {
    font: { bold: true },
    fill: hexFill('DCE6F1'),
    alignment: { horizontal: 'center', vertical: 'center' },
    border: borderAll()
  },
  top2Label: {
    font: { bold: true },
    fill: hexFill('DCE6F1'),
    alignment: { horizontal: 'left', vertical: 'center' },
    border: borderAll()
  },
  percentGreen: {
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
    font: { bold: true, color: { rgb: '7F6000' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    fill: hexFill('FFF2CC'),
    border: borderAll()
  },
  legendText: {
    alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    border: borderAll()
  },
  stringAccent: {
    font: { bold: true, color: { rgb: 'C55A11' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: borderAll()
  },
  textAccent: {
    font: { bold: true, color: { rgb: 'C55A11' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: borderAll()
  }
};

function isStrong2Plus(arr, index) {
  return Array.isArray(arr) && Array.isArray(arr[index]) && arr[index].length >= 2;
}

function withLettersText(value, letters) {
  const pct = Math.round((value || 0) * 100) + '%';
  if (!letters || !letters.length) return pct;
  return pct + ' ' + letters.join(',');
}

function applySheetRangeRef(ws, endRow, endCol) {
  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: endRow, c: endCol }
  });
}

function makeLegendRow(ws, row, codeText, descText, codeStyle) {
  setCell(ws, row, 1, codeText, codeStyle);
  setCell(ws, row, 2, descText, STYLES.legendText);
}

function scaleLabelsForBlock(key) {
  const map = {
    like: [
      '1 - Совсем не нравится',
      '2',
      '3',
      '4',
      '5 - Очень нравится'
    ],
    fitDish: [
      '1 - Точно не подходит',
      '2',
      '3',
      '4',
      '5 - Полностью подходит'
    ],
    fitBrand: [
      '1 - Точно не подходит',
      '2',
      '3',
      '4',
      '5 - Полностью подходит'
    ],
    visitBK: [
      '1 - Точно не посещу',
      '2',
      '3',
      '4',
      '5 - Точно посещу'
    ],
    buyDish: [
      '1 - Точно не куплю',
      '2',
      '3',
      '4',
      '5 - Точно куплю'
    ]
  };
  return map[key];
}

function blockTitleForKey(key) {
  const map = {
    like: 'НАСКОЛЬКО НРАВИТСЯ НАЗВАНИЕ',
    fitDish: 'НАСКОЛЬКО ПОДХОДИТ ДЛЯ ЭТОГО БЛЮДА',
    fitBrand: 'НАСКОЛЬКО ПОДХОДИТ ДЛЯ БРЕНДА БУРГЕР КИНГ В ЦЕЛОМ',
    visitBK: 'НАМЕРЕНИЕ ПОСЕТИТЬ БУРГЕР КИНГ, ЕСЛИ ПОЯВИТСЯ В МЕНЮ',
    buyDish: 'НАМЕРЕНИЕ КУПИТЬ ПО ПРИЕМЛЕМОЙ ЦЕНЕ'
  };
  return map[key];
}

function blockValueRows(stdRes, key) {
  const distArr = stdRes.scales[key];
  return [
    ['ТОП-2 (сумма 4+5)', distArr.map(d => d.top2), 'top2'],
    [scaleLabelsForBlock(key)[0], distArr.map(d => d['1']), '1'],
    [scaleLabelsForBlock(key)[1], distArr.map(d => d['2']), '2'],
    [scaleLabelsForBlock(key)[2], distArr.map(d => d['3']), '3'],
    [scaleLabelsForBlock(key)[3], distArr.map(d => d['4']), '4'],
    [scaleLabelsForBlock(key)[4], distArr.map(d => d['5']), '5']
  ];
}

// ---------- SUMMARY ----------

function makeSummarySheetStyled(stdRes, extraRes, concepts, signifRes) {
  const ws = {};

  ws['!cols'] = [
    { wch: 42 },
    { wch: 15 },
    { wch: 15 },
    { wch: 15 },
    { wch: 15 },
    { wch: 15 }
  ];

  const rows = [];
  rows.push(['САММАРИ: ТОП-2 (сумма оценок 4 и 5)']);
  rows.push(['Тестируемые варианты названий', ...concepts.map(c => c.label)]);
  rows.push([`База: n=${stdRes.n} респондентов | Все значения в %`]);
  rows.push(['', 'xx', 'значимо выше 2 и более других названий']);
  rows.push(['ОСНОВНЫЕ ПОКАЗАТЕЛИ']);
  rows.push(['Показатель', ...concepts.map(c => c.label)]);
  rows.push(['Нравится название', ...stdRes.top2.like]);
  rows.push(['Подходит для блюда', ...stdRes.top2.fitDish]);
  rows.push(['Подходит для бренда', ...stdRes.top2.fitBrand]);
  rows.push(['Намерение посетить БК', ...stdRes.top2.visitBK]);
  rows.push(['Намерение купить', ...stdRes.top2.buyDish]);
  rows.push([]);
  rows.push(['Прямое сравнение']);
  rows.push(['Нравится больше всего', ...stdRes.direct.likeMost.perConcept]);
  rows.push(['Куплю в первую очередь', ...stdRes.direct.buyFirst.perConcept]);
  rows.push([]);
  rows.push(['ИМИДЖЕВЫЙ БЛОК']);
  rows.push(['Показатель', ...concepts.map(c => c.label)]);

  const imageKeys = Object.keys(stdRes.image);
  imageKeys.forEach(k => rows.push([k, ...stdRes.image[k]]));

  XLSX.utils.sheet_add_aoa(ws, rows, { origin: 'A1' });
  applyPercentFormat(ws, 0);

  mergeRange(ws, 0, 0, 0, 4);
  mergeRange(ws, 2, 0, 2, 4);
  mergeRange(ws, 4, 0, 4, 4);
  mergeRange(ws, 12, 0, 12, 4);
  mergeRange(ws, 16, 0, 16, 4);
  mergeRange(ws, 3, 2, 3, 4);

  for (let c = 0; c <= 4; c++) {
    ensureCell(ws, 0, c);
    ws[cellRef(0, c)].s = STYLES.title;
  }

  setCell(ws, 1, 0, 'Тестируемые варианты названий', STYLES.headerCenter);
  for (let i = 0; i < concepts.length; i++) {
    setCell(ws, 1, i + 1, concepts[i].label, STYLES.headerCenter);
  }

  for (let c = 0; c <= 4; c++) {
    ensureCell(ws, 2, c);
    ws[cellRef(2, c)].s = STYLES.base;
  }

  setCell(ws, 3, 1, 'xx', STYLES.legendGreen);
  setCell(ws, 3, 2, 'значимо выше 2 и более других названий', STYLES.legendText);

  for (let c = 0; c <= 4; c++) {
    ensureCell(ws, 4, c);
    ws[cellRef(4, c)].s = STYLES.section;
  }

  setCell(ws, 5, 0, 'Показатель', STYLES.headerCenter);
  for (let i = 0; i < concepts.length; i++) {
    setCell(ws, 5, i + 1, concepts[i].label, STYLES.headerCenter);
  }

  const top2Map = [
    ['like', 6],
    ['fitDish', 7],
    ['fitBrand', 8],
    ['visitBK', 9],
    ['buyDish', 10]
  ];

  top2Map.forEach(([key, rowIndex]) => {
    ws[cellRef(rowIndex, 0)].s = STYLES.label;
    for (let i = 0; i < concepts.length; i++) {
      const cell = ws[cellRef(rowIndex, i + 1)];
      cell.s = isStrong2Plus(signifRes.top2[key], i) ? STYLES.percentGreen : STYLES.percent;
    }
  });

  for (let c = 0; c <= 4; c++) {
    ensureCell(ws, 12, c);
    ws[cellRef(12, c)].s = STYLES.section;
  }

  const directRows = [
    { row: 13, key: 'likeMost' },
    { row: 14, key: 'buyFirst' }
  ];

  directRows.forEach(({ row, key }) => {
    ws[cellRef(row, 0)].s = STYLES.label;
    const marks = signifRes.directMax[key];
    for (let i = 0; i < concepts.length; i++) {
      const cell = ws[cellRef(row, i + 1)];
      cell.s = marks[i] ? STYLES.percentGreen : STYLES.percent;
    }
  });

  for (let c = 0; c <= 4; c++) {
    ensureCell(ws, 16, c);
    ws[cellRef(16, c)].s = STYLES.section;
  }

  setCell(ws, 17, 0, 'Показатель', STYLES.headerCenter);
  for (let i = 0; i < concepts.length; i++) {
    setCell(ws, 17, i + 1, concepts[i].label, STYLES.headerCenter);
  }

  imageKeys.forEach((key, idx) => {
    const row = 18 + idx;
    ws[cellRef(row, 0)].s = STYLES.label;
    for (let i = 0; i < concepts.length; i++) {
      const cell = ws[cellRef(row, i + 1)];
      cell.s = isStrong2Plus(signifRes.image[key], i) ? STYLES.percentGreen : STYLES.percent;
    }
  });

  ws['!rows'] = [];
  ws['!rows'][0] = { hpt: 22 };
  ws['!rows'][1] = { hpt: 32 };

  applySheetRangeRef(ws, 18 + imageKeys.length - 1, 4);
  return ws;
}

// ---------- FULL TABLES ----------

function makeFullSheetStyled(stdRes, extraRes, concepts) {
  const ws = {};
  ws['!cols'] = [
    { wch: 46 },
    { wch: 15 },
    { wch: 15 },
    { wch: 15 },
    { wch: 15 },
    { wch: 15 }
  ];

  let row = 0;

  setCell(ws, row, 0, 'ПОЛНЫЕ ДАННЫЕ ПО ВСЕМ ВОПРОСАМ', STYLES.title);
  mergeRange(ws, row, 0, row, 4);
  for (let c = 1; c <= 4; c++) ensureCell(ws, row, c).s = STYLES.title;
  row++;

  setCell(ws, row, 0, 'Тестируемые варианты названий', STYLES.headerCenter);
  for (let i = 0; i < concepts.length; i++) {
    setCell(ws, row, i + 1, concepts[i].label, STYLES.headerCenter);
  }
  row++;

  setCell(ws, row, 0, `База: n=${stdRes.n} респондентов | Все значения в %`, STYLES.base);
  mergeRange(ws, row, 0, row, 4);
  for (let c = 1; c <= 4; c++) ensureCell(ws, row, c).s = STYLES.base;
  row++;

  const blockKeys = ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish'];

  blockKeys.forEach(key => {
    setCell(ws, row, 0, blockTitleForKey(key), STYLES.blockTitle);
    mergeRange(ws, row, 0, row, 4);
    for (let c = 1; c <= 4; c++) ensureCell(ws, row, c).s = STYLES.blockTitle;
    row++;

    const valueRows = blockValueRows(stdRes, key);
    valueRows.forEach(([label, vals, level]) => {
      setCell(ws, row, 0, label, level === 'top2' ? STYLES.top2Label : STYLES.label);
      for (let i = 0; i < concepts.length; i++) {
        setCell(ws, row, i + 1, vals[i], level === 'top2' ? STYLES.top2Row : STYLES.percent);
      }
      row++;
    });

    row++;
  });

  setCell(ws, row, 0, 'ИМИДЖЕВЫЙ БЛОК', STYLES.section);
  mergeRange(ws, row, 0, row, 4);
  for (let c = 1; c <= 4; c++) ensureCell(ws, row, c).s = STYLES.section;
  row++;

  Object.entries(stdRes.image).forEach(([k, vals]) => {
    setCell(ws, row, 0, k, STYLES.label);
    for (let i = 0; i < concepts.length; i++) {
      setCell(ws, row, i + 1, vals[i], STYLES.percent);
    }
    row++;
  });

  row++;

  setCell(ws, row, 0, 'ПРЯМОЕ СРАВНЕНИЕ', STYLES.section);
  mergeRange(ws, row, 0, row, 2);
  for (let c = 1; c <= 2; c++) ensureCell(ws, row, c).s = STYLES.section;
  row++;

  setCell(ws, row, 0, 'Название', STYLES.headerCenter);
  setCell(ws, row, 1, 'Нравится больше всего', STYLES.headerCenter);
  setCell(ws, row, 2, 'Куплю в первую очередь', STYLES.headerCenter);
  row++;

  concepts.forEach((c, i) => {
    setCell(ws, row, 0, c.label, STYLES.label);
    setCell(ws, row, 1, stdRes.direct.likeMost.perConcept[i] || 0, STYLES.percent);
    setCell(ws, row, 2, stdRes.direct.buyFirst.perConcept[i] || 0, STYLES.percent);
    row++;
  });

  setCell(ws, row, 0, 'Ни одно из них', STYLES.label);
  setCell(ws, row, 1, stdRes.direct.likeMost.none || 0, STYLES.percent);
  setCell(ws, row, 2, stdRes.direct.buyFirst.none || 0, STYLES.percent);
  row++;

  applyPercentFormat(ws, 0);
  applySheetRangeRef(ws, row, 4);
  return ws;
}

// ---------- SIGNIFICANCE ----------

function writeSignifBlock(ws, startRow, startCol, stdRes, concepts, signifRes, mode) {
  let row = startRow;
  const endCol = startCol + 4;

  setCell(ws, row, startCol, 'ПОЛНЫЕ ДАННЫЕ ПО ВСЕМ ВОПРОСАМ', STYLES.title);
  mergeRange(ws, row, startCol, row, endCol);
  for (let c = startCol + 1; c <= endCol; c++) ensureCell(ws, row, c).s = STYLES.title;
  row++;

  setCell(ws, row, startCol, 'Тестируемые варианты названий', STYLES.headerCenter);
  for (let i = 0; i < concepts.length; i++) {
    setCell(ws, row, startCol + 1 + i, concepts[i].label + (mode === 'letters' ? ` (${concepts[i].code})` : ''), STYLES.headerCenter);
  }
  row++;

  setCell(ws, row, startCol, `База: n=${stdRes.n} респондентов | Все значения в %`, STYLES.base);
  mergeRange(ws, row, startCol, row, endCol);
  for (let c = startCol + 1; c <= endCol; c++) ensureCell(ws, row, c).s = STYLES.base;
  row++;

  if (mode === 'green') {
    setCell(ws, row, startCol + 1, 'xx', STYLES.legendGreen);
    setCell(ws, row, startCol + 2, 'значимо выше 2 и более других названий', STYLES.legendText);
    mergeRange(ws, row, startCol + 2, row, endCol);
  } else {
    setCell(ws, row, startCol + 1, 'AB', STYLES.legendAccent);
    setCell(ws, row, startCol + 2, 'буквы рядом с процентом = значимо выше по сравнению с указанными названиями', STYLES.legendText);
    mergeRange(ws, row, startCol + 2, row, endCol);
  }
  row++;

  const blockKeys = ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish'];

  blockKeys.forEach(key => {
    setCell(ws, row, startCol, blockTitleForKey(key), STYLES.blockTitle);
    mergeRange(ws, row, startCol, row, endCol);
    for (let c = startCol + 1; c <= endCol; c++) ensureCell(ws, row, c).s = STYLES.blockTitle;
    row++;

    const valueRows = blockValueRows(stdRes, key);
    valueRows.forEach(([label, vals, level]) => {
      setCell(ws, row, startCol, label, level === 'top2' ? STYLES.top2Label : STYLES.label);

      for (let i = 0; i < concepts.length; i++) {
        if (mode === 'green') {
          setCell(
            ws,
            row,
            startCol + 1 + i,
            vals[i],
            isStrong2Plus(signifRes.scales[key][level], i) ? STYLES.percentGreen : (level === 'top2' ? STYLES.top2Row : STYLES.percent)
          );
        } else {
          const txt = withLettersText(vals[i], signifRes.scales[key][level][i]);
          setCell(ws, row, startCol + 1 + i, txt, STYLES.stringAccent);
        }
      }
      row++;
    });

    row++;
  });

  setCell(ws, row, startCol, 'ИМИДЖЕВЫЙ БЛОК', STYLES.section);
  mergeRange(ws, row, startCol, row, endCol);
  for (let c = startCol + 1; c <= endCol; c++) ensureCell(ws, row, c).s = STYLES.section;
  row++;

  Object.entries(stdRes.image).forEach(([k, vals]) => {
    setCell(ws, row, startCol, k, STYLES.label);
    for (let i = 0; i < concepts.length; i++) {
      if (mode === 'green') {
        setCell(
          ws,
          row,
          startCol + 1 + i,
          vals[i],
          isStrong2Plus(signifRes.image[k], i) ? STYLES.percentGreen : STYLES.percent
        );
      } else {
        setCell(ws, row, startCol + 1 + i, withLettersText(vals[i], signifRes.image[k][i]), STYLES.stringAccent);
      }
    }
    row++;
  });

  row++;

  setCell(ws, row, startCol, 'ПРЯМОЕ СРАВНЕНИЕ', STYLES.section);
  mergeRange(ws, row, startCol, row, startCol + 2);
  for (let c = startCol + 1; c <= startCol + 2; c++) ensureCell(ws, row, c).s = STYLES.section;
  row++;

  setCell(ws, row, startCol, 'Название', STYLES.headerCenter);
  setCell(ws, row, startCol + 1, 'Нравится больше всего', STYLES.headerCenter);
  setCell(ws, row, startCol + 2, 'Куплю в первую очередь', STYLES.headerCenter);
  row++;

  concepts.forEach((c, i) => {
    setCell(ws, row, startCol, c.label, STYLES.label);
    if (mode === 'green') {
      setCell(ws, row, startCol + 1, stdRes.direct.likeMost.perConcept[i] || 0, signifRes.directMax.likeMost[i] ? STYLES.percentGreen : STYLES.percent);
      setCell(ws, row, startCol + 2, stdRes.direct.buyFirst.perConcept[i] || 0, signifRes.directMax.buyFirst[i] ? STYLES.percentGreen : STYLES.percent);
    } else {
      setCell(ws, row, startCol + 1, Math.round((stdRes.direct.likeMost.perConcept[i] || 0) * 100) + '%', STYLES.stringAccent);
      setCell(ws, row, startCol + 2, Math.round((stdRes.direct.buyFirst.perConcept[i] || 0) * 100) + '%', STYLES.stringAccent);
    }
    row++;
  });

  setCell(ws, row, startCol, 'Ни одно из них', STYLES.label);
  if (mode === 'green') {
    setCell(ws, row, startCol + 1, stdRes.direct.likeMost.none || 0, STYLES.percent);
    setCell(ws, row, startCol + 2, stdRes.direct.buyFirst.none || 0, STYLES.percent);
  } else {
    setCell(ws, row, startCol + 1, Math.round((stdRes.direct.likeMost.none || 0) * 100) + '%', STYLES.stringAccent);
    setCell(ws, row, startCol + 2, Math.round((stdRes.direct.buyFirst.none || 0) * 100) + '%', STYLES.stringAccent);
  }
  row++;

  return row;
}

function makeSignifSheetStyled(stdRes, concepts, signifRes) {
  const ws = {};

  ws['!cols'] = [
    { wch: 40 },
    { wch: 15 },
    { wch: 15 },
    { wch: 15 },
    { wch: 15 },
    { wch: 4 },
    { wch: 40 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 }
  ];

  const endLeft = writeSignifBlock(ws, 0, 0, stdRes, concepts, signifRes, 'green');
  const endRight = writeSignifBlock(ws, 0, 6, stdRes, concepts, signifRes, 'letters');

  applyPercentFormat(ws, 0);
  applySheetRangeRef(ws, Math.max(endLeft, endRight), 10);
  return ws;
}

// ---------- AUDIENCE ----------

function makeAudienceSheet(audienceRes) {
  const ws = [];

  function addFreq(title, rows) {
    ws.push([title]);
    ws.push(['Категория', 'Доля']);
    rows.forEach(r => ws.push(r));
    ws.push([]);
  }

  ws.push(['АУДИТОРИЯ']);
  ws.push([`База: n=${audienceRes.n} респондентов, значения в долях (0–1)`]);
  ws.push([]);

  addFreq('Частота взятия новинок (горячие напитки)', audienceRes.freqNew);
  addFreq('Частота покупки капучино', audienceRes.freqProd);
  addFreq('Пол', audienceRes.sex);
  addFreq('Возраст', audienceRes.age);
  addFreq('Частота посещения Бургер Кинг', audienceRes.freqBK);

  return XLSX.utils.aoa_to_sheet(ws);
}
