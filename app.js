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
let parsed = null;      // { header, rows }
let autoMapping = null; // результат автопоиска
let userConfig = null;  // итоговая конфигурация после UI

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

      const summarySheet = makeSummarySheet(stdResults, extraResults, concepts);
      applyPercentFormat(summarySheet, 0);
      XLSX.utils.book_append_sheet(outWb, summarySheet, 'САММАРИ');

      const fullSheet = makeFullSheet(stdResults, extraResults, concepts);
      applyPercentFormat(fullSheet, 0);
      XLSX.utils.book_append_sheet(outWb, fullSheet, 'полные таблицы');

      const signifSheet = makeSignifSheet(signifRes, concepts);
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
    const t = (h || '').toLowerCase();

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

// ---------- UI: стандартные вопросы ----------

function renderStandardMappingUI(mapping, header) {
  const groups = [
    { key: 'like', label: 'Нравится название (шкала 1–5, Top‑2)', indexes: mapping.std.like },
    { key: 'fitDish', label: 'Подходит для блюда (шкала 1–5, Top‑2)', indexes: mapping.std.fitDish },
    { key: 'fitBrand', label: 'Подходит для бренда (шкала 1–5, Top‑2)', indexes: mapping.std.fitBrand },
    { key: 'visitBK', label: 'Намерение посетить БК (шкала 1–5, Top‑2)', indexes: mapping.std.visitBK },
    { key: 'buyDish', label: 'Намерение купить блюдо (шкала 1–5, Top‑2)', indexes: mapping.std.buyDish },
    { key: 'directLike', label: 'Прямое сравнение: нравится больше всего (single choice)', indexes: mapping.std.directLike },
    { key: 'directBuy', label: 'Прямое сравнение: куплю в первую очередь (single choice)', indexes: mapping.std.directBuy }
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
        item.innerHTML = `
          <input type="checkbox" id="${id}" data-std-key="${group.key}" data-col-idx="${idx}" checked>
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

// ---------- UI: доп. вопросы ----------

function renderExtraQuestionsUI(mapping, header) {
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

// ---------- СБОР КОНФИГА ИЗ UI ----------

function collectUserConfig(mapping, header) {
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
    if (parts.length > 1) {
      return parts[parts.length - 1].trim();
    }
    return text;
  });

  return labels.map((lab, i) => ({
    code: String.fromCharCode(65 + i),
    label: lab || `Название ${String.fromCharCode(65 + i)}`
  }));
}

// ---------- ВСПОМОГАТЕЛЬНОЕ ДЛЯ ДАННЫХ ----------

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

// ---------- РАСЧЁТ СТАНДАРТНЫХ МЕТРИК ----------

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
        if (v >= 1 && v <= 5) {
          const k = String(v);
          res[i][k] += 1;
        }
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

// ---------- РАСЧЁТ ДОП. ВОПРОСОВ ----------

function calcExtraBlocks(rows, config, concepts, header) {
  const n = rows.length;
  const result = [];

  config.extra.forEach(q => {
    const idx = q.idx;
    const type = q.type;

    if (type === 'scale5') {
      const counts = { '1':0,'2':0,'3':0,'4':0,'5':0 };

      rows.forEach(r => {
        const v = parseScaleValue(getCell(r, idx));
        if (v >= 1 && v <= 5) {
          counts[String(v)] += 1;
        }
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

// ---------- АУДИТОРИЯ ----------

function calcAudience(rows, config, header) {
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

// ---------- Z-ТЕСТ ДЛЯ ТОП-2 ----------

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
    extra: {}
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

  extraRes.forEach(er => {
    if (!er.where.includes('signif')) return;
    if (er.kind !== 'scale5') return;
    signif.extra[er.title] = er.dist.top2;
  });

  return signif;
}

// ---------- ФОРМАТИРОВАНИЕ EXCEL ----------

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

      if (cell.v >= 0 && cell.v <= 1) {
        formatPercentCell(cell, decimals);
      }
    }
  }

  return ws;
}

// ---------- ФОРМИРОВАНИЕ ЛИСТОВ ----------

function makeSummarySheet(stdRes, extraRes, concepts) {
  const ws = [];

  ws.push(['САММАРИ: ТОП-2 (сумма оценок 4 и 5)']);
  ws.push([]);
  ws.push(['Вариант названий', ...concepts.map(c => c.label)]);
  ws.push([`База: n=${stdRes.n} респондентов, все значения в долях (0–1)`]);
  ws.push([]);

  ws.push(['ОСНОВНЫЕ ПОКАЗАТЕЛИ']);
  ws.push(['Показатель', ...concepts.map(c => c.label)]);
  ws.push(['Нравится название (Top‑2)', ...stdRes.top2.like]);
  ws.push(['Подходит для блюда (Top‑2)', ...stdRes.top2.fitDish]);
  ws.push(['Подходит для бренда (Top‑2)', ...stdRes.top2.fitBrand]);
  ws.push(['Намерение посетить БК (Top‑2)', ...stdRes.top2.visitBK]);
  ws.push(['Намерение купить (Top‑2)', ...stdRes.top2.buyDish]);
  ws.push([]);

  ws.push(['Прямое сравнение']);
  ws.push(['Показатель', ...concepts.map(c => c.label), 'Ни одно из них']);
  ws.push([
    'Нравится больше всего',
    ...stdRes.direct.likeMost.perConcept,
    stdRes.direct.likeMost.none
  ]);
  ws.push([
    'Куплю в первую очередь',
    ...stdRes.direct.buyFirst.perConcept,
    stdRes.direct.buyFirst.none
  ]);
  ws.push([]);

  ws.push(['ИМИДЖЕВЫЙ БЛОК']);
  ws.push(['Показатель', ...concepts.map(c => c.label)]);

  Object.entries(stdRes.image).forEach(([k, vals]) => {
    ws.push([k, ...vals]);
  });

  const summaryExtras = extraRes.filter(x => x.where.includes('summary'));
  if (summaryExtras.length) {
    ws.push([]);
    ws.push(['ДОПОЛНИТЕЛЬНЫЕ МЕТРИКИ']);

    summaryExtras.forEach(er => {
      if (er.kind === 'scale5') {
        ws.push([er.title, er.dist.top2]);
      } else {
        ws.push([er.title]);
        ws.push(['Категория', 'Доля']);
        er.dist.forEach(row => ws.push([row.cat, row.p]));
      }
      ws.push([]);
    });
  }

  return XLSX.utils.aoa_to_sheet(ws);
}

function makeFullSheet(stdRes, extraRes, concepts) {
  const ws = [];

  function addScaleBlock(title, distArr) {
    ws.push([title]);
    ws.push(['Показатель', ...concepts.map(c => c.label)]);
    ws.push(['Top‑2 (4+5)', ...distArr.map(d => d.top2)]);
    ws.push(['1', ...distArr.map(d => d['1'])]);
    ws.push(['2', ...distArr.map(d => d['2'])]);
    ws.push(['3', ...distArr.map(d => d['3'])]);
    ws.push(['4', ...distArr.map(d => d['4'])]);
    ws.push(['5', ...distArr.map(d => d['5'])]);
    ws.push([]);
  }

  addScaleBlock('Нравится название', stdRes.scales.like);
  addScaleBlock('Подходит для блюда', stdRes.scales.fitDish);
  addScaleBlock('Подходит для бренда', stdRes.scales.fitBrand);
  addScaleBlock('Намерение посетить БК', stdRes.scales.visitBK);
  addScaleBlock('Намерение купить', stdRes.scales.buyDish);

  ws.push(['ИМИДЖЕВЫЙ БЛОК']);
  ws.push(['Показатель', ...concepts.map(c => c.label)]);
  Object.entries(stdRes.image).forEach(([k, vals]) => {
    ws.push([k, ...vals]);
  });
  ws.push([]);

  ws.push(['КУПИЛИ БЫ В ПЕРВУЮ ОЧЕРЕДЬ']);
  ws.push(['Название', 'нравится', 'куплю']);

  concepts.forEach((c, i) => {
    ws.push([
      c.label,
      stdRes.direct.likeMost.perConcept[i] || 0,
      stdRes.direct.buyFirst.perConcept[i] || 0
    ]);
  });

  ws.push(['Ни одно из них', stdRes.direct.likeMost.none || 0, stdRes.direct.buyFirst.none || 0]);
  ws.push([]);

  const fullExtras = extraRes.filter(x => x.where.includes('full'));
  if (fullExtras.length) {
    ws.push(['ДОПОЛНИТЕЛЬНЫЕ МЕТРИКИ']);
    fullExtras.forEach(er => {
      if (er.kind === 'scale5') {
        ws.push([er.title]);
        ws.push(['Показатель', 'Доля']);
        ws.push(['Top‑2 (4+5)', er.dist.top2]);
        ws.push(['1', er.dist['1']]);
        ws.push(['2', er.dist['2']]);
        ws.push(['3', er.dist['3']]);
        ws.push(['4', er.dist['4']]);
        ws.push(['5', er.dist['5']]);
      } else {
        ws.push([er.title]);
        ws.push(['Категория', 'Доля']);
        er.dist.forEach(row => ws.push([row.cat, row.p]));
      }
      ws.push([]);
    });
  }

  return XLSX.utils.aoa_to_sheet(ws);
}

function makeSignifSheet(signifRes, concepts) {
  const ws = [];

  ws.push(['ЗНАЧИМОСТИ (z‑тест, альфа=0.05, отмечены более сильные концепты)']);
  ws.push([]);

  const titles = {
    like: 'Нравится название',
    fitDish: 'Подходит для блюда',
    fitBrand: 'Подходит для бренда',
    visitBK: 'Намерение посетить БК',
    buyDish: 'Намерение купить'
  };

  Object.entries(titles).forEach(([key, title]) => {
    if (!signifRes.top2[key]) return;

    ws.push([title]);
    ws.push(['Концепция', 'Сильнее (значимо выше Top‑2, коды концепций)']);

    concepts.forEach((c, i) => {
      const stronger = (signifRes.top2[key][i] || []).join(', ');
      ws.push([`${c.label} (${c.code})`, stronger]);
    });

    ws.push([]);
  });

  if (Object.keys(signifRes.extra).length) {
    ws.push(['ДОП. МЕТРИКИ']);
    Object.entries(signifRes.extra).forEach(([title, val]) => {
      ws.push([title, val]);
    });
  }

  return XLSX.utils.aoa_to_sheet(ws);
}

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
