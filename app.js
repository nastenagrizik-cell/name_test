// app-3-fixed-v3.js

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

function debugLine(...parts) {
  const msg = parts
    .map(x => {
      if (typeof x === 'string') return x;
      try {
        return JSON.stringify(x);
      } catch (_e) {
        return String(x);
      }
    })
    .join(' ');
  console.log('[DEBUG]', ...parts);
  if (statusEl) {
    const prev = statusEl.textContent || '';
    statusEl.textContent = prev ? prev + '\n[DEBUG] ' + msg : '[DEBUG] ' + msg;
  }
}

function normalizeText(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[«»\"']/g, '')
    .replace(/\"/g, '')
    .replace(/[–—]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function looksLikeFitDishQuestion(h) {
  const t = normalizeText(h);
  if (t.includes('для бренда бургер кинг')) return false;

  const hasStem =
    t.includes('насколько каждое из этих названий') &&
    t.includes('подходит') &&
    t.includes('не подходит');

  const hasProductContext =
    t.includes('для этого') ||
    t.includes('для такого') ||
    t.includes('для воппера') ||
    t.includes('для такого воппера') ||
    t.includes('для бургера') ||
    t.includes('для такого бургера') ||
    t.includes('для напитка') ||
    t.includes('для такого напитка') ||
    t.includes('для капучино') ||
    t.includes('для такого капучино') ||
    t.includes('для набора') ||
    t.includes('для такого набора') ||
    t.includes('для соуса') ||
    t.includes('для такого соуса') ||
    t.includes('для продукта') ||
    t.includes('для такого продукта') ||
    t.includes('для блюда') ||
    t.includes('для этого блюда');

  return hasStem && hasProductContext;
}

function looksLikeShareIntentQuestion(h) {
  const t = normalizeText(h);
  return (
    (t.includes('для каждого названия') ||
      t.includes('оцените') ||
      t.includes('насколько вероятно')) &&
    (t.includes('расскажете') ||
      t.includes('рассказали бы') ||
      t.includes('рассказать') ||
      t.includes('поделитесь') ||
      t.includes('поделиться')) &&
    (t.includes('соцсет') ||
      t.includes('социальных сет') ||
      t.includes('друзьям'))
  );
}

function looksLikeDirectShareQuestion(h) {
  const t = normalizeText(h);
  return (
    (t.includes('с каким из этих названий') ||
      t.includes('какое из этих названий')) &&
    (t.includes('рассказали') ||
      t.includes('рассказала') ||
      t.includes('рассказать') ||
      t.includes('поделились') ||
      t.includes('поделиться')) &&
    (t.includes('в первую очередь') || t.includes('сначала'))
  );
}

const IMAGE_STATEMENT_CONFIG = [
  {
    key: 'Это оригинальный, необычный продукт',
    aliases: [
      'это оригинальный, необычный соус',
      'это оригинальный, необычный продукт',
      'это оригинальный, необычный бургер',
      'оригинальный, необычный',
    ],
  },
  {
    key: 'Ассоциируется со знакомым вкусом',
    aliases: [
      'этот соус ассоциируется со знакомым вкусом',
      'этот бургер ассоциируется со знакомым вкусом',
      'этот продукт ассоциируется со знакомым вкусом',
    ],
  },
  {
    key: 'Добавляет премиальности',
    aliases: [
      'это премиальный соус',
      'это премиальный продукт',
      'добавляет премиальности',
    ],
  },
  {
    key: 'Продукт с юмором',
    aliases: [
      'соус с юмором',
      'бургер с юмором',
      'продукт с юмором',
    ],
  },
  {
    key: 'Хочется попробовать',
    aliases: [
      'хочется попробовать воппер с таким соусом',
      'хочется попробовать такой бургер',
      'хочется попробовать такой капучино',
      'хочется попробовать такой напиток',
      'хочется попробовать такой продукт',
    ],
  },
  {
    key: 'Ассоциируется с приятным вкусом',
    aliases: [
      'этот соус ассоциируется с приятным вкусом',
      'этот бургер ассоциируется с приятным вкусом',
      'этот продукт ассоциируется с приятным вкусом',
    ],
  },
  {
    key: 'Уникальная новинка',
    aliases: [
      'это уникальная новинка',
      'это уникальный бургер',
      'это уникальный продукт',
      'оригинальное, отличается от других',
      'это уникальный напиток',
    ],
  },
  {
    key: 'Понятное и простое название',
    aliases: ['понятное и простое название'],
  },
  {
    key: 'Вызывает отторжение',
    aliases: [
      'этот соус вызывает отторжение',
      'этот продукт вызывает отторжение',
      'этот бургер вызывает отторжение',
    ],
  },
  {
    key: 'Понятно, какой будет вкус',
    aliases: [
      'понятно, с каким вкусом будет этот бургер',
      'понятно, с каким вкусом будет этот соус',
      'понятно, с каким вкусом будет этот продукт',
      'понятно какой будет вкус',
    ],
  },
  {
    key: 'Название легко запомнить',
    aliases: ['название легко запомнить'],
  },
  {
    key: 'Вызывает аппетит, звучит вкусно',
    aliases: [
      'вызывает аппетит',
      'вызывает аппетит, вкусно звучит',
      'вызывает аппетит, звучит вкусно',
      'название звучит вкусно и аппетитно',
    ],
  },
  {
    key: 'Вызывает доверие',
    aliases: ['вызывает у меня доверие', 'вызывает доверие'],
  },
  {
    key: 'Звучит как натуральный продукт',
    aliases: ['звучит как натуральный продукт'],
  },
  {
    key: 'Звучит как качественный продукт',
    aliases: ['звучит как качественный продукт'],
  },
  {
    key: 'По названию понятен маленький формат',
    aliases: [
      'по названию понятно, что это мини-бургеры',
      'по названию понятно, что это маленький формат',
      'по названию понятно, что это мини формат',
      'маленький формат',
    ],
  },
  {
    key: 'По названию понятно, что внутри несколько вкусов',
    aliases: [
      'по названию понятно, что внутри несколько разных бургеров',
      'по названию понятно, что внутри несколько разных вкусов',
      'внутри несколько разных вкусов',
    ],
  },
  {
    key: 'Хорошо передает идею набора',
    aliases: [
      'это название хорошо передает идею набора',
      'хорошо передает идею набора',
    ],
  },
  {
    key: 'Звучит хайпово и трендово',
    aliases: ['звучит хайпово и трендово'],
  },
  {
    key: 'Люди бы обсуждали такое название',
    aliases: ['люди бы обсуждали такое название'],
  },
  {
    key: 'Название звучит странно или отталкивающе',
    aliases: ['название звучит странно или отталкивающе'],
  },
  {
    key: 'Название из детского меню / продукт для детей',
    aliases: ['название из детского меню', 'продукт для детей'],
  },
  {
    key: 'Не сытно / не наешься',
    aliases: ['не сытно', 'не наешься'],
  },
  {
    key: 'Дешевый продукт',
    aliases: ['дешевый продукт'],
  },
  {
    key: 'Стоит своих денег',
    aliases: ['стоит своих денег'],
  },
  {
    key: 'Подходит для группового потребления',
    aliases: ['подходит для группового потребления'],
  },
  {
    key: 'Звучит старомодно',
    aliases: ['звучит старомодно'],
  },
];

function detectImageStatementKey(headerText) {
  const t = normalizeText(headerText);
  for (const item of IMAGE_STATEMENT_CONFIG) {
    if (item.aliases.some(alias => t.includes(normalizeText(alias)))) {
      return item.key;
    }
  }
  return null;
}

// === Инициализация и обработка файла ===

if (typeof XLSX !== 'object') {
  if (statusEl) {
    statusEl.textContent =
      'Ошибка: библиотека XLSX не загружена. Попробуйте обновить страницу.';
    statusEl.className = 'status error';
  }
} else if (!baseInput) {
  if (statusEl) {
    statusEl.textContent =
      'Ошибка: на странице не найден input с id="baseFile".';
    statusEl.className = 'status error';
  }
} else {
  baseInput.addEventListener('click', () => {
    baseInput.value = '';
  });

  baseInput.addEventListener('change', async e => {
    debugLine('change fired');
    baseFile = e.target.files[0] || null;
    debugLine(
      'baseFile =',
      baseFile
        ? { name: baseFile.name, size: baseFile.size, type: baseFile.type }
        : null,
    );
    resetState();
    debugLine('resetState done');

    if (!baseFile) {
      status('Файл не выбран');
      debugLine('no file selected');
      return;
    }

    status('Читаю базу...\nЭто может занять до минуты.');
    debugLine('starting read');

    try {
      const arrayBuffer = await baseFile.arrayBuffer();
      debugLine('arrayBuffer ok, bytes =', arrayBuffer.byteLength);

      const wb = XLSX.read(arrayBuffer, { type: 'array' });
      debugLine('xlsx read ok, sheets =', wb.SheetNames);

      const sheetName = wb.SheetNames[wb.SheetNames.length - 1];
      debugLine('selected sheet =', sheetName);

      const sheet = wb.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: '',
      });
      debugLine('sheet_to_json ok, rows =', data.length);

      const { header, rows } = splitHeaderRows(data);
      debugLine(
        'splitHeaderRows ok, header len = ' + header.length + ', rows len = ' + rows.length,
      );
      debugLine('header sample = ' + JSON.stringify(header.slice(0, 8)));

      parsed = { header, rows };
      autoMapping = autoDetectMapping(header);
      debugLine(
        'autoDetectMapping ok = ' +
          JSON.stringify({
            like: autoMapping.std.like.length,
            fitDish: autoMapping.std.fitDish.length,
            fitBrand: autoMapping.std.fitBrand.length,
            visitBK: autoMapping.std.visitBK.length,
            buyDish: autoMapping.std.buyDish.length,
            shareIntent: autoMapping.std.shareIntent.length,
            image: autoMapping.std.image.length,
            directLike: autoMapping.std.directLike.length,
            directBuy: autoMapping.std.directBuy.length,
            shareDirect: autoMapping.std.directShare.length,
            extraCandidates: autoMapping.extraCandidates.length,
          }),
      );

      debugLine(
        'matched like headers = ' +
          JSON.stringify(autoMapping.std.like.map(i => header[i]).slice(0, 5)),
      );
      debugLine(
        'matched fitDish headers = ' +
          JSON.stringify(autoMapping.std.fitDish.map(i => header[i]).slice(0, 5)),
      );
      debugLine(
        'matched fitBrand headers = ' +
          JSON.stringify(autoMapping.std.fitBrand.map(i => header[i]).slice(0, 5)),
      );
      debugLine(
        'matched buyDish headers = ' +
          JSON.stringify(autoMapping.std.buyDish.map(i => header[i]).slice(0, 5)),
      );
      debugLine(
        'matched shareIntent headers = ' +
          JSON.stringify(autoMapping.std.shareIntent.map(i => header[i]).slice(0, 5)),
      );
      debugLine(
        'matched extra headers = ' +
          JSON.stringify(autoMapping.extraCandidates.map(x => x.header).slice(0, 10)),
      );

      renderStandardMappingUI(autoMapping, header);
      debugLine('renderStandardMappingUI ok');
      renderExtraQuestionsUI(autoMapping);
      debugLine('renderExtraQuestionsUI ok');

      if (mappingSection) mappingSection.style.display = '';
      if (extraSection) extraSection.style.display = '';
      if (runSection) runSection.style.display = '';
      debugLine('sections shown');

      status(
        'База загружена. Проверьте найденные вопросы и доп.метрики, затем нажмите «Посчитать топлайн».',
        true,
      );
    } catch (e) {
      console.error(e);
      debugLine('READ ERROR =', e && e.message ? e.message : String(e));
      status(
        'Ошибка при чтении файла: ' +
          (e && e.message ? e.message : String(e)),
        false,
        true,
      );
    }
  });

  runBtn.addEventListener('click', () => {
    if (!parsed || !autoMapping) {
      status(
        'Сначала загрузите файл и дождитесь определения вопросов.',
        false,
        true,
      );
      return;
    }

    try {
      userConfig = collectUserConfig(autoMapping);
    } catch (e) {
      status(
        'Нужно завершить настройку вопросов: ' + e.message,
        false,
        true,
      );
      return;
    }

    try {
      runBtn.disabled = true;
      status('Считаю топлайн...\nПодождите, формируется Excel.');

      const { header, rows } = parsed;
      const concepts = inferConcepts(header, userConfig);
      const stdResults = calcStandardBlocks(rows, userConfig, concepts, header);
      const extraResults = calcExtraBlocks(rows, userConfig, concepts, header);
      const audienceRes = calcAudience(rows, userConfig);
      const signifRes = calcSignificance(stdResults, concepts, rows.length);

      const outWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(
        outWb,
        makeSummarySheetStyled(stdResults, concepts, signifRes, extraResults),
        'САММАРИ',
      );
      XLSX.utils.book_append_sheet(
        outWb,
        makeFullSheetStyled(stdResults, concepts, extraResults),
        'полные таблицы',
      );
      XLSX.utils.book_append_sheet(
        outWb,
        makeSignifSheetStyled(stdResults, concepts, signifRes),
        'значимости',
      );
      XLSX.utils.book_append_sheet(
        outWb,
        makeAudienceSheetStyled(audienceRes),
        'Аудитория',
      );

      // Новый расчёт: разбивки по возрасту
      const ageRes = calcAgeBreakdown(rows, userConfig, concepts, header);
      if (ageRes) {
        XLSX.utils.book_append_sheet(
          outWb,
          makeAgeBreakdownSheetStyled(ageRes),
          'разбивки по возрасту',
        );
      }

      const outName =
        'Topline_' + (baseFile.name.replace(/\.[^.]+$/, '') || 'output') + '.xlsx';
      XLSX.writeFile(outWb, outName);

      status('Готово. Файл ' + outName + ' сохранён.', true);
    } catch (e) {
      console.error(e);
      status(
        'Ошибка при расчете: ' +
          (e && e.message ? e.message : String(e)),
        false,
        true,
      );
    } finally {
      runBtn.disabled = false;
    }
  });
}

// === Служебные функции состояния ===

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

// === Парсинг входного листа ===

function splitHeaderRows(data) {
  if (!data || data.length < 2) return { header: [], rows: [] };

  const questionTexts = (data[0] || []).map(v => String(v || '').trim());
  const varNames = (data[1] || []).map(v => String(v || '').trim());

  const header = questionTexts.map((txt, i) => txt || varNames[i] || `col_${i}`);
  const rows = data
    .slice(2)
    .filter(r => r && r.some(v => v !== null && v !== ''));

  return { header, rows };
}

// === Авто-детект вопросов и аудиторных переменных ===

function autoDetectMapping(header) {
  const std = {
    like: [],
    fitDish: [],
    fitBrand: [],
    visitBK: [],
    buyDish: [],
    shareIntent: [],
    image: [],
    directLike: [],
    directBuy: [],
    directShare: [],
    audience: {
      sex: null,
      age: null,
      freqNew: null,
      freqProd: null,
      freqBK: null,
    },
  };

  header.forEach((h, idx) => {
    const text = String(h || '').trim();
    const t = normalizeText(text);
    if (!text) return;

    if (
      text.includes(
        'Оцените, пожалуйста, насколько вам нравится или не нравится каждое из этих названий',
      ) ||
      t.includes('нравится или не нравится каждое из этих названий')
    ) {
      std.like.push(idx);
    }

    if (looksLikeFitDishQuestion(text)) {
      std.fitDish.push(idx);
    }

    if (
      text.includes(
        'А теперь оцените, насколько каждое из этих названий подходит или не подходит для бренда Бургер Кинг',
      ) ||
      t.includes('подходит или не подходит для бренда бургер кинг')
    ) {
      std.fitBrand.push(idx);
    }

    if (
      text.includes(
        'Скажите, насколько вероятно, что Вы посетите ресторан Бургер Кинг',
      ) ||
      t.includes('насколько вероятно, что вы посетите ресторан бургер кинг')
    ) {
      std.visitBK.push(idx);
    }

    if (
      text.includes(
        'Для каждого названия укажите, насколько вероятно, что Вы купите',
      ) ||
      t.includes('для каждого названия укажите, насколько вероятно, что вы купите')
    ) {
      std.buyDish.push(idx);
    }

    if (looksLikeShareIntentQuestion(text)) {
      std.shareIntent.push(idx);
    }

    const imageKey = detectImageStatementKey(text);
    if (imageKey) {
      std.image.push({ key: imageKey, idx });
    }

    if (
      t.includes('какое из перечисленных ниже названий') ||
      t.includes('какое из этих названий')
    ) {
      std.directLike.push(idx);
    }

    if (
      t.includes('с каким из этих названий вы бы купили') ||
      t.includes('с каким из этих названий вы купили бы') ||
      (t.includes('с каким из этих названий') &&
        t.includes('купили') &&
        t.includes('в первую очередь'))
    ) {
      std.directBuy.push(idx);
    }

    if (looksLikeDirectShareQuestion(text)) {
      std.directShare.push(idx);
    }

    if (text.includes('Укажите Ваш пол')) std.audience.sex = idx;
    if (text.includes('Укажите Ваш возраст')) std.audience.age = idx;
    if (
      text.includes('Как часто Вы берете новинки') ||
      text.includes('Как часто вы берете новинки')
    )
      std.audience.freqNew = idx;
    if (
      (text.includes('Как часто вы покупаете') ||
        text.includes('Как часто Вы покупаете')) &&
      std.audience.freqProd == null
    )
      std.audience.freqProd = idx;
    if (text.includes('Как часто вы посещаете Бургер Кинг'))
      std.audience.freqBK = idx;
  });

  const used = new Set(
    [
      ...std.like,
      ...std.fitDish,
      ...std.fitBrand,
      ...std.visitBK,
      ...std.buyDish,
      ...std.shareIntent,
      ...std.image.map(x => x.idx),
      ...std.directLike,
      ...std.directBuy,
      ...std.directShare,
      std.audience.sex,
      std.audience.age,
      std.audience.freqNew,
      std.audience.freqProd,
      std.audience.freqBK,
    ].filter(v => v !== null && v !== undefined),
  );

  const extraCandidates = header
    .map((h, idx) => {
      if (used.has(idx)) return null;
      if (!h) return null;
      const lower = normalizeText(h);
      const looksClosed =
        lower.includes('насколько') ||
        lower.includes('оцените') ||
        lower.includes('выберите') ||
        lower.includes('какое из перечисленных') ||
        lower.includes('с каким из этих названий') ||
        lower.includes('насколько вероятно') ||
        lower.includes('какой из этих') ||
        lower.includes('что из перечисленного');
      if (!looksClosed) return null;
      return { idx, header: h };
    })
    .filter(Boolean);

  return { std, extraCandidates };
}

// === UI для выбора вопросов ===

function renderStandardMappingUI(mapping, header) {
  const groups = [
    {
      key: 'like',
      label: 'Нравится название (шкала 1–5, Top‑2)',
      indexes: mapping.std.like,
    },
    {
      key: 'fitDish',
      label: 'Подходит для блюда / продукта (шкала 1–5, Top‑2)',
      indexes: mapping.std.fitDish,
    },
    {
      key: 'fitBrand',
      label: 'Подходит для бренда (шкала 1–5, Top‑2)',
      indexes: mapping.std.fitBrand,
    },
    {
      key: 'visitBK',
      label: 'Намерение посетить БК (шкала 1–5, Top‑2)',
      indexes: mapping.std.visitBK,
    },
    {
      key: 'buyDish',
      label: 'Намерение купить (шкала 1–5, Top‑2)',
      indexes: mapping.std.buyDish,
    },
    {
      key: 'shareIntent',
      label: 'Намерение рассказать / поделиться (шкала 1–5, Top‑2)',
      indexes: mapping.std.shareIntent,
    },
    {
      key: 'directCompare',
      label: 'Прямое сравнение',
      indexes: [
        ...mapping.std.directLike,
        ...mapping.std.directBuy,
        ...mapping.std.directShare,
      ],
    },
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
      empty.innerHTML = 'Колонки не найдены по ключевым словам';
      list.appendChild(empty);
    } else {
      group.indexes.forEach(idx => {
        const item = document.createElement('div');
        item.className = 'mapping-item';
        const id = `std-${group.key}-${idx}`;
        let stdKey = group.key;
        if (group.key === 'directCompare') {
          if (mapping.std.directLike.includes(idx)) stdKey = 'directLike';
          else if (mapping.std.directBuy.includes(idx)) stdKey = 'directBuy';
          else stdKey = 'directShare';
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
    extraQuestionsEl.innerHTML =
      '<div class="status">Доп.метрики не найдены по ключевым словам.</div>';
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
          <span>${q.header}</span>
        </label>
      </div>
      <div class="row">
        <div class="col-half">
          <div class="field">
            <label>Как назвать метрику в отчете?</label>
            <input type="text" data-extra-idx="${q.idx}" data-role="title" placeholder="Название метрики">
          </div>
          <div class="field">
            <label>Тип вопроса</label>
            <select data-extra-idx="${q.idx}" data-role="type">
              <option value="scale5">Шкала 1–5</option>
              <option value="single">Выбор одного варианта</option>
            </select>
          </div>
        </div>
        <div class="col-half">
          <div class="field">
            <label>Где показывать эту метрику</label>
            <div class="pill-checkboxes" data-extra-idx="${q.idx}" data-role="where">
              <label><input type="checkbox" value="summary" checked> Саммари</label>
              <label><input type="checkbox" value="full" checked> Полные таблицы</label>
              <label><input type="checkbox" value="signif"> Значимости</label>
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
    shareIntent: [],
    directLike: [],
    directBuy: [],
    directShare: [],
    image: mapping.std.image.slice(),
    audience: mapping.std.audience,
  };

  document
    .querySelectorAll('input[type=checkbox][data-std-key]')
    .forEach(cb => {
      if (!cb.checked) return;
      const key = cb.getAttribute('data-std-key');
      const idx = Number(cb.getAttribute('data-col-idx'));
      stdSelected[key].push(idx);
    });

  const extra = [];

  mapping.extraCandidates.forEach(q => {
    const enabledCb = document.querySelector(
      `input[type=checkbox][data-extra-idx="${q.idx}"]`,
    );
    if (!enabledCb || !enabledCb.checked) return;

    const titleInput = document.querySelector(
      `input[data-extra-idx="${q.idx}"][data-role="title"]`,
    );
    const typeSelect = document.querySelector(
      `select[data-extra-idx="${q.idx}"][data-role="type"]`,
    );
    const whereWrap = document.querySelector(
      `div[data-extra-idx="${q.idx}"][data-role="where"]`,
    );

    const title = titleInput?.value?.trim();
    if (!title) {
      throw new Error(
        `Для доп.метрики "${q.header}" нужно задать название для отчета.`,
      );
    }

    const qtype = typeSelect?.value || 'scale5';

    const where = [];
    whereWrap
      .querySelectorAll('input[type=checkbox]')
      .forEach(cb => cb.checked && where.push(cb.value));
    if (!where.length) {
      throw new Error(
        `Для доп.метрики "${title}" нужно выбрать хотя бы один блок отчета.`,
      );
    }

    extra.push({
      idx: q.idx,
      header: q.header,
      title,
      type: qtype,
      where,
    });
  });

  return { std: stdSelected, extra };
}

// === Концепты / названия ===

function inferConcepts(header, config) {
  const sourceCols =
    config.std.like.length
      ? config.std.like
      : config.std.fitDish.length
      ? config.std.fitDish
      : config.std.fitBrand.length
      ? config.std.fitBrand
      : config.std.buyDish.length
      ? config.std.buyDish
      : config.std.shareIntent.length
      ? config.std.shareIntent
      : config.std.visitBK.length
      ? config.std.visitBK
      : [];

  if (!sourceCols.length) {
    return [
      { code: 'A', label: 'A' },
      { code: 'B', label: 'B' },
    ];
  }

  const labels = sourceCols.map(colIdx => {
    const text = String(header[colIdx] || '').trim();
    const parts = text.split(' - ');
    return parts.length > 1 ? parts[parts.length - 1].trim() : text;
  });

  return labels.map((label, i) => ({
    code: String.fromCharCode(65 + i),
    label,
  }));
}

// === Работа с ячейками и шкалами ===

function getCell(row, idx) {
  if (idx == null || idx < 0) return null;
  return row[idx];
}

function parseScaleValue(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return v;
  const s = String(v).trim();
  const m = s.match(/^([1-5])$/);
  return m ? Number(m[1]) : null;
}

// === Расчет стандартных блоков ===

function normalizeConceptLabel(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[^a-zа-я0-9]+/g, ' ')
    .trim();
}

function findConceptIndexByHeader(headerText, concepts) {
  const text = normalizeText(headerText);
  for (let i = 0; i < concepts.length; i++) {
    if (text.endsWith(normalizeText(concepts[i].label))) return i;
  }
  for (let i = 0; i < concepts.length; i++) {
    if (text.includes(normalizeText(concepts[i].label))) return i;
  }

  const normalizedHeader = normalizeConceptLabel(text);
  for (let i = 0; i < concepts.length; i++) {
    const conceptNorm = normalizeConceptLabel(concepts[i].label);
    if (conceptNorm && normalizedHeader.includes(conceptNorm)) return i;
  }

  return -1;
}

function calcStandardBlocks(rows, config, concepts, header) {
  const n = rows.length;

  function top2ByCols(cols) {
    if (!cols.length) return null;
    const res = Array(concepts.length).fill(0);
    rows.forEach(r => {
      cols.forEach((col, i) => {
        if (i >= concepts.length) return;
        const v = parseScaleValue(getCell(r, col));
        if (v === 4 || v === 5) res[i]++;
      });
    });
    return res.map(v => (n ? v / n : 0));
  }

  function dist5(cols) {
    if (!cols.length) return null;
    const arr = Array.from({ length: concepts.length }, () => ({
      1: 0,
      2: 0,
      3: 0,
      4: 0,
      5: 0,
    }));
    rows.forEach(r => {
      cols.forEach((col, i) => {
        if (i >= concepts.length) return;
        const v = parseScaleValue(getCell(r, col));
        if (v >= 1 && v <= 5) arr[i][String(v)]++;
      });
    });
    return arr.map(d => ({
      1: n ? d['1'] / n : 0,
      2: n ? d['2'] / n : 0,
      3: n ? d['3'] / n : 0,
      4: n ? d['4'] / n : 0,
      5: n ? d['5'] / n : 0,
      top2: n ? (d['4'] + d['5']) / n : 0,
    }));
  }

  function imageBlock() {
    const res = {};
    IMAGE_STATEMENT_CONFIG.forEach(item => {
      res[item.key] = Array(concepts.length).fill(0);
    });

    rows.forEach(r => {
      config.std.image.forEach(({ key, idx }) => {
        const val = String(getCell(r, idx) || '').trim();
        if (!val) return;
        const conceptIndex = findConceptIndexByHeader(header[idx], concepts);
        if (conceptIndex < 0 || conceptIndex >= concepts.length) return;
        res[key][conceptIndex]++;
      });
    });

    Object.keys(res).forEach(k => {
      res[k] = res[k].map(v => (n ? v / n : 0));
    });

    return res;
  }

  function directSingle(cols) {
    if (!cols.length) return null;

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
      const key = Object.keys(counts).find(k =>
        normalizeText(k).includes(normalizeText(c.label)),
      );
      perConcept[i] = key ? counts[key] / n : 0;
    });

    const noneKey = Object.keys(counts).find(k => {
      const t = normalizeText(k);
      return (
        t.includes('не выберу ни одно') ||
        t.includes('ни одно') ||
        t.includes('никакое') ||
        t.includes('не расскажу ни про одно') ||
        t.includes('ни про одно')
      );
    });

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
      buyDish: dist5(config.std.buyDish),
      shareIntent: dist5(config.std.shareIntent),
    },
    top2: {
      like: top2ByCols(config.std.like),
      fitDish: top2ByCols(config.std.fitDish),
      fitBrand: top2ByCols(config.std.fitBrand),
      visitBK: top2ByCols(config.std.visitBK),
      buyDish: top2ByCols(config.std.buyDish),
      shareIntent: top2ByCols(config.std.shareIntent),
    },
    image: imageBlock(),
    direct: {
      likeMost: directSingle(config.std.directLike),
      buyFirst: directSingle(config.std.directBuy),
      shareFirst: directSingle(config.std.directShare),
    },
  };
}

// === Доп.метрики ===

function calcExtraBlocks(rows, config, concepts, header) {
  const result = [];

  const handledScaleTitles = new Set();

  config.extra.forEach(q => {
    if (q.type === 'scale5') {
      if (handledScaleTitles.has(q.title)) return;
      handledScaleTitles.add(q.title);

      const sameTitle = config.extra.filter(
        x => x.type === 'scale5' && x.title === q.title,
      );

      const dists = Array.from({ length: concepts.length }, () => ({
        1: 0,
        2: 0,
        3: 0,
        4: 0,
        5: 0,
        base: 0,
      }));

      let matched = 0;

      sameTitle.forEach(item => {
        const conceptIndex = findConceptIndexByHeader(
          header[item.idx],
          concepts,
        );
        if (conceptIndex < 0 || conceptIndex >= concepts.length) return;
        matched++;

        rows.forEach(r => {
          const v = parseScaleValue(getCell(r, item.idx));
          if (v >= 1 && v <= 5) {
            dists[conceptIndex][String(v)]++;
            dists[conceptIndex].base++;
          }
        });
      });

      if (matched === 0) {
        result.push({
          kind: 'scale5',
          title: q.title,
          where: q.where,
          dist: Array.from({ length: concepts.length }, () => ({
            1: 0,
            2: 0,
            3: 0,
            4: 0,
            5: 0,
            top2: 0,
          })),
        });
        return;
      }

      result.push({
        kind: 'scale5byconcept',
        title: q.title,
        where: q.where,
        dist: dists.map(counts => {
          const n = counts.base || 1;
          return {
            1: counts['1'] / n,
            2: counts['2'] / n,
            3: counts['3'] / n,
            4: counts['4'] / n,
            5: counts['5'] / n,
            top2: (counts['4'] + counts['5']) / n,
          };
        }),
      });
    } else {
      const n = rows.length;
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
          .map(([cat, c]) => ({ cat, p: n ? c / n : 0 }))
          .sort((a, b) => b.p - a.p),
      });
    }
  });

  return result;
}

// === Аудитория ===

function calcAudience(rows, config) {
  const n = rows.length;

  function freq(idx) {
    if (idx == null || idx < 0) return { title: '', rows: [] };
    const counts = {};
    rows.forEach(r => {
      const v = String(getCell(r, idx) || '').trim();
      if (!v) return;
      counts[v] = (counts[v] || 0) + 1;
    });

    const items = Object.entries(counts)
      .map(([label, c]) => ({ label, p: n ? c / n : 0 }))
      .sort((a, b) => b.p - a.p);

    return { title: String(config.std.audience.age === idx ? 'Возраст' : ''), rows: items };
  }

  const sexBlock = freq(config.std.audience.sex);
  sexBlock.title = 'Пол';

  const ageBlock = freq(config.std.audience.age);
  ageBlock.title = 'Возраст';

  const freqNewBlock = freq(config.std.audience.freqNew);
  freqNewBlock.title = 'Частота взятия новинок';

  const freqProdBlock = freq(config.std.audience.freqProd);
  freqProdBlock.title = 'Частота покупки продукта';

  const freqBKBlock = freq(config.std.audience.freqBK);
  freqBKBlock.title = 'Частота посещения Бургер Кинг';

  return {
    n,
    sex: [sexBlock],
    age: [ageBlock],
    freqNew: [freqNewBlock],
    freqProd: [freqProdBlock],
    freqBK: [freqBKBlock],
  };
}

// === Значимости (общие) ===

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
    directMax: {
      likeMost: null,
      buyFirst: null,
      shareFirst: null,
    },
  };

  function labelsFor(arr) {
    if (!arr) return null;
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

  ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish', 'shareIntent'].forEach(
    key => {
      if (!stdRes.top2[key]) return;
      signif.top2[key] = labelsFor(stdRes.top2[key]);
    },
  );

  ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish', 'shareIntent'].forEach(
    key => {
      const distArr = stdRes.scales[key];
      if (!distArr) return;
      const levels = ['top2', '1', '2', '3', '4', '5'];
      const obj = {};
      levels.forEach(level => {
        obj[level] = labelsFor(distArr.map(d => d[level]));
      });
      signif.scales[key] = obj;
    },
  );

  Object.entries(stdRes.image).forEach(([k, vals]) => {
    signif.image[k] = labelsFor(vals);
  });

  function maxMask(arr) {
    if (!arr) return null;
    const max = Math.max(...arr);
    return arr.map(v => (max > 0 && v === max ? max : 0));
  }

  signif.directMax.likeMost = stdRes.direct.likeMost
    ? maxMask(stdRes.direct.likeMost.perConcept)
    : null;
  signif.directMax.buyFirst = stdRes.direct.buyFirst
    ? maxMask(stdRes.direct.buyFirst.perConcept)
    : null;
  signif.directMax.shareFirst = stdRes.direct.shareFirst
    ? maxMask(stdRes.direct.shareFirst.perConcept)
    : null;

  return signif;
}

// === Стили и утилиты листов ===

function cellRef(r, c) {
  return XLSX.utils.encode_cell({ r, c });
}

function setCell(ws, r, c, value, style = null) {
  const addr = cellRef(r, c);
  ws[addr] = {
    t: typeof value === 'number' ? 'n' : 's',
    v: value,
  };
  if (style) ws[addr].s = JSON.parse(JSON.stringify(style));
  return ws[addr];
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
    e: { r: eRow, c: eCol },
  });
}

function applySheetRangeRef(ws, endRow, endCol) {
  ws['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: endRow, c: endCol },
  });
}

function borderAll() {
  return {
    top: { style: 'thin', color: { rgb: '000000' } },
    bottom: { style: 'thin', color: { rgb: '000000' } },
    left: { style: 'thin', color: { rgb: '000000' } },
    right: { style: 'thin', color: { rgb: '000000' } },
  };
}

function hexFill(rgb) {
  return {
    patternType: 'solid',
    fgColor: { rgb },
  };
}

const STYLES = {
  title: {
    font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 14 },
    fill: hexFill('244C73'),
    alignment: { horizontal: 'left', vertical: 'center' },
    border: borderAll(),
  },
  section: {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: hexFill('244C73'),
    alignment: { horizontal: 'left', vertical: 'center' },
    border: borderAll(),
  },
  blockTitle: {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: hexFill('5E86B4'),
    alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    border: borderAll(),
  },
  headerCenter: {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: hexFill('244C73'),
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: borderAll(),
  },
  base: {
    font: { italic: true, color: { rgb: '333333' } },
    fill: hexFill('D9D9D9'),
    alignment: { horizontal: 'left', vertical: 'center' },
    border: borderAll(),
  },
  label: {
    alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    border: borderAll(),
  },
  top2Label: {
    font: { bold: true },
    fill: hexFill('DCE6F1'),
    alignment: { horizontal: 'left', vertical: 'center' },
    border: borderAll(),
  },
  top2Row: {
    font: { bold: true },
    fill: hexFill('DCE6F1'),
    alignment: { horizontal: 'center', vertical: 'center' },
    border: borderAll(),
  },
  percent: {
    alignment: { horizontal: 'center', vertical: 'center' },
    border: borderAll(),
  },
  percentGreen: {
    alignment: { horizontal: 'center', vertical: 'center' },
    fill: hexFill('70AD47'),
    border: borderAll(),
  },
  signifTextGreen: {
    font: { bold: true, color: { rgb: '000000' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    fill: hexFill('70AD47'),
    border: borderAll(),
  },
  legendGreen: {
    alignment: { horizontal: 'center', vertical: 'center' },
    fill: hexFill('92D050'),
    border: borderAll(),
  },
  legendAccent: {
    font: { bold: true, color: { rgb: 'C55A11' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    fill: hexFill('FFF2CC'),
    border: borderAll(),
  },
  legendText: {
    alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    border: borderAll(),
  },
  ageHeader: {
    font: { bold: true, color: { rgb: 'FFFFFF' } },
    fill: hexFill('5E86B4'),
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: borderAll(),
  },
  ageLabel: {
    alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
    border: borderAll(),
  },
  agePercent: {
    alignment: { horizontal: 'center', vertical: 'center' },
    border: borderAll(),
  },
  agePercentUp: {
    alignment: { horizontal: 'center', vertical: 'center' },
    fill: hexFill('C6EFCE'),
    border: borderAll(),
  },
  agePercentDown: {
    alignment: { horizontal: 'center', vertical: 'center' },
    fill: hexFill('FFC7CE'),
    border: borderAll(),
  },
};

// === Проверка наличия метрик ===

function isStrong2Plus(arr, index) {
  return Array.isArray(arr) && Array.isArray(arr[index]) && arr[index].length >= 2;
}

function hasMetric(stdRes, key) {
  return !!stdRes.top2[key] && !!stdRes.scales[key];
}

function blockTitleForKey(key) {
  return {
    like: 'Нравится название',
    fitDish: 'Подходит для блюда / продукта',
    fitBrand: 'Подходит для бренда',
    visitBK: 'Намерение посетить БК',
    buyDish: 'Намерение купить',
    shareIntent: 'Намерение поделиться / рассказать',
  }[key] || key;
}

function scaleLabelsForBlock(key) {
  const base = ['1 - совсем не нравится', '2', '3', '4', '5 - очень нравится'];
  return {
    like: ['Top‑2 (4–5)', ...base],
    fitDish: ['Top‑2 (4–5)', ...base],
    fitBrand: ['Top‑2 (4–5)', ...base],
    visitBK: ['Top‑2 (4–5)', ...base],
    buyDish: ['Top‑2 (4–5)', ...base],
    shareIntent: ['Top‑2 (4–5)', ...base],
  }[key] || ['Top‑2', ...base];
}

function blockValueRows(stdRes, key) {
  const distArr = stdRes.scales[key];
  const labels = scaleLabelsForBlock(key);
  return [
    ['Top‑2 (4–5)', distArr.map(d => d.top2), 'top2'],
    [labels[1], distArr.map(d => d['1']), '1'],
    [labels[2], distArr.map(d => d['2']), '2'],
    [labels[3], distArr.map(d => d['3']), '3'],
    [labels[4], distArr.map(d => d['4']), '4'],
    [labels[5], distArr.map(d => d['5']), '5'],
  ];
}

// === Лист САММАРИ ===

function makeSummarySheetStyled(stdRes, concepts, signifRes, extraResults) {
  const ws = {};
  const lastCol = concepts.length;

  ws['!cols'] = [{ wch: 42 }, ...Array.from({ length: concepts.length }, () => ({ wch: 15 }))];

  let row = 0;

  setCell(ws, row, 0, 'Топлайн по названиям (Top‑2, прямые вопросы)', STYLES.title);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  setCell(ws, row, 0, 'Название', STYLES.headerCenter);
  concepts.forEach((c, i) => setCell(ws, row, i + 1, c.label, STYLES.headerCenter));
  row++;

  setCell(ws, row, 0, `n=${stdRes.n}`, STYLES.base);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  setCell(ws, row, 1, 'xx', STYLES.legendGreen);
  setCell(ws, row, 2, '2+ буквы = значимо выше других', STYLES.legendText);
  if (lastCol > 2) mergeRange(ws, row, 2, row, lastCol);
  row++;

  setCell(ws, row, 0, 'Шкальные вопросы (Top‑2)', STYLES.section);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  setCell(ws, row, 0, 'Показатель', STYLES.headerCenter);
  concepts.forEach((c, i) => setCell(ws, row, i + 1, c.label, STYLES.headerCenter));
  row++;

  ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish', 'shareIntent'].forEach(
    key => {
      if (!hasMetric(stdRes, key)) return;
      const vals = stdRes.top2[key];
      setCell(ws, row, 0, blockTitleForKey(key), STYLES.label);
      vals.forEach((v, i) =>
        setPercent(
          ws,
          row,
          i + 1,
          v,
          isStrong2Plus(signifRes.top2[key], i)
            ? STYLES.percentGreen
            : STYLES.percent,
        ),
      );
      row++;
    },
  );

  const hasDirectLike = !!stdRes.direct.likeMost;
  const hasDirectBuy = !!stdRes.direct.buyFirst;
  const hasDirectShare = !!stdRes.direct.shareFirst;

  if (hasDirectLike || hasDirectBuy || hasDirectShare) {
    setCell(ws, row, 0, 'Прямые вопросы (выбор одного названия)', STYLES.section);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    if (hasDirectLike) {
      setCell(ws, row, 0, 'Название нравится больше всего', STYLES.label);
      stdRes.direct.likeMost.perConcept.forEach((v, i) =>
        setPercent(
          ws,
          row,
          i + 1,
          v,
          signifRes.directMax.likeMost?.[i]
            ? STYLES.percentGreen
            : STYLES.percent,
        ),
      );
      row++;
    }

    if (hasDirectBuy) {
      setCell(ws, row, 0, 'С каким названием купили бы в первую очередь', STYLES.label);
      stdRes.direct.buyFirst.perConcept.forEach((v, i) =>
        setPercent(
          ws,
          row,
          i + 1,
          v,
          signifRes.directMax.buyFirst?.[i]
            ? STYLES.percentGreen
            : STYLES.percent,
        ),
      );
      row++;
    }

    if (hasDirectShare) {
      setCell(ws, row, 0, 'Про какое название рассказали бы в первую очередь', STYLES.label);
      stdRes.direct.shareFirst.perConcept.forEach((v, i) =>
        setPercent(
          ws,
          row,
          i + 1,
          v,
          signifRes.directMax.shareFirst?.[i]
            ? STYLES.percentGreen
            : STYLES.percent,
        ),
      );
      row++;
    }
  }

  const imageEntries = Object.entries(stdRes.image);
  if (imageEntries.length) {
    setCell(ws, row, 0, 'Имиджевые характеристики', STYLES.section);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    setCell(ws, row, 0, 'Утверждение', STYLES.headerCenter);
    concepts.forEach((c, i) => setCell(ws, row, i + 1, c.label, STYLES.headerCenter));
    row++;

    imageEntries.forEach(([label, vals]) => {
      setCell(ws, row, 0, label, STYLES.label);
      vals.forEach((v, i) =>
        setPercent(
          ws,
          row,
          i + 1,
          v,
          isStrong2Plus(signifRes.image[label], i)
            ? STYLES.percentGreen
            : STYLES.percent,
        ),
      );
      row++;
    });
  }

  const summaryExtras = extraResults.filter(x => x.where.includes('summary'));
  if (summaryExtras.length) {
    setCell(ws, row, 0, 'Дополнительные метрики (саммари)', STYLES.section);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    summaryExtras.forEach(item => {
      if (item.kind === 'scale5') {
        setCell(ws, row, 0, item.title, STYLES.label);
        setPercent(ws, row, 1, item.dist.top2, STYLES.percent);
        if (lastCol > 1) mergeRange(ws, row, 1, row, lastCol);
        row++;
      } else {
        setCell(ws, row, 0, item.title, STYLES.label);
        const top3 = item.dist.slice(0, 3);
        const txt = top3.map(x => `${x.cat}: ${Math.round(x.p * 100)}%`).join(', ');
        setCell(ws, row, 1, txt, STYLES.label);
        if (lastCol > 1) mergeRange(ws, row, 1, row, lastCol);
        row++;
      }
    });
  }

  applySheetRangeRef(ws, row, lastCol);
  return ws;
}

// === Лист полных таблиц ===

function makeFullSheetStyled(stdRes, concepts, extraResults) {
  const ws = {};
  const lastCol = concepts.length;

  ws['!cols'] = [{ wch: 46 }, ...Array.from({ length: concepts.length }, () => ({ wch: 15 }))];

  let row = 0;

  setCell(ws, row, 0, 'Полные распределения по шкалам и доп.метрикам', STYLES.title);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  setCell(ws, row, 0, 'Название', STYLES.headerCenter);
  concepts.forEach((c, i) => setCell(ws, row, i + 1, c.label, STYLES.headerCenter));
  row++;

  setCell(ws, row, 0, `n=${stdRes.n}`, STYLES.base);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish', 'shareIntent'].forEach(
    key => {
      if (!hasMetric(stdRes, key)) return;

      setCell(ws, row, 0, blockTitleForKey(key), STYLES.blockTitle);
      mergeRange(ws, row, 0, row, lastCol);
      row++;

      blockValueRows(stdRes, key).forEach(([label, vals, level]) => {
        setCell(
          ws,
          row,
          0,
          label,
          level === 'top2' ? STYLES.top2Label : STYLES.label,
        );
        vals.forEach((v, i) =>
          setPercent(
            ws,
            row,
            i + 1,
            v,
            level === 'top2' ? STYLES.top2Row : STYLES.percent,
          ),
        );
        row++;
      });
    },
  );

  const imageEntries = Object.entries(stdRes.image);
  if (imageEntries.length) {
    setCell(ws, row, 0, 'Имиджевые характеристики', STYLES.section);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    imageEntries.forEach(([label, vals]) => {
      setCell(ws, row, 0, label, STYLES.label);
      vals.forEach((v, i) =>
        setPercent(ws, row, i + 1, v, STYLES.percent),
      );
      row++;
    });
  }

  const fullExtras = extraResults.filter(x => x.where.includes('full'));
  if (fullExtras.length) {
    setCell(ws, row, 0, 'Дополнительные метрики', STYLES.section);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    fullExtras.forEach(item => {
      setCell(ws, row, 0, item.title, STYLES.blockTitle);
      mergeRange(ws, row, 0, row, lastCol);
      row++;

      if (item.kind === 'scale5byconcept') {
        const rowsToWrite = [
          ['Top‑2 (4–5)', item.dist.map(d => d.top2), true],
          ['1', item.dist.map(d => d['1']), false],
          ['2', item.dist.map(d => d['2']), false],
          ['3', item.dist.map(d => d['3']), false],
          ['4', item.dist.map(d => d['4']), false],
          ['5', item.dist.map(d => d['5']), false],
        ];
        rowsToWrite.forEach(([label, values, isTop]) => {
          setCell(
            ws,
            row,
            0,
            label,
            isTop ? STYLES.top2Label : STYLES.label,
          );
          values.forEach((value, i) =>
            setPercent(
              ws,
              row,
              i + 1,
              value,
              isTop ? STYLES.top2Row : STYLES.percent,
            ),
          );
          row++;
        });
      } else if (item.kind === 'scale5') {
        const rowsToWrite = [
          ['Top‑2 (4–5)', item.dist.top2],
          ['1', item.dist['1']],
          ['2', item.dist['2']],
          ['3', item.dist['3']],
          ['4', item.dist['4']],
          ['5', item.dist['5']],
        ];
        rowsToWrite.forEach(([label, value]) => {
          setCell(
            ws,
            row,
            0,
            label,
            label.startsWith('Top‑2') ? STYLES.top2Label : STYLES.label,
          );
          setPercent(
            ws,
            row,
            1,
            value,
            label.startsWith('Top‑2') ? STYLES.top2Row : STYLES.percent,
          );
          if (lastCol > 1) mergeRange(ws, row, 1, row, lastCol);
          row++;
        });
      } else {
        item.dist.forEach(x => {
          setCell(ws, row, 0, x.cat, STYLES.label);
          setPercent(ws, row, 1, x.p, STYLES.percent);
          if (lastCol > 1) mergeRange(ws, row, 1, row, lastCol);
          row++;
        });
      }
    });
  }

  const hasDirectLike = !!stdRes.direct.likeMost;
  const hasDirectBuy = !!stdRes.direct.buyFirst;
  const hasDirectShare = !!stdRes.direct.shareFirst;

  if (hasDirectLike || hasDirectBuy || hasDirectShare) {
    const directCols =
      1 +
      (hasDirectLike ? 1 : 0) +
      (hasDirectBuy ? 1 : 0) +
      (hasDirectShare ? 1 : 0);

    setCell(ws, row, 0, 'Прямые вопросы', STYLES.section);
    mergeRange(ws, row, 0, row, directCols - 1);
    row++;

    setCell(ws, row, 0, 'Название', STYLES.headerCenter);
    let hdr = 1;
    if (hasDirectLike)
      setCell(ws, row, hdr++, 'Больше всего нравится', STYLES.headerCenter);
    if (hasDirectBuy)
      setCell(ws, row, hdr++, 'Купили бы в первую очередь', STYLES.headerCenter);
    if (hasDirectShare)
      setCell(ws, row, hdr++, 'Рассказали бы в первую очередь', STYLES.headerCenter);
    row++;

    concepts.forEach((c, i) => {
      setCell(ws, row, 0, c.label, STYLES.label);
      let col = 1;
      if (hasDirectLike)
        setPercent(ws, row, col++, stdRes.direct.likeMost.perConcept[i] || 0, STYLES.percent);
      if (hasDirectBuy)
        setPercent(ws, row, col++, stdRes.direct.buyFirst.perConcept[i] || 0, STYLES.percent);
      if (hasDirectShare)
        setPercent(ws, row, col++, stdRes.direct.shareFirst.perConcept[i] || 0, STYLES.percent);
      row++;
    });

    setCell(ws, row, 0, 'Никто / ни одно', STYLES.label);
    let col = 1;
    if (hasDirectLike)
      setPercent(ws, row, col++, stdRes.direct.likeMost.none || 0, STYLES.percent);
    if (hasDirectBuy)
      setPercent(ws, row, col++, stdRes.direct.buyFirst.none || 0, STYLES.percent);
    if (hasDirectShare)
      setPercent(ws, row, col++, stdRes.direct.shareFirst.none || 0, STYLES.percent);
    row++;
  }

  applySheetRangeRef(ws, row, Math.max(lastCol, 3));
  return ws;
}

// === Лист значимостей ===

function signifCellText(value, letters) {
  const pct = Math.round((value || 0) * 100);
  return letters && letters.length ? `${pct}% (${letters.join('')})` : `${pct}%`;
}

function writeSignifBlock(ws, startRow, startCol, stdRes, concepts, signifRes, mode) {
  let row = startRow;
  const lastCol = startCol + concepts.length;

  setCell(ws, row, startCol, 'Значимые отличия по шкалам', STYLES.title);
  mergeRange(ws, row, startCol, row, lastCol);
  row++;

  setCell(ws, row, startCol, 'Показатель', STYLES.headerCenter);
  concepts.forEach((c, i) =>
    setCell(ws, row, startCol + 1 + i, c.label, STYLES.headerCenter),
  );
  row++;

  setCell(ws, row, startCol, `n=${stdRes.n}`, STYLES.base);
  mergeRange(ws, row, startCol, row, lastCol);
  row++;

  if (mode === 'green') {
    setCell(ws, row, startCol + 1, 'xx', STYLES.legendGreen);
    setCell(
      ws,
      row,
      startCol + 2,
      '2+ букв = значимо выше других',
      STYLES.legendText,
    );
    if (lastCol > startCol + 2)
      mergeRange(ws, row, startCol + 2, row, lastCol);
  } else {
    setCell(ws, row, startCol + 1, 'Б', STYLES.legendAccent);
    setCell(
      ws,
      row,
      startCol + 2,
      'Буквы показывают значимо более высокие варианты',
      STYLES.legendText,
    );
    if (lastCol > startCol + 2)
      mergeRange(ws, row, startCol + 2, row, lastCol);
  }
  row++;

  ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish', 'shareIntent'].forEach(
    key => {
      if (!hasMetric(stdRes, key)) return;

      setCell(ws, row, startCol, blockTitleForKey(key), STYLES.blockTitle);
      mergeRange(ws, row, startCol, row, lastCol);
      row++;

      blockValueRows(stdRes, key).forEach(([label, vals, level]) => {
        setCell(
          ws,
          row,
          startCol,
          label,
          level === 'top2' ? STYLES.top2Label : STYLES.label,
        );
        vals.forEach((v, i) => {
          if (mode === 'green') {
            const style = isStrong2Plus(
              signifRes.scales[key][level],
              i,
            )
              ? STYLES.percentGreen
              : level === 'top2'
              ? STYLES.top2Row
              : STYLES.percent;
            setPercent(ws, row, startCol + 1 + i, v, style);
          } else {
            const letters = signifRes.scales[key][level][i];
            const style =
              letters && letters.length
                ? STYLES.signifTextGreen
                : level === 'top2'
                ? STYLES.top2Row
                : STYLES.percent;
            setCell(
              ws,
              row,
              startCol + 1 + i,
              signifCellText(v, letters),
              style,
            );
          }
        });
        row++;
      });
    },
  );

  const imageEntries = Object.entries(stdRes.image);
  if (imageEntries.length) {
    setCell(ws, row, startCol, 'Имиджевые характеристики', STYLES.section);
    mergeRange(ws, row, startCol, row, lastCol);
    row++;

    imageEntries.forEach(([label, vals]) => {
      setCell(ws, row, startCol, label, STYLES.label);
      vals.forEach((v, i) => {
        if (mode === 'green') {
          setPercent(
            ws,
            row,
            startCol + 1 + i,
            v,
            isStrong2Plus(signifRes.image[label], i)
              ? STYLES.percentGreen
              : STYLES.percent,
          );
        } else {
          const letters = signifRes.image[label][i];
          setCell(
            ws,
            row,
            startCol + 1 + i,
            signifCellText(v, letters),
            letters && letters.length
              ? STYLES.signifTextGreen
              : STYLES.percent,
          );
        }
      });
      row++;
    });
  }

  const hasDirectLike = !!stdRes.direct.likeMost;
  const hasDirectBuy = !!stdRes.direct.buyFirst;
  const hasDirectShare = !!stdRes.direct.shareFirst;

  if (hasDirectLike || hasDirectBuy || hasDirectShare) {
    const directCols =
      1 +
      (hasDirectLike ? 1 : 0) +
      (hasDirectBuy ? 1 : 0) +
      (hasDirectShare ? 1 : 0);

    setCell(ws, row, startCol, 'Прямые вопросы', STYLES.section);
    mergeRange(ws, row, startCol, row, startCol + directCols - 1);
    row++;

    setCell(ws, row, startCol, 'Название', STYLES.headerCenter);
    let hdr = startCol + 1;
    if (hasDirectLike)
      setCell(ws, row, hdr++, 'Нравится больше всего', STYLES.headerCenter);
    if (hasDirectBuy)
      setCell(ws, row, hdr++, 'Купили бы в первую очередь', STYLES.headerCenter);
    if (hasDirectShare)
      setCell(ws, row, hdr++, 'Рассказали бы в первую очередь', STYLES.headerCenter);
    row++;

    concepts.forEach((c, i) => {
      setCell(ws, row, startCol, c.label, STYLES.label);
      if (mode === 'green') {
        let col = startCol + 1;
        if (hasDirectLike)
          setPercent(
            ws,
            row,
            col++,
            stdRes.direct.likeMost.perConcept[i] || 0,
            signifRes.directMax.likeMost?.[i]
              ? STYLES.percentGreen
              : STYLES.percent,
          );
        if (hasDirectBuy)
          setPercent(
            ws,
            row,
            col++,
            stdRes.direct.buyFirst.perConcept[i] || 0,
            signifRes.directMax.buyFirst?.[i]
              ? STYLES.percentGreen
              : STYLES.percent,
          );
        if (hasDirectShare)
          setPercent(
            ws,
            row,
            col++,
            stdRes.direct.shareFirst.perConcept[i] || 0,
            signifRes.directMax.shareFirst?.[i]
              ? STYLES.percentGreen
              : STYLES.percent,
          );
      } else {
        let col = startCol + 1;
        if (hasDirectLike)
          setCell(
            ws,
            row,
            col++,
            Math.round((stdRes.direct.likeMost.perConcept[i] || 0) * 100),
            STYLES.percent,
          );
        if (hasDirectBuy)
          setCell(
            ws,
            row,
            col++,
            Math.round((stdRes.direct.buyFirst.perConcept[i] || 0) * 100),
            STYLES.percent,
          );
        if (hasDirectShare)
          setCell(
            ws,
            row,
            col++,
            Math.round((stdRes.direct.shareFirst.perConcept[i] || 0) * 100),
            STYLES.percent,
          );
      }
      row++;
    });

    setCell(ws, row, startCol, 'Никто / ни одно', STYLES.label);
    if (mode === 'green') {
      let col = startCol + 1;
      if (hasDirectLike)
        setPercent(ws, row, col++, stdRes.direct.likeMost.none || 0, STYLES.percent);
      if (hasDirectBuy)
        setPercent(ws, row, col++, stdRes.direct.buyFirst.none || 0, STYLES.percent);
      if (hasDirectShare)
        setPercent(ws, row, col++, stdRes.direct.shareFirst.none || 0, STYLES.percent);
    } else {
      let col = startCol + 1;
      if (hasDirectLike)
        setCell(
          ws,
          row,
          col++,
          Math.round((stdRes.direct.likeMost.none || 0) * 100),
          STYLES.percent,
        );
      if (hasDirectBuy)
        setCell(
          ws,
          row,
          col++,
          Math.round((stdRes.direct.buyFirst.none || 0) * 100),
          STYLES.percent,
        );
      if (hasDirectShare)
        setCell(
          ws,
          row,
          col++,
          Math.round((stdRes.direct.shareFirst.none || 0) * 100),
          STYLES.percent,
        );
    }
    row++;
  }

  return { endRow: row, endCol: lastCol };
}

function makeSignifSheetStyled(stdRes, concepts, signifRes) {
  const ws = {};

  const leftCols = [{ wch: 42 }, ...Array.from({ length: concepts.length }, () => ({ wch: 15 }))];
  const spacer = [{ wch: 4 }];
  const rightCols = [{ wch: 42 }, ...Array.from({ length: concepts.length }, () => ({ wch: 18 }))];

  ws['!cols'] = [...leftCols, ...spacer, ...rightCols];

  const left = writeSignifBlock(ws, 0, 0, stdRes, concepts, signifRes, 'green');
  const rightStart = leftCols.length + spacer.length;
  const right = writeSignifBlock(ws, 0, rightStart, stdRes, concepts, signifRes, 'letters');

  applySheetRangeRef(
    ws,
    Math.max(left.endRow, right.endRow),
    Math.max(left.endCol, right.endCol),
  );

  return ws;
}

// === Лист аудитории ===

function writeAudienceBlock(ws, startRow, sectionTitle, blocks) {
  let row = startRow;

  setCell(ws, row, 0, sectionTitle, STYLES.section);
  mergeRange(ws, row, 0, row, 1);
  row++;

  blocks.forEach(block => {
    if (!block.rows || !block.rows.length) return;

    setCell(ws, row, 0, block.title, STYLES.blockTitle);
    mergeRange(ws, row, 0, row, 1);
    row++;

    setCell(ws, row, 0, 'Категория', STYLES.headerCenter);
    setCell(ws, row, 1, '%', STYLES.headerCenter);
    row++;

    block.rows.forEach(item => {
      setCell(ws, row, 0, item.label, STYLES.label);
      setPercent(ws, row, 1, item.p, STYLES.percent);
      row++;
    });
  });

  return row;
}

function makeAudienceSheetStyled(audienceRes) {
  const ws = {};

  ws['!cols'] = [{ wch: 42 }, { wch: 18 }];

  let row = 0;

  setCell(ws, row, 0, 'Аудитория теста', STYLES.title);
  mergeRange(ws, row, 0, row, 1);
  row++;

  setCell(ws, row, 0, `n=${audienceRes.n}`, STYLES.base);
  mergeRange(ws, row, 0, row, 1);
  row++;

  row = writeAudienceBlock(ws, row, 'Демография', [
    ...audienceRes.sex,
    ...audienceRes.age,
  ]);

  row = writeAudienceBlock(ws, row, 'Поведение', [
    ...audienceRes.freqNew,
    ...audienceRes.freqProd,
    ...audienceRes.freqBK,
  ]);

  applySheetRangeRef(ws, row, 1);
  return ws;
}

// === НОВЫЙ блок: разбивки по возрасту ===

// Группы возраста соответствуют твоему формату: 18-24, 25-34, 35-44, 45+
function buildAgeGroups(rows, ageColIndex) {
  const groups = {};
  const totalKey = 'Тотал';

  rows.forEach((row, rIdx) => {
    const ageRaw = String(row[ageColIndex] || '').trim();
    if (!ageRaw) return;

    // Берем категорию как есть (18-24, 25-34, 35-44, 45+)
    const groupName = ageRaw;

    if (!groups[groupName]) groups[groupName] = [];
    groups[groupName].push(rIdx);
  });

  const totalIndexes = rows.map((_, idx) => idx);
  return { groups, totalKey, totalIndexes };
}

// Z‑тест долей: группа vs тотал
function zTestProportion(pGroup, nGroup, pTotal, nTotal) {
  if (!isFinite(pGroup) || !isFinite(pTotal)) return 0;
  if (nGroup <= 0 || nTotal <= 0) return 0;
  const pPool =
    (pGroup * nGroup + pTotal * nTotal) / (nGroup + nTotal);
  const se = Math.sqrt(
    pPool * (1 - pPool) * (1 / nGroup + 1 / nTotal),
  );
  if (se <= 0) return 0;
  return (pGroup - pTotal) / se;
}

function isSignificant(pGroup, nGroup, pTotal, nTotal, alpha = 0.05) {
  const z = zTestProportion(pGroup, nGroup, pTotal, nTotal);
  const zCrit = 1.96; // 95% доверительная вероятность
  if (z > zCrit) return 'up';
  if (z < -zCrit) return 'down';
  return null;
}

// Расчет доли Top‑2 по возрастным группам по всем вопросам
function calcAgeBreakdown(rows, userConfig, concepts, header) {
  const ageCol = userConfig.std.audience.age;
  if (ageCol == null) return null;

  const { groups, totalKey, totalIndexes } = buildAgeGroups(rows, ageCol);
  const groupNames = Object.keys(groups);
  if (!groupNames.length) return null;

  const questionSpecs = [];

  const std = userConfig.std;

  const pushStdBlock = (key, label, indexes) => {
    indexes.forEach(idx => {
      questionSpecs.push({
        type: 'std',
        block: key,
        colIndex: idx,
        label: `${label} – ${header[idx]}`,
      });
    });
  };

  pushStdBlock('like', 'Нравится (Top‑2)', std.like || []);
  pushStdBlock('fitDish', 'Подходит для блюда (Top‑2)', std.fitDish || []);
  pushStdBlock('fitBrand', 'Подходит для бренда (Top‑2)', std.fitBrand || []);
  pushStdBlock('visitBK', 'Намерение посетить БК (Top‑2)', std.visitBK || []);
  pushStdBlock('buyDish', 'Намерение купить (Top‑2)', std.buyDish || []);
  pushStdBlock('shareIntent', 'Намерение поделиться (Top‑2)', std.shareIntent || []);

  if (std.image && std.image.length) {
    std.image.forEach(item => {
      questionSpecs.push({
        type: 'image',
        key: item.key,
        colIndex: item.idx,
        label: 'Имидж: ' + item.key,
      });
    });
  }

  if (userConfig.extra && userConfig.extra.length) {
    userConfig.extra.forEach(extraItem => {
      if (extraItem.type === 'scale5') {
        questionSpecs.push({
          type: 'extraScale',
          colIndex: extraItem.idx,
          label: 'Доп.метрика – ' + extraItem.title,
        });
      }
    });
  }

  function calcTop2ShareForIndexes(colIdx, rowIndexes) {
    let n = 0;
    let top2 = 0;
    rowIndexes.forEach(rIdx => {
      const v = rows[rIdx][colIdx];
      if (v === null || v === '' || v === undefined) return;
      const num = Number(v);
      if (!isFinite(num)) return;
      n++;
      if (num >= 4) top2++;
    });
    if (!n) return { n: 0, share: NaN };
    return { n, share: top2 / n };
  }

  const resultTable = {
    groupNames: [...groupNames, totalKey],
    questions: questionSpecs,
    data: {}, // data[groupName][qIdx] = { share, n, signif }
  };

  const totalShares = questionSpecs.map(q =>
    calcTop2ShareForIndexes(q.colIndex, totalIndexes),
  );

  groupNames.forEach(groupName => {
    const idxs = groups[groupName];
    resultTable.data[groupName] = {};
    questionSpecs.forEach((q, qIdx) => {
      const gRes = calcTop2ShareForIndexes(q.colIndex, idxs);
      const tRes = totalShares[qIdx];
      const signif = isSignificant(
        gRes.share,
        gRes.n,
        tRes.share,
        tRes.n,
        0.05,
      );
      resultTable.data[groupName][qIdx] = {
        share: gRes.share,
        n: gRes.n,
        signif,
      };
    });
  });

  resultTable.data[totalKey] = {};
  totalShares.forEach((tRes, qIdx) => {
    resultTable.data[totalKey][qIdx] = {
      share: tRes.share,
      n: tRes.n,
      signif: null,
    };
  });

  return resultTable;
}

// Лист «разбивки по возрасту» с подсветкой значимостей
function makeAgeBreakdownSheetStyled(ageRes) {
  if (!ageRes) {
    return XLSX.utils.aoa_to_sheet([['Нет данных по возрасту']]);
  }

  const headerRow = ['Возраст / Тотал'];
  ageRes.questions.forEach(q => {
    headerRow.push(q.label);
  });

  const aoa = [headerRow];

  ageRes.groupNames.forEach(groupName => {
    const row = [groupName];
    ageRes.questions.forEach((q, qIdx) => {
      const cellData = ageRes.data[groupName][qIdx];
      const sharePct = isFinite(cellData.share)
        ? Math.round(cellData.share * 1000) / 10
        : '';
      row.push(sharePct);
    });
    aoa.push(row);
  });

  const sheet = XLSX.utils.aoa_to_sheet(aoa);

  ageRes.groupNames.forEach((groupName, gIdx) => {
    if (groupName === 'Тотал') return;
    const excelRow = gIdx + 1; // aoa: header = 0, первая группа = 1
    ageRes.questions.forEach((q, qIdx) => {
      const cellAddr = XLSX.utils.encode_cell({
        r: excelRow,
        c: qIdx + 1,
      });
      const cellData = ageRes.data[groupName][qIdx];
      const signif = cellData.signif;
      if (!signif) return;
      const cell = sheet[cellAddr] || {};
      cell.s = cell.s || {};
      cell.s.fill = { patternType: 'solid' };
      if (signif === 'up') {
        cell.s.fill.fgColor = { rgb: 'C6EFCE' };
      } else if (signif === 'down') {
        cell.s.fill.fgColor = { rgb: 'FFC7CE' };
      }
      sheet[cellAddr] = cell;
    });
  });

  sheet['!cols'] = [
    { wch: 18 },
    ...ageRes.questions.map(() => ({ wch: 20 })),
  ];

  sheet['!ref'] = XLSX.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: ageRes.groupNames.length, c: ageRes.questions.length },
  });

  return sheet;
}
