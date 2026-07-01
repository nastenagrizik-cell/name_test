// app.js
const baseInput = document.getElementById('baseFile');
const statusEl = document.getElementById('status');
const mappingSection = document.getElementById('mappingSection');
const extraSection = document.getElementById('extraSection');
const runSection = document.getElementById('runSection');
const standardGroupsEl = document.getElementById('standardGroups');
const extraQuestionsEl = document.getElementById('extraQuestions');
const runBtn = document.getElementById('runBtn');
const ageToggle = document.getElementById('ageToggle');

let baseFile;
let parsed = null;
let autoMapping = null;
let userConfig = null;

function debugLine(...parts) {
  const msg = parts.map(x => {
    if (typeof x === 'string') return x;
    try { return JSON.stringify(x); } catch (_) { return String(x); }
  }).join(' ');

  console.log('[DEBUG]', ...parts);

  if (statusEl) {
    const prev = statusEl.textContent || '';
    statusEl.textContent = prev ? (prev + '\n[DEBUG] ' + msg) : ('[DEBUG] ' + msg);
  }
}

function normalizeText(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[«»"']/g, '')
    .replace(/"/g, '')
    .replace(/[–—]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}


function stripTrailingConceptName(headerText) {
  const text = String(headerText || '').trim();
  const parts = text.split(/\s+-\s+/);
  if (parts.length < 2) return text;
  return parts.slice(0, -1).join(' - ').trim();
}

function canonicalExtraMetricName(headerText) {
  return normalizeText(stripTrailingConceptName(headerText))
    .replace(/^из каждого названия\s*/, 'насколько из каждого названия ')
    .replace(/^для каждого названия\s*/, '')
    .replace(/^оцените,? пожалуйста,?\s*/, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function prettyExtraGroupTitle(headerText) {
  const raw = stripTrailingConceptName(headerText).trim();
  if (!raw) return '';
  return raw.charAt(0).toUpperCase() + raw.slice(1);
}

function groupExtraCandidates(extraCandidates) {
  const groups = new Map();
  (extraCandidates || []).forEach(item => {
    const key = canonicalExtraMetricName(item.header) || normalizeText(item.header);
    if (!groups.has(key)) groups.set(key, { key, items: [] });
    groups.get(key).items.push(item);
  });
  return Array.from(groups.values()).map(group => ({
    key: group.key,
    title: prettyExtraGroupTitle(group.items[0]?.header || ''),
    items: group.items
  }));
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
    (
      t.includes('для каждого названия') ||
      t.includes('оцените') ||
      t.includes('насколько вероятно')
    ) &&
    (
      t.includes('расскажете') ||
      t.includes('рассказали бы') ||
      t.includes('рассказать') ||
      t.includes('поделитесь') ||
      t.includes('поделиться')
    ) &&
    (
      t.includes('соцсет') ||
      t.includes('социальных сет') ||
      t.includes('друзьям')
    )
  );
}

function looksLikeDirectShareQuestion(h) {
  const t = normalizeText(h);
  return (
    (
      t.includes('с каким из этих названий') ||
      t.includes('какое из этих названий')
    ) &&
    (
      t.includes('рассказали') ||
      t.includes('рассказала') ||
      t.includes('рассказать') ||
      t.includes('поделились') ||
      t.includes('поделиться')
    ) &&
    (
      t.includes('в первую очередь') ||
      t.includes('сначала')
    )
  );
}

const IMAGE_STATEMENT_CONFIG = [
  {
    key: 'Это оригинальный, необычный продукт',
    aliases: [
      'это оригинальный, необычный соус',
      'это оригинальный, необычный продукт',
      'это оригинальный, необычный бургер',
      'оригинальный, необычный'
    ]
  },
  {
    key: 'Ассоциируется со знакомым вкусом',
    aliases: [
      'этот соус ассоциируется со знакомым вкусом',
      'этот бургер ассоциируется со знакомым вкусом',
      'этот продукт ассоциируется со знакомым вкусом'
    ]
  },
  {
    key: 'Добавляет премиальности',
    aliases: [
      'это премиальный соус',
      'это премиальный продукт',
      'добавляет премиальности'
    ]
  },
  {
    key: 'Продукт с юмором',
    aliases: [
      'соус с юмором',
      'бургер с юмором',
      'продукт с юмором'
    ]
  },
  {
    key: 'Хочется попробовать',
    aliases: [
      'хочется попробовать воппер с таким соусом',
      'хочется попробовать такой бургер',
      'хочется попробовать такой капучино',
      'хочется попробовать такой напиток',
      'хочется попробовать такой продукт'
    ]
  },
  {
    key: 'Ассоциируется с приятным вкусом',
    aliases: [
      'этот соус ассоциируется с приятным вкусом',
      'этот бургер ассоциируется с приятным вкусом',
      'этот продукт ассоциируется с приятным вкусом'
    ]
  },
  {
    key: 'Уникальная новинка',
    aliases: [
      'это уникальная новинка',
      'это уникальный бургер',
      'это уникальный продукт',
      'оригинальное, отличается от других',
      'это уникальный напиток'
    ]
  },
  {
    key: 'Понятное и простое название',
    aliases: [
      'понятное и простое название'
    ]
  },
  {
    key: 'Вызывает отторжение',
    aliases: [
      'этот соус вызывает отторжение',
      'этот продукт вызывает отторжение',
      'этот бургер вызывает отторжение'
    ]
  },
  {
    key: 'Понятно, какой будет вкус',
    aliases: [
      'понятно, с каким вкусом будет этот бургер',
      'понятно, с каким вкусом будет этот соус',
      'понятно, с каким вкусом будет этот продукт',
      'понятно какой будет вкус'
    ]
  },
  {
    key: 'Название легко запомнить',
    aliases: [
      'название легко запомнить'
    ]
  },
  {
    key: 'Вызывает аппетит, звучит вкусно',
    aliases: [
      'вызывает аппетит',
      'вызывает аппетит, вкусно звучит',
      'вызывает аппетит, звучит вкусно',
      'название звучит вкусно и аппетитно'
    ]
  },
  {
    key: 'Вызывает доверие',
    aliases: [
      'вызывает у меня доверие',
      'вызывает доверие'
    ]
  },
  {
    key: 'Звучит как натуральный продукт',
    aliases: [
      'звучит как натуральный продукт'
    ]
  },
  {
    key: 'Звучит как качественный продукт',
    aliases: [
      'звучит как качественный продукт'
    ]
  },
  {
    key: 'По названию понятен маленький формат',
    aliases: [
      'по названию понятно, что это мини-бургеры',
      'по названию понятно, что это маленький формат',
      'по названию понятно, что это мини формат',
      'маленький формат'
    ]
  },
  {
    key: 'По названию понятно, что внутри несколько вкусов',
    aliases: [
      'по названию понятно, что внутри несколько разных бургеров',
      'по названию понятно, что внутри несколько разных вкусов',
      'внутри несколько разных вкусов'
    ]
  },
  {
    key: 'Хорошо передает идею набора',
    aliases: [
      'это название хорошо передает идею набора',
      'хорошо передает идею набора'
    ]
  },
  {
    key: 'Звучит хайпово и трендово',
    aliases: [
      'звучит хайпово и трендово'
    ]
  },
  {
    key: 'Люди бы обсуждали такое название',
    aliases: [
      'люди бы обсуждали такое название'
    ]
  },
  {
    key: 'Название звучит странно или отталкивающе',
    aliases: [
      'название звучит странно или отталкивающе'
    ]
  },
  {
    key: 'Название из детского меню / продукт для детей',
    aliases: [
      'название из детского меню',
      'продукт для детей'
    ]
  },
  {
    key: 'Не сытно / не наешься',
    aliases: [
      'не сытно',
      'не наешься'
    ]
  },
  {
    key: 'Дешевый продукт',
    aliases: [
      'дешевый продукт'
    ]
  },
  {
    key: 'Стоит своих денег',
    aliases: [
      'стоит своих денег'
    ]
  },
  {
    key: 'Подходит для группового потребления',
    aliases: [
      'подходит для группового потребления'
    ]
  },
  {
    key: 'Звучит старомодно',
    aliases: [
      'звучит старомодно'
    ]
  }
];
function detectImageStatementKey(headerText) {
  const t = normalizeText(headerText);
  for (const item of IMAGE_STATEMENT_CONFIG) {
    if (item.aliases.some(alias => t.includes(normalizeText(alias)))) return item.key;
  }
  return null;
}

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
  baseInput.addEventListener('click', () => {
    baseInput.value = '';
  });

  baseInput.addEventListener('change', async e => {
    debugLine('change fired');

    baseFile = e.target.files[0] || null;
    debugLine('baseFile =', baseFile ? { name: baseFile.name, size: baseFile.size, type: baseFile.type } : null);

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
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      debugLine('sheet_to_json ok, rows =', data.length);

      const { header, rows } = splitHeaderRows(data);
      debugLine('splitHeaderRows ok, header len = ' + header.length + ', rows len = ' + rows.length);
      debugLine('header sample = ' + JSON.stringify(header.slice(0, 8)));

      parsed = { header, rows };
      debugLine('parsed assigned');

      autoMapping = autoDetectMapping(header);
      debugLine('autoDetectMapping ok = ' + JSON.stringify({
        like: autoMapping.std.like.length,
        fitDish: autoMapping.std.fitDish.length,
        fitBrand: autoMapping.std.fitBrand.length,
        visitBK: autoMapping.std.visitBK.length,
        buyDish: autoMapping.std.buyDish.length,
        shareIntent: autoMapping.std.shareIntent.length,
        image: autoMapping.std.image.length,
        directLike: autoMapping.std.directLike.length,
        directBuy: autoMapping.std.directBuy.length,
        extraCandidates: autoMapping.extraCandidates.length
      }));

      debugLine('matched like headers = ' + JSON.stringify(autoMapping.std.like.map(i => header[i]).slice(0, 5)));
      debugLine('matched fitDish headers = ' + JSON.stringify(autoMapping.std.fitDish.map(i => header[i]).slice(0, 5)));
      debugLine('matched fitBrand headers = ' + JSON.stringify(autoMapping.std.fitBrand.map(i => header[i]).slice(0, 5)));
      debugLine('matched buyDish headers = ' + JSON.stringify(autoMapping.std.buyDish.map(i => header[i]).slice(0, 5)));
      debugLine('matched shareIntent headers = ' + JSON.stringify(autoMapping.std.shareIntent.map(i => header[i]).slice(0, 5)));
      debugLine('matched extra headers = ' + JSON.stringify(autoMapping.extraCandidates.map(x => x.header).slice(0, 10)));

      renderStandardMappingUI(autoMapping, header);
      debugLine('renderStandardMappingUI ok');

      renderExtraQuestionsUI(autoMapping, header);
      debugLine('renderExtraQuestionsUI ok');

      if (mappingSection) mappingSection.style.display = '';
      if (extraSection) extraSection.style.display = '';
      if (runSection) runSection.style.display = '';
      debugLine('sections shown');

      status('База загружена. Проверьте найденные вопросы и доп.метрики, затем нажмите «Посчитать топлайн».', true);
    } catch (e) {
      console.error(e);
      debugLine('READ ERROR =', e && e.message ? e.message : String(e));
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
      const extraResults = calcExtraBlocks(rows, userConfig, concepts, header);
      const audienceRes = calcAudience(rows, userConfig);
      const signifRes = calcSignificance(stdResults, concepts, rows.length);

      const outWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(outWb, makeSummarySheetStyled(stdResults, concepts, signifRes, extraResults), 'САММАРИ');
      XLSX.utils.book_append_sheet(outWb, makeFullSheetStyled(stdResults, concepts, extraResults), 'полные таблицы');
      XLSX.utils.book_append_sheet(outWb, makeSignifSheetStyled(stdResults, concepts, signifRes), 'значимости');
      XLSX.utils.book_append_sheet(outWb, makeAudienceSheetStyled(audienceRes), 'Аудитория');

      if (ageToggle && ageToggle.checked) {
        if (userConfig.std.audience.age == null || userConfig.std.audience.age < 0) {
          status('Внимание: столбец возраста не найден в базе — лист по возрастам не добавлен.', false, true);
        } else {
          const ageData = calcAgeBreakdown(rows, userConfig, concepts, header);
          if (ageData) {
            const ageSignif = calcAgeSignificance(ageData, stdResults, concepts);
            XLSX.utils.book_append_sheet(outWb, makeAgeSheetStyled(ageData, ageSignif, stdResults, extraResults, concepts), 'Возраст');
          }
        }
      }

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

  const questionTexts = (data[0] || []).map(v => String(v || '').trim());
  const varNames = (data[1] || []).map(v => String(v || '').trim());
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
    shareIntent: [],
    image: [],
    directLike: [],
    directBuy: [],
    directShare: [],
    audience: { sex: null, age: null, freqNew: null, freqProd: null, freqBK: null }
  };

  header.forEach((h, idx) => {
    const text = String(h || '').trim();
    const t = normalizeText(text);
    if (!text) return;

    if (
      text.includes('Оцените, пожалуйста, насколько вам нравится или не нравится каждое из этих названий') ||
      t.includes('нравится или не нравится каждое из этих названий')
    ) {
      std.like.push(idx);
    }

    if (looksLikeFitDishQuestion(text)) {
      std.fitDish.push(idx);
    }

    if (
      text.includes('А теперь оцените, насколько каждое из этих названий подходит или не подходит для бренда Бургер Кинг') ||
      t.includes('подходит или не подходит для бренда бургер кинг')
    ) {
      std.fitBrand.push(idx);
    }

    if (
      text.includes('Скажите, насколько вероятно, что Вы посетите ресторан Бургер Кинг') ||
      t.includes('насколько вероятно, что вы посетите ресторан бургер кинг')
    ) {
      std.visitBK.push(idx);
    }

    if (
      text.includes('Для каждого названия укажите, насколько вероятно, что Вы купите') ||
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
      (t.includes('с каким из этих названий') && t.includes('купили') && t.includes('в первую очередь'))
    ) {
      std.directBuy.push(idx);
    }

    if (looksLikeDirectShareQuestion(text)) {
      std.directShare.push(idx);
    }

    if (text.includes('Укажите Ваш пол')) std.audience.sex = idx;
    if (text.includes('Укажите Ваш возраст')) std.audience.age = idx;
    if (text.includes('Как часто Вы берете новинки') || text.includes('Как часто вы берете новинки')) std.audience.freqNew = idx;
    if ((text.includes('Как часто вы покупаете') || text.includes('Как часто Вы покупаете')) && std.audience.freqProd == null) std.audience.freqProd = idx;
    if (text.includes('Как часто вы посещаете Бургер Кинг')) std.audience.freqBK = idx;
  });

  const used = new Set([
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
    std.audience.freqBK
  ].filter(v => v !== null && v !== undefined));

  const extraCandidates = header.map((h, idx) => {
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
  }).filter(Boolean);

  return { std, extraCandidates };
}
function renderStandardMappingUI(mapping, header) {
  const groups = [
    { key: 'like', label: 'Нравится название', type: 'шкала 1–5, Top‑2', indexes: mapping.std.like },
    { key: 'fitDish', label: 'Подходит для блюда / продукта', type: 'шкала 1–5, Top‑2', indexes: mapping.std.fitDish },
    { key: 'fitBrand', label: 'Подходит для бренда', type: 'шкала 1–5, Top‑2', indexes: mapping.std.fitBrand },
    { key: 'visitBK', label: 'Намерение посетить БК', type: 'шкала 1–5, Top‑2', indexes: mapping.std.visitBK },
    { key: 'buyDish', label: 'Намерение купить', type: 'шкала 1–5, Top‑2', indexes: mapping.std.buyDish },
    { key: 'shareIntent', label: 'Намерение рассказать / поделиться', type: 'шкала 1–5, Top‑2', indexes: mapping.std.shareIntent },
    { key: 'directCompare', label: 'Прямое сравнение', type: 'single choice', indexes: [...mapping.std.directLike, ...mapping.std.directBuy, ...mapping.std.directShare] }
  ];

  standardGroupsEl.innerHTML = '';

  groups.forEach(group => {
    const found = group.indexes.length;
    const expected = group.key === 'directCompare' ? 3 : 1;
    const badgeClass = found === 0 ? 'missing' : (found >= expected ? 'found' : 'partial');
    const badgeText = found === 0 ? 'не найдено' : `${found} ${found === 1 ? 'колонка' : 'колонки'}`;

    const card = document.createElement('div');
    card.className = 'metric-group';

    const head = document.createElement('div');
    head.className = 'metric-group-head';
    head.innerHTML = `
      <div class="metric-group-title">${group.label}<span style="font-weight:400;color:var(--muted);font-size:13px"> — ${group.type}</span></div>
      <div class="metric-group-meta">
        <span class="metric-badge ${badgeClass}">${badgeText}</span>
        <span class="metric-chevron">⌄</span>
      </div>
    `;

    const body = document.createElement('div');
    body.className = 'metric-group-body';

    if (found === 0) {
      body.innerHTML = '<div class="metric-line"><span class="metric-line-dot missing"></span><span class="metric-line-text">Колонки не найдены по ключевым словам. Проверьте базу или добавьте метрику вручную.</span></div>';
    } else {
      group.indexes.forEach(idx => {
        let stdKey = group.key;
        if (group.key === 'directCompare') {
          if (mapping.std.directLike.includes(idx)) stdKey = 'directLike';
          else if (mapping.std.directBuy.includes(idx)) stdKey = 'directBuy';
          else stdKey = 'directShare';
        }
        const id = `std-${group.key}-${idx}`;
        const line = document.createElement('div');
        line.className = 'metric-line';
        line.innerHTML = `
          <span class="metric-line-dot"></span>
          <input type="checkbox" id="${id}" data-std-key="${stdKey}" data-col-idx="${idx}" checked style="margin-top:4px;accent-color:var(--primary-2)">
          <label for="${id}" class="metric-line-text"><strong>${header[idx]}</strong></label>
        `;
        body.appendChild(line);
      });
    }

    head.addEventListener('click', () => {
      card.classList.toggle('open');
    });

    card.appendChild(head);
    card.appendChild(body);
    standardGroupsEl.appendChild(card);
  });
}

function renderExtraQuestionsUI(mapping) {
  extraQuestionsEl.innerHTML = '';
  const groups = groupExtraCandidates(mapping.extraCandidates || []);

  if (!groups.length) {
    extraQuestionsEl.innerHTML = '<div class="status">Дополнительные закрытые вопросы не найдены.</div>';
    return;
  }

  groups.forEach((group, i) => {
    const countLabel = group.items.length === 1 ? '1 колонка найдена' : `${group.items.length} колонок найдено`;
    const card = document.createElement('div');
    card.className = 'extra-card-compact';

    const head = document.createElement('div');
    head.className = 'extra-card-compact-head';
    head.innerHTML = `
      <div>
        <div class="metric-group-title" style="font-size:14px">${group.title}</div>
        <div class="panel-subtitle" style="font-size:13px">${countLabel}</div>
      </div>
      <div class="metric-group-meta">
        <span class="metric-badge found">доп. метрика</span>
        <span class="metric-chevron">⌄</span>
      </div>
    `;

    const body = document.createElement('div');
    body.className = 'extra-card-compact-body';
    body.innerHTML = `
      <div class="field">
        <label>Название метрики в topline</label>
        <input type="text" data-extra-group-name="${group.key}" placeholder="Напр. Понятно, что речь о наборе из трёх мини-бургеров">
      </div>
      <div class="field">
        <label>Куда выводить</label>
        <div class="pill-checkboxes" data-extra-group-where="${group.key}">
          <label><input type="checkbox" value="summary" checked> САММАРИ</label>
          <label><input type="checkbox" value="full" checked> полные таблицы</label>
          <label><input type="checkbox" value="signif" checked> значимости</label>
          <label><input type="checkbox" value="age" checked> возраст</label>
        </div>
      </div>
      <div class="mapping-list">
        ${group.items.map(item => `<label class="mapping-item"><input type="checkbox" data-extra-index="${item.idx}" data-extra-group="${group.key}" checked><span><strong>${item.header}</strong><br><small>Колонка ${item.idx + 1}</small></span></label>`).join('')}
      </div>
    `;

    head.addEventListener('click', (e) => {
      if (e.target.tagName === 'INPUT' || e.target.tagName === 'SELECT') return;
      card.classList.toggle('open');
    });

    card.appendChild(head);
    card.appendChild(body);
    extraQuestionsEl.appendChild(card);
  });
}

function collectUserConfig(mapping) {
  const stdSelected = {
    like: [], fitDish: [], fitBrand: [], visitBK: [], buyDish: [], shareIntent: [],
    directLike: [], directBuy: [], directShare: [], image: mapping.std.image.slice(), audience: mapping.std.audience
  };

  document.querySelectorAll('input[type="checkbox"][data-std-key]').forEach(cb => {
    if (!cb.checked) return;
    const key = cb.getAttribute('data-std-key');
    const idx = Number(cb.getAttribute('data-col-idx'));
    stdSelected[key].push(idx);
  });

  const extra = [];
  const grouped = groupExtraCandidates(mapping.extraCandidates || []);

  grouped.forEach(group => {
    const selected = Array.from(extraQuestionsEl.querySelectorAll(`[data-extra-group="${group.key}"]:checked`));
    if (!selected.length) return;

    const titleInput = extraQuestionsEl.querySelector(`[data-extra-group-name="${group.key}"]`);
    const title = (titleInput?.value || '').trim();
    if (!title) throw new Error('У доп.метрики "' + group.title + '" не задано название.');

    const whereWrap = extraQuestionsEl.querySelector(`[data-extra-group-where="${group.key}"]`);
    const where = [];
    if (whereWrap) {
      whereWrap.querySelectorAll('input[type="checkbox"]').forEach(cb => {
        if (cb.checked) where.push(cb.value);
      });
    }
    if (!where.length) throw new Error('У доп.метрики "' + title + '" не выбрано, куда выводить.');

    selected.forEach(el => {
      extra.push({
        idx: Number(el.getAttribute('data-extra-index')),
        header: '',
        title,
        type: 'scale5',
        where,
        groupKey: group.key
      });
    });
  });

  return { std: stdSelected, extra };
}

function inferConcepts(header, config) {
  const sourceCols =
    config.std.like.length ? config.std.like :
    config.std.fitDish.length ? config.std.fitDish :
    config.std.fitBrand.length ? config.std.fitBrand :
    config.std.buyDish.length ? config.std.buyDish :
    config.std.shareIntent.length ? config.std.shareIntent :
    config.std.visitBK.length ? config.std.visitBK : [];

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

const AGE_GROUPS = [
  { key: '18-24', label: '18-24', min: 18, max: 24 },
  { key: '25-34', label: '25-34', min: 25, max: 34 },
  { key: '35-44', label: '35-44', min: 35, max: 44 },
  { key: '45+', label: '45+', min: 45, max: 999 }
];

function parseAgeValue(v) {
  if (v === null || v === undefined || v === '') return null;
  if (typeof v === 'number') return v;
  const s = String(v).trim();
  const m = s.match(/(\d+)/);
  return m ? Number(m[1]) : null;
}

function ageGroupForValue(age) {
  if (age == null || age < 18) return null;
  return AGE_GROUPS.find(g => age >= g.min && age <= g.max) || null;
}

function splitRowsByAge(rows, ageColIdx) {
  const groups = {};
  AGE_GROUPS.forEach(g => { groups[g.key] = []; });
  let unassigned = 0;

  rows.forEach(r => {
    const age = parseAgeValue(getCell(r, ageColIdx));
    const g = ageGroupForValue(age);
    if (g) groups[g.key].push(r);
    else unassigned++;
  });

  return { groups, unassigned };
}

function normalizeConceptLabel(s) {
  return normalizeText(s)
    .replace(/\bвоппер\b/g, '')
    .replace(/\bбургер\b/g, '')
    .replace(/\bкапучино\b/g, '')
    .replace(/\bнапиток\b/g, '')
    .replace(/\bсоус\b/g, '')
    .replace(/\bнабор\b/g, '')
    .replace(/\bпродукт\b/g, '')
    .replace(/\s+/g, ' ')
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

    return res.map(v => v / n);
  }

  function dist5(cols) {
    if (!cols.length) return null;
    const arr = Array.from({ length: concepts.length }, () => ({ '1': 0, '2': 0, '3': 0, '4': 0, '5': 0 }));

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
    const res = {};
    IMAGE_STATEMENT_CONFIG.forEach(item => {
      res[item.key] = Array(concepts.length).fill(0);
    });

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

    return Object.fromEntries(Object.entries(res).filter(([, vals]) => vals.some(v => v > 0)));
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
      const key = Object.keys(counts).find(k => normalizeText(k).includes(normalizeText(c.label)));
      perConcept[i] = key ? counts[key] / n : 0;
    });

    const noneKey = Object.keys(counts).find(k => {
      const t = normalizeText(k);
      return t.includes('ни одно') || t.includes('ничего из перечисленного') || t.includes('не купил') || t.includes('не рассказал');
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
      shareIntent: dist5(config.std.shareIntent)
    },
    top2: {
      like: top2ByCols(config.std.like),
      fitDish: top2ByCols(config.std.fitDish),
      fitBrand: top2ByCols(config.std.fitBrand),
      visitBK: top2ByCols(config.std.visitBK),
      buyDish: top2ByCols(config.std.buyDish),
      shareIntent: top2ByCols(config.std.shareIntent)
    },
    image: imageBlock(),
    direct: {
      likeMost: directSingle(config.std.directLike),
      buyFirst: directSingle(config.std.directBuy),
      shareFirst: directSingle(config.std.directShare)
    }
  };
}

function calcExtraBlocks(rows, config, concepts, header) {
  const result = [];
  const handledScaleTitles = new Set();

  config.extra.forEach(q => {
    if (q.type === 'scale5') {
      if (handledScaleTitles.has(q.title)) return;
      handledScaleTitles.add(q.title);

      const sameTitle = config.extra.filter(x => x.type === 'scale5' && x.title === q.title);
      const dists = Array.from({ length: concepts.length }, () => ({ '1': 0, '2': 0, '3': 0, '4': 0, '5': 0, base: 0 }));
      let matched = 0;

      sameTitle.forEach(item => {
        const conceptIndex = findConceptIndexByHeader(header[item.idx], concepts);
        if (conceptIndex < 0) return;
        matched++;
        rows.forEach(r => {
          const v = parseScaleValue(getCell(r, item.idx));
          if (v >= 1 && v <= 5) {
            dists[conceptIndex][String(v)]++;
            dists[conceptIndex].base++;
          }
        });
      });

      if (matched > 0) {
        result.push({
          kind: 'scale5_by_concept',
          title: q.title,
          where: q.where,
          dist: dists.map(counts => ({
            '1': counts.base ? counts['1'] / counts.base : 0,
            '2': counts.base ? counts['2'] / counts.base : 0,
            '3': counts.base ? counts['3'] / counts.base : 0,
            '4': counts.base ? counts['4'] / counts.base : 0,
            '5': counts.base ? counts['5'] / counts.base : 0,
            top2: counts.base ? (counts['4'] + counts['5']) / counts.base : 0
          }))
        });
        return;
      }

      const n = rows.length;
      const counts = { '1': 0, '2': 0, '3': 0, '4': 0, '5': 0 };
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
        dist: Object.entries(counts).map(([cat, c]) => ({ cat, p: c / n })).sort((a, b) => b.p - a.p)
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
  const signif = { top2: {}, scales: {}, image: {}, directMax: {} };

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

  ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish', 'shareIntent'].forEach(k => {
    if (stdRes.top2[k]) signif.top2[k] = labelsFor(stdRes.top2[k]);
  });

  ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish', 'shareIntent'].forEach(k => {
    if (!stdRes.scales[k]) return;
    signif.scales[k] = {};
    ['top2', '1', '2', '3', '4', '5'].forEach(level => {
      signif.scales[k][level] = labelsFor(stdRes.scales[k].map(d => d[level]));
    });
  });

  Object.entries(stdRes.image).forEach(([k, vals]) => {
    signif.image[k] = labelsFor(vals);
  });

  function maxMask(arr) {
    if (!arr) return null;
    const max = Math.max(...arr);
    return arr.map(v => max > 0 && v === max);
  }

  signif.directMax.likeMost = stdRes.direct.likeMost ? maxMask(stdRes.direct.likeMost.perConcept) : null;
  signif.directMax.buyFirst = stdRes.direct.buyFirst ? maxMask(stdRes.direct.buyFirst.perConcept) : null;
  signif.directMax.shareFirst = stdRes.direct.shareFirst ? maxMask(stdRes.direct.shareFirst.perConcept) : null;

  return signif;
}

function cellRef(r, c) {
  return XLSX.utils.encode_cell({ r, c });
}

function setCell(ws, r, c, value, style = null) {
  ws[cellRef(r, c)] = { t: typeof value === 'number' ? 'n' : 's', v: value };
  if (style) ws[cellRef(r, c)].s = JSON.parse(JSON.stringify(style));
  return ws[cellRef(r, c)];
}

function setPercent(ws, r, c, value, style) {
  const cell = setCell(ws, r, c, Number(value || 0), style);
  cell.z = '0%';
  return cell;
}

function mergeRange(ws, sRow, sCol, eRow, eCol) {
  if (!ws['!merges']) ws['!merges'] = [];
  ws['!merges'].push({ s: { r: sRow, c: sCol }, e: { r: eRow, c: eCol } });
}

function applySheetRangeRef(ws, endRow, endCol) {
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: endRow, c: endCol } });
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
  title: { font: { bold: true, color: { rgb: 'FFFFFF' }, sz: 14 }, fill: hexFill('244C73'), alignment: { horizontal: 'left', vertical: 'center' }, border: borderAll() },
  section: { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: hexFill('244C73'), alignment: { horizontal: 'left', vertical: 'center' }, border: borderAll() },
  blockTitle: { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: hexFill('5E86B4'), alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: borderAll() },
  headerCenter: { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: hexFill('244C73'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: borderAll() },
  base: { font: { italic: true, color: { rgb: '333333' } }, fill: hexFill('D9D9D9'), alignment: { horizontal: 'left', vertical: 'center' }, border: borderAll() },
  label: { alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: borderAll() },
  top2Label: { font: { bold: true }, fill: hexFill('DCE6F1'), alignment: { horizontal: 'left', vertical: 'center' }, border: borderAll() },
  top2Row: { font: { bold: true }, fill: hexFill('DCE6F1'), alignment: { horizontal: 'center', vertical: 'center' }, border: borderAll() },
  percent: { alignment: { horizontal: 'center', vertical: 'center' }, border: borderAll() },
  percentGreen: { alignment: { horizontal: 'center', vertical: 'center' }, fill: hexFill('70AD47'), border: borderAll() },
  signifTextGreen: { font: { bold: true, color: { rgb: '000000' } }, alignment: { horizontal: 'center', vertical: 'center' }, fill: hexFill('70AD47'), border: borderAll() },
  legendGreen: { alignment: { horizontal: 'center', vertical: 'center' }, fill: hexFill('92D050'), border: borderAll() },
  legendAccent: { font: { bold: true, color: { rgb: 'C55A11' } }, alignment: { horizontal: 'center', vertical: 'center' }, fill: hexFill('FFF2CC'), border: borderAll() },
  legendText: { alignment: { horizontal: 'left', vertical: 'center', wrapText: true }, border: borderAll() }
};

function isStrong2Plus(arr, index) {
  return Array.isArray(arr) && Array.isArray(arr[index]) && arr[index].length >= 2;
}

function hasMetric(stdRes, key) {
  return !!(stdRes.top2[key] && stdRes.scales[key]);
}

function blockTitleForKey(key) {
  return {
    like: 'НАСКОЛЬКО НРАВИТСЯ НАЗВАНИЕ',
    fitDish: 'НАСКОЛЬКО ПОДХОДИТ ДЛЯ ЭТОГО ПРОДУКТА',
    fitBrand: 'НАСКОЛЬКО ПОДХОДИТ ДЛЯ БРЕНДА БУРГЕР КИНГ В ЦЕЛОМ',
    visitBK: 'НАМЕРЕНИЕ ПОСЕТИТЬ БУРГЕР КИНГ, ЕСЛИ ПОЯВИТСЯ В МЕНЮ',
    buyDish: 'НАМЕРЕНИЕ КУПИТЬ ПО ПРИЕМЛЕМОЙ ЦЕНЕ',
    shareIntent: 'НАМЕРЕНИЕ РАССКАЗАТЬ / ПОДЕЛИТЬСЯ'
  }[key];
}

function scaleLabelsForBlock(key) {
  return {
    like: ['1 - Совсем не нравится', '2', '3', '4', '5 - Очень нравится'],
    fitDish: ['1 - Точно не подходит', '2', '3', '4', '5 - Полностью подходит'],
    fitBrand: ['1 - Точно не подходит', '2', '3', '4', '5 - Полностью подходит'],
    visitBK: ['1 - Точно не посещу', '2', '3', '4', '5 - Точно посещу'],
    buyDish: ['1 - Точно не куплю', '2', '3', '4', '5 - Точно куплю'],
    shareIntent: ['1 - Точно не рассказал(а) бы', '2', '3', '4', '5 - Точно рассказал(а) бы']
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

function makeSummarySheetStyled(stdRes, concepts, signifRes, extraResults = []) {
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
    ['Нравится название', 'like'],
    ['Подходит для блюда / продукта', 'fitDish'],
    ['Подходит для бренда', 'fitBrand'],
    ['Намерение посетить БК', 'visitBK'],
    ['Намерение купить', 'buyDish'],
    ['Намерение рассказать / поделиться', 'shareIntent']
  ].forEach(([label, key]) => {
    if (!hasMetric(stdRes, key)) return;
    const vals = stdRes.top2[key];
    setCell(ws, row, 0, label, STYLES.label);
    vals.forEach((v, i) => {
      setPercent(ws, row, i + 1, v, isStrong2Plus(signifRes.top2[key], i) ? STYLES.percentGreen : STYLES.percent);
    });
    row++;
  });

  const hasDirectLike = !!stdRes.direct.likeMost;
  const hasDirectBuy = !!stdRes.direct.buyFirst;
  const hasDirectShare = !!stdRes.direct.shareFirst;

  if (hasDirectLike || hasDirectBuy || hasDirectShare) {
    row++;
    setCell(ws, row, 0, 'Прямое сравнение', STYLES.section);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    if (hasDirectLike) {
      setCell(ws, row, 0, 'Нравится больше всего', STYLES.label);
      stdRes.direct.likeMost.perConcept.forEach((v, i) => {
        setPercent(ws, row, i + 1, v, signifRes.directMax.likeMost?.[i] ? STYLES.percentGreen : STYLES.percent);
      });
      row++;
    }

    if (hasDirectBuy) {
      setCell(ws, row, 0, 'Куплю в первую очередь', STYLES.label);
      stdRes.direct.buyFirst.perConcept.forEach((v, i) => {
        setPercent(ws, row, i + 1, v, signifRes.directMax.buyFirst?.[i] ? STYLES.percentGreen : STYLES.percent);
      });
      row++;
    }

    if (hasDirectShare) {
      setCell(ws, row, 0, 'Рассказал(а) бы в первую очередь', STYLES.label);
      stdRes.direct.shareFirst.perConcept.forEach((v, i) => {
        setPercent(ws, row, i + 1, v, signifRes.directMax.shareFirst?.[i] ? STYLES.percentGreen : STYLES.percent);
      });
      row++;
    }
  }

  const imageEntries = Object.entries(stdRes.image || {});
  if (imageEntries.length) {
    row++;
    setCell(ws, row, 0, 'ИМИДЖЕВЫЙ БЛОК', STYLES.section);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    setCell(ws, row, 0, 'Показатель', STYLES.headerCenter);
    concepts.forEach((c, i) => setCell(ws, row, i + 1, c.label, STYLES.headerCenter));
    row++;

    imageEntries.forEach(([label, vals]) => {
      setCell(ws, row, 0, label, STYLES.label);
      vals.forEach((v, i) => {
        setPercent(ws, row, i + 1, v, isStrong2Plus(signifRes.image[label], i) ? STYLES.percentGreen : STYLES.percent);
      });
      row++;
    });
  }

  const summaryExtras = extraResults.filter(x => x.where.includes('summary'));
  if (summaryExtras.length) {
    row++;
    setCell(ws, row, 0, 'ДОПОЛНИТЕЛЬНЫЕ МЕТРИКИ', STYLES.section);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    summaryExtras.forEach(item => {
      if (item.kind === 'scale5') {
        setCell(ws, row, 0, item.title, STYLES.label);
        setPercent(ws, row, 1, item.dist.top2, STYLES.percent);
        if (lastCol >= 2) mergeRange(ws, row, 1, row, lastCol);
        row++;
      } else {
        setCell(ws, row, 0, item.title, STYLES.label);
        setCell(ws, row, 1, item.dist.slice(0, 3).map(x => `${x.cat}: ${Math.round(x.p * 100)}%`).join(' | '), STYLES.label);
        if (lastCol >= 1) mergeRange(ws, row, 1, row, lastCol);
        row++;
      }
    });
  }

  applySheetRangeRef(ws, row, lastCol);
  return ws;
}

function makeFullSheetStyled(stdRes, concepts, extraResults = []) {
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

  ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish', 'shareIntent'].forEach(key => {
    if (!hasMetric(stdRes, key)) return;

    setCell(ws, row, 0, blockTitleForKey(key), STYLES.blockTitle);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    blockValueRows(stdRes, key).forEach(([label, vals, level]) => {
      setCell(ws, row, 0, label, level === 'top2' ? STYLES.top2Label : STYLES.label);
      vals.forEach((v, i) => {
        setPercent(ws, row, i + 1, v, level === 'top2' ? STYLES.top2Row : STYLES.percent);
      });
      row++;
    });

    row++;
  });

  const imageEntries = Object.entries(stdRes.image || {});
  if (imageEntries.length) {
    setCell(ws, row, 0, 'ИМИДЖЕВЫЙ БЛОК', STYLES.section);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    imageEntries.forEach(([label, vals]) => {
      setCell(ws, row, 0, label, STYLES.label);
      vals.forEach((v, i) => setPercent(ws, row, i + 1, v, STYLES.percent));
      row++;
    });

    row++;
  }

  const fullExtras = extraResults.filter(x => x.where.includes('full'));
  if (fullExtras.length) {
    setCell(ws, row, 0, 'ДОПОЛНИТЕЛЬНЫЕ МЕТРИКИ', STYLES.section);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    fullExtras.forEach(item => {
      setCell(ws, row, 0, item.title, STYLES.blockTitle);
      mergeRange(ws, row, 0, row, lastCol);
      row++;

      if (item.kind === 'scale5_by_concept') {
        [
          ['ТОП-2 (4+5)', item.dist.map(d => d.top2), true],
          ['1', item.dist.map(d => d['1']), false],
          ['2', item.dist.map(d => d['2']), false],
          ['3', item.dist.map(d => d['3']), false],
          ['4', item.dist.map(d => d['4']), false],
          ['5', item.dist.map(d => d['5']), false]
        ].forEach(([label, values, isTop]) => {
          setCell(ws, row, 0, label, isTop ? STYLES.top2Label : STYLES.label);
          values.forEach((value, i) => setPercent(ws, row, i + 1, value, isTop ? STYLES.top2Row : STYLES.percent));
          row++;
        });
      } else if (item.kind === 'scale5') {
        [
          ['ТОП-2 (4+5)', item.dist.top2],
          ['1', item.dist['1']],
          ['2', item.dist['2']],
          ['3', item.dist['3']],
          ['4', item.dist['4']],
          ['5', item.dist['5']]
        ].forEach(([label, value]) => {
          setCell(ws, row, 0, label, label.startsWith('ТОП') ? STYLES.top2Label : STYLES.label);
          setPercent(ws, row, 1, value, label.startsWith('ТОП') ? STYLES.top2Row : STYLES.percent);
          if (lastCol >= 1) mergeRange(ws, row, 1, row, lastCol);
          row++;
        });
      } else {
        item.dist.forEach(x => {
          setCell(ws, row, 0, x.cat, STYLES.label);
          setPercent(ws, row, 1, x.p, STYLES.percent);
          if (lastCol >= 1) mergeRange(ws, row, 1, row, lastCol);
          row++;
        });
      }

      row++;
    });
  }

  const hasDirectLike = !!stdRes.direct.likeMost;
  const hasDirectBuy = !!stdRes.direct.buyFirst;
  const hasDirectShare = !!stdRes.direct.shareFirst;

  if (hasDirectLike || hasDirectBuy || hasDirectShare) {
    const directCols = 1 + (hasDirectLike ? 1 : 0) + (hasDirectBuy ? 1 : 0) + (hasDirectShare ? 1 : 0);
    setCell(ws, row, 0, 'ПРЯМОЕ СРАВНЕНИЕ', STYLES.section);
    mergeRange(ws, row, 0, row, directCols - 1);
    row++;

    setCell(ws, row, 0, 'Название', STYLES.headerCenter);
    let hdr = 1;
    if (hasDirectLike) setCell(ws, row, hdr++, 'Нравится больше всего', STYLES.headerCenter);
    if (hasDirectBuy) setCell(ws, row, hdr++, 'Куплю в первую очередь', STYLES.headerCenter);
    if (hasDirectShare) setCell(ws, row, hdr++, 'Рассказал(а) бы в первую очередь', STYLES.headerCenter);
    row++;

    concepts.forEach((c, i) => {
      setCell(ws, row, 0, c.label, STYLES.label);
      let col = 1;
      if (hasDirectLike) setPercent(ws, row, col++, stdRes.direct.likeMost.perConcept[i] || 0, STYLES.percent);
      if (hasDirectBuy) setPercent(ws, row, col++, stdRes.direct.buyFirst.perConcept[i] || 0, STYLES.percent);
      if (hasDirectShare) setPercent(ws, row, col++, stdRes.direct.shareFirst.perConcept[i] || 0, STYLES.percent);
      row++;
    });

    setCell(ws, row, 0, 'Ни одно из них', STYLES.label);
    let col = 1;
    if (hasDirectLike) setPercent(ws, row, col++, stdRes.direct.likeMost.none || 0, STYLES.percent);
    if (hasDirectBuy) setPercent(ws, row, col++, stdRes.direct.buyFirst.none || 0, STYLES.percent);
    if (hasDirectShare) setPercent(ws, row, col++, stdRes.direct.shareFirst.none || 0, STYLES.percent);
  }

  applySheetRangeRef(ws, row, Math.max(lastCol, 3));
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

  ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish', 'shareIntent'].forEach(key => {
    if (!hasMetric(stdRes, key)) return;

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

  const imageEntries = Object.entries(stdRes.image || {});
  if (imageEntries.length) {
    setCell(ws, row, startCol, 'ИМИДЖЕВЫЙ БЛОК', STYLES.section);
    mergeRange(ws, row, startCol, row, lastCol);
    row++;

    imageEntries.forEach(([label, vals]) => {
      setCell(ws, row, startCol, label, STYLES.label);
      vals.forEach((v, i) => {
        if (mode === 'green') {
          setPercent(ws, row, startCol + 1 + i, v, isStrong2Plus(signifRes.image[label], i) ? STYLES.percentGreen : STYLES.percent);
        } else {
          const letters = signifRes.image[label][i];
          setCell(ws, row, startCol + 1 + i, signifCellText(v, letters), letters && letters.length ? STYLES.signifTextGreen : STYLES.percent);
        }
      });
      row++;
    });

    row++;
  }

  const hasDirectLike = !!stdRes.direct.likeMost;
  const hasDirectBuy = !!stdRes.direct.buyFirst;
  const hasDirectShare = !!stdRes.direct.shareFirst;

  if (hasDirectLike || hasDirectBuy || hasDirectShare) {
    const directCols = 1 + (hasDirectLike ? 1 : 0) + (hasDirectBuy ? 1 : 0) + (hasDirectShare ? 1 : 0);
    setCell(ws, row, startCol, 'ПРЯМОЕ СРАВНЕНИЕ', STYLES.section);
    mergeRange(ws, row, startCol, row, startCol + directCols - 1);
    row++;

    setCell(ws, row, startCol, 'Название', STYLES.headerCenter);
    let hdr = startCol + 1;
    if (hasDirectLike) setCell(ws, row, hdr++, 'Нравится больше всего', STYLES.headerCenter);
    if (hasDirectBuy) setCell(ws, row, hdr++, 'Куплю в первую очередь', STYLES.headerCenter);
    if (hasDirectShare) setCell(ws, row, hdr++, 'Рассказал(а) бы в первую очередь', STYLES.headerCenter);
    row++;

    concepts.forEach((c, i) => {
      setCell(ws, row, startCol, c.label, STYLES.label);
      if (mode === 'green') {
        let col = startCol + 1;
        if (hasDirectLike) setPercent(ws, row, col++, stdRes.direct.likeMost.perConcept[i] || 0, signifRes.directMax.likeMost?.[i] ? STYLES.percentGreen : STYLES.percent);
        if (hasDirectBuy) setPercent(ws, row, col++, stdRes.direct.buyFirst.perConcept[i] || 0, signifRes.directMax.buyFirst?.[i] ? STYLES.percentGreen : STYLES.percent);
        if (hasDirectShare) setPercent(ws, row, col++, stdRes.direct.shareFirst.perConcept[i] || 0, signifRes.directMax.shareFirst?.[i] ? STYLES.percentGreen : STYLES.percent);
      } else {
        let col = startCol + 1;
        if (hasDirectLike) setCell(ws, row, col++, Math.round((stdRes.direct.likeMost.perConcept[i] || 0) * 100) + '%', STYLES.percent);
        if (hasDirectBuy) setCell(ws, row, col++, Math.round((stdRes.direct.buyFirst.perConcept[i] || 0) * 100) + '%', STYLES.percent);
        if (hasDirectShare) setCell(ws, row, col++, Math.round((stdRes.direct.shareFirst.perConcept[i] || 0) * 100) + '%', STYLES.percent);
      }
      row++;
    });

    setCell(ws, row, startCol, 'Ни одно из них', STYLES.label);
    if (mode === 'green') {
      let col = startCol + 1;
      if (hasDirectLike) setPercent(ws, row, col++, stdRes.direct.likeMost.none || 0, STYLES.percent);
      if (hasDirectBuy) setPercent(ws, row, col++, stdRes.direct.buyFirst.none || 0, STYLES.percent);
      if (hasDirectShare) setPercent(ws, row, col++, stdRes.direct.shareFirst.none || 0, STYLES.percent);
    } else {
      let col = startCol + 1;
      if (hasDirectLike) setCell(ws, row, col++, Math.round((stdRes.direct.likeMost.none || 0) * 100) + '%', STYLES.percent);
      if (hasDirectBuy) setCell(ws, row, col++, Math.round((stdRes.direct.buyFirst.none || 0) * 100) + '%', STYLES.percent);
      if (hasDirectShare) setCell(ws, row, col++, Math.round((stdRes.direct.shareFirst.none || 0) * 100) + '%', STYLES.percent);
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
  const rightStart = leftCols.length + 1;
  const right = writeSignifBlock(ws, 0, rightStart, stdRes, concepts, signifRes, 'letters');

  applySheetRangeRef(ws, Math.max(left.endRow, right.endRow), Math.max(left.endCol, right.endCol));
  return ws;
}

function calcAgeBreakdown(rows, config, concepts, header) {
  const ageColIdx = config.std.audience.age;
  if (ageColIdx == null || ageColIdx < 0) return null;

  const { groups, unassigned } = splitRowsByAge(rows, ageColIdx);
  const totalN = rows.length;

  const groupResults = {};
  AGE_GROUPS.forEach(g => {
    const gRows = groups[g.key];
    if (!gRows.length) {
      groupResults[g.key] = null;
      return;
    }
    const stdRes = calcStandardBlocks(gRows, config, concepts, header);
    const extraRes = calcExtraBlocks(gRows, config, concepts, header);
    groupResults[g.key] = { n: gRows.length, stdRes, extraRes };
  });

  return { groups: groupResults, totalN, unassigned };
}

function calcAgeSignificance(ageData, totalStdRes, concepts) {
  const alphaZ = 1.96;
  const result = {};

  AGE_GROUPS.forEach(g => {
    const gd = ageData.groups[g.key];
    if (!gd) { result[g.key] = null; return; }

    const n1 = gd.n;
    const signif = { top2: {}, image: {}, directMax: {} };

    function compareVsTotal(arr, totalArr) {
      if (!arr || !totalArr) return null;
      return arr.map((p, i) => {
        const flags = { higher: false, lower: false };
        const pTotal = totalArr[i];
        if (pTotal != null && n1 > 0) {
          const z = zTest(p, pTotal, n1, ageData.totalN);
          if (z > alphaZ) flags.higher = true;
          if (z < -alphaZ) flags.lower = true;
        }
        return flags;
      });
    }

    ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish', 'shareIntent'].forEach(k => {
      if (gd.stdRes.top2[k]) {
        signif.top2[k] = compareVsTotal(gd.stdRes.top2[k], totalStdRes.top2[k]);
      }
    });

    Object.entries(gd.stdRes.image || {}).forEach(([k, vals]) => {
      const totalVals = totalStdRes.image?.[k];
      signif.image[k] = compareVsTotal(vals, totalVals);
    });

    function maxMask(arr) {
      if (!arr) return null;
      const max = Math.max(...arr);
      return arr.map(v => max > 0 && v === max);
    }
    signif.directMax.likeMost = gd.stdRes.direct.likeMost ? maxMask(gd.stdRes.direct.likeMost.perConcept) : null;
    signif.directMax.buyFirst = gd.stdRes.direct.buyFirst ? maxMask(gd.stdRes.direct.buyFirst.perConcept) : null;
    signif.directMax.shareFirst = gd.stdRes.direct.shareFirst ? maxMask(gd.stdRes.direct.shareFirst.perConcept) : null;

    result[g.key] = signif;
  });

  return result;
}

const AGE_STYLES = {
  ageHeader: { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: hexFill('3C78B5'), alignment: { horizontal: 'center', vertical: 'center' }, border: borderAll() },
  ageGroupHeader: { font: { bold: true, color: { rgb: 'FFFFFF' } }, fill: hexFill('5E86B4'), alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, border: borderAll() },
  percentGreen: { alignment: { horizontal: 'center', vertical: 'center' }, fill: hexFill('70AD47'), font: { bold: true, color: { rgb: 'FFFFFF' } }, border: borderAll() },
  percentRed: { alignment: { horizontal: 'center', vertical: 'center' }, fill: hexFill('C84D4D'), font: { bold: true, color: { rgb: 'FFFFFF' } }, border: borderAll() },
  legendGreen: { alignment: { horizontal: 'center', vertical: 'center' }, fill: hexFill('70AD47'), font: { bold: true, color: { rgb: 'FFFFFF' } }, border: borderAll() },
  legendRed: { alignment: { horizontal: 'center', vertical: 'center' }, fill: hexFill('C84D4D'), font: { bold: true, color: { rgb: 'FFFFFF' } }, border: borderAll() }
};

function makeAgeSheetStyled(ageData, ageSignif, totalStdRes, totalExtraRes, concepts) {
  const ws = {};
  const numConcepts = concepts.length;
  const numAgeGroups = AGE_GROUPS.length;
  const lastCol = 1 + numConcepts * numAgeGroups;

  ws['!cols'] = [{ wch: 38 }, ...Array.from({ length: numConcepts * numAgeGroups }, () => ({ wch: 13 }))];

  let row = 0;

  setCell(ws, row, 0, 'РАЗБИВКА ПО ВОЗРАСТАМ: ТОП-2 (сумма оценок 4 и 5)', STYLES.title);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  setCell(ws, row, 0, `База: n=${ageData.totalN} респондентов | Все значения в %`, STYLES.base);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  setCell(ws, row, 0, 'Легенда:', STYLES.label);
  setCell(ws, row, 1, 'значимо выше тотал', AGE_STYLES.legendGreen);
  mergeRange(ws, row, 1, row, Math.floor(lastCol / 2));
  setCell(ws, row, Math.floor(lastCol / 2) + 1, 'значимо ниже тотал', AGE_STYLES.legendRed);
  mergeRange(ws, row, Math.floor(lastCol / 2) + 1, row, lastCol);
  row++;
  row++;

  // Two-row header: age group names spanning concept columns, concept labels underneath
  setCell(ws, row, 0, 'Метрика', STYLES.headerCenter);
  let col = 1;
  AGE_GROUPS.forEach(g => {
    setCell(ws, row, col, g.label, AGE_STYLES.ageGroupHeader);
    mergeRange(ws, row, col, row, col + numConcepts - 1);
    col += numConcepts;
  });
  row++;

  setCell(ws, row, 0, '', STYLES.headerCenter);
  col = 1;
  AGE_GROUPS.forEach(() => {
    concepts.forEach(c => {
      setCell(ws, row, col, c.label, AGE_STYLES.ageHeader);
      col++;
    });
  });
  row++;
  row++;

  function writeAgeMetricBlock(blockTitle, metricKey, totalArr) {
    if (!totalArr) return;

    setCell(ws, row, 0, blockTitle, STYLES.blockTitle);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    setCell(ws, row, 0, 'ТОП-2 (4+5)', STYLES.top2Label);
    col = 1;
    AGE_GROUPS.forEach(g => {
      const gd = ageData.groups[g.key];
      const signif = ageSignif[g.key];
      concepts.forEach((c, i) => {
        let style = STYLES.top2Row;
        let val = 0;
        if (gd && gd.stdRes.top2[metricKey]) {
          val = gd.stdRes.top2[metricKey][i];
          if (signif && signif.top2[metricKey] && signif.top2[metricKey][i]) {
            if (signif.top2[metricKey][i].higher) style = AGE_STYLES.percentGreen;
            else if (signif.top2[metricKey][i].lower) style = AGE_STYLES.percentRed;
          }
        }
        setPercent(ws, row, col, val, style);
        col++;
      });
    });
    row++;
    row++;
  }

  [
    ['Нравится название', 'like'],
    ['Подходит для блюда / продукта', 'fitDish'],
    ['Подходит для бренда', 'fitBrand'],
    ['Намерение посетить БК', 'visitBK'],
    ['Намерение купить', 'buyDish'],
    ['Намерение рассказать / поделиться', 'shareIntent']
  ].forEach(([label, key]) => {
    if (hasMetric(totalStdRes, key)) {
      writeAgeMetricBlock(label, key, totalStdRes.top2[key]);
    }
  });

  // Image block
  const imageEntries = Object.entries(totalStdRes.image || {});
  if (imageEntries.length) {
    setCell(ws, row, 0, 'ИМИДЖЕВЫЙ БЛОК', STYLES.section);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    imageEntries.forEach(([label, totalVals]) => {
      setCell(ws, row, 0, label, STYLES.label);
      col = 1;
      AGE_GROUPS.forEach(g => {
        const gd = ageData.groups[g.key];
        const signif = ageSignif[g.key];
        concepts.forEach((c, i) => {
          let style = STYLES.percent;
          let val = 0;
          if (gd && gd.stdRes.image[label]) {
            val = gd.stdRes.image[label][i];
            if (signif && signif.image[label] && signif.image[label][i]) {
              if (signif.image[label][i].higher) style = AGE_STYLES.percentGreen;
              else if (signif.image[label][i].lower) style = AGE_STYLES.percentRed;
            }
          }
          setPercent(ws, row, col, val, style);
          col++;
        });
      });
      row++;
    });
    row++;
  }

  // Direct comparison
  const hasDirectLike = !!totalStdRes.direct.likeMost;
  const hasDirectBuy = !!totalStdRes.direct.buyFirst;
  const hasDirectShare = !!totalStdRes.direct.shareFirst;

  if (hasDirectLike || hasDirectBuy || hasDirectShare) {
    setCell(ws, row, 0, 'ПРЯМОЕ СРАВНЕНИЕ', STYLES.section);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    if (hasDirectLike) {
      setCell(ws, row, 0, 'Нравится больше всего', STYLES.label);
      col = 1;
      AGE_GROUPS.forEach(g => {
        const gd = ageData.groups[g.key];
        const signif = ageSignif[g.key];
        concepts.forEach((c, i) => {
          let style = STYLES.percent;
          let val = 0;
          if (gd && gd.stdRes.direct.likeMost) {
            val = gd.stdRes.direct.likeMost.perConcept[i];
            if (signif && signif.directMax.likeMost && signif.directMax.likeMost[i]) style = AGE_STYLES.percentGreen;
          }
          setPercent(ws, row, col, val, style);
          col++;
        });
      });
      row++;
    }

    if (hasDirectBuy) {
      setCell(ws, row, 0, 'Куплю в первую очередь', STYLES.label);
      col = 1;
      AGE_GROUPS.forEach(g => {
        const gd = ageData.groups[g.key];
        const signif = ageSignif[g.key];
        concepts.forEach((c, i) => {
          let style = STYLES.percent;
          let val = 0;
          if (gd && gd.stdRes.direct.buyFirst) {
            val = gd.stdRes.direct.buyFirst.perConcept[i];
            if (signif && signif.directMax.buyFirst && signif.directMax.buyFirst[i]) style = AGE_STYLES.percentGreen;
          }
          setPercent(ws, row, col, val, style);
          col++;
        });
      });
      row++;
    }

    if (hasDirectShare) {
      setCell(ws, row, 0, 'Рассказал(а) бы в первую очередь', STYLES.label);
      col = 1;
      AGE_GROUPS.forEach(g => {
        const gd = ageData.groups[g.key];
        const signif = ageSignif[g.key];
        concepts.forEach((c, i) => {
          let style = STYLES.percent;
          let val = 0;
          if (gd && gd.stdRes.direct.shareFirst) {
            val = gd.stdRes.direct.shareFirst.perConcept[i];
            if (signif && signif.directMax.shareFirst && signif.directMax.shareFirst[i]) style = AGE_STYLES.percentGreen;
          }
          setPercent(ws, row, col, val, style);
          col++;
        });
      });
      row++;
    }
    row++;
  }

  // Extra metrics
  if (totalExtraRes && totalExtraRes.length) {
    setCell(ws, row, 0, 'ДОПОЛНИТЕЛЬНЫЕ МЕТРИКИ', STYLES.section);
    mergeRange(ws, row, 0, row, lastCol);
    row++;

    totalExtraRes.forEach((ex, exIdx) => {
      const exLabel = ex.title || `Метрика ${exIdx + 1}`;
      setCell(ws, row, 0, exLabel, STYLES.label);
      col = 1;
      AGE_GROUPS.forEach(g => {
        const gd = ageData.groups[g.key];
        concepts.forEach((c, i) => {
          let style = STYLES.percent;
          let val = 0;
          if (gd && gd.extraRes && gd.extraRes[exIdx]) {
            const exData = gd.extraRes[exIdx];
            if (exData.kind === 'scale5_by_concept' && Array.isArray(exData.dist)) {
              val = exData.dist[i] ? exData.dist[i].top2 : 0;
            } else if (exData.kind === 'scale5' && exData.dist) {
              val = exData.dist.top2 || 0;
            } else if (exData.kind === 'single' && Array.isArray(exData.dist)) {
              val = exData.dist[0] ? exData.dist[0].p : 0;
            }
          }
          setPercent(ws, row, col, val, style);
          col++;
        });
      });
      row++;
    });
    row++;
  }

  // Sample sizes per age group
  setCell(ws, row, 0, 'Размер выборки по возрастным группам', STYLES.blockTitle);
  mergeRange(ws, row, 0, row, lastCol);
  row++;

  setCell(ws, row, 0, 'Группа', STYLES.headerCenter);
  setCell(ws, row, 1, 'n', STYLES.headerCenter);
  mergeRange(ws, row, 1, row, lastCol);
  row++;

  AGE_GROUPS.forEach(g => {
    const gd = ageData.groups[g.key];
    setCell(ws, row, 0, g.label, STYLES.label);
    setCell(ws, row, 1, gd ? gd.n : 0, STYLES.percent);
    mergeRange(ws, row, 1, row, lastCol);
    row++;
  });

  if (ageData.unassigned > 0) {
    row++;
    setCell(ws, row, 0, `Не распределено по возрасту: ${ageData.unassigned} респондентов`, STYLES.base);
    mergeRange(ws, row, 0, row, lastCol);
    row++;
  }

  applySheetRangeRef(ws, row, lastCol);
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
    { title: 'Частота взятия новинок', rows: audienceRes.freqNew },
    { title: 'Частота покупки продукта', rows: audienceRes.freqProd },
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
    if (!block.rows || !block.rows.length) return;

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
