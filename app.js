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

  if (!mapping.extraCandidates.length) {
    extraQuestionsEl.innerHTML = '<div class="status">Дополнительные закрытые вопросы не найдены.</div>';
    return;
  }

  mapping.extraCandidates.forEach(q => {
    const card = document.createElement('div');
    card.className = 'extra-card-compact';

    const head = document.createElement('div');
    head.className = 'extra-card-compact-head';
    head.innerHTML = `
      <div class="metric-group-title" style="font-size:14px">${q.header}</div>
      <div class="metric-group-meta">
        <span class="metric-badge found">доп. метрика</span>
        <span class="metric-chevron">⌄</span>
      </div>
    `;

    const body = document.createElement('div');
    body.className = 'extra-card-compact-body';

    body.innerHTML = `
      <div class="field">
        <label>
          <input type="checkbox" data-extra-idx="${q.idx}" checked style="accent-color:var(--primary-2)">
          Использовать этот вопрос как доп.метрику
        </label>
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
  const conceptWidth = 2;
  const lastCol = concepts.length * conceptWidth;
  ws['!cols'] = [{ wch: 42 }, ...Array.from({ length: concepts.length }, () => [{ wch: 14 }, { wch: 14 }]).flat()];
  let row = 0;

  const metricRows = [
    ['Нравится название', 'like'],
    ['Подходит для блюда / продукта', 'fitDish'],
    ['Подходит для бренда', 'fitBrand'],
    ['Намерение посетить БК', 'visitBK'],
    ['Намерение купить', 'buyDish'],
    ['Намерение рассказать / поделиться', 'shareIntent']
  ].filter(([, key]) => hasMetric(stdRes, key));
  const extraRows = extraResults.filter(x => x.kind === 'scale5_by_concept').map(x => [x.title, x]);

  setCell(ws, row, 0, 'САММАРИ: ОСНОВНЫЕ И ДОПОЛНИТЕЛЬНЫЕ МЕТРИКИ', STYLES.title);
  mergeRange(ws, row, 0, row, lastCol);
  row++;
  setCell(ws, row, 0, `База: n=${stdRes.n} | Внутри каждого названия: Total и рейтинг`, STYLES.base);
  mergeRange(ws, row, 0, row, lastCol);
  row += 2;

  setCell(ws, row, 0, 'Показатель', STYLES.headerCenter);
  concepts.forEach((c, i) => {
    const start = 1 + i * conceptWidth;
    setCell(ws, row, start, c.label, STYLES.headerCenter);
    mergeRange(ws, row, start, row, start + 1);
  });
  row++;
  setCell(ws, row, 0, '', STYLES.headerCenter);
  concepts.forEach((_, i) => {
    const start = 1 + i * conceptWidth;
    setCell(ws, row, start, 'Total', STYLES.base);
    setCell(ws, row, start + 1, 'Ранг', STYLES.top2Label);
  });
  row++;

  function rankingFor(values, idx) {
    const sorted = values.map((v, i) => ({ v, i })).sort((a, b) => b.v - a.v);
    return sorted.findIndex(x => x.i === idx) + 1;
  }

  metricRows.forEach(([label, key]) => {
    const vals = stdRes.top2[key];
    setCell(ws, row, 0, label, STYLES.label);
    vals.forEach((v, i) => {
      const start = 1 + i * conceptWidth;
      setPercent(ws, row, start, v, isStrong2Plus(signifRes.top2[key], i) ? STYLES.percentGreen : STYLES.percent);
      setCell(ws, row, start + 1, String(rankingFor(vals, i)), STYLES.top2Row);
    });
    row++;
  });

  extraRows.forEach(([label, item]) => {
    const vals = item.dist.map(x => x.top2);
    setCell(ws, row, 0, label, STYLES.label);
    vals.forEach((v, i) => {
      const start = 1 + i * conceptWidth;
      setPercent(ws, row, start, v, STYLES.percent);
      setCell(ws, row, start + 1, String(rankingFor(vals, i)), STYLES.top2Row);
    });
    row++;
  });

  row++;
  setCell(ws, row, 0, 'ИМИДЖЕВЫЕ ВЫСКАЗЫВАНИЯ', STYLES.section);
  mergeRange(ws, row, 0, row, lastCol);
  row++;
  setCell(ws, row, 0, 'Высказывание', STYLES.headerCenter);
  concepts.forEach((c, i) => {
    const start = 1 + i * conceptWidth;
    setCell(ws, row, start, c.label, STYLES.headerCenter);
    mergeRange(ws, row, start, row, start + 1);
  });
  row++;
  setCell(ws, row, 0, '', STYLES.headerCenter);
  concepts.forEach((_, i) => {
    const start = 1 + i * conceptWidth;
    setCell(ws, row, start, 'Total', STYLES.base);
    setCell(ws, row, start + 1, 'Ранг', STYLES.top2Label);
  });
  row++;
  Object.entries(stdRes.image || {}).forEach(([label, vals]) => {
    setCell(ws, row, 0, label, STYLES.label);
    vals.forEach((v, i) => {
      const start = 1 + i * conceptWidth;
      setPercent(ws, row, start, v, isStrong2Plus(signifRes.image[label], i) ? STYLES.percentGreen : STYLES.percent);
      setCell(ws, row, start + 1, String(rankingFor(vals, i)), STYLES.top2Row);
    });
    row++;
  });

  applySheetRangeRef(ws, row, lastCol);
  return ws;
}

function makeFullSheetStyled(stdRes, concepts, extraResults = []) {
  const ws = {};
  const lastCol = concepts.length;
  ws['!cols'] = [{ wch: 46 }, ...Array.from({ length: concepts.length }, () => ({ wch: 15 }))];
  let row = 0;

  function writeTop2Section(title, values) {
    setCell(ws, row, 0, title, STYLES.blockTitle);
    mergeRange(ws, row, 0, row, lastCol);
    row++;
    setCell(ws, row, 0, 'Показатель', STYLES.headerCenter);
    concepts.forEach((c, i) => setCell(ws, row, i + 1, c.label, STYLES.headerCenter));
    row++;
    setCell(ws, row, 0, 'ТОП-2 (4+5)', STYLES.top2Label);
    values.forEach((v, i) => setPercent(ws, row, i + 1, v, STYLES.top2Row));
    row += 2;
  }

  setCell(ws, row, 0, 'ПОЛНЫЕ ТАБЛИЦЫ', STYLES.title);
  mergeRange(ws, row, 0, row, lastCol);
  row++;
  setCell(ws, row, 0, `База: n=${stdRes.n}`, STYLES.base);
  mergeRange(ws, row, 0, row, lastCol);
  row += 2;

  ['like', 'fitDish', 'fitBrand', 'visitBK', 'buyDish', 'shareIntent'].forEach(key => {
    if (!hasMetric(stdRes, key)) return;
    writeTop2Section(blockTitleForKey(key), stdRes.top2[key]);
    blockValueRows(stdRes, key).slice(1).forEach(([label, vals]) => {
      setCell(ws, row, 0, label, STYLES.label);
      vals.forEach((v, i) => setPercent(ws, row, i + 1, v, STYLES.percent));
      row++;
    });
    row++;
    const extrasAfter = extraResults.filter(x => x.kind === 'scale5_by_concept');
    extrasAfter.forEach(item => {
      setCell(ws, row, 0, item.title, STYLES.blockTitle);
      mergeRange(ws, row, 0, row, lastCol);
      row++;
      [['ТОП-2 (4+5)', item.dist.map(d => d.top2)], ['1', item.dist.map(d => d['1'])], ['2', item.dist.map(d => d['2'])], ['3', item.dist.map(d => d['3'])], ['4', item.dist.map(d => d['4'])], ['5', item.dist.map(d => d['5'])]].forEach(([label, values]) => {
        setCell(ws, row, 0, label, label.startsWith('ТОП') ? STYLES.top2Label : STYLES.label);
        values.forEach((v, i) => setPercent(ws, row, i + 1, v, label.startsWith('ТОП') ? STYLES.top2Row : STYLES.percent));
        row++;
      });
      row++;
    });
  });

  setCell(ws, row, 0, 'ИМИДЖЕВЫЕ ВЫСКАЗЫВАНИЯ', STYLES.section);
  mergeRange(ws, row, 0, row, lastCol);
  row++;
  Object.entries(stdRes.image || {}).forEach(([label, vals]) => {
    setCell(ws, row, 0, label, STYLES.label);
    vals.forEach((v, i) => setPercent(ws, row, i + 1, v, STYLES.percent));
    row++;
  });

  applySheetRangeRef(ws, row, lastCol);
  return ws;
}

function makeSignifSheetStyled(stdRes, concepts, signifRes, extraResults = []) {
  const ws = {};
  ws['!cols'] = [{ wch: 42 }, { wch: 14 }, { wch: 24 }];
  let row = 0;
  setCell(ws, row, 0, 'ЗНАЧИМОСТИ', STYLES.title);
  mergeRange(ws, row, 0, row, 2);
  row++;
  setCell(ws, row, 0, 'Для каждой метрики видно значение и какие названия значимо слабее текущего.', STYLES.base);
  mergeRange(ws, row, 0, row, 2);
  row += 2;

  const blocks = [
    ['Нравится название', 'like', stdRes.top2.like],
    ['Подходит для блюда / продукта', 'fitDish', stdRes.top2.fitDish],
    ['Подходит для бренда', 'fitBrand', stdRes.top2.fitBrand],
    ['Намерение посетить БК', 'visitBK', stdRes.top2.visitBK],
    ['Намерение купить', 'buyDish', stdRes.top2.buyDish],
    ['Намерение рассказать / поделиться', 'shareIntent', stdRes.top2.shareIntent]
  ].filter(([, , vals]) => vals);

  blocks.forEach(([title, key, vals]) => {
    setCell(ws, row, 0, title, STYLES.blockTitle);
    mergeRange(ws, row, 0, row, 2);
    row++;
    setCell(ws, row, 0, 'Название', STYLES.headerCenter);
    setCell(ws, row, 1, 'Total', STYLES.base);
    setCell(ws, row, 2, 'Лучше, чем', STYLES.headerCenter);
    row++;
    concepts.forEach((c, i) => {
      setCell(ws, row, 0, c.label, STYLES.label);
      setPercent(ws, row, 1, vals[i], STYLES.percent);
      setCell(ws, row, 2, (signifRes.top2[key]?.[i] || []).join(', '), STYLES.label);
      row++;
    });
    row++;
  });

  extraResults.filter(x => x.kind === 'scale5_by_concept').forEach(item => {
    setCell(ws, row, 0, item.title, STYLES.blockTitle);
    mergeRange(ws, row, 0, row, 2);
    row++;
    setCell(ws, row, 0, 'Название', STYLES.headerCenter);
    setCell(ws, row, 1, 'Total', STYLES.base);
    setCell(ws, row, 2, 'Ранг', STYLES.headerCenter);
    row++;
    const vals = item.dist.map(d => d.top2);
    const sorted = vals.map((v, i) => ({ v, i })).sort((a, b) => b.v - a.v);
    concepts.forEach((c, i) => {
      setCell(ws, row, 0, c.label, STYLES.label);
      setPercent(ws, row, 1, vals[i], STYLES.percent);
      setCell(ws, row, 2, String(sorted.findIndex(x => x.i === i) + 1), STYLES.top2Row);
      row++;
    });
    row++;
  });

  applySheetRangeRef(ws, row, 2);
  return ws;
}

function makeAgeSheetStyled(ageData, ageSignif, totalStdRes, totalExtraRes, concepts) {
  const ws = {};
  const subCols = ['Total', '18-24', '25-34', '35-44', '45+'];
  const width = subCols.length;
  const lastCol = concepts.length * width;
  ws['!cols'] = [{ wch: 40 }, ...Array.from({ length: concepts.length * width }, () => ({ wch: 13 }))];
  let row = 0;

  const sampleLine = 'Выборка: total n=' + ageData.total + '; ' + ageData.groups.map(g => `${g.label}: n=${g.n}`).join('; ');
  setCell(ws, row, 0, 'ВОЗРАСТ', STYLES.title);
  mergeRange(ws, row, 0, row, lastCol);
  row++;
  setCell(ws, row, 0, sampleLine, STYLES.base);
  mergeRange(ws, row, 0, row, lastCol);
  row += 2;

  function groupValue(metricKey, title, isExtra = false) {
    const vals = [];
    concepts.forEach((_, ci) => {
      vals.push(isExtra ? (totalExtraRes.find(x => x.title === title)?.dist?.[ci]?.top2 || 0) : (totalStdRes.top2[metricKey]?.[ci] || 0));
      ageData.groups.forEach(group => {
        if (isExtra) {
          const found = group.extraRes.find(x => x.title === title);
          vals.push(found?.dist?.[ci]?.top2 || 0);
        } else {
          vals.push(group.stdRes.top2[metricKey]?.[ci] || 0);
        }
      });
    });
    return vals;
  }

  function writeGroupedHeader(sectionTitle) {
    setCell(ws, row, 0, sectionTitle, STYLES.section);
    mergeRange(ws, row, 0, row, lastCol);
    row++;
    setCell(ws, row, 0, 'Показатель', STYLES.headerCenter);
    concepts.forEach((c, i) => {
      const start = 1 + i * width;
      setCell(ws, row, start, c.label, STYLES.headerCenter);
      mergeRange(ws, row, start, row, start + width - 1);
    });
    row++;
    setCell(ws, row, 0, '', STYLES.headerCenter);
    concepts.forEach((_, i) => {
      const start = 1 + i * width;
      subCols.forEach((label, j) => setCell(ws, row, start + j, label, label === 'Total' ? STYLES.base : STYLES.top2Label));
    });
    row++;
  }

  writeGroupedHeader('ОСНОВНЫЕ И ДОПОЛНИТЕЛЬНЫЕ МЕТРИКИ');
  [['Нравится название', 'like'], ['Подходит для блюда / продукта', 'fitDish'], ['Подходит для бренда', 'fitBrand'], ['Намерение посетить БК', 'visitBK'], ['Намерение купить', 'buyDish'], ['Намерение рассказать / поделиться', 'shareIntent']].forEach(([label, key]) => {
    if (!totalStdRes.top2[key]) return;
    setCell(ws, row, 0, label, STYLES.label);
    const vals = groupValue(key, label, false);
    vals.forEach((v, idx) => setPercent(ws, row, idx + 1, v, (idx % width === 0) ? STYLES.top2Row : STYLES.percent));
    row++;
  });
  totalExtraRes.filter(x => x.kind === 'scale5_by_concept').forEach(item => {
    setCell(ws, row, 0, item.title, STYLES.label);
    const vals = groupValue('', item.title, true);
    vals.forEach((v, idx) => setPercent(ws, row, idx + 1, v, (idx % width === 0) ? STYLES.top2Row : STYLES.percent));
    row++;
  });

  row++;
  writeGroupedHeader('ИМИДЖЕВЫЕ ВЫСКАЗЫВАНИЯ');
  Object.entries(totalStdRes.image || {}).forEach(([label, totalVals]) => {
    setCell(ws, row, 0, label, STYLES.label);
    let out = [];
    concepts.forEach((_, ci) => {
      out.push(totalVals[ci] || 0);
      ageData.groups.forEach(group => out.push(group.stdRes.image?.[label]?.[ci] || 0));
    });
    out.forEach((v, idx) => setPercent(ws, row, idx + 1, v, (idx % width === 0) ? STYLES.top2Row : STYLES.percent));
    row++;
  });

  row++;
  setCell(ws, row, 0, 'РАНКИНГ ПО ВОЗРАСТАМ', STYLES.section);
  mergeRange(ws, row, 0, row, 2);
  row++;
  setCell(ws, row, 0, 'Группа', STYLES.headerCenter);
  setCell(ws, row, 1, 'Топ названий', STYLES.headerCenter);
  mergeRange(ws, row, 1, row, 2);
  row++;
  ageData.groups.forEach(group => {
    const vals = group.stdRes.top2.like || [];
    const ranking = concepts.map((c, i) => ({ label: c.label, v: vals[i] || 0 })).sort((a, b) => b.v - a.v).map((x, i) => `${i + 1}. ${x.label} (${Math.round(x.v * 100)}%)`).join(' | ');
    setCell(ws, row, 0, group.label, STYLES.label);
    setCell(ws, row, 1, ranking, STYLES.label);
    mergeRange(ws, row, 1, row, 2);
    row++;
  });

  applySheetRangeRef(ws, row, Math.max(lastCol, 2));
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
