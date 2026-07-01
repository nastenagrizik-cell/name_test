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
    .replace(/[–—]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function canonicalExtraMetricName(headerText) {
  return normalizeText(headerText)
    .replace(/^оцените,? пожалуйста,?\s*/, '')
    .replace(/^для каждого названия\s*/, '')
    .replace(/^насколько\s*/, '')
    .replace(/^скажите,?\s*/, '')
    .replace(/\bкажд(ое|ого|ым|ому)\b/g, '')
    .replace(/\bиз этих названий\b/g, '')
    .replace(/\bдля этого продукта\b/g, 'продукт')
    .replace(/\bдля такого продукта\b/g, 'продукт')
    .replace(/\bдля этого блюда\b/g, 'блюдо')
    .replace(/\bдля такого блюда\b/g, 'блюдо')
    .replace(/\bдля этого напитка\b/g, 'напиток')
    .replace(/\bдля такого напитка\b/g, 'напиток')
    .replace(/\bдля этого соуса\b/g, 'соус')
    .replace(/\bдля такого соуса\b/g, 'соус')
    .replace(/\bдля этого бургера\b/g, 'бургер')
    .replace(/\bдля такого бургера\b/g, 'бургер')
    .replace(/[?.!]+$/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function titleCaseMetricName(s) {
  const text = String(s || '').trim();
  return text ? text.charAt(0).toUpperCase() + text.slice(1) : '';
}

function groupExtraCandidates(extraCandidates) {
  const groups = new Map();
  (extraCandidates || []).forEach(item => {
    const key = canonicalExtraMetricName(item.header) || normalizeText(item.header);
    if (!groups.has(key)) groups.set(key, { key, items: [] });
    groups.get(key).items.push(item);
  });
  return Array.from(groups.values()).map(group => ({
    ...group,
    sampleHeader: group.items[0]?.header || '',
    label: titleCaseMetricName(canonicalExtraMetricName(group.items[0]?.header || '') || group.items[0]?.header || '')
  }));
}

function looksLikeFitDishQuestion(h) {
  const t = normalizeText(h);
  if (t.includes('для бренда бургер кинг')) return false;
  const hasStem = t.includes('насколько каждое из этих названий') && t.includes('подходит') && t.includes('не подходит');
  const hasProductContext = ['для этого','для такого','для воппера','для бургера','для напитка','для капучино','для набора','для соуса','для продукта','для блюда'].some(x => t.includes(x));
  return hasStem && hasProductContext;
}

function looksLikeShareIntentQuestion(h) {
  const t = normalizeText(h);
  return ((t.includes('для каждого названия') || t.includes('оцените') || t.includes('насколько вероятно')) &&
          (t.includes('расскажете') || t.includes('рассказали бы') || t.includes('рассказать') || t.includes('поделитесь') || t.includes('поделиться')) &&
          (t.includes('соцсет') || t.includes('социальных сет') || t.includes('друзьям')));
}

function looksLikeDirectShareQuestion(h) {
  const t = normalizeText(h);
  return ((t.includes('с каким из этих названий') || t.includes('какое из этих названий')) &&
          (t.includes('рассказали') || t.includes('рассказать') || t.includes('поделились') || t.includes('поделиться')) &&
          (t.includes('в первую очередь') || t.includes('сначала')));
}

const IMAGE_STATEMENT_CONFIG = [
  { key: 'Это оригинальный, необычный продукт', aliases: ['это оригинальный, необычный соус','это оригинальный, необычный продукт','это оригинальный, необычный бургер','оригинальный, необычный'] },
  { key: 'Ассоциируется со знакомым вкусом', aliases: ['этот соус ассоциируется со знакомым вкусом','этот бургер ассоциируется со знакомым вкусом','этот продукт ассоциируется со знакомым вкусом'] },
  { key: 'Добавляет премиальности', aliases: ['это премиальный соус','это премиальный продукт','добавляет премиальности'] },
  { key: 'Продукт с юмором', aliases: ['соус с юмором','бургер с юмором','продукт с юмором'] },
  { key: 'Хочется попробовать', aliases: ['хочется попробовать воппер с таким соусом','хочется попробовать такой бургер','хочется попробовать такой капучино','хочется попробовать такой напиток','хочется попробовать такой продукт'] },
  { key: 'Ассоциируется с приятным вкусом', aliases: ['этот соус ассоциируется с приятным вкусом','этот бургер ассоциируется с приятным вкусом','этот продукт ассоциируется с приятным вкусом'] },
  { key: 'Уникальная новинка', aliases: ['это уникальная новинка','это уникальный бургер','это уникальный продукт','оригинальное, отличается от других','это уникальный напиток'] },
  { key: 'Понятное и простое название', aliases: ['понятное и простое название'] },
  { key: 'Вызывает отторжение', aliases: ['этот соус вызывает отторжение','этот продукт вызывает отторжение','этот бургер вызывает отторжение'] },
  { key: 'Понятно, какой будет вкус', aliases: ['понятно, с каким вкусом будет этот бургер','понятно, с каким вкусом будет этот соус','понятно, с каким вкусом будет этот продукт','понятно какой будет вкус'] },
  { key: 'Название легко запомнить', aliases: ['название легко запомнить'] },
  { key: 'Вызывает аппетит, звучит вкусно', aliases: ['вызывает аппетит','вызывает аппетит, вкусно звучит','вызывает аппетит, звучит вкусно','название звучит вкусно и аппетитно'] },
  { key: 'Вызывает доверие', aliases: ['вызывает у меня доверие','вызывает доверие'] },
  { key: 'Звучит как натуральный продукт', aliases: ['звучит как натуральный продукт'] },
  { key: 'Звучит как качественный продукт', aliases: ['звучит как качественный продукт'] },
  { key: 'По названию понятен маленький формат', aliases: ['по названию понятно, что это мини-бургеры','по названию понятно, что это маленький формат','по названию понятно, что это мини формат','маленький формат'] },
  { key: 'По названию понятно, что внутри несколько вкусов', aliases: ['по названию понятно, что внутри несколько разных бургеров','по названию понятно, что внутри несколько разных вкусов','внутри несколько разных вкусов'] },
  { key: 'Хорошо передает идею набора', aliases: ['это название хорошо передает идею набора','хорошо передает идею набора'] },
  { key: 'Звучит хайпово и трендово', aliases: ['звучит хайпово и трендово'] },
  { key: 'Люди бы обсуждали такое название', aliases: ['люди бы обсуждали такое название'] },
  { key: 'Название звучит странно или отталкивающе', aliases: ['название звучит странно или отталкивающе'] },
  { key: 'Название из детского меню / продукт для детей', aliases: ['название из детского меню','продукт для детей'] },
  { key: 'Не сытно / не наешься', aliases: ['не сытно','не наешься'] },
  { key: 'Дешевый продукт', aliases: ['дешевый продукт'] },
  { key: 'Стоит своих денег', aliases: ['стоит своих денег'] },
  { key: 'Подходит для группового потребления', aliases: ['подходит для группового потребления'] },
  { key: 'Звучит старомодно', aliases: ['звучит старомодно'] }
];

function detectImageStatementKey(headerText) {
  const t = normalizeText(headerText);
  for (const item of IMAGE_STATEMENT_CONFIG) {
    if (item.aliases.some(alias => t.includes(normalizeText(alias)))) return item.key;
  }
  return null;
}

if (typeof XLSX !== 'object') {
  status('Ошибка: библиотека XLSX не загружена. Попробуйте обновить страницу.', false, true);
} else if (!baseInput) {
  status('Ошибка: на странице не найден input с id="baseFile".', false, true);
} else {
  baseInput.addEventListener('click', () => { baseInput.value = ''; });
  baseInput.addEventListener('change', async e => {
    baseFile = e.target.files[0] || null;
    resetState();
    if (!baseFile) return status('Файл не выбран');
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
      renderExtraQuestionsUI(autoMapping);
      mappingSection.style.display = '';
      extraSection.style.display = '';
      runSection.style.display = '';
      status('База загружена. Проверьте найденные вопросы и доп. метрики, затем нажмите «Посчитать topline».', true);
    } catch (e) {
      console.error(e);
      status('Ошибка при чтении файла: ' + (e && e.message ? e.message : String(e)), false, true);
    }
  });

  runBtn.addEventListener('click', () => {
    if (!parsed || !autoMapping) return status('Сначала загрузите файл и дождитесь определения вопросов.', false, true);
    try {
      userConfig = collectUserConfig(autoMapping);
    } catch (e) {
      return status('Нужно завершить настройку вопросов: ' + e.message, false, true);
    }

    try {
      runBtn.disabled = true;
      status('Считаю topline...\nПодождите, формируется Excel.');
      const { header, rows } = parsed;
      const concepts = inferConcepts(header, userConfig);
      const stdResults = calcStandardBlocks(rows, userConfig, concepts, header);
      const extraResults = calcExtraBlocks(rows, userConfig, concepts, header);
      const audienceRes = calcAudience(rows, userConfig);
      const signifRes = calcSignificance(stdResults, concepts, rows.length);
      const outWb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(outWb, makeSummarySheetStyled(stdResults, concepts, signifRes, extraResults), 'САММАРИ');
      XLSX.utils.book_append_sheet(outWb, makeFullSheetStyled(stdResults, concepts, extraResults), 'полные таблицы');
      XLSX.utils.book_append_sheet(outWb, makeSignifSheetStyled(stdResults, concepts, signifRes, extraResults), 'значимости');
      XLSX.utils.book_append_sheet(outWb, makeAudienceSheetStyled(audienceRes), 'Аудитория');
      if (ageToggle && ageToggle.checked) {
        if (userConfig.std.audience.age == null || userConfig.std.audience.age < 0) {
          status('Внимание: столбец возраста не найден в базе — лист по возрастам не добавлен.', false, true);
        } else {
          const ageData = calcAgeBreakdown(rows, userConfig, concepts, header);
          if (ageData) {
            const ageSignif = calcAgeSignificance(ageData, stdResults, concepts, extraResults);
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
  parsed = null; autoMapping = null; userConfig = null;
  mappingSection.style.display = 'none'; extraSection.style.display = 'none'; runSection.style.display = 'none';
  standardGroupsEl.innerHTML = ''; extraQuestionsEl.innerHTML = '';
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
  const std = { like: [], fitDish: [], fitBrand: [], visitBK: [], buyDish: [], shareIntent: [], image: [], directLike: [], directBuy: [], directShare: [], audience: { sex: null, age: null, freqNew: null, freqProd: null, freqBK: null } };
  header.forEach((h, idx) => {
    const text = String(h || '').trim();
    const t = normalizeText(text);
    if (!text) return;
    if (text.includes('Оцените, пожалуйста, насколько вам нравится или не нравится каждое из этих названий') || t.includes('нравится или не нравится каждое из этих названий')) std.like.push(idx);
    if (looksLikeFitDishQuestion(text)) std.fitDish.push(idx);
    if (text.includes('А теперь оцените, насколько каждое из этих названий подходит или не подходит для бренда Бургер Кинг') || t.includes('подходит или не подходит для бренда бургер кинг')) std.fitBrand.push(idx);
    if (text.includes('Скажите, насколько вероятно, что Вы посетите ресторан Бургер Кинг') || t.includes('насколько вероятно, что вы посетите ресторан бургер кинг')) std.visitBK.push(idx);
    if (text.includes('Для каждого названия укажите, насколько вероятно, что Вы купите') || t.includes('для каждого названия укажите, насколько вероятно, что вы купите')) std.buyDish.push(idx);
    if (looksLikeShareIntentQuestion(text)) std.shareIntent.push(idx);
    const imageKey = detectImageStatementKey(text); if (imageKey) std.image.push({ key: imageKey, idx });
    if (t.includes('какое из перечисленных ниже названий') || t.includes('какое из этих названий')) std.directLike.push(idx);
    if (t.includes('с каким из этих названий вы бы купили') || t.includes('с каким из этих названий вы купили бы') || (t.includes('с каким из этих названий') && t.includes('купили') && t.includes('в первую очередь'))) std.directBuy.push(idx);
    if (looksLikeDirectShareQuestion(text)) std.directShare.push(idx);
    if (text.includes('Укажите Ваш пол')) std.audience.sex = idx;
    if (text.includes('Укажите Ваш возраст')) std.audience.age = idx;
    if (text.includes('Как часто Вы берете новинки') || text.includes('Как часто вы берете новинки')) std.audience.freqNew = idx;
    if ((text.includes('Как часто вы покупаете') || text.includes('Как часто Вы покупаете')) && std.audience.freqProd == null) std.audience.freqProd = idx;
    if (text.includes('Как часто вы посещаете Бургер Кинг')) std.audience.freqBK = idx;
  });
  const used = new Set([...std.like,...std.fitDish,...std.fitBrand,...std.visitBK,...std.buyDish,...std.shareIntent,...std.image.map(x => x.idx),...std.directLike,...std.directBuy,...std.directShare,std.audience.sex,std.audience.age,std.audience.freqNew,std.audience.freqProd,std.audience.freqBK].filter(v => v != null));
  const extraCandidates = header.map((h, idx) => {
    if (used.has(idx) || !h) return null;
    const lower = normalizeText(h);
    const looksClosed = ['насколько','оцените','выберите','какое из перечисленных','с каким из этих названий','насколько вероятно','какой из этих','что из перечисленного'].some(x => lower.includes(x));
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
    head.innerHTML = `<div class="metric-group-title">${group.label}<span style="font-weight:400;color:var(--muted);font-size:13px"> — ${group.type}</span></div><div class="metric-group-meta"><span class="metric-badge ${badgeClass}">${badgeText}</span><span class="metric-chevron">⌄</span></div>`;
    const body = document.createElement('div');
    body.className = 'metric-group-body';
    if (!found) {
      body.innerHTML = '<div class="metric-line"><span class="metric-line-dot missing"></span><span class="metric-line-text">Колонки не найдены по ключевым словам. Проверьте базу или добавьте метрику вручную.</span></div>';
    } else {
      body.innerHTML = group.indexes.map(idx => `<label class="mapping-item"><input type="checkbox" data-std-key="${group.key === 'directCompare' ? inferDirectKey(header[idx]) : group.key}" data-col-idx="${idx}" checked><span><strong>${header[idx]}</strong><br><small>Колонка ${idx + 1}</small></span></label>`).join('');
    }
    head.addEventListener('click', () => card.classList.toggle('open'));
    card.append(head, body);
    standardGroupsEl.appendChild(card);
  });
}

function inferDirectKey(headerText) {
  const t = normalizeText(headerText);
  if (t.includes('купили')) return 'directBuy';
  if (t.includes('рассказ')) return 'directShare';
  return 'directLike';
}

function renderExtraQuestionsUI(mapping) {
  const grouped = groupExtraCandidates(mapping.extraCandidates || []);
  extraQuestionsEl.innerHTML = '';
  if (!grouped.length) {
    extraQuestionsEl.innerHTML = '<div class="status">Дополнительные закрытые вопросы не найдены.</div>';
    return;
  }
  grouped.forEach((group, i) => {
    const card = document.createElement('div');
    card.className = 'extra-card-compact';
    card.innerHTML = `<div class="extra-card-compact-head"><div><div class="metric-group-title">${group.label || group.sampleHeader}</div><div class="panel-subtitle" style="font-size:13px">${group.items.length > 1 ? `Найдено похожих формулировок: ${group.items.length}` : 'Одна формулировка'}</div></div><span class="metric-chevron">⌄</span></div><div class="extra-card-compact-body"><div class="field"><label for="extraGroupName_${i}">Название метрики в topline</label><input id="extraGroupName_${i}" type="text" data-extra-group-name="${group.key}" value="${group.label || group.sampleHeader}" /></div><div class="mapping-list">${group.items.map(item => `<label class="mapping-item"><input type="checkbox" data-extra-index="${item.idx}" data-extra-group="${group.key}" checked><span><strong>${item.header}</strong><br><small>Колонка ${item.idx + 1}</small></span></label>`).join('')}</div></div>`;
    card.querySelector('.extra-card-compact-head').addEventListener('click', () => card.classList.toggle('open'));
    extraQuestionsEl.appendChild(card);
  });
}

function collectUserConfig(mapping) {
  const stdSelected = { like: [], fitDish: [], fitBrand: [], visitBK: [], buyDish: [], shareIntent: [], directLike: [], directBuy: [], directShare: [], image: mapping.std.image.slice(), audience: mapping.std.audience };
  document.querySelectorAll('input[type="checkbox"][data-std-key]').forEach(cb => {
    if (!cb.checked) return;
    const key = cb.getAttribute('data-std-key');
    stdSelected[key].push(Number(cb.getAttribute('data-col-idx')));
  });
  const extra = Array.from(extraQuestionsEl.querySelectorAll('[data-extra-index]:checked')).map(el => {
    const groupKey = el.getAttribute('data-extra-group');
    const nameInput = extraQuestionsEl.querySelector(`[data-extra-group-name="${groupKey}"]`);
    const title = (nameInput?.value || '').trim();
    if (!title) throw new Error('Не задано название одной из дополнительных метрик.');
    return { idx: Number(el.getAttribute('data-extra-index')), header: '', title, type: 'scale5', where: ['summary','full','signif','age'], groupKey };
  });
  return { std: stdSelected, extra };
}

function inferConcepts(header, config) {
  const sourceCols = config.std.like.length ? config.std.like : config.std.fitDish.length ? config.std.fitDish : config.std.fitBrand.length ? config.std.fitBrand : config.std.buyDish.length ? config.std.buyDish : config.std.shareIntent.length ? config.std.shareIntent : config.std.visitBK.length ? config.std.visitBK : [];
  if (!sourceCols.length) return [{ code: 'A', label: 'Название A' }];
  const labels = sourceCols.map(colIdx => {
    const text = String(header[colIdx] || '').trim();
    const parts = text.split(' - ');
    return parts.length > 1 ? parts[parts.length - 1].trim() : text;
  });
  return labels.map((label, i) => ({ code: String.fromCharCode(65 + i), label: label || `Название ${String.fromCharCode(65 + i)}` }));
}

function getCell(row, idx) { return idx == null || idx < 0 ? null : row[idx]; }
function parseScaleValue(v) { if (v === null || v === undefined || v === '') return null; if (typeof v === 'number') return v; const m = String(v).trim().match(/^([1-5])/); return m ? Number(m[1]) : null; }
const AGE_GROUPS = [{ key: '18-24', label: '18-24', min: 18, max: 24 }, { key: '25-34', label: '25-34', min: 25, max: 34 }, { key: '35-44', label: '35-44', min: 35, max: 44 }, { key: '45+', label: '45+', min: 45, max: 999 }];
function parseAgeValue(v) { if (v === null || v === undefined || v === '') return null; if (typeof v === 'number') return v; const m = String(v).trim().match(/(\d+)/); return m ? Number(m[1]) : null; }
function ageGroupForValue(age) { if (age == null || age < 18) return null; return AGE_GROUPS.find(g => age >= g.min && age <= g.max) || null; }
function splitRowsByAge(rows, ageColIdx) { const groups = {}; AGE_GROUPS.forEach(g => { groups[g.key] = []; }); let unassigned = 0; rows.forEach(r => { const g = ageGroupForValue(parseAgeValue(getCell(r, ageColIdx))); if (g) groups[g.key].push(r); else unassigned++; }); return { groups, unassigned }; }
function normalizeConceptLabel(s) { return normalizeText(s).replace(/\bвоппер\b|\bбургер\b|\bкапучино\b|\bнапиток\b|\bсоус\b|\bнабор\b|\bпродукт\b/g, '').replace(/\s+/g, ' ').trim(); }
function findConceptIndexByHeader(headerText, concepts) { const text = normalizeText(headerText); for (let i=0;i<concepts.length;i++) if (text.endsWith(normalizeText(concepts[i].label))) return i; for (let i=0;i<concepts.length;i++) if (text.includes(normalizeText(concepts[i].label))) return i; const normalizedHeader = normalizeConceptLabel(text); for (let i=0;i<concepts.length;i++) { const conceptNorm = normalizeConceptLabel(concepts[i].label); if (conceptNorm && normalizedHeader.includes(conceptNorm)) return i; } return -1; }

function calcStandardBlocks(rows, config, concepts, header) {
  const n = rows.length;
  function top2ByCols(cols) {
    if (!cols.length) return null;
    const res = Array(concepts.length).fill(0);
    rows.forEach(r => cols.forEach((col, i) => { if (i < concepts.length) { const v = parseScaleValue(getCell(r, col)); if (v === 4 || v === 5) res[i]++; } }));
    return res.map(v => v / n);
  }
  function dist5(cols) {
    if (!cols.length) return null;
    const arr = Array.from({ length: concepts.length }, () => ({ '1': 0, '2': 0, '3': 0, '4': 0, '5': 0 }));
    rows.forEach(r => cols.forEach((col, i) => { if (i < concepts.length) { const v = parseScaleValue(getCell(r, col)); if (v >= 1 && v <= 5) arr[i][String(v)]++; } }));
    return arr.map(d => ({ '1': d['1'] / n, '2': d['2'] / n, '3': d['3'] / n, '4': d['4'] / n, '5': d['5'] / n, top2: (d['4'] + d['5']) / n }));
  }
  function imageBlock() {
    const res = {}; IMAGE_STATEMENT_CONFIG.forEach(item => { res[item.key] = Array(concepts.length).fill(0); });
    rows.forEach(r => config.std.image.forEach(({ key, idx }) => { const val = String(getCell(r, idx) || '').trim(); if (!val) return; const conceptIndex = findConceptIndexByHeader(header[idx], concepts); if (conceptIndex !== -1 && res[key]) res[key][conceptIndex]++; }));
    Object.keys(res).forEach(k => { res[k] = res[k].map(v => v / n); });
    return Object.fromEntries(Object.entries(res).filter(([, vals]) => vals.some(v => v > 0)));
  }
  function directSingle(cols) {
    if (!cols.length) return null;
    const counts = {};
    cols.forEach(idx => rows.forEach(r => { const v = String(getCell(r, idx) || '').trim(); if (v) counts[v] = (counts[v] || 0) + 1; }));
    const perConcept = Array(concepts.length).fill(0); let none = 0;
    concepts.forEach((c, i) => { const key = Object.keys(counts).find(k => normalizeText(k).includes(normalizeText(c.label))); perConcept[i] = key ? counts[key] / n : 0; });
    const noneKey = Object.keys(counts).find(k => ['ни одно','ничего из перечисленного','не купил','не рассказал'].some(x => normalizeText(k).includes(x))); if (noneKey) none = counts[noneKey] / n;
    return { perConcept, none };
  }
  return {
    n,
    scales: { like: dist5(config.std.like), fitDish: dist5(config.std.fitDish), fitBrand: dist5(config.std.fitBrand), visitBK: dist5(config.std.visitBK), buyDish: dist5(config.std.buyDish), shareIntent: dist5(config.std.shareIntent) },
    top2: { like: top2ByCols(config.std.like), fitDish: top2ByCols(config.std.fitDish), fitBrand: top2ByCols(config.std.fitBrand), visitBK: top2ByCols(config.std.visitBK), buyDish: top2ByCols(config.std.buyDish), shareIntent: top2ByCols(config.std.shareIntent) },
    image: imageBlock(),
    direct: { likeMost: directSingle(config.std.directLike), buyFirst: directSingle(config.std.directBuy), shareFirst: directSingle(config.std.directShare) }
  };
}

function calcExtraBlocks(rows, config, concepts, header) {
  const grouped = new Map();
  config.extra.forEach(item => {
    const key = item.groupKey || item.title;
    if (!grouped.has(key)) grouped.set(key, { title: item.title, items: [] });
    grouped.get(key).items.push(item);
  });
  return Array.from(grouped.values()).map(group => {
    const dist = Array.from({ length: concepts.length }, () => ({ '1': 0, '2': 0, '3': 0, '4': 0, '5': 0, base: 0 }));
    group.items.forEach(item => {
      const conceptIndex = findConceptIndexByHeader(header[item.idx], concepts);
      if (conceptIndex < 0) return;
      rows.forEach(r => {
        const v = parseScaleValue(getCell(r, item.idx));
        if (v >= 1 && v <= 5) { dist[conceptIndex][String(v)]++; dist[conceptIndex].base++; }
      });
    });
    return {
      kind: 'scale5_by_concept',
      title: group.title,
      dist: dist.map(counts => ({ '1': counts.base ? counts['1'] / counts.base : 0, '2': counts.base ? counts['2'] / counts.base : 0, '3': counts.base ? counts['3'] / counts.base : 0, '4': counts.base ? counts['4'] / counts.base : 0, '5': counts.base ? counts['5'] / counts.base : 0, top2: counts.base ? (counts['4'] + counts['5']) / counts.base : 0 }))
    };
  });
}

function calcAudience(rows, config) {
  const result = [];
  const defs = [
    { key: 'sex', label: 'Пол', idx: config.std.audience.sex },
    { key: 'age', label: 'Возраст', idx: config.std.audience.age },
    { key: 'freqNew', label: 'Частота покупки новинок', idx: config.std.audience.freqNew },
    { key: 'freqProd', label: 'Частота покупки продукта', idx: config.std.audience.freqProd },
    { key: 'freqBK', label: 'Частота посещения БК', idx: config.std.audience.freqBK }
  ];
  defs.forEach(def => {
    if (def.idx == null || def.idx < 0) return;
    const counts = new Map();
    rows.forEach(r => { const val = String(getCell(r, def.idx) || '').trim(); if (val) counts.set(val, (counts.get(val) || 0) + 1); });
    result.push({ label: def.label, rows: Array.from(counts.entries()).map(([name, count]) => ({ name, count, share: count / rows.length })) });
  });
  return result;
}

function zTestProp(p1, n1, p2, n2) {
  if (!n1 || !n2) return null;
  const pooled = (p1 * n1 + p2 * n2) / (n1 + n2);
  const se = Math.sqrt(pooled * (1 - pooled) * (1 / n1 + 1 / n2));
  if (!se) return null;
  return (p1 - p2) / se;
}

function calcSignificance(stdResults, concepts, totalN) {
  const out = {};
  Object.entries(stdResults.top2).forEach(([metricKey, vals]) => {
    if (!vals) return;
    out[metricKey] = vals.map((v, i) => {
      let betterThan = [];
      for (let j = 0; j < vals.length; j++) {
        if (j === i) continue;
        const z = zTestProp(v, totalN, vals[j], totalN);
        if (z != null && z > 1.96) betterThan.push(concepts[j].code);
      }
      return betterThan.join(', ');
    });
  });
  return out;
}

function calcAgeBreakdown(rows, config, concepts, header) {
  const ageColIdx = config.std.audience.age;
  if (ageColIdx == null || ageColIdx < 0) return null;
  const split = splitRowsByAge(rows, ageColIdx);
  const groups = AGE_GROUPS.map(g => {
    const gRows = split.groups[g.key] || [];
    const stdRes = calcStandardBlocks(gRows.length ? gRows : [], config, concepts, header);
    const extraRes = calcExtraBlocks(gRows.length ? gRows : [], config, concepts, header);
    return { key: g.key, label: g.label, n: gRows.length, stdRes, extraRes };
  });
  return { total: rows.length, groups, unassigned: split.unassigned };
}

function calcAgeSignificance(ageData, totalStdRes, concepts, totalExtraRes = []) {
  const byMetric = {};
  const blocks = [
    { key: 'like', title: 'Нравится название', getter: x => x.stdRes.top2.like },
    { key: 'fitDish', title: 'Подходит для блюда / продукта', getter: x => x.stdRes.top2.fitDish },
    { key: 'fitBrand', title: 'Подходит для бренда', getter: x => x.stdRes.top2.fitBrand },
    { key: 'visitBK', title: 'Намерение посетить БК', getter: x => x.stdRes.top2.visitBK },
    { key: 'buyDish', title: 'Намерение купить', getter: x => x.stdRes.top2.buyDish },
    { key: 'shareIntent', title: 'Намерение рассказать / поделиться', getter: x => x.stdRes.top2.shareIntent }
  ];
  blocks.forEach(block => {
    byMetric[block.key] = AGE_GROUPS.map((_, gi) => {
      const arr = block.getter(ageData.groups[gi]) || [];
      return arr.map((p, ci) => {
        const total = (totalStdRes.top2[block.key] || [])[ci];
        const z = zTestProp(p || 0, ageData.groups[gi].n || 0, total || 0, ageData.total || 0);
        return z != null && z > 1.96 ? 'выше тотала' : '';
      });
    });
  });
  totalExtraRes.forEach(extra => {
    byMetric['extra__' + extra.title] = AGE_GROUPS.map((_, gi) => {
      const match = ageData.groups[gi].extraRes.find(x => x.title === extra.title);
      return (match?.dist || []).map((p, ci) => {
        const total = (extra.dist || [])[ci]?.top2 || 0;
        const z = zTestProp((p?.top2 || 0), ageData.groups[gi].n || 0, total, ageData.total || 0);
        return z != null && z > 1.96 ? 'выше тотала' : '';
      });
    });
  });
  return byMetric;
}

const STYLES = {
  title: { font: { bold: true, sz: 14, color: { rgb: 'FFFFFF' } }, fill: { fgColor: { rgb: '244C73' } }, alignment: { vertical: 'center', horizontal: 'center', wrapText: true }, border: boxBorder('AEBFD1') },
  blockHead: { font: { bold: true, sz: 12, color: { rgb: '16324A' } }, fill: { fgColor: { rgb: 'EAF2FB' } }, alignment: { vertical: 'center', horizontal: 'left', wrapText: true }, border: boxBorder('C7D6E5') },
  subHead: { font: { bold: true, sz: 11, color: { rgb: '16324A' } }, fill: { fgColor: { rgb: 'EDF1F5' } }, alignment: { vertical: 'center', horizontal: 'center', wrapText: true }, border: boxBorder('D8E2EC') },
  left: { font: { sz: 10, color: { rgb: '16324A' } }, alignment: { vertical: 'center', horizontal: 'left', wrapText: true }, border: boxBorder('E2E8F0') },
  metric: { font: { bold: true, sz: 10, color: { rgb: '16324A' } }, fill: { fgColor: { rgb: 'F8FBFF' } }, alignment: { vertical: 'center', horizontal: 'left', wrapText: true }, border: boxBorder('E2E8F0') },
  totalCell: { font: { bold: true, sz: 10, color: { rgb: '16324A' } }, fill: { fgColor: { rgb: 'EDF1F5' } }, alignment: { vertical: 'center', horizontal: 'center' }, border: boxBorder('D8E2EC') },
  val: { font: { sz: 10, color: { rgb: '16324A' } }, alignment: { vertical: 'center', horizontal: 'center' }, border: boxBorder('E2E8F0') },
  signif: { font: { sz: 9, italic: true, color: { rgb: '326B12' } }, alignment: { vertical: 'center', horizontal: 'center', wrapText: true }, border: boxBorder('E2E8F0') },
  note: { font: { italic: true, sz: 10, color: { rgb: '6B7F92' } }, alignment: { vertical: 'center', horizontal: 'left', wrapText: true } },
  rank: { font: { bold: true, sz: 10, color: { rgb: '16324A' } }, fill: { fgColor: { rgb: 'EEF8E8' } }, alignment: { vertical: 'center', horizontal: 'left', wrapText: true }, border: boxBorder('D7E7CA') }
};
function boxBorder(color) { return { top: { style: 'thin', color: { rgb: color } }, bottom: { style: 'thin', color: { rgb: color } }, left: { style: 'thin', color: { rgb: color } }, right: { style: 'thin', color: { rgb: color } } }; }
function pct(v) { return v == null || Number.isNaN(v) ? '' : Math.round(v * 100) + '%'; }
function sheetFromAOA(aoa, merges, cols, styles = []) { const ws = XLSX.utils.aoa_to_sheet(aoa); ws['!merges'] = merges; ws['!cols'] = cols; styles.forEach(x => { if (ws[x.ref]) ws[x.ref].s = x.style; }); return ws; }
function pushCellStyle(styles, r, c, style) { styles.push({ ref: XLSX.utils.encode_cell({ r, c }), style }); }

function metricDefinitions(stdRes, extraResults) {
  const defs = [
    { key: 'like', title: 'Нравится название', source: stdRes.top2.like },
    { key: 'fitDish', title: 'Подходит для блюда / продукта', source: stdRes.top2.fitDish },
    { key: 'fitBrand', title: 'Подходит для бренда', source: stdRes.top2.fitBrand },
    { key: 'visitBK', title: 'Намерение посетить БК', source: stdRes.top2.visitBK },
    { key: 'buyDish', title: 'Намерение купить', source: stdRes.top2.buyDish },
    { key: 'shareIntent', title: 'Намерение рассказать / поделиться', source: stdRes.top2.shareIntent }
  ].filter(x => x.source);
  extraResults.forEach(extra => defs.push({ key: 'extra__' + extra.title, title: extra.title, source: extra.dist.map(x => x.top2), isExtra: true }));
  return defs;
}

function imageMetricDefinitions(stdRes) {
  return Object.entries(stdRes.image || {}).map(([title, values]) => ({ key: 'image__' + title, title, source: values }));
}

function buildConceptColumnGroups(aoa, styles, merges, startRow, concepts, groupLabels, valueAccessor, signifAccessor) {
  let row = startRow;
  aoa[row] = ['Метрика'];
  let col = 1;
  concepts.forEach((concept, ci) => {
    aoa[row][col] = concept.label;
    merges.push({ s: { r: row, c: col }, e: { r: row, c: col + groupLabels.length - 1 } });
    for (let cc = col; cc < col + groupLabels.length; cc++) pushCellStyle(styles, row, cc, STYLES.blockHead);
    col += groupLabels.length;
  });
  pushCellStyle(styles, row, 0, STYLES.title);
  row++;
  aoa[row] = [''];
  col = 1;
  concepts.forEach(() => {
    groupLabels.forEach(label => { aoa[row][col] = label; pushCellStyle(styles, row, col, label === 'Total' ? STYLES.totalCell : STYLES.subHead); col++; });
  });
  pushCellStyle(styles, row, 0, STYLES.subHead);
  row++;
  return { row, writeMetric: (title, valuesByConcept) => {
    aoa[row] = [title];
    pushCellStyle(styles, row, 0, STYLES.metric);
    let c = 1;
    concepts.forEach((_, ci) => {
      groupLabels.forEach(label => {
        const value = valueAccessor(valuesByConcept, ci, label);
        aoa[row][c] = value;
        pushCellStyle(styles, row, c, label === 'Total' ? STYLES.totalCell : STYLES.val);
        c++;
      });
    });
    row++;
    if (signifAccessor) {
      aoa[row] = ['Значимость'];
      pushCellStyle(styles, row, 0, STYLES.left);
      c = 1;
      concepts.forEach((_, ci) => {
        groupLabels.forEach(label => {
          aoa[row][c] = signifAccessor(valuesByConcept, ci, label) || '';
          pushCellStyle(styles, row, c, STYLES.signif);
          c++;
        });
      });
      row++;
    }
  }, endRow: () => row };
}

function rankText(labelsAndValues) {
  return labelsAndValues.sort((a,b) => b.value - a.value).map((x, i) => `${i+1}. ${x.label} (${pct(x.value)})`).join('   ');
}

function makeSummarySheetStyled(stdRes, concepts, signifRes, extraResults = []) {
  const aoa = [['САММАРИ']]; const styles = []; const merges = [];
  pushCellStyle(styles, 0, 0, STYLES.title);
  aoa.push([`Выборка: n=${stdRes.n}`]); pushCellStyle(styles, 1, 0, STYLES.note);
  let row = 3;
  aoa[row] = ['Основные и дополнительные метрики']; pushCellStyle(styles, row, 0, STYLES.blockHead); row += 2;
  const defs = metricDefinitions(stdRes, extraResults);
  const builder = buildConceptColumnGroups(aoa, styles, merges, row, concepts, ['Total'], (vals, ci) => pct(vals[ci]), (vals, ci) => {
    const def = defs.find(d => d.source === vals);
    return def && signifRes[def.key?.replace('extra__','')] ? signifRes[def.key.replace('extra__','')][ci] : '';
  });
  defs.forEach(def => builder.writeMetric(def.title, def.source));
  row = builder.endRow() + 1;
  aoa[row] = ['Имиджевые высказывания']; pushCellStyle(styles, row, 0, STYLES.blockHead); row += 2;
  const imgBuilder = buildConceptColumnGroups(aoa, styles, merges, row, concepts, ['Total'], (vals, ci) => pct(vals[ci]));
  imageMetricDefinitions(stdRes).forEach(def => imgBuilder.writeMetric(def.title, def.source));
  row = imgBuilder.endRow() + 1;
  aoa[row] = ['Ранкинг названий']; pushCellStyle(styles, row, 0, STYLES.blockHead); row++;
  defs.forEach(def => { aoa[row] = [def.title, rankText(concepts.map((c, i) => ({ label: c.label, value: def.source[i] || 0 })))]; pushCellStyle(styles, row, 0, STYLES.metric); pushCellStyle(styles, row, 1, STYLES.rank); row++; });
  const cols = [{ wch: 38 }, ...Array(concepts.length).fill({ wch: 16 })];
  return sheetFromAOA(aoa, merges, cols, styles);
}

function makeFullSheetStyled(stdRes, concepts, extraResults = []) {
  const aoa = [['Полные таблицы']]; const styles = []; const merges = [];
  pushCellStyle(styles, 0, 0, STYLES.title);
  aoa.push([`Выборка: n=${stdRes.n}`]); pushCellStyle(styles, 1, 0, STYLES.note);
  let row = 3;
  const defs = metricDefinitions(stdRes, extraResults);
  defs.forEach(def => {
    aoa[row] = [def.title]; pushCellStyle(styles, row, 0, STYLES.blockHead); row += 2;
    const builder = buildConceptColumnGroups(aoa, styles, merges, row, concepts, ['Total'], (vals, ci) => pct(vals[ci]));
    builder.writeMetric('Top-2', def.source);
    row = builder.endRow() + 1;
  });
  aoa[row] = ['Имиджевые высказывания']; pushCellStyle(styles, row, 0, STYLES.blockHead); row += 2;
  const imgBuilder = buildConceptColumnGroups(aoa, styles, merges, row, concepts, ['Total'], (vals, ci) => pct(vals[ci]));
  imageMetricDefinitions(stdRes).forEach(def => imgBuilder.writeMetric(def.title, def.source));
  const cols = [{ wch: 38 }, ...Array(concepts.length).fill({ wch: 16 })];
  return sheetFromAOA(aoa, merges, cols, styles);
}

function makeSignifSheetStyled(stdRes, concepts, signifRes, extraResults = []) {
  const aoa = [['Значимости']]; const styles = []; const merges = [];
  pushCellStyle(styles, 0, 0, STYLES.title);
  aoa.push(['Показывает Top-2 и названия, которые значимо уступают текущему.']); pushCellStyle(styles, 1, 0, STYLES.note);
  let row = 3;
  const defs = metricDefinitions(stdRes, extraResults);
  defs.forEach(def => {
    aoa[row] = [def.title]; pushCellStyle(styles, row, 0, STYLES.blockHead); row += 1;
    aoa[row] = ['Название', 'Top-2', 'Лучше, чем'];
    [0,1,2].forEach(c => pushCellStyle(styles, row, c, STYLES.subHead)); row++;
    concepts.forEach((concept, i) => {
      aoa[row] = [concept.label, pct(def.source[i]), (signifRes[def.key?.replace('extra__','')] || [])[i] || ''];
      pushCellStyle(styles, row, 0, STYLES.metric); pushCellStyle(styles, row, 1, STYLES.val); pushCellStyle(styles, row, 2, STYLES.signif); row++;
    });
    row++;
  });
  const cols = [{ wch: 38 }, { wch: 12 }, { wch: 22 }];
  return sheetFromAOA(aoa, merges, cols, styles);
}

function makeAudienceSheetStyled(audienceRes) {
  const aoa = [['Аудитория']]; const styles = []; const merges = [];
  pushCellStyle(styles, 0, 0, STYLES.title);
  let row = 2;
  audienceRes.forEach(block => {
    aoa[row] = [block.label]; pushCellStyle(styles, row, 0, STYLES.blockHead); row++;
    aoa[row] = ['Категория', 'Кол-во', '%']; [0,1,2].forEach(c => pushCellStyle(styles, row, c, STYLES.subHead)); row++;
    block.rows.forEach(item => { aoa[row] = [item.name, item.count, pct(item.share)]; pushCellStyle(styles, row, 0, STYLES.left); pushCellStyle(styles, row, 1, STYLES.val); pushCellStyle(styles, row, 2, STYLES.val); row++; });
    row++;
  });
  return sheetFromAOA(aoa, merges, [{ wch: 40 }, { wch: 12 }, { wch: 12 }], styles);
}

function makeAgeSheetStyled(ageData, ageSignif, totalStdRes, totalExtraRes, concepts) {
  const aoa = [['Возраст']]; const styles = []; const merges = [];
  pushCellStyle(styles, 0, 0, STYLES.title);
  const sampleLine = 'Выборка: total n=' + ageData.total + '; ' + ageData.groups.map(g => `${g.label}: n=${g.n}`).join('; ');
  aoa.push([sampleLine]); pushCellStyle(styles, 1, 0, STYLES.note);
  let row = 3;

  const defs = metricDefinitions(totalStdRes, totalExtraRes);
  aoa[row] = ['Основные и дополнительные метрики']; pushCellStyle(styles, row, 0, STYLES.blockHead); row += 2;
  const ageLabels = ['Total', ...AGE_GROUPS.map(g => g.label)];
  const builder = buildConceptColumnGroups(
    aoa, styles, merges, row, concepts, ageLabels,
    (vals, ci, label) => {
      if (label === 'Total') return pct(vals[ci]);
      const group = ageData.groups.find(x => x.label === label);
      return pct((group?.metricMap?.[vals]?.[ci]) || 0);
    },
    (vals, ci, label) => {
      if (label === 'Total') return '';
      const def = defs.find(d => d.source === vals);
      const gi = AGE_GROUPS.findIndex(g => g.label === label);
      return (ageSignif[def.key] || [])[gi]?.[ci] || '';
    }
  );

  ageData.groups.forEach(group => {
    group.metricMap = {};
    defs.forEach(def => {
      if (def.key.startsWith('extra__')) {
        const match = group.extraRes.find(x => x.title === def.title);
        group.metricMap[def.source] = (match?.dist || []).map(x => x.top2);
      } else {
        group.metricMap[def.source] = group.stdRes.top2[def.key] || [];
      }
    });
  });

  defs.forEach(def => builder.writeMetric(def.title, def.source));
  row = builder.endRow() + 1;

  aoa[row] = ['Имиджевые высказывания']; pushCellStyle(styles, row, 0, STYLES.blockHead); row += 2;
  const imgDefs = imageMetricDefinitions(totalStdRes);
  ageData.groups.forEach(group => {
    group.imageMap = {};
    imgDefs.forEach(def => { group.imageMap[def.source] = (group.stdRes.image || {})[def.title] || []; });
  });
  const imgBuilder = buildConceptColumnGroups(
    aoa, styles, merges, row, concepts, ageLabels,
    (vals, ci, label) => {
      if (label === 'Total') return pct(vals[ci]);
      const group = ageData.groups.find(x => x.label === label);
      return pct((group?.imageMap?.[vals]?.[ci]) || 0);
    }
  );
  imgDefs.forEach(def => imgBuilder.writeMetric(def.title, def.source));
  row = imgBuilder.endRow() + 1;

  aoa[row] = ['Ранкинг по возрастам']; pushCellStyle(styles, row, 0, STYLES.blockHead); row++;
  AGE_GROUPS.forEach((groupDef, gi) => {
    aoa[row] = [groupDef.label, rankText(concepts.map((c, i) => ({ label: c.label, value: ageData.groups[gi].stdRes.top2.like?.[i] || 0 })))];
    pushCellStyle(styles, row, 0, STYLES.metric); pushCellStyle(styles, row, 1, STYLES.rank); row++;
  });

  const cols = [{ wch: 38 }, ...Array(concepts.length * ageLabels.length).fill({ wch: 14 })];
  return sheetFromAOA(aoa, merges, cols, styles);
}
