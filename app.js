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

function show(msg, type = '') {
  statusEl.textContent = msg;
  statusEl.className = type ? `status ${type}` : 'status';
}

function append(msg) {
  statusEl.textContent += `\n${msg}`;
}

show('Страница загружена. XLSX = ' + typeof XLSX);

if (!baseInput) {
  show('Ошибка: input #baseFile не найден', 'error');
} else {
  append('input #baseFile найден');
}

baseInput?.addEventListener('change', async (e) => {
  show('Событие change сработало');

  const file = e.target.files && e.target.files[0];
  if (!file) {
    show('Файл не выбран', 'error');
    return;
  }

  baseFile = file;
  append('Файл выбран: ' + file.name);
  append('Размер: ' + file.size + ' байт');

  try {
    append('Читаю arrayBuffer...');
    const buffer = await file.arrayBuffer();
    append('arrayBuffer ок: ' + buffer.byteLength + ' байт');

    append('Запускаю XLSX.read...');
    const workbook = XLSX.read(buffer, { type: 'array' });

    append('Листы: ' + workbook.SheetNames.join(', '));

    if (!workbook.SheetNames.length) {
      throw new Error('В файле нет листов');
    }

    const firstSheetName = workbook.SheetNames[0];
    const firstSheet = workbook.Sheets[firstSheetName];
    append('Первый лист: ' + firstSheetName);

    const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
    append('Строк считано: ' + rows.length);

    parsed = {
      header: rows[0] || [],
      rows
    };

    mappingSection.style.display = 'block';
    extraSection.style.display = 'block';
    runSection.style.display = 'block';

    standardGroupsEl.innerHTML = `
      <div class="col-half">
        <div class="mapping-group">
          <div class="mapping-group-title">Файл считан успешно</div>
          <div class="mapping-list">
            <div class="mapping-item">
              <label>
                Название файла: <strong>${file.name}</strong><br>
                Листов: <strong>${workbook.SheetNames.length}</strong><br>
                Строк на первом листе: <strong>${rows.length}</strong>
              </label>
            </div>
          </div>
        </div>
      </div>
    `;

    extraQuestionsEl.innerHTML = `
      <div class="field">
        <label>Результат проверки</label>
        <div class="mapping-list">
          <div class="mapping-item">
            <label>
              Если ты видишь этот блок — файл точно читается, и проблема была не в XLSX.
            </label>
          </div>
        </div>
      </div>
    `;

    show(
      'Готово.\n' +
      'XLSX = ' + typeof XLSX + '\n' +
      'Файл выбран: ' + file.name + '\n' +
      'Листов: ' + workbook.SheetNames.length + '\n' +
      'Строк на первом листе: ' + rows.length,
      'ok'
    );
  } catch (err) {
    show(
      'Ошибка при чтении файла:\n' +
      (err && err.message ? err.message : String(err)),
      'error'
    );
  }
});

runBtn?.addEventListener('click', () => {
  show('Кнопка "Посчитать топлайн" нажалась. Значит интерфейс живой.', 'ok');
});
