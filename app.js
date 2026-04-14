console.log('app.js loaded, XLSX =', typeof XLSX);

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

console.log('elements:', {
  baseInput: !!baseInput,
  statusEl: !!statusEl,
  mappingSection: !!mappingSection,
  extraSection: !!extraSection,
  runSection: !!runSection,
  standardGroupsEl: !!standardGroupsEl,
  extraQuestionsEl: !!extraQuestionsEl,
  runBtn: !!runBtn
});

function setStatus(text, type = '') {
  if (!statusEl) return;
  statusEl.textContent = text;
  statusEl.className = type ? `status ${type}` : 'status';
}

function resetUI() {
  console.log('resetUI called');

  if (mappingSection) mappingSection.style.display = 'none';
  if (extraSection) extraSection.style.display = 'none';
  if (runSection) runSection.style.display = 'none';
  if (standardGroupsEl) standardGroupsEl.innerHTML = '';
  if (extraQuestionsEl) extraQuestionsEl.innerHTML = '';

  parsed = null;
  autoMapping = null;
  userConfig = null;
}

function renderSuccessInfo(rows, workbook) {
  console.log('renderSuccessInfo called');

  if (mappingSection) mappingSection.style.display = 'block';
  if (extraSection) extraSection.style.display = 'block';
  if (runSection) runSection.style.display = 'block';

  if (standardGroupsEl) {
    standardGroupsEl.innerHTML = `
      <div class="col-half">
        <div class="mapping-group">
          <div class="mapping-group-title">Файл считан успешно</div>
          <div class="mapping-list">
            <div class="mapping-item">
              <label>
                Найдено строк: <strong>${rows.length}</strong><br>
                Листов: <strong>${workbook.SheetNames.length}</strong><br>
                Первый лист: <strong>${workbook.SheetNames[0] || '-'}</strong>
              </label>
            </div>
          </div>
        </div>
      </div>
    `;
  }

  if (extraQuestionsEl) {
    extraQuestionsEl.innerHTML = `
      <div class="field">
        <label>Отладочная информация</label>
        <div class="mapping-list">
          <div class="mapping-item">
            <label>
              Если ты видишь этот блок, значит:
              <br>1. файл выбрался,
              <br>2. событие change сработало,
              <br>3. XLSX.read отработал,
              <br>4. интерфейс умеет показывать следующий шаг.
            </label>
          </div>
        </div>
      </div>
    `;
  }

  setStatus('Файл прочитан. Если дальше пусто, значит проблема уже в основной логике построения интерфейса.', 'ok');
}

if (!baseInput) {
  console.error('baseInput not found');
  setStatus('Ошибка: на странице не найден input #baseFile', 'error');
} else {
  baseInput.addEventListener('change', async (e) => {
    console.log('change fired');

    resetUI();

    const file = e.target.files && e.target.files[0];
    console.log('selected file =', file);

    if (!file) {
      console.log('no file selected');
      setStatus('Файл не выбран.', 'error');
      return;
    }

    baseFile = file;
    setStatus('Читаю файл...');

    try {
      console.log('before arrayBuffer');
      const buffer = await file.arrayBuffer();
      console.log('arrayBuffer ok, bytes =', buffer.byteLength);

      console.log('before XLSX.read');
      const workbook = XLSX.read(buffer, { type: 'array' });
      console.log('workbook =', workbook);
      console.log('sheet names =', workbook.SheetNames);

      if (!workbook.SheetNames || !workbook.SheetNames.length) {
        throw new Error('В Excel не найдено ни одного листа');
      }

      const firstSheetName = workbook.SheetNames[0];
      const firstSheet = workbook.Sheets[firstSheetName];
      console.log('first sheet name =', firstSheetName);
      console.log('first sheet =', firstSheet);

      const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
      console.log('rows count =', rows.length);
      console.log('first 10 rows =', rows.slice(0, 10));

      parsed = {
        header: rows[0] || [],
        rows
      };

      console.log('parsed =', parsed);

      renderSuccessInfo(rows, workbook);
    } catch (err) {
      console.error('read failed:', err);
      setStatus(
        'Ошибка при чтении файла: ' + (err && err.message ? err.message : String(err)),
        'error'
      );
    }
  });

  console.log('change listener attached');
}

if (runBtn) {
  runBtn.addEventListener('click', () => {
    console.log('runBtn clicked');
    setStatus('Кнопка нажалась. Значит UI живой, а проблема не в клике.', 'ok');
  });
}
