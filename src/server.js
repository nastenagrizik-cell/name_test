import express from 'express';
import multer from 'multer';
import { parseSurvey } from './parser.js';
import { analyzeSurvey } from './analyzer.js';
import { buildWorkbook } from './excel-builder.js';
import { slugify } from './utils.js';

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });
const port = process.env.PORT || 3000;

app.use(express.static('public'));
app.use(express.json());

app.get('/api/health', (_, res) => {
  res.json({ ok: true });
});

app.post('/api/inspect', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Файл не загружен.' });
    const parsed = parseSurvey(req.file.buffer);
    res.json({
      fileName: req.file.originalname,
      rows: parsed.rowCount,
      sheets: parsed.sheetNames,
      questionBlocks: parsed.questionBlocks.slice(0, 25),
      audienceColumns: parsed.audienceColumns.map(c => c.header)
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/generate', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: 'Файл не загружен.' });
    const parsed = parseSurvey(req.file.buffer);
    const report = analyzeSurvey(parsed);
    const fileName = slugify(req.body.projectName || req.file.originalname.replace(/\.[^.]+$/, ''));
    const buffer = await buildWorkbook(report, fileName);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}.xlsx"`);
    res.send(Buffer.from(buffer));
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.listen(port, () => {
  console.log(`Topline Generator listening on port ${port}`);
});
