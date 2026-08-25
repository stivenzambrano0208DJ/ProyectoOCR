import express from 'express';
import multer from 'multer';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { procesarCruce, precalentarPool } from './core.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const PORT = 3578;
const HOST = '127.0.0.1';

const uploadsDir = path.join(__dirname, 'uploads');
const salidasDir = path.join(__dirname, 'salidas');
const imagenesDir = path.join(__dirname, 'imagenes');
fs.mkdirSync(uploadsDir, { recursive: true });
fs.mkdirSync(salidasDir, { recursive: true });
fs.mkdirSync(imagenesDir, { recursive: true });

const upload = multer({ dest: uploadsDir });

const trabajos = new Map(); // jobId -> estado

const app = express();
app.use(express.static(path.join(__dirname, 'public')));
// Vistas previas de las páginas del PDF: /img/<jobId>/pag_<n>.png
app.use('/img', express.static(imagenesDir));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/procesar', upload.fields([{ name: 'pdf', maxCount: 1 }, { name: 'xlsx', maxCount: 1 }]), (req, res) => {
  const pdfFile = req.files?.pdf?.[0];
  const xlsxFile = req.files?.xlsx?.[0];

  if (!pdfFile || !xlsxFile) {
    for (const f of [pdfFile, xlsxFile]) {
      if (f) fs.unlink(f.path, () => {});
    }
    return res.status(400).json({ error: 'Se requieren ambos archivos: pdf y xlsx.' });
  }

  const paginas = typeof req.body?.paginas === 'string' ? req.body.paginas.trim() : '';

  const jobId = crypto.randomUUID();
  const outPath = path.join(salidasDir, `informe_${jobId}.xlsx`);

  trabajos.set(jobId, {
    status: 'procesando',
    paginaActual: 0,
    totalProcesar: 0,
    indice: 0,
    inicio: Date.now(),
    fin: null,
    resumen: null,
    detalle: null,
    outPath,
    error: null,
  });

  res.json({ jobId });

  procesarCruce({
    pdfPath: pdfFile.path,
    xlsxPath: xlsxFile.path,
    outPath,
    paginas: paginas || undefined,
    imagenesDir: path.join(imagenesDir, jobId),
    onProgress: ({ paginaActual, indice, totalProcesar }) => {
      const estado = trabajos.get(jobId);
      if (!estado) return;
      estado.paginaActual = paginaActual;
      estado.indice = indice;
      estado.totalProcesar = totalProcesar;
    },
  })
    .then((resultado) => {
      const estado = trabajos.get(jobId);
      if (!estado) return;
      estado.status = 'listo';
      estado.fin = Date.now();
      estado.resumen = resultado.resumen;
      estado.detalle = resultado.detalle;
      estado.totalPaginasPDF = resultado.totalPaginasPDF;
      estado.totalRegistrosBD = resultado.totalRegistrosBD;
    })
    .catch((err) => {
      const estado = trabajos.get(jobId);
      if (!estado) return;
      estado.status = 'error';
      estado.fin = Date.now();
      estado.error = err?.message || String(err);
    })
    .finally(() => {
      for (const f of [pdfFile, xlsxFile]) {
        fs.unlink(f.path, () => {});
      }
    });
});

app.get('/estado/:jobId', (req, res) => {
  const estado = trabajos.get(req.params.jobId);
  if (!estado) return res.status(404).json({ error: 'Trabajo no encontrado.' });
  const { outPath, inicio, fin, ...publico } = estado;
  publico.transcurridoMs = (fin ?? Date.now()) - inicio;
  res.json(publico);
});

app.get('/descargar/:jobId', (req, res) => {
  const estado = trabajos.get(req.params.jobId);
  if (!estado || estado.status !== 'listo') {
    return res.status(404).json({ error: 'Informe no disponible.' });
  }
  res.download(estado.outPath, 'informe_cruce.xlsx');
});

app.listen(PORT, HOST, () => {
  console.log(`Servidor local escuchando en http://${HOST}:${PORT} (solo accesible desde esta máquina)`);
  // Precalienta el pool de OCR en segundo plano para que el primer trabajo sea rápido.
  console.log('Precalentando el motor de OCR…');
  precalentarPool()
    .then(() => console.log('Motor de OCR listo.'))
    .catch((err) => console.error('No se pudo precalentar el OCR:', err?.message || err));
});
