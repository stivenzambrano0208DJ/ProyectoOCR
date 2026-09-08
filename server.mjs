import express from 'express';
import multer from 'multer';
import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { procesarCruce, precalentarPool, escribirInformeConRevisiones } from './core.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
// Puerto y host configurables por variables de entorno (necesario en Docker/Dokploy:
// dentro de un contenedor hay que escuchar en 0.0.0.0 para ser accesible).
const PORT = process.env.PORT || 3578;
const HOST = process.env.HOST || '0.0.0.0';

const uploadsDir = path.join(__dirname, 'uploads');
const salidasDir = path.join(__dirname, 'salidas');
const imagenesDir = path.join(__dirname, 'imagenes');
const revisionesDir = path.join(__dirname, 'revisiones');
fs.mkdirSync(uploadsDir, { recursive: true });
fs.mkdirSync(salidasDir, { recursive: true });
fs.mkdirSync(imagenesDir, { recursive: true });
fs.mkdirSync(revisionesDir, { recursive: true });

const upload = multer({ dest: uploadsDir });

const trabajos = new Map(); // jobId -> estado

// --- Persistencia de la revisión (correcciones + "revisado") en disco --------
const JOBID_RE = /^[0-9a-fA-F-]{36}$/;
const revPath = (jobId) => path.join(revisionesDir, `${jobId}.json`);
function leerRevisiones(jobId) {
  try { return JSON.parse(fs.readFileSync(revPath(jobId), 'utf8')); } catch { return {}; }
}
function guardarRevisiones(jobId, obj) {
  fs.writeFileSync(revPath(jobId), JSON.stringify(obj));
}

// --- Limpieza automática de trabajos y archivos viejos -----------------------
// Evita que el disco y la memoria crezcan sin límite. Retención configurable por
// la variable de entorno RETENCION_HORAS (por defecto 7 días).
const RETENCION_MS = (Number(process.env.RETENCION_HORAS) || 168) * 60 * 60 * 1000;

function borrarArtefactos(jobId) {
  fs.rm(path.join(imagenesDir, jobId), { recursive: true, force: true }, () => {});
  fs.rm(path.join(salidasDir, `informe_${jobId}.xlsx`), { force: true }, () => {});
  fs.rm(revPath(jobId), { force: true }, () => {});
}

function limpiarViejos() {
  const corte = Date.now() - RETENCION_MS;
  let trabajosBorrados = 0;
  let archivosBorrados = 0;

  // 1) Trabajos en memoria antiguos (y sus archivos).
  for (const [jobId, estado] of trabajos) {
    if ((estado.inicio || 0) < corte) {
      trabajos.delete(jobId);
      borrarArtefactos(jobId);
      trabajosBorrados++;
    }
  }

  // 2) Archivos huérfanos en disco (por fecha de modificación) — cubre lo que
  //    quedó de trabajos ya no presentes en memoria (p. ej. tras un reinicio).
  const barrer = (dir, recursivo) => {
    let entradas;
    try { entradas = fs.readdirSync(dir); } catch { return; }
    for (const nombre of entradas) {
      const ruta = path.join(dir, nombre);
      try {
        if (fs.statSync(ruta).mtimeMs < corte) {
          fs.rmSync(ruta, { recursive: recursivo, force: true });
          archivosBorrados++;
        }
      } catch { /* ignorar */ }
    }
  };
  barrer(imagenesDir, true);
  barrer(salidasDir, false);
  barrer(revisionesDir, false);
  barrer(uploadsDir, false); // subidas temporales huérfanas

  if (trabajosBorrados || archivosBorrados) {
    console.log(`Limpieza: ${trabajosBorrados} trabajo(s) y ${archivosBorrados} archivo(s)/carpeta(s) viejos eliminados.`);
  }
}

const app = express();
app.set('trust proxy', 1); // detrás de Traefik: para que req.secure refleje HTTPS
app.use(express.json());

// --- Login con formulario + sesión por cookie firmada -----------------------
// Se activa solo si defines APP_PASSWORD (variable de entorno en Dokploy).
// Credenciales: APP_USER (por defecto "admin") y APP_PASSWORD. La landing queda
// pública; la herramienta y las imágenes de cédulas requieren iniciar sesión.
const APP_USER = process.env.APP_USER || 'admin';
const APP_PASSWORD = process.env.APP_PASSWORD || '';
const COOKIE = 'ocr_sesion';
const SESION_MAX_MS = 12 * 60 * 60 * 1000; // 12 horas
const SESION_SECRET = process.env.SESION_SECRET
  || (APP_PASSWORD ? crypto.createHash('sha256').update('cruceocr:' + APP_PASSWORD).digest('hex') : 'dev');

const firmar = (v) => crypto.createHmac('sha256', SESION_SECRET).update(v).digest('hex');
function crearToken() {
  const payload = 'ok.' + (Date.now() + SESION_MAX_MS);
  return payload + '.' + firmar(payload);
}
function tokenValido(token) {
  if (!token) return false;
  const p = token.split('.');
  if (p.length !== 3) return false;
  const payload = p[0] + '.' + p[1];
  if (firmar(payload) !== p[2]) return false;
  return Date.now() <= Number(p[1]);
}
function leerCookie(req, nombre) {
  for (const par of (req.headers.cookie || '').split(';')) {
    const i = par.indexOf('=');
    if (i !== -1 && par.slice(0, i).trim() === nombre) return decodeURIComponent(par.slice(i + 1));
  }
  return null;
}
const estaAutenticado = (req) => !APP_PASSWORD || tokenValido(leerCookie(req, COOKIE));
function ponerCookie(req, res, valor, maxSeg) {
  const secure = req.secure ? '; Secure' : '';
  res.setHeader('Set-Cookie', `${COOKIE}=${valor}; HttpOnly; SameSite=Lax; Path=/; Max-Age=${maxSeg}${secure}`);
}

// Páginas de la herramienta: si no hay sesión, al login.
app.use(['/app', '/app.html'], (req, res, next) => estaAutenticado(req) ? next() : res.redirect('/login'));
// API e imágenes: si no hay sesión, 401.
app.use(['/procesar', '/estado', '/descargar', '/revision', '/img'], (req, res, next) => estaAutenticado(req) ? next() : res.status(401).json({ error: 'Sesión requerida.' }));

// Vista de login.
app.get('/login', (req, res) => {
  if (estaAutenticado(req)) return res.redirect('/app');
  res.sendFile(path.join(__dirname, 'public', 'login.html'));
});
app.post('/login', (req, res) => {
  if (!APP_PASSWORD) return res.json({ ok: true });
  const usuario = String(req.body?.usuario || '').trim();
  const password = String(req.body?.password || '');
  if (usuario === APP_USER && password === APP_PASSWORD) {
    ponerCookie(req, res, crearToken(), Math.floor(SESION_MAX_MS / 1000));
    return res.json({ ok: true });
  }
  res.status(401).json({ ok: false, error: 'Usuario o contraseña incorrectos.' });
});
app.get('/logout', (req, res) => { ponerCookie(req, res, '', 0); res.redirect('/'); });

app.use(express.static(path.join(__dirname, 'public')));
// Vistas previas de las páginas del PDF: /img/<jobId>/pag_<n>.jpg
// Las imágenes de un trabajo no cambian, así que se cachean fuerte en el navegador.
app.use('/img', express.static(imagenesDir, { maxAge: '7d', immutable: true }));

// Landing pública (explica qué hace la herramienta).
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// La herramienta de cruce (la app en sí).
app.get('/app', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'app.html'));
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
  const dosCaras = req.body?.dosCaras === 'true' || req.body?.dosCaras === 'on' || req.body?.dosCaras === '1';

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
    dosCaras,
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
    dosCaras,
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
      estado.totalPersonas = resultado.totalPersonas;
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

// Lee el estado de revisión guardado (correcciones + "revisado") de un trabajo.
app.get('/revision/:jobId', (req, res) => {
  const { jobId } = req.params;
  if (!JOBID_RE.test(jobId)) return res.status(400).json({ error: 'jobId inválido.' });
  res.json(leerRevisiones(jobId));
});

// Guarda/actualiza la revisión de una persona (por su id = página del frente).
app.post('/revision/:jobId', (req, res) => {
  const { jobId } = req.params;
  if (!JOBID_RE.test(jobId)) return res.status(400).json({ error: 'jobId inválido.' });
  const id = req.body?.id;
  if (id === undefined || id === null) return res.status(400).json({ error: 'Falta el id de la persona.' });

  const campos = {};
  for (const k of ['documento', 'nombre', 'estado', 'notas']) {
    if (typeof req.body?.[k] === 'string') campos[k] = req.body[k];
  }
  if (typeof req.body?.revisado === 'boolean') campos.revisado = req.body.revisado;

  // Datos de la cédula editados (sexo, fechas, lugares, RH, estatura).
  if (req.body?.datos && typeof req.body.datos === 'object') {
    const dl = {};
    for (const k of ['sexo', 'nacimiento', 'lugarNacimiento', 'expedicion', 'lugarExpedicion', 'rh', 'estatura']) {
      if (typeof req.body.datos[k] === 'string') dl[k] = req.body.datos[k];
    }
    campos.datos = dl;
  }

  const revs = leerRevisiones(jobId);
  revs[String(id)] = { ...(revs[String(id)] || {}), ...campos };
  guardarRevisiones(jobId, revs);
  res.json({ ok: true, id: String(id), revision: revs[String(id)] });
});

app.get('/descargar/:jobId', (req, res) => {
  const { jobId } = req.params;
  if (!JOBID_RE.test(jobId)) return res.status(400).json({ error: 'jobId inválido.' });
  const estado = trabajos.get(jobId);
  if (!estado || estado.status !== 'listo') {
    return res.status(404).json({ error: 'Informe no disponible.' });
  }
  // Regenera el informe aplicando las correcciones de la revisión, si las hay.
  try {
    const revisiones = leerRevisiones(jobId);
    escribirInformeConRevisiones(estado.outPath, estado.detalle, revisiones);
  } catch (err) {
    console.error('No se pudo regenerar el informe con revisiones:', err?.message || err);
  }
  res.download(estado.outPath, 'informe_cruce.xlsx');
});

app.listen(PORT, HOST, () => {
  console.log(`Servidor escuchando en http://${HOST}:${PORT}`);
  // Precalienta el pool de OCR en segundo plano para que el primer trabajo sea rápido.
  console.log('Precalentando el motor de OCR…');
  precalentarPool()
    .then(() => console.log('Motor de OCR listo.'))
    .catch((err) => console.error('No se pudo precalentar el OCR:', err?.message || err));

  // Limpieza de trabajos/archivos viejos: al arrancar y cada 6 horas.
  console.log(`Retención de trabajos: ${Math.round(RETENCION_MS / 3600000)} h.`);
  limpiarViejos();
  const tarea = setInterval(limpiarViejos, 6 * 60 * 60 * 1000);
  if (tarea.unref) tarea.unref();
});
