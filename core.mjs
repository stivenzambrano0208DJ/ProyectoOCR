import * as mupdf from 'mupdf';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createWorker } from 'tesseract.js';
import XLSX from 'xlsx';

const UMBRAL_BINARIZACION = 150;
// Resolución del render para OCR. Más alto = lee mejor la letra pequeña (RH,
// estatura), pero usa más memoria. Configurable con OCR_ESCALA (por defecto 4).
const OCR_ESCALA = Math.min(6, Math.max(2, Number(process.env.OCR_ESCALA) || 4));
const UMBRAL_SIMILITUD_NOMBRE = 0.5;
const STOPWORDS_NOMBRE = new Set(['DE', 'DEL', 'LA', 'LAS', 'LOS', 'Y']);

// ---------------------------------------------------------------------------
// Normalización
// ---------------------------------------------------------------------------

function normalizarDocumento(raw) {
  const soloDigitos = String(raw ?? '').replace(/\D/g, '');
  if (!soloDigitos) return null;
  const sinCeros = soloDigitos.replace(/^0+/, '');
  return sinCeros || '0';
}

function normalizarTexto(s) {
  return String(s ?? '')
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toUpperCase();
}

function tokenizarNombre(nombre) {
  return normalizarTexto(nombre)
    .split(/[^A-Z0-9Ñ]+/)
    .filter((t) => t.length > 1 && !STOPWORDS_NOMBRE.has(t));
}

// Distancia de edición (Levenshtein) acotada: cuántos cambios de letra hay
// entre dos palabras. Se usa para tolerar errores de lectura del OCR.
function distancia(a, b) {
  if (Math.abs(a.length - b.length) > 3) return 99; // muy distintas
  let prev = Array.from({ length: b.length + 1 }, (_, i) => i);
  for (let i = 1; i <= a.length; i++) {
    const cur = [i];
    for (let j = 1; j <= b.length; j++) {
      const costo = a[i - 1] === b[j - 1] ? 0 : 1;
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + costo);
    }
    prev = cur;
  }
  return prev[b.length];
}

function similitudNombre(nombreBD, textoOCR) {
  const tokens = tokenizarNombre(nombreBD);
  if (tokens.length === 0) return 0;
  const textoNorm = normalizarTexto(textoOCR);
  const textoSinEspacios = textoNorm.replace(/\s+/g, '');
  const palabrasOCR = textoNorm.split(/[^A-Z0-9Ñ]+/).filter((w) => w.length > 1);

  let coincididos = 0;
  for (const tok of tokens) {
    // 1) Coincidencia exacta (rápida): palabra completa o subcadena sin espacios.
    if (new RegExp(`\\b${tok}\\b`).test(textoNorm) || textoSinEspacios.includes(tok)) { coincididos++; continue; }
    // 2) Coincidencia difusa: alguna palabra del OCR se parece, tolerando errores
    //    de lectura (aprox. 1 error de letra por cada 4 letras del nombre).
    const tolerancia = Math.max(1, Math.floor(tok.length / 4));
    let mejor = 99;
    for (const w of palabrasOCR) {
      const d = distancia(tok, w);
      if (d < mejor) mejor = d;
      if (mejor === 0) break;
    }
    if (mejor <= tolerancia) coincididos++;
  }
  return coincididos / tokens.length;
}

function extracto(texto, maxLen = 500) {
  return String(texto ?? '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, maxLen);
}

// ---------------------------------------------------------------------------
// Extracción de datos de la persona desde el texto OCR de la cédula
// ---------------------------------------------------------------------------
// Best-effort: depende de la calidad del OCR. Los campos que no se logran leer
// quedan en null y en la interfaz se muestran como "—".

const MESES = 'ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|SET|OCT|NOV|DIC';

function ymdALegible(ymd) {
  // "20120628" -> "28-06-2012" (valida mes/día). Devuelve null si no es fecha real.
  if (!/^\d{8}$/.test(ymd)) return null;
  const a = ymd.slice(0, 4), me = +ymd.slice(4, 6), d = +ymd.slice(6, 8);
  if (+a < 1900 || +a > 2100 || me < 1 || me > 12 || d < 1 || d > 31) return null;
  return `${String(d).padStart(2, '0')}-${String(me).padStart(2, '0')}-${a}`;
}

function extraerDatosCedula(textoOCR) {
  const T = normalizarTexto(textoOCR).replace(/\s+/g, ' ');
  const datos = { sexo: null, nacimiento: null, lugarNacimiento: null, expedicion: null, lugarExpedicion: null, estatura: null, rh: null };

  // === 1) Código de barras impreso: el dato más fiable de la cédula ===
  //     Formato típico: ...-M-1117489876-20120628...  (sexo - documento - AAAAMMDD)
  const bc = T.match(/-\s*([MF])\s*-\s*[\d ]{6,14}-\s*(\d{6,10})/);
  if (bc) {
    datos.sexo = bc[1] === 'M' ? 'Masculino' : 'Femenino';
    const fx = ymdALegible(bc[2].slice(0, 8));
    if (fx) datos.expedicion = fx; // fecha de expedición codificada en el barcode
  }

  // === 2) Fechas escritas (DD-MMM-AAAA o DD/MM/AAAA) por proximidad a etiquetas ===
  const reFecha = new RegExp(`\\b(\\d{1,2})\\s*[-\\/. ]\\s*(?:(${MESES})[A-Z]*|(\\d{1,2}))\\s*[-\\/. ]\\s*(\\d{4})\\b`, 'g');
  const fechas = [];
  let m;
  while ((m = reFecha.exec(T))) {
    fechas.push({ idx: m.index, texto: `${m[1].padStart(2, '0')}-${m[2] || m[3].padStart(2, '0')}-${m[4]}` });
  }
  const idxNac = T.search(/NACIMIENTO/);
  const idxExp = T.search(/EXPEDIC/);
  const primeraTras = (idx) => (idx < 0 ? null : (fechas.find((f) => f.idx >= idx) || null));
  if (!datos.nacimiento) datos.nacimiento = (primeraTras(idxNac) || fechas[0] || {}).texto || null;
  if (!datos.expedicion) datos.expedicion = (primeraTras(idxExp) || (fechas.length > 1 ? fechas[fechas.length - 1] : null) || {}).texto || null;

  // === 3) Sexo escrito (si el barcode no lo dio) ===
  if (!datos.sexo) {
    if (/\bMASCULINO\b/.test(T)) datos.sexo = 'Masculino';
    else if (/\bFEMENINO\b/.test(T)) datos.sexo = 'Femenino';
    else {
      const s = T.match(/SEXO\s*[:.\-]?\s*([MF])\b/) || T.match(/\b([MF])\s+SEXO\b/);
      if (s) datos.sexo = s[1] === 'M' ? 'Masculino' : 'Femenino';
    }
  }

  // === 4) Estatura (ej. 1.64). Si no viene con punto, se busca cerca de la
  //        etiqueta ESTATURA un "1XX" (la cédula la imprime como 1.XX). ===
  const est = T.match(/\b1[.,]\d{2}\b/);
  if (est) datos.estatura = est[0].replace(',', '.') + ' m';
  if (!datos.estatura) {
    const iEst = T.search(/ESTATURA/);
    if (iEst >= 0) {
      const zona = T.slice(Math.max(0, iEst - 30), iEst + 30);
      const e2 = zona.match(/1[.,]?([0-9]\d)\b|1[.,]?([0-9]\d)/);
      const dd = e2 && (e2[1] || e2[2]);
      if (dd && +dd >= 40 && +dd <= 99) datos.estatura = '1.' + dd + ' m';
    }
  }

  // === 5) Grupo sanguíneo / RH (ej. O+, A-, AB+); evita confundir con el barcode "A-4400..." ===
  const rh = T.match(/\b(AB|A|B|O)\s*(POSITIVO|NEGATIVO|RH\s*[+-]|[+-])(?!\s*\d{3})/);
  if (rh) datos.rh = rh[1] + (/\+|POS/.test(rh[2]) ? '+' : '-');

  // === 6) Lugar de nacimiento ===
  // 6a) Patrón limpio "CIUDAD (DEPARTAMENTO)".
  const lug = T.match(/\b([A-ZÑ]{4,25})\s*\(\s*([A-ZÑ]{3,25})\s*\)/);
  if (lug) datos.lugarNacimiento = `${lug[1].trim()} (${lug[2].trim()})`;
  // 6b) Si no, la ciudad escrita tras la fecha de nacimiento (ignorando ruido corto).
  if (!datos.lugarNacimiento && datos.nacimiento) {
    const iFecha = T.indexOf(datos.nacimiento);
    if (iFecha >= 0) {
      const resto = T.slice(iFecha + datos.nacimiento.length, iFecha + datos.nacimiento.length + 60);
      const ETIQUETAS = /^(FECHA|LUGAR|SEXO|ESTATURA|EXPEDIC|NACIMIENTO|REPUBLICA|COLOMBIA|CEDULA|NUMERO|IDENTIF|PERSONAL|CIUDADANIA|INDICE|DERECH|IZQUIERD|HUELLA|FIRMA|PREPARAC|REGISTR|GRUPO|SANGUIN|DACTILAR|APELLIDOS|NOMBRES)/;
      const palabras = (resto.match(/[A-ZÑ]{4,}/g) || []).filter((w) => !/^(.)\1+$/.test(w) && !ETIQUETAS.test(w));
      if (palabras.length) datos.lugarNacimiento = palabras.slice(0, 2).join(' ');
    }
  }

  return datos;
}

// ---------------------------------------------------------------------------
// Lectura de la base de datos de referencia (xlsx/xls)
// ---------------------------------------------------------------------------

export function leerBD(xlsxPath) {
  const wb = XLSX.readFile(xlsxPath);
  const registros = [];

  for (const nombreHoja of wb.SheetNames) {
    const ws = wb.Sheets[nombreHoja];
    const filas = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });

    let headerIdx = -1;
    let colDoc = -1;
    let colNombre = -1;
    let colEstado = -1;

    for (let i = 0; i < filas.length; i++) {
      const fila = filas[i].map((c) => String(c).trim().toLowerCase());
      const idxDoc = fila.findIndex((c) => c.includes('identificaci'));
      const idxNombre = fila.findIndex((c) => c === 'nombre' || c.includes('nombre'));
      if (idxDoc !== -1 && idxNombre !== -1) {
        headerIdx = i;
        colDoc = idxDoc;
        colNombre = idxNombre;
        colEstado = fila.findIndex((c) => c.includes('estado'));
        break;
      }
    }

    if (headerIdx === -1) continue;

    for (let i = headerIdx + 1; i < filas.length; i++) {
      const fila = filas[i];
      const rawDoc = String(fila[colDoc] ?? '').trim();
      const nombre = String(fila[colNombre] ?? '').trim();
      if (!rawDoc || !nombre) continue;
      const documento = normalizarDocumento(rawDoc);
      if (!documento) continue;
      registros.push({
        documento,
        nombre,
        estado: colEstado !== -1 ? String(fila[colEstado] ?? '').trim() : '',
        hoja: nombreHoja,
      });
    }
  }

  return registros;
}

function construirIndiceBD(registrosBD) {
  const bdPorDocumento = new Map();
  for (const r of registrosBD) {
    if (bdPorDocumento.has(r.documento)) {
      const existente = bdPorDocumento.get(r.documento);
      if (!existente.hoja.includes(r.hoja)) existente.hoja += `, ${r.hoja}`;
    } else {
      bdPorDocumento.set(r.documento, { ...r });
    }
  }
  return bdPorDocumento;
}

// ---------------------------------------------------------------------------
// Extracción del número de documento desde el texto OCR (fallback en capas)
// ---------------------------------------------------------------------------

function candidatosBarcode(texto) {
  const candidatos = [];
  const re = /-([MF])-\s*([\d\s]{6,14})-\s*\d{6,8}/g;
  let m;
  while ((m = re.exec(texto))) {
    const doc = normalizarDocumento(m[2]);
    if (doc) candidatos.push(doc);
  }
  return candidatos;
}

function candidatosEtiqueta(texto) {
  const candidatos = [];
  // Cédula antigua: "NUMERO"/"NÚMERO". Cédula digital nueva: "NUIP".
  const re = /(?:N[UÚ]MERO|NUIP)\s*[:\-]?\s*([\d][\d.\s]{4,20}\d)/gi;
  let m;
  while ((m = re.exec(texto))) {
    const doc = normalizarDocumento(m[1]);
    if (doc) candidatos.push(doc);
  }
  return candidatos;
}

function candidatosMRZ(texto) {
  const candidatos = [];
  const re = /(\d{6,10})</g;
  let m;
  while ((m = re.exec(texto))) {
    const doc = normalizarDocumento(m[1]);
    if (doc) candidatos.push(doc);
  }
  return candidatos;
}

function candidatosDigitosSueltos(texto) {
  const candidatos = [];
  // Último recurso: cualquier corrida de dígitos (con o sin puntos de miles,
  // ej. "1.117.489.876") de 6 a 10 dígitos. Se acepta como válida más abajo
  // solo si cruza de forma única contra la BD, así que el riesgo de falso
  // positivo es mínimo aunque el patrón sea amplio.
  const patrones = [/\d{6,10}/g, /\d[\d.]{4,14}\d/g];
  for (const re of patrones) {
    let m;
    while ((m = re.exec(texto))) {
      const doc = normalizarDocumento(m[0]);
      if (doc) candidatos.push(doc);
    }
  }
  return candidatos;
}

function unicoValidoEnBD(candidatos, docsValidos) {
  const validos = [...new Set(candidatos.filter((c) => docsValidos.has(c)))];
  return validos.length === 1 ? validos[0] : null;
}

function extraerDocumento(textoCombinado, docsValidos) {
  const capas = [
    { nombre: 'barcode', candidatos: candidatosBarcode(textoCombinado) },
    { nombre: 'etiqueta', candidatos: candidatosEtiqueta(textoCombinado) },
    { nombre: 'mrz', candidatos: candidatosMRZ(textoCombinado) },
    { nombre: 'digitos-sueltos', candidatos: candidatosDigitosSueltos(textoCombinado) },
  ];

  // 1) Preferir, en orden de capa, un candidato que cruce de forma única contra la BD.
  for (const capa of capas) {
    const doc = unicoValidoEnBD(capa.candidatos, docsValidos);
    if (doc) return { documento: doc, capa: capa.nombre, enBD: true };
  }

  // 2) Si ninguna capa cruzó contra la BD, aceptar un candidato inequívoco de las
  //    capas estructuradas (barcode/etiqueta) aunque no exista en la BD, para
  //    poder reportarlo como "documento leído pero no encontrado en la BD".
  for (const capa of capas.slice(0, 2)) {
    const unicos = [...new Set(capa.candidatos)];
    if (unicos.length === 1) return { documento: unicos[0], capa: capa.nombre, enBD: false };
  }

  // 3) No se pudo identificar un documento fiable. Aun así, para mostrarlo en el
  //    informe ("Solo en el PDF"), guardamos el candidato más largo leído: suele
  //    ser el número de cédula completo, aunque no esté en la base.
  let tentativo = null;
  for (const capa of capas) {
    for (const c of capa.candidatos) {
      if (!tentativo || c.length > tentativo.length) tentativo = c;
    }
  }

  return { documento: null, capa: null, enBD: false, tentativo };
}

// ---------------------------------------------------------------------------
// Utilidades de páginas / rangos
// ---------------------------------------------------------------------------

export function parsearRangoPaginas(spec, totalPaginas) {
  if (!spec) {
    return Array.from({ length: totalPaginas }, (_, i) => i + 1);
  }
  const paginas = new Set();
  for (const parte of String(spec).split(',')) {
    const trozo = parte.trim();
    if (!trozo) continue;
    const rango = trozo.match(/^(\d+)\s*-\s*(\d+)$/);
    if (rango) {
      const desde = Math.max(1, parseInt(rango[1], 10));
      const hasta = Math.min(totalPaginas, parseInt(rango[2], 10));
      for (let p = desde; p <= hasta; p++) paginas.add(p);
    } else {
      const n = parseInt(trozo, 10);
      if (n >= 1 && n <= totalPaginas) paginas.add(n);
    }
  }
  return [...paginas].sort((a, b) => a - b);
}

// ---------------------------------------------------------------------------
// Pool de workers de Tesseract
// ---------------------------------------------------------------------------
// Se crea una sola vez y se reutiliza entre trabajos (crear/destruir un worker
// es costoso). El tamaño se ajusta a los núcleos del CPU para procesar varias
// páginas en paralelo. Un semáforo garantiza que cada worker atienda una sola
// página a la vez, incluso si llegan varios trabajos simultáneos.

const TAM_POOL = Math.max(1, Math.min(6, (os.cpus()?.length || 2) - 1));

const workersLibres = [];
const enEspera = [];
let poolInicializado = null;

async function inicializarPool() {
  if (!poolInicializado) {
    poolInicializado = (async () => {
      const workers = await Promise.all(
        Array.from({ length: TAM_POOL }, () => createWorker('spa'))
      );
      workersLibres.push(...workers);
    })();
  }
  return poolInicializado;
}

async function adquirirWorker() {
  await inicializarPool();
  if (workersLibres.length > 0) return workersLibres.pop();
  return new Promise((resolve) => enEspera.push(resolve));
}

// Precalienta el pool (crea los workers y carga el modelo) por adelantado, para
// que el primer trabajo no pague ese costo de arranque. Se llama al iniciar el
// servidor; es seguro llamarla varias veces (inicializarPool es idempotente).
export async function precalentarPool() {
  await inicializarPool();
}

function liberarWorker(worker) {
  const siguiente = enEspera.shift();
  if (siguiente) siguiente(worker);
  else workersLibres.push(worker);
}

// Cierra todos los workers del pool. Útil para procesos de un solo uso (CLI),
// donde hay que liberar el event loop para que el proceso termine. El servidor
// no lo llama: mantiene el pool vivo para reutilizarlo entre trabajos.
export async function cerrarPool() {
  if (!poolInicializado) return;
  await poolInicializado;
  const workers = workersLibres.splice(0, workersLibres.length);
  poolInicializado = null;
  await Promise.all(workers.map((w) => w.terminate()));
}

// ---------------------------------------------------------------------------
// Renderizado de una página a PNG (crudo + binarizado)
// ---------------------------------------------------------------------------
// El renderizado es sincrónico y barato comparado con el OCR; generamos ambas
// versiones de una vez, pero el reconocimiento de la binarizada se hace solo
// cuando hace falta (ver analizarPagina).

function renderizarPagina(doc, numeroPagina) {
  const page = doc.loadPage(numeroPagina - 1);
  const matrix = mupdf.Matrix.scale(OCR_ESCALA, OCR_ESCALA);
  const pixmap = page.toPixmap(matrix, mupdf.ColorSpace.DeviceGray, false, true);

  const rawPng = Buffer.from(pixmap.asPNG());

  const pixels = pixmap.getPixels();
  const numComponentes = pixmap.getNumberOfComponents();
  for (let i = 0; i < pixels.length; i += numComponentes) {
    pixels[i] = pixels[i] < UMBRAL_BINARIZACION ? 0 : 255;
  }
  const thrPng = Buffer.from(pixmap.asPNG());

  return { rawPng, thrPng };
}

async function ocrTexto(worker, png) {
  const { data } = await worker.recognize(png);
  return data.text;
}

// Renderiza una vista previa en color (escala moderada) para mostrar la cédula
// en la interfaz de revisión. Se guarda como JPEG (calidad 72): ~8x más liviano
// que PNG con la misma resolución, así carga mucho más rápido.
function renderizarPreview(doc, numeroPagina) {
  const page = doc.loadPage(numeroPagina - 1);
  const matrix = mupdf.Matrix.scale(1.5, 1.5);
  const pixmap = page.toPixmap(matrix, mupdf.ColorSpace.DeviceRGB, false, true);
  return Buffer.from(pixmap.asJPEG(72));
}

// ---------------------------------------------------------------------------
// Escritura del informe de salida
// ---------------------------------------------------------------------------

const CAT_LABEL = {
  coincidencias: 'Coincidencia',
  revisarNombre: 'Revisar nombre',
  soloPdf: 'Solo en el PDF',
  duplicados: 'Duplicado',
};

// Escribe el informe. `revisiones` es un mapa opcional { [frente]: {documento,
// nombre, estado, notas, revisado} } con las correcciones hechas en la revisión;
// si viene, sus valores tienen prioridad sobre lo detectado automáticamente.
function escribirInforme(outPath, detalle, revisiones = {}) {
  const wb = XLSX.utils.book_new();
  const personas = detalle.personas || [];
  const ov = (f) => (revisiones && revisiones[f]) || {};
  const usa = (o, r) => (o !== undefined && o !== null && o !== '' ? o : r);

  // Hoja maestra "Revisión": todas las personas con correcciones aplicadas.
  const filas = personas.map((p) => {
    const o = ov(p.frente);
    const d = { ...(p.datos || {}), ...(o.datos || {}) }; // datos OCR + correcciones
    return {
      Frente: p.frente,
      Reverso: p.reverso ?? '',
      Documento: usa(o.documento, p.documento || p.documentoDetectado || ''),
      Nombre: usa(o.nombre, p.nombreBD || ''),
      Estado: usa(o.estado, p.estado || ''),
      Categoría: CAT_LABEL[p.categoria] || p.categoria,
      Similitud: p.similitudPct || '',
      Sexo: d.sexo || '',
      'F. nacimiento': d.nacimiento || '',
      'Lugar nacimiento': d.lugarNacimiento || '',
      'F. expedición': d.expedicion || '',
      'Lugar expedición': d.lugarExpedicion || '',
      RH: d.rh || '',
      Estatura: d.estatura || '',
      Revisado: o.revisado ? 'Sí' : 'No',
      Notas: o.notas || '',
      'Extracto OCR': p.extractoOCR || '',
    };
  });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filas), 'Revisión');

  const hojaSoloBd = XLSX.utils.json_to_sheet(
    (detalle.soloBd || []).map((r) => ({ Documento: r.documento, Nombre: r.nombreBD, Estado: r.estado, Hoja: r.hoja }))
  );
  XLSX.utils.book_append_sheet(wb, hojaSoloBd, 'Solo en inscritos');

  const hojaDup = XLSX.utils.json_to_sheet(
    (detalle.duplicados || []).map((r) => ({ Documento: r.documento, 'Nombre (inscrito)': r.nombreBD, 'Frentes (págs)': r.paginas, Hoja: r.hoja }))
  );
  XLSX.utils.book_append_sheet(wb, hojaDup, 'Duplicados');

  XLSX.writeFile(wb, outPath);
}

// Exportada para regenerar el informe al descargar, aplicando las correcciones.
export function escribirInformeConRevisiones(outPath, detalle, revisiones) {
  escribirInforme(outPath, detalle, revisiones);
}

// ---------------------------------------------------------------------------
// Función de alto nivel
// ---------------------------------------------------------------------------

export async function procesarCruce({ pdfPath, xlsxPath, outPath, onProgress, paginas, imagenesDir, dosCaras = false }) {
  const registrosBD = leerBD(xlsxPath);
  const bdPorDocumento = construirIndiceBD(registrosBD);
  const docsValidos = new Set(bdPorDocumento.keys());

  const pdfBuf = fs.readFileSync(pdfPath);
  const doc = mupdf.Document.openDocument(pdfBuf, 'application/pdf');
  const totalPaginasPDF = doc.countPages();
  const rangoPaginas = parsearRangoPaginas(paginas, totalPaginasPDF);

  if (imagenesDir) fs.mkdirSync(imagenesDir, { recursive: true });

  // Procesa una página tomando prestado un worker del pool. Con varios workers,
  // hasta TAM_POOL páginas se analizan en paralelo; el resto espera turno.
  //
  // Optimización: la 2ª pasada de OCR (imagen binarizada) es costosa, así que
  // solo se ejecuta si la 1ª pasada (cruda) no bastó — es decir, si no cruzó un
  // documento contra la BD, o si cruzó pero el nombre no coincidió lo suficiente.
  // En páginas limpias esto ahorra la mitad del OCR.
  const analizarPagina = async (numeroPagina) => {
    const { rawPng, thrPng } = renderizarPagina(doc, numeroPagina);

    // Guarda una vista previa en color para la comparación visual (si se pidió).
    if (imagenesDir) {
      try {
        const preview = renderizarPreview(doc, numeroPagina);
        await fs.promises.writeFile(path.join(imagenesDir, `pag_${numeroPagina}.jpg`), preview);
      } catch { /* la vista previa es opcional: si falla, seguimos con el OCR */ }
    }

    const worker = await adquirirWorker();
    try {
      const textoRaw = await ocrTexto(worker, rawPng);
      let textoCombinado = textoRaw;
      let res = extraerDocumento(textoCombinado, docsValidos);

      let similitud = null;
      if (res.documento && res.enBD) {
        similitud = similitudNombre(bdPorDocumento.get(res.documento).nombre, textoCombinado);
      }

      // ¿Necesitamos la 2ª pasada? Sí si no cruzó en BD, o si el nombre quedó bajo el umbral.
      const necesitaSegunda = !(res.documento && res.enBD) || similitud < UMBRAL_SIMILITUD_NOMBRE;
      if (necesitaSegunda) {
        const textoThr = await ocrTexto(worker, thrPng);
        textoCombinado = `${textoRaw}\n${textoThr}`;
        res = extraerDocumento(textoCombinado, docsValidos);
        similitud = res.documento && res.enBD
          ? similitudNombre(bdPorDocumento.get(res.documento).nombre, textoCombinado)
          : null;
      }

      const { documento, capa, enBD } = res;
      const nombreBD = documento && enBD ? bdPorDocumento.get(documento).nombre : null;

      return { pagina: numeroPagina, documento, capa, enBD, nombreBD, similitud, tentativo: res.tentativo ?? null, textoOCR: textoCombinado };
    } finally {
      liberarWorker(worker);
    }
  };

  const paginasInfo = new Array(rangoPaginas.length);
  let completadas = 0;
  await Promise.all(
    rangoPaginas.map(async (p, idx) => {
      paginasInfo[idx] = await analizarPagina(p);
      completadas++;
      onProgress?.({ paginaActual: p, indice: completadas, totalProcesar: rangoPaginas.length });
    })
  );

  // --- Agrupación en personas (frente + reverso) y clasificación --------
  // En modo "dos caras" cada persona son 2 páginas consecutivas (frente y
  // reverso); si no, cada página es una persona. Para cada persona combinamos
  // el texto OCR de sus caras y extraemos un único documento.
  const paso = dosCaras ? 2 : 1;
  const personas = [];
  for (let i = 0; i < paginasInfo.length; i += paso) {
    const frente = paginasInfo[i];
    const reverso = dosCaras ? (paginasInfo[i + 1] ?? null) : null;
    const textos = [frente?.textoOCR, reverso?.textoOCR].filter(Boolean).join('\n');
    const res = extraerDocumento(textos, docsValidos);

    let nombreBD = null, estado = '', hoja = '', similitud = null;
    if (res.documento && res.enBD) {
      const reg = bdPorDocumento.get(res.documento);
      nombreBD = reg.nombre; estado = reg.estado; hoja = reg.hoja;
      similitud = similitudNombre(reg.nombre, textos);
    }

    personas.push({
      id: frente.pagina,
      frente: frente.pagina,
      reverso: reverso ? reverso.pagina : null,
      documento: res.documento ?? null,
      documentoDetectado: res.documento ?? res.tentativo ?? null,
      enBD: res.enBD,
      nombreBD, estado, hoja,
      similitud,
      datos: extraerDatosCedula(textos),
      extractoOCR: extracto(textos),
    });
  }

  // Duplicados: mismo documento en varias personas.
  const conteoDoc = new Map();
  for (const p of personas) if (p.documento && p.enBD) conteoDoc.set(p.documento, (conteoDoc.get(p.documento) || 0) + 1);

  for (const p of personas) {
    if (!p.documento || !p.enBD) p.categoria = 'soloPdf';
    else if (conteoDoc.get(p.documento) > 1) p.categoria = 'duplicados';
    else if ((p.similitud ?? 0) >= UMBRAL_SIMILITUD_NOMBRE) p.categoria = 'coincidencias';
    else p.categoria = 'revisarNombre';
    p.similitudPct = p.similitud != null ? `${Math.round(p.similitud * 100)}%` : null;
  }

  // Derivar listas por categoría (para tabla e informe), ahora por persona.
  const coincidencias = [];
  const revisarNombre = [];
  const soloPdf = [];
  const duplicados = [];
  const dupReportados = new Set();

  for (const p of personas) {
    const base = { frente: p.frente, reverso: p.reverso };
    if (p.categoria === 'soloPdf') {
      soloPdf.push({ ...base, documentoDetectado: p.documentoDetectado ?? '', extractoOCR: p.extractoOCR });
    } else if (p.categoria === 'duplicados') {
      if (!dupReportados.has(p.documento)) {
        dupReportados.add(p.documento);
        const frentes = personas.filter((q) => q.documento === p.documento && q.enBD).map((q) => q.frente).join(', ');
        duplicados.push({ documento: p.documento, nombreBD: p.nombreBD, paginas: frentes, hoja: p.hoja });
      }
    } else if (p.categoria === 'coincidencias') {
      coincidencias.push({ ...base, documento: p.documento, nombreBD: p.nombreBD, hoja: p.hoja });
    } else {
      revisarNombre.push({ ...base, documento: p.documento, nombreBD: p.nombreBD, similitud: p.similitudPct, extractoOCR: p.extractoOCR });
    }
  }

  const documentosEncontradosEnPDF = new Set(personas.filter((p) => p.documento && p.enBD).map((p) => p.documento));
  const soloBd = [];
  for (const [documento, registro] of bdPorDocumento) {
    if (!documentosEncontradosEnPDF.has(documento)) {
      soloBd.push({ documento, nombreBD: registro.nombre, estado: registro.estado, hoja: registro.hoja });
    }
  }

  const resumen = {
    coincidencias: coincidencias.length,
    revisarNombre: revisarNombre.length,
    soloPdf: soloPdf.length,
    soloBd: soloBd.length,
    duplicados: duplicados.length,
  };

  const detalle = { coincidencias, revisarNombre, soloPdf, soloBd, duplicados, personas };

  if (outPath) escribirInforme(outPath, detalle);

  return {
    resumen,
    detalle,
    dosCaras,
    totalPaginasPDF,
    paginasProcesadas: rangoPaginas.length,
    totalPersonas: personas.length,
    totalRegistrosBD: bdPorDocumento.size,
  };
}
