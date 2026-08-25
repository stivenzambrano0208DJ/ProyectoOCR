import * as mupdf from 'mupdf';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createWorker } from 'tesseract.js';
import XLSX from 'xlsx';

const UMBRAL_BINARIZACION = 150;
const UMBRAL_SIMILITUD_NOMBRE = 0.6;
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

function similitudNombre(nombreBD, textoOCR) {
  const tokens = tokenizarNombre(nombreBD);
  if (tokens.length === 0) return 0;
  const textoNorm = normalizarTexto(textoOCR);
  const textoSinEspacios = textoNorm.replace(/\s+/g, '');
  let coincididos = 0;
  for (const tok of tokens) {
    const re = new RegExp(`\\b${tok}\\b`);
    if (re.test(textoNorm) || textoSinEspacios.includes(tok)) coincididos++;
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
  const matrix = mupdf.Matrix.scale(3, 3);
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
// en la interfaz de comparación. Es independiente del render de OCR.
function renderizarPreview(doc, numeroPagina) {
  const page = doc.loadPage(numeroPagina - 1);
  const matrix = mupdf.Matrix.scale(1.5, 1.5);
  const pixmap = page.toPixmap(matrix, mupdf.ColorSpace.DeviceRGB, false, true);
  return Buffer.from(pixmap.asPNG());
}

// ---------------------------------------------------------------------------
// Escritura del informe de salida
// ---------------------------------------------------------------------------

function escribirInforme(outPath, detalle) {
  const wb = XLSX.utils.book_new();

  const hojaCoincidencias = XLSX.utils.json_to_sheet(
    detalle.coincidencias.map((r) => ({
      Página: r.pagina,
      Documento: r.documento,
      'Nombre (BD)': r.nombreBD,
      Hoja: r.hoja,
    }))
  );
  XLSX.utils.book_append_sheet(wb, hojaCoincidencias, 'Coincidencias');

  const hojaRevisarNombre = XLSX.utils.json_to_sheet(
    detalle.revisarNombre.map((r) => ({
      Página: r.pagina,
      Documento: r.documento,
      'Nombre (BD)': r.nombreBD,
      Similitud: r.similitud,
      'Extracto OCR': r.extractoOCR,
    }))
  );
  XLSX.utils.book_append_sheet(wb, hojaRevisarNombre, 'Revisar nombre');

  const hojaSoloPdf = XLSX.utils.json_to_sheet(
    detalle.soloPdf.map((r) => ({
      Página: r.pagina,
      'Documento detectado': r.documentoDetectado,
      'Extracto OCR': r.extractoOCR,
    }))
  );
  XLSX.utils.book_append_sheet(wb, hojaSoloPdf, 'Solo en PDF');

  const hojaSoloBd = XLSX.utils.json_to_sheet(
    detalle.soloBd.map((r) => ({
      Documento: r.documento,
      Nombre: r.nombreBD,
      Estado: r.estado,
      Hoja: r.hoja,
    }))
  );
  XLSX.utils.book_append_sheet(wb, hojaSoloBd, 'Solo en BD');

  const hojaDuplicados = XLSX.utils.json_to_sheet(
    detalle.duplicados.map((r) => ({
      Documento: r.documento,
      'Nombre (BD)': r.nombreBD,
      Páginas: r.paginas,
      Hoja: r.hoja,
    }))
  );
  XLSX.utils.book_append_sheet(wb, hojaDuplicados, 'Duplicados');

  XLSX.writeFile(wb, outPath);
}

// ---------------------------------------------------------------------------
// Función de alto nivel
// ---------------------------------------------------------------------------

export async function procesarCruce({ pdfPath, xlsxPath, outPath, onProgress, paginas, imagenesDir }) {
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
        await fs.promises.writeFile(path.join(imagenesDir, `pag_${numeroPagina}.png`), preview);
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

  // --- Clasificación ---------------------------------------------------

  const paginasPorDocumento = new Map();
  for (const info of paginasInfo) {
    if (!info.documento || !info.enBD) continue;
    if (!paginasPorDocumento.has(info.documento)) paginasPorDocumento.set(info.documento, []);
    paginasPorDocumento.get(info.documento).push(info.pagina);
  }

  const coincidencias = [];
  const revisarNombre = [];
  const soloPdf = [];
  const duplicados = [];
  const documentosDuplicadosReportados = new Set();

  for (const info of paginasInfo) {
    if (!info.documento || !info.enBD) {
      soloPdf.push({
        pagina: info.pagina,
        documentoDetectado: info.documento ?? info.tentativo ?? '',
        extractoOCR: extracto(info.textoOCR),
      });
      continue;
    }

    const paginasConEsteDoc = paginasPorDocumento.get(info.documento);
    if (paginasConEsteDoc.length > 1) {
      if (!documentosDuplicadosReportados.has(info.documento)) {
        documentosDuplicadosReportados.add(info.documento);
        const registro = bdPorDocumento.get(info.documento);
        duplicados.push({
          documento: info.documento,
          nombreBD: registro.nombre,
          paginas: paginasConEsteDoc.join(', '),
          hoja: registro.hoja,
        });
      }
      continue;
    }

    if (info.similitud >= UMBRAL_SIMILITUD_NOMBRE) {
      const registro = bdPorDocumento.get(info.documento);
      coincidencias.push({
        pagina: info.pagina,
        documento: info.documento,
        nombreBD: info.nombreBD,
        hoja: registro.hoja,
      });
    } else {
      revisarNombre.push({
        pagina: info.pagina,
        documento: info.documento,
        nombreBD: info.nombreBD,
        similitud: `${Math.round((info.similitud ?? 0) * 100)}%`,
        extractoOCR: extracto(info.textoOCR),
      });
    }
  }

  const documentosEncontradosEnPDF = new Set(
    paginasInfo.filter((i) => i.documento && i.enBD).map((i) => i.documento)
  );
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

  const detalle = { coincidencias, revisarNombre, soloPdf, soloBd, duplicados };

  if (outPath) escribirInforme(outPath, detalle);

  return {
    resumen,
    detalle,
    totalPaginasPDF,
    paginasProcesadas: rangoPaginas.length,
    totalRegistrosBD: bdPorDocumento.size,
  };
}
