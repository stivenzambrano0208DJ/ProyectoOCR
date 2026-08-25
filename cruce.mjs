import { procesarCruce, cerrarPool } from './core.mjs';

function parsearArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i++) {
    if (argv[i].startsWith('--')) {
      const clave = argv[i].slice(2);
      const valor = argv[i + 1] && !argv[i + 1].startsWith('--') ? argv[++i] : true;
      args[clave] = valor;
    }
  }
  return args;
}

function imprimirAyuda() {
  console.log(`
Uso:
  node cruce.mjs --pdf "documento.pdf" --xlsx "base.xlsx" --out "informe_cruce.xlsx" [--paginas "1-10"]

Opciones:
  --pdf       Ruta al PDF escaneado (obligatorio)
  --xlsx      Ruta a la base de referencia .xls/.xlsx (obligatorio)
  --out       Ruta del informe de salida .xlsx (obligatorio)
  --paginas   Rango de páginas a procesar, ej. "1-10,15,20-25" (opcional, por defecto todas)
`);
}

const args = parsearArgs(process.argv.slice(2));

if (!args.pdf || !args.xlsx || !args.out) {
  imprimirAyuda();
  process.exit(1);
}

const inicio = Date.now();

const resultado = await procesarCruce({
  pdfPath: args.pdf,
  xlsxPath: args.xlsx,
  outPath: args.out,
  paginas: typeof args.paginas === 'string' ? args.paginas : undefined,
  onProgress: ({ paginaActual, indice, totalProcesar }) => {
    process.stdout.write(`\rProcesando página ${paginaActual} (${indice}/${totalProcesar})...`.padEnd(60));
  },
});

const segundos = ((Date.now() - inicio) / 1000).toFixed(1);
console.log(`\n\nListo en ${segundos}s. Páginas PDF: ${resultado.totalPaginasPDF} (procesadas: ${resultado.paginasProcesadas}). Registros BD: ${resultado.totalRegistrosBD}.\n`);
console.log('Resumen:');
console.log(`  Coincidencias:    ${resultado.resumen.coincidencias}`);
console.log(`  Revisar nombre:   ${resultado.resumen.revisarNombre}`);
console.log(`  Solo en el PDF:   ${resultado.resumen.soloPdf}`);
console.log(`  Solo en la BD:    ${resultado.resumen.soloBd}`);
console.log(`  Duplicados:       ${resultado.resumen.duplicados}`);
console.log(`\nInforme escrito en: ${args.out}`);

await cerrarPool();
