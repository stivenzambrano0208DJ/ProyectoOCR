# ProyectoOCR — Cruce de documentación vs. base de inscritos

Aplicación web local que sube un **PDF escaneado** de documentos de identidad y una
**base de referencia en Excel** (.xlsx/.xls), extrae el número de identificación de
cada página con **OCR (Tesseract)**, lo cruza contra la base y clasifica los resultados.

Todo el procesamiento ocurre **en la propia máquina** (no se envía nada a servicios externos).

## Características

- OCR en español página por página (doble pasada: imagen cruda + binarizada).
- Extracción del documento en capas: código de barras, etiqueta (NÚMERO/NUIP), MRZ y dígitos sueltos.
- Cruce contra la base y clasificación en 5 categorías: **Coincidencias**, **Revisar nombre**,
  **Solo en el PDF**, **Solo en la BD** y **Duplicados**.
- Procesamiento en paralelo con un **pool de workers** reutilizable (precalentado al arrancar).
- Interfaz web con progreso en vivo (tiempo transcurrido y ETA) y **vista de comparación**
  desplegable: cada cédula junto a su cajita de datos.
- Informe descargable en **.xlsx** con una hoja por categoría.

## Requisitos

- Node.js 18 o superior (recomendado 20 LTS).

## Instalación

```bash
npm install
```

## Uso (servidor web)

```bash
npm start
```

Luego abre en el navegador: **http://127.0.0.1:3578**

Sube el PDF y el Excel, opcionalmente indica un rango de páginas (ej. `1-10, 15`) y procesa.

## Uso por línea de comandos

```bash
node cruce.mjs --pdf "documento.pdf" --xlsx "base.xlsx" --out "informe.xlsx" [--paginas "1-10"]
```

## Tecnologías

- [Express](https://expressjs.com/) · [Multer](https://github.com/expressjs/multer)
- [mupdf](https://www.npmjs.com/package/mupdf) (renderizado de PDF)
- [tesseract.js](https://tesseract.projectnaptha.com/) (OCR)
- [SheetJS/xlsx](https://sheetjs.com/) (lectura/escritura de Excel)

## Despliegue

Para publicarlo en un VPS con dominio y HTTPS, consulta
[`deploy/DESPLIEGUE.md`](deploy/DESPLIEGUE.md).

## Nota de privacidad

Este proyecto procesa **datos personales (cédulas)**. Los archivos de prueba, subidas,
salidas e informes están excluidos del repositorio en [`.gitignore`](.gitignore).
No subas documentos reales al control de versiones.
