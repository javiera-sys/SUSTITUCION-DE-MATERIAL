# Sustituidor de Materiales Excel

Herramienta web local para sustituir nombres de materiales dentro de archivos Excel (.xlsx), preservando el formato original de las celdas.

## Uso

1. Abre `index.html` en tu navegador (o accede a la URL de GitHub Pages).
2. Carga uno o varios archivos `.xlsx` arrastrándolos o seleccionándolos.
3. Agrega reglas de sustitución (material a buscar → material sustituto).
4. Haz clic en **ANALIZAR ARCHIVOS** para ver la vista previa.
5. Si todo es correcto, haz clic en **APLICAR SUSTITUCIONES**.
6. Descarga los archivos modificados individualmente o en ZIP.

## Limitaciones

- **Solo .xlsx**: los archivos `.xls` antiguos deben convertirse a `.xlsx` en Excel.
- **Fórmulas**: no se recalculan automáticamente (el valor cacheado se preserva).
- **Data validations complejas**: pueden verse afectadas por limitaciones de ExcelJS.

## Tecnología

- ExcelJS 4.4.0 (lectura/escritura de .xlsx preservando formato)
- JSZip 3.10.1 (generación de ZIP)

## Privacidad

Todos los archivos se procesan **en tu navegador**. Nada se envía a ningún servidor.
