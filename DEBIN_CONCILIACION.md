# Conciliación DEBIN COELSA vs Core

Script: `scripts_debin_conciliacion.py`

## Qué hace
1. (Opcional) Descarga archivos `.txt` desde una carpeta FTP.
2. Lee movimientos de COELSA y del Core bancario.
3. Conciliación en dos etapas:
   - **Primera prioridad:** `ID DEBIN` extraído del campo concepto (o columna ID DEBIN si existe).
   - **Segunda prioridad:** combinación de `comnro + cuenta + importe`.
4. Genera un **Excel** con:
   - Resumen de importes y cantidades COELSA vs CORE.
   - Diferencias.
   - Detalle de conciliación.
   - Movimientos que faltan en CORE y movimientos que sobran en CORE.

## Dependencia
```bash
pip install openpyxl
```

## Ejecutar

### Con archivos locales
```bash
python scripts_debin_conciliacion.py \
  --core-csv "Conciliacion DEBIN 2025-08-20.csv" \
  --local-input-dir ./inputs \
  --coelsa-pattern "*.txt" \
  --output-dir ./outputs \
  --output-excel "reporte_conciliacion_debin.xlsx"
```

### Descargando desde FTP
```bash
python scripts_debin_conciliacion.py \
  --core-csv "Conciliacion DEBIN 2025-08-20.csv" \
  --local-input-dir ./inputs \
  --coelsa-pattern "*.txt" \
  --output-dir ./outputs \
  --ftp-host ftp.midominio.com \
  --ftp-user usuario \
  --ftp-pass clave \
  --ftp-dir /conciliaciones/debin
```

## Hojas del Excel
- `Resumen`
- `Detalle Conciliacion`
- `Faltan_en_CORE`
- `Sobran_en_CORE`
