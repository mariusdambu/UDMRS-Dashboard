# UDMRS-Hydra-TwoHeads-Stable-v1.0-20260618

Fecha: 2026-06-18

Build: `UDMRS Build 2026.06.18-H2H-v1.0`

Estado: stable release candidate funcional para el ciclo actual de UDMRS Dashboard.

## Capacidades principales

- Primera entrada: organizar bibliotecas normales hacia `Year\Quarter`.
- Segunda entrada: importar galerías de proveedor mediante `GoogleTakeout`, `ApplePhotos` y `XmpSidecarLibrary`.
- Apple Photos / iCloud soporta exportaciones multiparte, `Photo Details.csv`, CSV de álbumes, papelera, vídeos y candidatos Live Photo.
- Google Photos / Takeout usa sidecars JSON, álbumes, papelera, confianza de metadata y deduplicación exacta por SHA256.
- XMP / Sidecar Library interpreta sidecars `.xmp`, `.json`, `.yaml` y `.yml` de forma conservadora, con fallback clásico.
- CaptureDateMaterialization materializa fechas fiables en metadata visible, fechas de sistema e índice cuando es seguro.
- MetadataAudit y MetadataRepair quedan disponibles desde mantenimiento avanzado.
- Dashboard ajusta tamaño inicial a pantallas pequeñas y mantiene compatibilidad con Windows PowerShell 5.1 y PowerShell 7.

## Límites conocidos

- Samsung Gallery e Immich permanecen planificados/deshabilitados hasta disponer de muestras reales suficientes.
- PNG/GIF y ciertos formatos con metadata no estándar requieren auditorías futuras específicas antes de ampliar reglas de escritura.
- Los conflictos de provider frente a metadata embebida se envían a revisión en lugar de resolverse automáticamente.
- Las galerías cloud-only siguen procesándose de forma conservadora: placeholders se omiten para evitar hidratación automática.

## Validación de cierre

- No se ejecutó Apply sobre galerías reales durante el cierre.
- Sintaxis PS5.1/PS7 validada en los scripts principales.
- Revisión documental actualizada para Importar galería, MetadataAudit, MetadataRepair y CaptureDateMaterialization.
- La copia GitHub debe sincronizarse desde la instalación validada antes del commit.

## Mantenimiento posterior - 2026-06-20

- La cuarentena de duplicados confirmados pasa a ser una red de seguridad temporal administrada por `Limpieza técnica` / `RetentionCleanup`.
- DedupeCleanup y Recovery guardan manifiestos por ejecución con la relación entre copia en cuarentena, canónico y SHA256 confirmado.
- Tras 45 días, `RetentionCleanup` recalcula los SHA256 actuales y solo elimina si existe un único canónico local accesible y ambos contenidos siguen siendo idénticos.
- Se conservan siempre la última ejecución correcta, ejecuciones con errores, RAW/DNG, cloud-only, conflictos, hashes distintos y cuarentenas históricas sin manifiesto.
- No se añadieron tareas al arranque ni procesos en segundo plano; el coste solo existe cuando el usuario inicia explícitamente la limpieza técnica.


