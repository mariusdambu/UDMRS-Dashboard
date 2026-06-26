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
- CaptureDateMaterialization materializa fechas fiables en metadata embebida solo cuando la ausencia está demostrada; las fechas del sistema y el índice se tratan como estados separados.
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

## Separación ProviderDate / EmbeddedDate - 2026-06-23

- `CaptureDateMaterialization` ya no interpreta metadata no leída como metadata ausente.
- Se introduce `EmbeddedCaptureDateProbe` con estados `PresentValid`, `Absent`, `Conflict`, `Unreadable`, `Unsupported` y `NotChecked`.
- Solo `Absent`, confirmado por una lectura correcta, permite escritura automática de metadata embebida. La sincronización de fechas de sistema se reporta por separado y no equivale a reescritura EXIF/QuickTime/XMP.
- Los assets `ProviderTrusted` que omiten EXIF mantienen el rendimiento actual, quedan como `NotChecked` y se preservan sin reescritura.
- No cambia todavía la política de prioridad entre Google, Apple, XMP y metadata embebida.

## Contrato de archivo sano y certificación física - 2026-06-26

- Un archivo sano se define como contenido legible con fecha de captura embebida válida, metadata embebida coherente y sin necesidad real de reparación embebida.
- La consecuencia documental queda fijada: metadata embebida sana es intocable. ProviderTrusted puede organizar y clasificar, pero no sobrescribe una fecha embebida válida.
- `CreationTime` y `LastWriteTime` son fechas del sistema de archivos y se documentan separadas de EXIF/QuickTime/XMP/PNG metadata.
- `PhysicalMetadataCertification` certifica una comprobación física reciente de metadata; no sustituye al índice ni equivale a provider fiable.
- `MetadataRepair` y `Materialize` pueden usar la certificación para evitar inspecciones innecesarias cuando sigue vigente.
- `NormalizeExistingFolders` sigue recorriendo físicamente la biblioteca, pero puede reutilizar fechas certificadas para evitar lecturas EXIF/QuickTime innecesarias.
