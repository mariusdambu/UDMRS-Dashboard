# UDMRS Single-Head Stable Release Candidate 1

Build: `UDMRS Build 2026.05.30-SH-RC1`

Fecha: `2026-05-30`

Estado: `Stable Release Candidate`

Validacion final: `16/16 PASS`, `FAIL_COUNT = 0`

## Alcance

Esta release congela la version estable de una sola entrada de UDMRS Dashboard y el motor `PhotoOrganizer.ps1`.

La futura fase de importacion de proveedores externos queda fuera de esta build.

## Capacidades principales

- Dashboard portable con ayuda, ajustes, modo avanzado y consola tecnica controlada.
- Organizacion oficial `Year\Quarter`.
- Organize DryRun / Apply.
- Organize + RepairExif limitado a los archivos del run actual.
- RepairOnlyExistingOrganizedLibrary como mantenimiento explicito de biblioteca completa.
- NormalizeExistingFolders como reestructuracion visual y limpieza estructural.
- ReconcileProcessedDatabase para sincronizar y reparar el indice.
- PurgeMissingFromProcessedDatabase como accion avanzada protegida.
- DedupeCleanup para duplicados exactos confirmados.
- RetentionCleanup para backups EXIF antiguos y duplicados confirmados antiguos.
- Carpetas internas protegidas, carpetas excluidas por usuario y presets vendor-managed.
- Comportamiento cloud-aware: procesar contenido local verificable y saltar placeholders cloud-only.
- Retencion automatica ligera de logs al arrancar el dashboard.
- Backups tecnicos del indice separados en `%APPDATA%\PhotoOrganizer\IndexBackups`.

## Limitaciones conocidas

- Import Providers / segunda entrada no esta implementado en esta release.
- Los placeholders cloud-only no se hidratan automaticamente; algunas validaciones profundas quedan pendientes hasta que el archivo exista localmente.
- DedupeCleanup trabaja sobre duplicados exactos confirmados por hash, no sobre duplicados visuales o probables.
- RetentionCleanup no limpia la galeria normal ni `_Duplicados_Para_Revisar`.
- En bibliotecas grandes se recomienda revisar DryRun antes de cualquier Apply real.

## Validacion

Prueba de estres final single-head:

```text
16/16 PASS
FAIL_COUNT = 0
```

La build queda congelada como referencia estable para evolucionar posteriormente la entrada Import Providers.


