# UDMRS Dashboard

Documento principal del paquete público actual.

PhotoOrganizer es el motor PowerShell. UDMRS Dashboard es la interfaz visual para usarlo sin escribir comandos.

## Estado de release

Release estable: `UDMRS-Hydra-TwoHeads-Stable-v1.0-20260618`.

Build: `UDMRS Build 2026.06.18-H2H-v1.0`.

Validacion final: sintaxis PS5.1/PS7, dashboard, providers disponibles y flujos principales revisados sin Apply sobre galerias reales.

Esta release congela el flujo clasico estable y la entrada `Importar galeria` como segunda entrada productiva.

Estado actual: `Google Photos / Takeout`, `Apple Photos / iCloud`, `Movistar Cloud Export` y `XMP / Sidecar Library` estan disponibles. `Samsung Gallery` e `Immich` aparecen como providers planificados y deshabilitados hasta disponer de muestras reales.

## Branding visual

La identidad visual recomendada del proyecto queda documentada en:

```text
Branding\UDMRS-Branding.md
```
Resumen:

```text
UDMRS Dashboard
Universal Digital Memory Recovery System
Recover · Repair · Organize · Preserve
```
Esta marca es solo visual/documental. No cambia nombres tecnicos internos: el motor sigue siendo `PhotoOrganizer.ps1`, el lanzador sigue siendo `Start-PhotoOrganizer.cmd` y no se renombran logs, JSON, aliases, runtime, carpetas internas ni comandos por branding.

## Abrir

Ejecuta:

```text
Start-PhotoOrganizer.cmd
```
La herramienta es portable cuando se abre desde el `.cmd` de la carpeta copiada. Si copias UDMRS Dashboard a otra ruta, no reutilices un acceso directo `.lnk` antiguo: Windows guarda dentro del acceso directo la ruta absoluta original. Abre primero `Start-PhotoOrganizer.cmd` desde la nueva carpeta; el lanzador reparará el acceso directo local si existe.

El dashboard abre la ayuda correcta desde `Ajustes > Ayuda` según el idioma activo.

## Referencia de comandos

La referencia oficial para comandos manuales está en:

```text
Docs\CommandReference.html
```
Ese HTML contiene comandos listos para copiar con la plantilla portable:

```text
Script: <CarpetaUDMRS>\App\PhotoOrganizer.ps1
Source: %USERPROFILE%\OneDrive\Imágenes
Destination: %USERPROFILE%\OneDrive\Imágenes\Fotos_Organizadas
Language: es
```
Incluye Test Scan, Organize, Normalize, Reconcile, Purge, DedupeCleanup, RepairOnly, MetadataAudit, MetadataRepair, Recovery, traducción de carpetas internas, comandos con `-Diagnostic`, ETA/progress y búsquedas útiles en logs.

## Manuales vivos

- Español: `Docs\Manuals\Manual_ES.md`
- Română: `Docs\Manuals\Manual_RO.md`
- Arquitectura futura: `Docs\Future\UDMRS-Future-ImportProviders.md`

Roadmap de la segunda entrada: `GoogleTakeout`, `ApplePhotos` y `XmpSidecarLibrary` disponibles; `SamsungGallery` e `Immich` definidos como planned/sample-gated. Servicios que solo entregan carpetas normales de archivos multimedia siguen usando `Organize`.

`README-OrganizePhotoLibrary.md` pertenece al script antiguo `Organize-PhotoLibrary.ps1` y no describe el flujo actual.

## Flujo oficial del dashboard

```text
Inicio
↓
Elegir origen y destino
↓
Elegir modo
↓
Revisar resumen dinámico
↓
Ejecutar / Cancelar / Prueba de escaneo
↓
Progreso
↓
Logs y resumen
↓
Resultados
↓
Ajustes
```
Para uso normal no necesitas lanzar herramientas internas manualmente. El dashboard ejecuta el flujo principal y deja las herramientas avanzadas para mantenimiento, recuperación o diagnóstico.

## Importar galería

La pestaña `Importar galería` es la entrada para exportaciones de proveedores que traen más contexto que archivos multimedia sueltos: álbumes, sidecars, JSON, papelera o relaciones entre elementos.

Disponible actualmente:

- `Google Photos / Takeout`: analiza una exportación Google Takeout seleccionada por el usuario, usa sidecars JSON, álbumes, papelera, confianza de metadata y deduplicación por assets lógicos, y copia assets limpios al destino `Year\Quarter`.
- `Apple Photos / iCloud`: analiza una exportación iCloud seleccionada por el usuario, usa `Photo Details.csv`, CSV de álbumes, flags de papelera, fechas de proveedor, vídeos y candidatos Live Photo, y copia assets limpios al destino `Year\Quarter`.
- `Movistar Cloud Export`: reconoce `profile.json` y las áreas activas `Pictures`/`Videos`, omite y contabiliza `Trash`, y reutiliza EXIF, QuickTime, filename fiable, SHA256 y el índice común. El perfil solo identifica la estructura; sus datos personales no se incluyen en reportes.
- `XMP / Sidecar Library`: intenta interpretar galerías desconocidas con sidecars XMP, JSON o YAML. Si la relación media/sidecar es clara usa metadata; si es ambigua manda a revisión; si no hay metadata útil usa fallback clásico.

Planificados y deshabilitados:

- `Samsung Gallery`: sample-gated hasta disponer de exportación real.
- `Immich`: requiere muestras/export adecuados.

Las exportaciones de provider se consideran fuentes temporales. Tras un `Apply` correcto, el dashboard puede preguntar si quieres eliminar la carpeta de exportación seleccionada. Esa eliminación es opcional, requiere confirmación explícita, nunca ocurre en DryRun y queda registrada en log/reporte.

Para carpetas normales con JPG/PNG/HEIC/MP4/MOV sin semántica adicional usa `Inicio` / `Organize`.

## Opciones avanzadas / Modo avanzado

En `Ajustes` existe el botón `Modo avanzado`. Abre un panel técnico sin recargar la pantalla principal.

Ese panel expone modos del motor que antes requerían comandos PowerShell largos:

- `RetentionCleanup`: limpia backups EXIF antiguos y revalida duplicados confirmados vencidos antes de eliminarlos. No limpia la galería normal, no limpia carpetas vacías, no reorganiza fotos y no toca `_Duplicados_Para_Revisar`.
- `RecoverFromWrongDuplicateMove`: recupera archivos enviados erróneamente a duplicados usando logs e índice.
- `RenameExistingFoldersToCurrentLanguage`: traduce carpetas trimestrales existentes al idioma activo sin leer contenido.
- `RenameInternalFoldersToCurrentLanguage`: traduce contenedores internos conocidos al idioma activo sin hidratar contenido cloud-only.
- `NormalizeExistingFolders`: reestructuración visual a año/trimestre y limpieza segura.
- `DedupeCleanup`: duplicados exactos por hash; no limpia carpetas vacías.
- `RepairOnlyExistingOrganizedLibrary`: reparación EXIF in-place dentro de la biblioteca organizada.
- `MetadataAudit`: auditoría segura de fechas visibles. Genera reporte CSV; no escribe metadata aunque el dashboard esté en Apply.
- `MetadataRepair`: materializa fechas fiables en metadata embebida y fechas de sistema. Crea backup, recalcula hash y actualiza el índice.
- `Migrar UDMRS a otro PC`: crea un paquete con ZIP de instalación compartida, ZIP de estado del usuario actual y una guía de migración. No incluye logs ni runtime.

Los botones avanzados respetan `Simulación` por defecto. Si activas `Aplicar cambios reales`, el dashboard pide confirmación específica y muestra los switches que lanzará.

Cuando lanzas una acción desde `Modo avanzado`, el dashboard abre una consola técnica persistente (`pwsh` si existe, si no Windows PowerShell). Esa consola se reutiliza para acciones avanzadas posteriores, no se cierra al terminar cada acción y permite ver salida, errores y progreso en tiempo real sin bloquear el dashboard principal.

La sesión técnica pertenece al dashboard. Si cierras el dashboard, se cierra también la consola técnica controlada, sus runners y el estado temporal de `%APPDATA%\PhotoOrganizer\Runtime\TechnicalConsole`. El objetivo es no dejar ventanas PowerShell huérfanas ni estados `Running` falsos.

### OneDrive / Explorer tras operaciones estructurales

Después de `RenameExistingFoldersToCurrentLanguage`, `RenameInternalFoldersToCurrentLanguage` o `NormalizeExistingFolders -Apply`, Windows Explorer y el proveedor cloud pueden quedarse temporalmente sin refrescar, sincronizando cambios o mostrando `No responde`.

Esto es normal durante unos minutos en bibliotecas grandes o cloud-backed. No significa por sí mismo corrupción, hidratación masiva, pérdida de datos o fallo del motor. Si el log terminó con mensajes como `No content scan`, `no EXIF`, `no hash`, `no cloud traversal`, el motor no leyó fotos ni descargó deliberadamente placeholders.

Recomendación: espera unos minutos y deja que Explorer/OneDrive reindexen. Evita lanzar otra operación inmediatamente. Si Explorer sigue bloqueado demasiado tiempo, reinicia Explorer o Windows.


## Fechas visibles y CaptureDateMaterialization

UDMRS distingue tres decisiones que antes podían parecer una sola:

```text
resolver fecha fiable
↓
organizar ruta/nombre
↓
materializar fecha visible
```
`CaptureDateMaterialization` es la capa común que usa esa fecha fiable para escribir metadata visible. Antes de cualquier escritura exige un estado explícito de `EmbeddedCaptureDateProbe`: `PresentValid`, `Absent`, `Conflict`, `Unreadable`, `Unsupported` o `NotChecked`. Solo `Absent`, confirmado por una lectura correcta, autoriza escritura embebida o sincronización de fechas de sistema.

En ImportProvider, una lectura EXIF omitida por confianza del provider queda como `NotChecked`, no como metadata ausente. El provider puede seguir decidiendo fecha, clasificación y destino, pero la copia permanece sin reescritura automática hasta que una lectura confirme `Absent`. Los estados se registran en reportes e índice. Esta separación no decide todavía si Google, Apple, XMP o EXIF debe ganar un conflicto.

- `MetadataAudit`: revisa la biblioteca organizada y genera un CSV con candidatos. Siempre es auditoría; no escribe metadata ni cambia fechas de sistema.
- `MetadataRepair`: actúa solo sobre candidatos con fecha fiable y ausencia embebida confirmada. Escribe metadata, sincroniza fechas de sistema, crea backup, recalcula hash y actualiza `ProcessedFiles.json`.
- `NormalizeExistingFolders`: no sustituye a MetadataRepair. Normalize arregla estructura; MetadataRepair arregla visibilidad temporal dentro del archivo.
- `RepairExif`: sigue existiendo como opción de organización/reparación, pero la política madura de fechas visibles se centraliza en CaptureDateMaterialization.

Este flujo ayuda especialmente cuando Microsoft Photos, OneDrive o Windows muestran recuerdos antiguos en una fecha de sistema reciente porque el archivo no contiene una fecha de captura visible.

## Arquitectura RC madura

La estructura oficial única es:

```text
Año
↓
Trimestre
```
El perfil oficial es `QuarterlyFolders`. El modelo estructural anterior queda retirado y ya no se documenta como opción válida.

Responsabilidades actuales:

- `Organize`: construye o reconstruye la biblioteca, clasifica, deduplica, repara EXIF si se pide y registra el índice.
- `NormalizeExistingFolders`: reestructuración visual y limpieza. Mueve/renombra a año/trimestre, limpia carpetas vacías, junk-only y ramas residuales.
- `ReconcileProcessedDatabase`: sincroniza y repara `ProcessedFiles.json`. Es la herramienta principal para arreglar rutas tras movimientos manuales o Normalize.
- `PurgeMissingFromProcessedDatabase`: acción avanzada para purgar historial desaparecido después de revalidar.
- `DedupeCleanup`: deduplicación por hash. No limpia carpetas vacías.

## Reglas rápidas de seguridad

- Sin `-Apply`, el motor trabaja en simulación o análisis.
- `DryRun` no mueve, no copia y no modifica EXIF.
- `Aplicar cambios reales` pide confirmación inicial desde el dashboard y luego trabaja de forma automática.
- Las carpetas internas de la aplicación y las carpetas excluidas por el usuario quedan protegidas antes de escanear contenido.
- `Reparar índice` / `ReconcileProcessedDatabase` sincroniza `ProcessedFiles.json` con la galería real y no mueve fotos.
- `Purgar entradas desaparecidas` es avanzado: revalida antes y solo actúa sobre entradas realmente desaparecidas.
- `NormalizeExistingFolders -Apply` reestructura visualmente y ejecuta validación/reconcile al final; no es el mantenedor principal del índice movimiento a movimiento.

## Carpetas excluidas y protegidas

La RC distingue tres conceptos:

```text
InternalProtectedFolders
= carpetas propias de la aplicación

UserExcludedFolders
= carpetas que el usuario declara intocables

VendorManagedFolders
= presets/ayudas para carpetas típicas de apps externas
```
La configuración portable vive en:

```text
%APPDATA%\PhotoOrganizer\Config\UserExcludedFolders.json
```
El dashboard muestra un resumen visible en `Inicio` con el número de carpetas protegidas activas, encontradas y no encontradas. Desde ese panel puedes abrir directamente la gestión de carpetas excluidas. También está disponible desde `Ajustes`: añadir carpetas, quitarlas, activar/desactivar entradas, revisar etiqueta/motivo y restaurar defaults. El usuario no necesita editar JSON manualmente.

Una carpeta excluida por usuario es una zona protegida: no se organiza, no se normaliza, no se repara EXIF, no se deduplica, no se purga, no se limpia, no se calcula hash, no se lee EXIF y no se hidrata cloud-only. Si la ruta no existe, la entrada se conserva en la configuración y se muestra como no encontrada.

`Samsung Gallery` y `Camera Roll` pueden aparecer como presets sugeridos, pero no son exclusiones activas globales. Cada usuario decide qué carpetas proteger desde el dashboard; el motor no hereda exclusiones desde la instalación compartida.

## Limpieza de carpetas

No existe un parámetro oficial `-Cleanup`.

La limpieza segura de carpetas vacías, carpetas junk-only y ramas residuales se ejecuta durante:

```powershell
-NormalizeExistingFolders
```
Para permitir limpieza durante Normalize:

```powershell
-NormalizeExistingFolders `
-KeepEmptyFolders:$false
```
Importante:

```text
-KeepEmptyFolders NO limpia carpetas vacías.
-KeepEmptyFolders conserva carpetas vacías.
```
## RetentionCleanup no es limpieza general

`RetentionCleanup` no limpia la galería normal. No elimina carpetas vacías, no reestructura, no normaliza, no purga el índice y no deduplica pendientes.

Solo actúa sobre contenido temporal o confirmado:

```text
%LOCALAPPDATA%\PhotoOrganizer\_Backup_Metadate
_Carantina_Duplicate_Confirmate
```
También reconoce aliases históricos como `_CopiaSeguridadMetadatos`, `_MetadataBackup`, `_Cuarentena_Duplicados_Confirmados` o `_Confirmed_Duplicates_Quarantine`.

`_Duplicados_Para_Revisar` y sus equivalentes de idioma no se tocan jamás por `RetentionCleanup`, porque pueden contener casos que requieren revisión humana.

Por defecto la retención depende de antigüedad: backups de metadata 30 días y duplicados confirmados 45 días. Las nuevas cuarentenas confirmadas guardan un manifiesto técnico por ejecución. Cuando el usuario lanza `RetentionCleanup`, UDMRS vuelve a calcular el SHA256 de la copia en cuarentena y de su copia canónica local antes de borrar. También comprueba que no sea RAW/DNG, cloud-only, conflictiva, perteneciente a la última ejecución correcta o a una ejecución con errores.

Las cuarentenas históricas sin manifiesto se conservan: no se convierten retroactivamente en borrado automático. Si cualquier revalidación falla, el archivo permanece en cuarentena y el log explica el motivo. Si todavía no se alcanzó el umbral de días, no habrá candidatos; que no se borre nada es un resultado normal.

## Estado interno

El estado portable está en:

```text
%APPDATA%\PhotoOrganizer\
```
La configuración de carpetas excluidas es específica de cada usuario y vive en:

```text
%APPDATA%\PhotoOrganizer\Config\UserExcludedFolders.json
```
Si este archivo no existe, UDMRS lo crea automáticamente con una lista vacía de exclusiones activas y presets sugeridos. La carpeta compartida de la aplicación no es fuente activa de exclusiones de usuario.

Los backups temporales de metadatos están en:

```text
%LOCALAPPDATA%\PhotoOrganizer\_CopiaSeguridadMetadatos\
```
Consulta `Docs\Manuals\Manual_ES.md` y `Docs\CommandReference.html` para detalles completos.

## Migrar la aplicación a otro PC

UDMRS separa instalación compartida y estado de usuario. Para moverlo a otro equipo:

La forma recomendada es usar `Ajustes` -> `Modo avanzado` -> `Migrar UDMRS a otro PC`. El asistente crea en `%USERPROFILE%\Downloads\UDMRS-MigrationPackages` y abre esa carpeta al terminar:

- un ZIP de la instalación compartida actual, resuelto desde la raíz real de la aplicación
- un ZIP del estado del usuario actual, resuelto desde `%APPDATA%\PhotoOrganizer`
- `MigrationGuide.txt` con los pasos de restauración

El ZIP de instalación incluye la aplicación, documentación, herramientas, branding, plantillas, releases y lanzadores existentes. El ZIP de usuario incluye `ProcessedFiles.json`, `Config`, `IndexBackups`, ajustes del dashboard y otros JSON de estado útiles. No incluye `Logs`, `Runtime`, `*.progress.json`, colas, PID/status temporales ni backups EXIF locales.

1. Copia la carpeta completa de instalación, por ejemplo `<CarpetaUDMRS>`, al nuevo PC.
2. Abre siempre `Start-PhotoOrganizer.cmd` desde la carpeta copiada.
3. No reutilices accesos directos `.lnk` antiguos: Windows guarda rutas absolutas dentro del acceso directo.
4. Si quieres continuar el historial incremental de un usuario, copia también `%APPDATA%\PhotoOrganizer\` del usuario antiguo al mismo perfil del nuevo equipo.
5. No copies `%LOCALAPPDATA%\PhotoOrganizer\` salvo para investigar una ejecución reciente; contiene backups temporales y no es estado estable.
6. Deja que OneDrive/Dropbox/iCloud/Google Drive termine de sincronizar antes de ejecutar operaciones grandes.
7. Ejecuta `Sincronizar índice / reparar cambios manuales` si moviste carpetas manualmente, cambiaste de PC o restauraste una biblioteca parcialmente.

Los logs, progress y reportes HTML viven dentro de `%APPDATA%\PhotoOrganizer\Logs`. La acción avanzada `Limpieza técnica` puede limpiar artefactos operativos antiguos de esa carpeta con una retención aproximada de 7 días: logs, `*.progress.json` y reportes HTML. Siempre evita archivos modificados en la última hora y rutas activas detectadas.

Los backups técnicos de `ProcessedFiles.json` viven separados en `%APPDATA%\PhotoOrganizer\IndexBackups`. No son logs normales. `Limpieza técnica` conserva siempre el backup más reciente, conserva backups recientes hasta un máximo aproximado de 10 copias y purga el resto. Si existen backups antiguos en `Logs\JsonBackups`, esa acción los migra a `IndexBackups`.

La infraestructura temporal del modo avanzado vive separada en `%APPDATA%\PhotoOrganizer\Runtime\TechnicalConsole`, porque no es documentación ni log. Esa carpeta no forma parte de la limpieza de logs. Los backups temporales de metadatos viven en `%LOCALAPPDATA%\PhotoOrganizer\` y no son necesarios para migrar el estado incremental.

El `.cmd` usa `%~dp0` para arrancar desde su propia carpeta. Los scripts usan `$PSScriptRoot` para resolver recursos, idiomas, manuales y ExifTool dentro de la instalación actual; logs, runtime y configuración mutable se resuelven por usuario en `%APPDATA%\PhotoOrganizer`.

## Bibliotecas sincronizadas y una sola operación activa

La aplicación puede convivir con bibliotecas sincronizadas, pero debe existir una sola operación activa del motor a la vez.

- Puedes usar un destino dentro de OneDrive u otro sistema cloud-backed.
- UDMRS Dashboard procesa contenido local verificable y salta placeholders cloud-only.
- Para galerías grandes, marca la biblioteca como disponible sin conexión, por ejemplo `Always keep on this device`, antes de Organize/Repair/Normalize masivos.
- No ejecutes dos dashboards, dos consolas técnicas o dos comandos manuales contra la misma biblioteca al mismo tiempo.
- No lances una acción normal mientras `Modo avanzado` ejecuta otra acción avanzada.









