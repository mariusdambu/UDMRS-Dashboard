# Manual de UDMRS Dashboard

Este manual describe el paquete público actual. Está escrito para usar la herramienta sin conocer la historia interna del proyecto.

Release estable: `UDMRS-Hydra-TwoHeads-Stable-v1.0-20260618`.

Build: `UDMRS Build 2026.06.18-H2H-v1.0`.

Validacion final: sintaxis PS5.1/PS7, dashboard, providers disponibles y flujos principales revisados sin Apply sobre galerías reales.

Estado público actual: el flujo clásico estable sigue siendo la entrada principal para carpetas normales. `Importar galería` está disponible para `Google Photos / Takeout`, `Apple Photos / iCloud` y `XMP / Sidecar Library`.

## 1. Qué es cada pieza

- `PhotoOrganizer.ps1`: motor PowerShell. Hace el trabajo real.
- `PhotoOrganizerDashboard.ps1`: interfaz gráfica. Lanza el motor con parámetros y muestra progreso.
- `Start-PhotoOrganizer.cmd`: forma recomendada de abrir el dashboard.
- `ProcessedFiles.json`: índice interno incremental. Guarda hashes y rutas actuales conocidas.

El dashboard no duplica la lógica del motor.

## 2. Abrir la herramienta

Ejecuta:

```text
Start-PhotoOrganizer.cmd
```

Portabilidad: abre siempre el `.cmd` que está dentro de la carpeta actual de UDMRS Dashboard. Un acceso directo `.lnk` copiado desde otra carpeta puede seguir apuntando a la ruta antigua, porque Windows guarda rutas absolutas dentro del acceso directo. Al abrir `Start-PhotoOrganizer.cmd`, el lanzador actualiza el acceso directo local si existe.

El flujo oficial es:

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

Para uso normal trabaja desde `Inicio`.

## Referencia oficial de comandos

Los comandos manuales listos para copiar están en:

```text
Docs\CommandReference.html
```

Ese archivo usa la plantilla portable:

```text
Script: <CarpetaUDMRS>\App\PhotoOrganizer.ps1
Source: %USERPROFILE%\OneDrive\Imágenes
Destination: %USERPROFILE%\OneDrive\Imágenes\Fotos_Organizadas
Language: es
```

No documenta parámetros inexistentes como `-Cleanup`.

## 3. Inicio y resumen dinámico

En `Inicio` eliges:

- carpeta origen
- carpeta destino
- `DryRun`
- `Aplicar cambios reales`
- `Repair EXIF`
- modo de rendimiento
- workers máximos, si quieres limitar el equipo

El texto junto a `Ejecutar` cambia según la combinación seleccionada. Ese resumen es la explicación de lo que se va a lanzar.

## 4. Acciones principales

### Ejecutar

Usa la configuración actual.

- Con `DryRun`: analiza y simula. No toca archivos.
- Con `Aplicar cambios reales`: mueve o copia según configuración.
- Con `Repair EXIF`: puede reparar metadatos si también está activo `Aplicar cambios reales`.

### Cancelar

Detiene el proceso hijo de PowerShell. El motor registra la cancelación y conserva lo necesario para poder revisar logs.

### Prueba de escaneo

Es segura.

Hace:

- escanea archivos compatibles
- aplica exclusiones
- valida acceso
- muestra encontrados, omitidos e inaccesibles

No hace:

- no lee EXIF
- no genera hashes
- no mueve
- no copia
- no modifica EXIF
- no crea estructura final

Úsala antes de bibliotecas grandes para confirmar que la ruta origen es correcta.

## 4.1 Importar galería

`Importar galería` es una entrada principal para exportaciones de proveedores que contienen más contexto que archivos multimedia sueltos: álbumes, sidecars, JSON, papelera o relaciones entre elementos.

Disponible:

- `Google Photos / Takeout`: analiza una exportación Google Takeout seleccionada por el usuario, interpreta sidecars JSON, álbumes, papelera, confianza de metadata y duplicados físicos, y copia assets limpios al destino `Año\Trimestre`.
- `Apple Photos / iCloud`: analiza una exportación iCloud seleccionada por el usuario, interpreta `Photo Details.csv`, CSV de álbumes, flags de papelera, fechas de proveedor, vídeos y candidatos Live Photo, y copia assets limpios al destino `Año\Trimestre`.
- `XMP / Sidecar Library`: intenta interpretar galerías desconocidas con sidecars XMP, JSON o YAML. Si la relación media/sidecar es clara usa metadata; si es ambigua manda a revisión; si no hay metadata útil usa fallback clásico.

Planificados y deshabilitados:

- `Samsung Gallery`: bloqueado hasta disponer de exportación real.
- `Immich`: requiere muestras/export adecuados.

Para cualquier provider, trata la carpeta exportada como fuente temporal: primero ejecuta Simulación, revisa el reporte HTML/JSON y después Apply. Al terminar un Apply correcto, el dashboard puede pedir confirmación para eliminar la carpeta de exportación seleccionada. No se elimina en Simulación ni si el import falla.

Para XMP / Sidecar Library, el reporte muestra media encontrados, sidecars encontrados/usados, sidecars huérfanos, media sin sidecar, confianza High/Medium/Low, conflictos y fallback clásico.

Para carpetas normales con fotos y vídeos usa `Inicio` / `Organize`.

## 4.2 Opciones avanzadas / Modo avanzado

En `Ajustes`, el botón `Modo avanzado` abre herramientas técnicas sin mezclarlas con el flujo normal de `Inicio`.

No duplica `Ejecutar`, `Prueba de escaneo`, `Sincronizar índice` ni `Purgar entradas desaparecidas` cuando esas acciones ya tienen botón visible.

Herramientas disponibles:

- `Limpieza técnica`: limpia backups EXIF, logs/progress/reportes e IndexBackups antiguos; revalida duplicados confirmados vencidos antes de eliminarlos. No limpia la galería normal, no limpia carpetas vacías y no toca `_Duplicados_Para_Revisar`.
- `Recuperar duplicados mal movidos`: intenta recuperar archivos enviados erróneamente a duplicados usando logs e índice.
- `Traducir carpetas existentes`: adapta carpetas trimestrales existentes al idioma activo.
- `Traducir carpetas internas`: renombra o consolida contenedores internos conocidos al idioma activo sin leer fotos ni hidratar nube.
- `Normalizar estructura`: reestructura visualmente a `Año\Trimestre` y limpia ramas vacías/junk/residuales.
- `Limpieza de duplicados`: analiza duplicados exactos por hash y en Apply mueve confirmados a cuarentena.
- `Reparar EXIF in-place`: repara EXIF dentro de la biblioteca organizada sin mover, deduplicar ni renombrar.
- `Auditar fechas visibles`: detecta archivos con fecha fiable conocida pero metadata visible o fechas de sistema pendientes. Siempre es simulaci├│n.
- `Reparar fechas visibles`: escribe metadata visible y fechas de sistema cuando falta una fecha de captura ├║til. Requiere Apply y confirmaci├│n espec├¡fica.
- `Migrar UDMRS a otro PC`: crea los ZIP necesarios para mover la instalación compartida y el estado del usuario actual a otro equipo. No migra logs ni runtime.

Todas respetan `Simulación` por defecto. Si activas `Aplicar cambios reales`, el dashboard pide confirmación específica y muestra los switches que va a ejecutar.

Al lanzar una herramienta desde `Modo avanzado`, el dashboard abre una consola técnica persistente. Esa consola usa PowerShell 7 si está disponible y se reutiliza para las siguientes acciones avanzadas. No se cierra al terminar cada acción, para que puedas revisar salida, errores y progreso sin perder contexto. El dashboard principal queda usable.

La consola técnica pertenece a la sesión del dashboard. Si cierras el dashboard, se cierra también la consola técnica controlada, sus procesos hijos y el estado temporal de `%APPDATA%\PhotoOrganizer\Runtime\TechnicalConsole`. Así se evitan ventanas PowerShell huérfanas, locks falsos y estados `Running` antiguos.

## 5. DryRun y Apply

`DryRun` es el modo de simulación. No escribe cambios reales.

`Aplicar cambios reales` permite modificar la biblioteca después de una confirmación inicial. No hay confirmaciones por archivo.

`DryRun + Repair EXIF` solo informa de posibles reparaciones; no modifica metadatos.

Para una primera pasada real sobre una biblioteca grande o sin procesar se recomienda:

```text
Aplicar cambios reales + Repair EXIF
```

Las fechas EXIF reparadas mejoran la organización desde el principio.

## 6. PowerShell

El dashboard detecta automáticamente `pwsh.exe`.

- Si existe PowerShell 7, lo usa.
- Si no existe, usa Windows PowerShell 5.1.

PowerShell 7 es recomendado para bibliotecas grandes.

## 7. Organización actual

El perfil oficial único es `QuarterlyFolders`.

La estructura madura de la RC es:

```text
Año
↓
Trimestre
```

El modelo estructural anterior queda retirado. No hay dos filosofías activas de organización.

Ejemplo en español:

```text
Fotos_Organizadas
  2025
    Ene-Mar
    Abr-Jun
    Jul-Sep
    Oct-Dic
```

Dentro del trimestre se conservan nombres de archivo de forma conservadora. `NormalizeExistingFolders` no debe inventar eventos ni reescribir nombres masivamente con EXIF.

Responsabilidades:

- `Organize`: ingestión, clasificación, deduplicación, reparación EXIF opcional e índice.
- `NormalizeExistingFolders`: reestructuración visual y limpieza.
- `ReconcileProcessedDatabase`: sincronización fina de `ProcessedFiles.json`.
- `PurgeMissingFromProcessedDatabase`: limpieza avanzada de historial desaparecido.
- `DedupeCleanup`: duplicados por hash; no limpia carpetas vacías.

## 8. Carpetas internas

Con idioma español, las carpetas internas principales son:

```text
Fotos_Organizadas
_NecesitaRevision
_Duplicados_Para_Revisar
_Cuarentena_Duplicados_Confirmados
Logs
MediaMetadataIssues
```

El modo `Traducir carpetas internas` afecta solo a contenedores internos conocidos del proyecto. Puede renombrar o consolidar aliases como `Fotos_Organizadas`, `Poze_Organizate`, `Organized_Photos`, `_NecesitaRevision`, `_NeedsReview`, `_Duplicados_Para_Revisar`, `_Duplicates_To_Review`, `_CopiaSeguridadMetadatos`, `_MetadataBackup` o cuarentenas antiguas hacia el nombre oficial del idioma activo.

En bibliotecas cloud-backed, este renombrado es una operación estructural de carpeta. No calcula hash, no lee EXIF, no enumera archivos internos y no debe hidratar placeholders cloud-only. En `DryRun` muestra los contenedores que renombraría; en `Apply` renombra el contenedor conocido o mueve solo sus hijos inmediatos si el destino oficial ya existe y no hay conflictos de nombre.

`RenameExistingFoldersToCurrentLanguage` se limita a carpetas trimestrales existentes, por ejemplo `Jan-Mar` a `Ene-Mar`, y también trabaja a nivel de carpeta sin leer contenido.

Después de renames estructurales, Explorer y OneDrive pueden tardar unos minutos en refrescar o sincronizar los cambios. Puede parecer que Explorer no responde o que la vista no se actualiza, especialmente con bibliotecas grandes, carpetas online-only o varios renames seguidos. Si el log indica que no hubo content scan, EXIF, hash ni cloud traversal, ese comportamiento de Explorer/OneDrive no implica corrupción, pérdida de datos ni hidratación masiva. Espera unos minutos antes de lanzar otra operación. Si Explorer sigue bloqueado demasiado tiempo, reinicia Explorer o Windows.

Las carpetas excluidas por el usuario se saltan antes de cualquier lectura de contenido. `Samsung Gallery` y `Camera Roll` pueden usarse como presets sugeridos, pero no son exclusiones activas globales: cada usuario decide qué proteger desde el dashboard.

## 8.1 Carpetas excluidas / protegidas externas

En `Inicio` hay un panel visible con el resumen de carpetas protegidas activas, encontradas y no encontradas. Desde ahí puedes abrir directamente la gestión de exclusiones.

En `Ajustes` también puedes gestionar:

```text
Carpetas excluidas / Carpetas protegidas externas
```

Estas carpetas son intocables para la aplicación:

- no se organizan
- no se normalizan
- no se reparan con EXIF
- no se deduplican
- no se purgan
- no se limpian como carpetas vacías/junk/residuales
- no se calcula hash
- no se lee EXIF
- no se hidratan si son cloud-only

Conceptos:

```text
InternalProtectedFolders
= carpetas propias de la aplicación

UserExcludedFolders
= carpetas que el usuario decide que la aplicación no debe tocar

VendorManagedFolders
= presets/ayudas para carpetas típicas de apps externas
```

La configuración se guarda de forma portable en:

```text
%APPDATA%\PhotoOrganizer\Config\UserExcludedFolders.json
```

Puede usar rutas absolutas o plantillas como `%USERPROFILE%`, `<SourcePath>` y `<DestinationPath>`. Si una ruta no existe, no se borra automáticamente: queda en la lista y se muestra como no encontrada.

Si el archivo de configuración no existe, UDMRS lo crea automáticamente en `%APPDATA%\PhotoOrganizer\Config\UserExcludedFolders.json` con una lista vacía de exclusiones activas. La instalación compartida no define exclusiones reales para todos los usuarios.

## 9. Reparar índice: ReconcileProcessedDatabase

En el dashboard aparece como:

```text
Sincronizar índice / reparar cambios manuales
```

o como acción equivalente de reparar índice.

Sirve para sobrevivir a movimientos manuales o accidentales:

- carpetas arrastradas por error
- archivos movidos fuera de `Fotos_Organizadas`
- carpetas temporales como `Primavera 2005`, `Revisar` o `Conflictos_Reparados_EXIF`

Comportamiento actual:

- no hereda exclusiones normales pensadas solo para Organize, pero sí respeta carpetas excluidas por el usuario y presets externos protegidos
- sí ve `Fotos_Organizadas`
- valida rutas registradas en `ProcessedFiles.json`
- si la ruta existe, conserva la entrada
- si la ruta antigua no existe, busca el mismo hash en el ámbito actual
- si encuentra una única ruta nueva, actualiza el índice
- si el archivo está fuera de la biblioteca organizada, puede marcarlo como `ManualMoved/ExternalLocation`
- si no aparece, lo marca como `Missing/Stale`
- si hay varias rutas actuales con el mismo hash, lo reporta como conflicto

No mueve fotos. No modifica EXIF. No reorganiza carpetas.

En `DryRun` solo informa. En `Apply` hace backup del JSON interno antes de guardar cambios.

## 10. Purgar entradas desaparecidas

`PurgeMissingFromProcessedDatabase` es una acción avanzada, no parte del flujo normal.

Antes de purgar, revalida cada entrada marcada como desaparecida:

- si la ruta existe de nuevo, la rescata
- si sigue desaparecida, puede purgar la entrada del JSON en `Apply`
- no borra fotos reales
- no mueve archivos
- no modifica EXIF

Úsalo solo cuando ya hayas reconciliado y quieras limpiar historial de archivos que realmente ya no existen.

Si no hay entradas desaparecidas reales, el log debe dejar claro que no hay nada que purgar.

## 11. NormalizeExistingFolders

`NormalizeExistingFolders` es una herramienta opcional para normalizar visualmente una biblioteca ya organizada.

Nuevo rol:

```text
NormalizeExistingFolders -Apply
↓
mueve/renombra a Año\Trimestre
↓
limpia carpetas vacías
↓
limpia junk-only folders
↓
limpia ramas residuales
↓
entra en validación/reconcile post-operación
```

Importante: Normalize ya no debe entenderse como mantenedor principal de `ProcessedFiles.json` movimiento a movimiento. La verdad final del índice corresponde a `ReconcileProcessedDatabase`.

Después de `Normalize Apply`, el motor entra en una fase separada de validación/reconcile para detectar rutas stale y reparar el índice. Si necesitas auditar o reparar cambios manuales, usa `ReconcileProcessedDatabase`.

Siempre revisa primero el reporte `DryRun` antes de aplicar Normalize.

En bibliotecas OneDrive u otros proveedores cloud-backed, `NormalizeExistingFolders -Apply` puede provocar mucha actividad de Explorer/sincronización porque mueve y renombra estructura. Si Explorer queda temporalmente sin refrescar o `No responde`, normalmente está reindexando cambios. El motor no debe leer EXIF ni calcular hashes extra solo por esa espera visual. Deja terminar la sincronización antes de iniciar otra operación.


## 11.1 Fechas visibles: MetadataAudit y MetadataRepair

`NormalizeExistingFolders` arregla estructura. `MetadataRepair` arregla la fecha visible dentro del archivo.

UDMRS puede conocer una fecha fiable por EXIF, provider, sidecar o patrón de nombre. Esa fecha debe ser coherente en:

- ruta/carpeta
- nombre final
- metadata embebida cuando el formato lo permite
- `CreationTime` / `LastWriteTime` cuando son fechas accidentales
- `ProcessedFiles.json` y reportes

`MetadataAudit` revisa la biblioteca organizada y genera un CSV con candidatos. No modifica nada, incluso si el dashboard estuviera en Apply. Úsalo para saber cuántos archivos tienen fecha fiable conocida pero no visible para Windows, OneDrive o Microsoft Photos.

`MetadataRepair` actúa solo sobre candidatos seguros. Crea backup, escribe metadata embebida según formato, sincroniza fechas de sistema si procede, recalcula hash y actualiza el índice.

Formatos cubiertos por política de materialización: JPG/JPEG, HEIC/HEIF, MP4/MOV/M4V/3GP, PNG, TIFF, WEBP y GIF cuando exista forma segura de escribir metadata útil. Si el formato no permite escribir la fecha esperada o existe conflicto con metadata válida, el log/report debe dejarlo como `DateKnownButMetadataNotWritten` o warning equivalente.

Reglas importantes:

- no sobrescribe metadata válida existente
- no resuelve conflictos automáticamente
- no mueve ni reorganiza carpetas
- no deduplica
- no sustituye a Reconcile
- puede tardar mucho en bibliotecas grandes porque lee metadata in-place

Ejemplo DryRun:

``powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" `
  -SourcePath "%USERPROFILE%\OneDrive\Imágenes" `
  -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Fotos_Organizadas" `
  -MetadataAudit `
  -Language es
``

Ejemplo Apply:

``powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" `
  -SourcePath "%USERPROFILE%\OneDrive\Imágenes" `
  -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Fotos_Organizadas" `
  -MetadataRepair `
  -Language es `
  -Apply
``
## 12. ProcessedFiles.json

Ruta habitual:

```text
%APPDATA%\PhotoOrganizer\ProcessedFiles.json
```

Guarda:

- hash
- ruta relativa original o nueva
- estado
- `lastSeen`
- `missingSince`
- versión de herramienta

Estados relevantes:

- entrada activa: alimenta el índice de duplicados
- `ManualMoved/ExternalLocation`: archivo encontrado fuera de la ubicación organizada; no se usa como duplicado canónico automático
- `Missing/Stale`: la ruta registrada no existe actualmente

El hash es la identidad principal. La ruta es actualizable.

## 13. Limpieza segura

El motor puede eliminar carpetas vacías o carpetas que solo contienen basura conocida.

La limpieza segura se ejecuta dentro de `NormalizeExistingFolders`. No existe un comando separado `-Cleanup`.

Importante:

```text
-KeepEmptyFolders NO limpia carpetas vacías.
-KeepEmptyFolders conserva carpetas vacías.
```

Para permitir limpieza durante Normalize usa:

```powershell
-NormalizeExistingFolders `
-KeepEmptyFolders:$false
```

Basura conocida incluye, entre otros:

```text
desktop.ini
Thumbs.db
ehthumbs.db
.DS_Store
.picasa.ini
.dropbox
._*
Icon?
.Trashes
.fseventsd
.Spotlight-V100
*.tmp
```

También puede retirar carpetas con pequeños marcadores seguros, siempre que no haya archivos potencialmente útiles.

Nunca se considera basura si hay extensiones como fotos, vídeos, RAW, JSON, XMP, TXT, PDF, DOC, enlaces, HTML, XML o CSV.

Después de Normalize, la limpieza intenta retirar ramas residuales vacías de forma ascendente.

## 13.1 RetentionCleanup

`RetentionCleanup` es una limpieza de retención, no una limpieza general de la galería.

No hace esto:

- no limpia carpetas vacías
- no ejecuta Normalize
- no reorganiza fotos
- no purga `ProcessedFiles.json`
- no busca duplicados nuevos
- no toca `_Duplicados_Para_Revisar`

Solo puede limpiar contenido temporal o confirmado:

```text
%LOCALAPPDATA%\PhotoOrganizer\_CopiaSeguridadMetadatos
_Cuarentena_Duplicados_Confirmados
```

También reconoce aliases de otros idiomas, como `_Backup_Metadate`, `_MetadataBackup`, `_Carantina_Duplicate_Confirmate` o `_Confirmed_Duplicates_Quarantine`.

`_Duplicados_Para_Revisar` no se toca jamás por `RetentionCleanup`, porque puede contener variantes, casos ambiguos o archivos que requieren revisión humana.

Para nuevas cuarentenas confirmadas, UDMRS registra un manifiesto técnico que relaciona cada copia con su canónico y el SHA256 confirmado. Al ejecutar `RetentionCleanup`, los duplicados que superan 45 días se revalidan con su contenido actual: deben existir ambas copias, estar disponibles localmente y tener el mismo SHA256 actual. RAW/DNG, cloud-only, conflictos de índice, hashes distintos, ejecuciones con errores y la última ejecución correcta siempre se conservan.

Las cuarentenas históricas sin manifiesto también se conservan de forma deliberada. El resultado `No se borró nada` puede ser completamente normal: quizá no se alcanzó la antigüedad o ningún candidato superó todas las comprobaciones. No hay tareas en segundo plano; esta limpieza solo ocurre al lanzar explícitamente `Limpieza técnica` / `RetentionCleanup`.

## 14. Diagnóstico EXIF lento

Si un lote EXIF tarda demasiado, el motor mantiene el fallback actual:

```text
batch timeout
↓
modo individual
```

Además registra observabilidad:

```text
Slow EXIF candidate: Path=...; Extension=...; Size=...; Time=...; Batch=...; Reason=...; Count=...
```

Contadores nuevos:

- `slowExifCandidates`: archivos únicos candidatos a lentitud EXIF
- `slowExifDetections`: total de detecciones

Motivos típicos:

- `Batch timeout/fallback`
- `timeout`
- `slow per-file read`
- `metadata warning`
- `fallback`

Un timeout de batch marca candidatos del lote. El archivo realmente sospechoso suele repetirse luego como timeout individual, lectura lenta o warning de metadata.

Búsqueda útil:

```powershell
Select-String -Path "%APPDATA%\PhotoOrganizer\Logs\*.log" -Pattern "Slow EXIF candidate"
```

## 15. Logs y progress.json

La pestaña `Logs y resumen` muestra salida reciente y permite abrir la carpeta `Logs` incluso durante ejecución.

Archivos habituales:

```text
PhotoOrganizer-YYYYMMDD-HHMMSS.log
PhotoOrganizer-YYYYMMDD-HHMMSS.progress.json
```

El dashboard ignora silenciosamente lecturas o escrituras bloqueadas temporalmente para no mostrar excepciones crudas de PowerShell.

La acción avanzada `Limpieza técnica` ejecuta la limpieza de logs antiguos en `%APPDATA%\PhotoOrganizer\Logs`. La retención aproximada es de 7 días para `*.log`, `*.progress.json` y reportes HTML. Esta limpieza no toca fotos, vídeos, `ProcessedFiles.json` real, backups EXIF ni cuarentenas; tampoco toca `%APPDATA%\PhotoOrganizer\Runtime\TechnicalConsole`.

Los backups técnicos de `ProcessedFiles.json` viven separados en `%APPDATA%\PhotoOrganizer\IndexBackups`. No son logs normales. `Limpieza técnica` conserva siempre el backup más reciente, conserva backups recientes hasta un máximo aproximado de 10 copias y purga el resto. Si existen backups antiguos en `Logs\JsonBackups`, esa acción los migra a `IndexBackups`.

Como medida conservadora, nunca borra archivos modificados en la última hora, protege logs/progress de ejecuciones activas cuando puede detectarlos y conserva siempre al menos el backup JSON más reciente. Si un archivo está bloqueado o hay permisos insuficientes, lo omite y la acción continúa.

`progress.json` puede incluir, entre otros:

- estado
- fase
- PID
- analizados
- movidos
- duplicados
- EXIF reparados
- ETA
- workers
- cola
- errores
- `slowExifCandidates`
- `slowExifDetections`

## 16. Backups de metadatos

Ruta habitual:

```text
%LOCALAPPDATA%\PhotoOrganizer\_CopiaSeguridadMetadatos\<RunId>\
```

Son copias completas de archivos antes de una reparación EXIF real. No son miniaturas ni simples metadatos.

Se crean cuando hay:

- `Aplicar cambios reales`
- `Repair EXIF`
- ExifTool disponible
- reparación EXIF necesaria
- confianza suficiente

No se crean por `DryRun` ni por `Prueba de escaneo`.

Si la ejecución termina correctamente, el motor intenta limpiar backups de esa ejecución. Si hay cancelación, apagón o verificación insegura, puede conservarlos.

No migres esta carpeta entre PCs. El estado portable está en `%APPDATA%\PhotoOrganizer\`.

## 17. Migrar la aplicación a otro PC

PhotoOrganizer está diseñado para ser portable como carpeta de aplicación.

La forma recomendada es abrir `Ajustes` -> `Modo avanzado` -> `Migrar UDMRS a otro PC`.

El asistente genera una carpeta en:

```text
%USERPROFILE%\Downloads\UDMRS-MigrationPackages\
```

Dentro deja:

- un ZIP de instalación compartida, creado desde la raíz real de UDMRS
- un ZIP de estado del usuario actual, creado desde `%APPDATA%\PhotoOrganizer`
- `MigrationGuide.txt`

El ZIP de instalación incluye `App`, `Docs`, `Tools`, `Branding`, `Templates`, `Releases`, `Config`, `README.md` y lanzadores existentes. El ZIP de usuario incluye `ProcessedFiles.json`, `Config`, `IndexBackups`, ajustes del dashboard y JSON de estado necesarios. No incluye `Logs`, `Runtime`, `*.progress.json`, colas, PID/status temporales ni backups EXIF locales.

Al terminar, el dashboard abre automáticamente la carpeta donde se generaron los paquetes.

Para mover la aplicación:

1. Copia la carpeta completa `<CarpetaUDMRS>` al nuevo PC.
2. Abre `Start-PhotoOrganizer.cmd` desde la carpeta copiada.
3. No reutilices accesos directos `.lnk` antiguos. Windows guarda rutas absolutas dentro del acceso directo y podría abrir una copia vieja.
4. Si quieres un acceso directo en el escritorio, créalo de nuevo desde la copia actual.

Para conservar progreso, hashes y estado incremental, copia también:

```text
%APPDATA%\PhotoOrganizer\
```

en la misma ubicación del perfil Windows del nuevo equipo.

Ese directorio contiene normalmente:

- `ProcessedFiles.json`
- `settings.json`
- otros JSON internos de estado

Los logs, reportes HTML y progress viven en:

```text
%APPDATA%\PhotoOrganizer\Logs\
```

Esa carpeta debe poder limpiarse cuando no hay ejecuciones activas. La infraestructura runtime del modo avanzado vive separada en:

```text
%APPDATA%\PhotoOrganizer\Runtime\TechnicalConsole\
```

No es log ni reporte; contiene runner, cola y estado temporal controlado por el dashboard. Los backups temporales de metadatos viven en `%LOCALAPPDATA%\PhotoOrganizer\` y no deben migrarse salvo que estés investigando una ejecución reciente.

El lanzador `.cmd` usa `%~dp0` para localizar su propia carpeta. Los scripts usan `$PSScriptRoot` para resolver recursos, idiomas, manuales y ExifTool dentro de la instalación actual. Logs, runtime técnico, ajustes y exclusiones de usuario viven en `%APPDATA%\PhotoOrganizer`. Por eso la aplicación no debe depender de rutas privadas fijas.

Si la biblioteca vive en OneDrive u otro proveedor cloud:

- espera a que la sincronización termine antes de procesos grandes
- marca la biblioteca como disponible sin conexión si vas a ejecutar Organize, Normalize o Repair masivos
- UDMRS Dashboard puede convivir con contenido sincronizado, pero solo procesa archivos localmente verificables
- los placeholders cloud-only se saltan y se reportan, no se descargan automáticamente

Después de migrar, restaurar o mover carpetas manualmente, ejecuta:

```text
Sincronizar índice / reparar cambios manuales
```

Esto revalida `ProcessedFiles.json`, actualiza rutas por hash cuando procede y evita falsos duplicados por rutas antiguas.

Regla importante:

```text
Una sola operación activa del motor a la vez.
```

No ejecutes dos dashboards, dos consolas técnicas ni dos comandos PowerShell contra la misma biblioteca al mismo tiempo. El dashboard bloquea acciones normales mientras `Modo avanzado` ejecuta una acción avanzada, pero también conviene evitar lanzamientos manuales paralelos.

## 18. Qué no hacer durante una ejecución

Mientras PhotoOrganizer está activo, no muevas ni borres:

- carpeta origen
- carpeta destino
- carpeta `Logs`
- carpeta `_CopiaSeguridadMetadatos`
- archivos que estén siendo procesados

Usa `Cancelar` desde el dashboard si necesitas detenerlo.

## 19. Comandos oficiales rápidos

La referencia completa con todos los comandos está en `Docs\CommandReference.html`.

### Test scan

```powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" -SourcePath "%USERPROFILE%\OneDrive\Imágenes" -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Fotos_Organizadas" -TestScan -Language es
```

### Organize DryRun

```powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" -SourcePath "%USERPROFILE%\OneDrive\Imágenes" -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Fotos_Organizadas" -Language es
```

### Organize Apply + RepairExif

```powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" -SourcePath "%USERPROFILE%\OneDrive\Imágenes" -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Fotos_Organizadas" -Apply -RepairExif -Language es
```

### NormalizeExistingFolders DryRun con limpieza segura

```powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" -SourcePath "%USERPROFILE%\OneDrive\Imágenes" -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Fotos_Organizadas" -NormalizeExistingFolders -KeepEmptyFolders:$false -Language es
```

### ReconcileProcessedDatabase DryRun

```powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" -SourcePath "%USERPROFILE%\OneDrive\Imágenes" -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Fotos_Organizadas" -ReconcileProcessedDatabase -Language es
```

### DedupeCleanup DryRun

```powershell
pwsh -ExecutionPolicy Bypass -NoProfile -File "<CarpetaUDMRS>\App\PhotoOrganizer.ps1" -SourcePath "%USERPROFILE%\OneDrive\Imágenes" -DestinationPath "%USERPROFILE%\OneDrive\Imágenes\Fotos_Organizadas" -DedupeCleanup -Language es
```

### Buscar problemas en logs

```powershell
Select-String -Path "%APPDATA%\PhotoOrganizer\Logs\*.log" -Pattern "Error","Slow EXIF candidate","DateInfo resolved","Missing/Stale","JSON conflict"
```

Los comandos Apply completos están en `Docs\CommandReference.html`. Revisa siempre DryRun y reportes HTML antes de aplicar cambios reales.









