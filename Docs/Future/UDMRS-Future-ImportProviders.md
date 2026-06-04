# UDMRS Dashboard - nota futura: Import Providers

Estado: arquitectura futura documentada, con Google Takeout y XMP / Sidecar Library disponibles.  
La build estable RC1 sigue congelada como referencia single-head; la segunda entrada no cambia el flujo clasico `Organize`.

## Objetivo

UDMRS debe mantener el flujo estable actual:

```text
Carpeta origen
↓
Organize clasico
↓
Fotos_Organizadas
↓
Year\Quarter
```

La segunda entrada especializada se introduce de forma gradual:

```text
Import Providers
```

La segunda cabeza crece por riqueza semantica exportada, no por marca. No se crean providers para servicios que solo entregan una carpeta normal de JPG/PNG/HEIC/MP4/MOV sin metadatos auxiliares relevantes.

Roadmap oficial inicial:

```text
Google Photos / Google Takeout
Apple Photos / iCloud Photos
Samsung Gallery
Immich
Generic XMP / Sidecar Library
```

La implementacion no debe basarse solo en documentacion publica o supuestos. Debe apoyarse principalmente en muestras reales de exportacion generadas por usuarios y analizadas antes de escribir logica de importacion.

No se planifican providers oficiales iniciales para:

```text
OneDrive Photos
Amazon Photos
Dropbox Photos
Google Drive folders
otros sistemas que simplemente entregan archivos multimedia normales
```

Para esos casos se usa `Organize` clasico.

## Registro de providers

| Provider | Estado | Gate | Justificacion |
|---|---|---|---|
| `GoogleTakeout` | Available | Implemented | Sidecars JSON, albumes, papelera, timestamps provider y duplicados fisicos por album. |
| `ApplePhotos` | Planned | SampleRequired | Live Photos como pares imagen/video, originales HEIC/MOV, variantes editadas y XMP/IPTC opcional. |
| `SamsungGallery` | Planned | SampleGated | Adopcion enorme y posible semantica de albumes/stories/trash, pero requiere exportacion real antes de implementar. |
| `Immich` | Planned | SampleRequired | XMP sidecars, tags, ratings, descripciones, albumes y metadata de biblioteca. |
| `XmpSidecarLibrary` | Available | Implemented | Provider generico y conservador para bibliotecas desconocidas con sidecars XMP/JSON/YAML. |

`SamsungGallery` queda contemplado oficialmente, pero deshabilitado hasta disponer de una muestra real adecuada.

## Estado Google Takeout

La primera entrada disponible es:

```text
ImportProvider GoogleTakeout
```

Acceso:

```text
Dashboard -> Importar galería -> Google Photos / Takeout
```

El usuario selecciona la ruta de exportacion. No hay rutas fijas. El provider analiza `Google Fotos` / `Google Photos` / nombres localizados equivalentes, carpetas por año, carpetas de album, papelera, medios JPG/PNG/MP4 y sidecars Google `*.supp*.json`, incluidas variantes truncadas generadas por Takeout.

Comportamiento:

- separa identidad logica del asset de sus ocurrencias fisicas;
- deduplica ocurrencias por SHA256 para no importar copias de album como recuerdos nuevos;
- usa `photoTakenTime` de Google como fuente fuerte cuando es coherente;
- trata contradicciones proveedor/EXIF como conflicto conservador;
- no importa automaticamente assets que solo estan en `Papelera`;
- genera reporte HTML/JSON;
- en Apply copia assets limpios al destino `Year\Quarter` y registra `ProcessedFiles.json`;
- puede eliminar opcionalmente la carpeta Takeout seleccionada tras un Apply correcto, solo con confirmacion explicita del usuario y dejando constancia en log/reporte.

## Estado XMP / Sidecar Library

La segunda entrada disponible es:

```text
ImportProvider XmpSidecarLibrary
```

Acceso:

```text
Dashboard -> Importar galería -> XMP / Sidecar Library
```

El usuario selecciona la carpeta de la galeria/exportacion. No hay rutas fijas. El provider intenta asociar media con sidecars en la misma carpeta usando patrones claros:

```text
media.ext.xmp
media.xmp
media.ext.json
media.json
media.ext.yaml / media.yaml
media.ext.yml / media.yml
metadata.json / metadata.yaml como contexto de carpeta
```

Comportamiento:

- lee sidecars XMP mediante ExifTool cuando esta disponible;
- interpreta sidecars JSON y YAML simples con campos comunes de fecha, titulo, descripcion, tags, rating y ubicacion;
- si una relacion media/sidecar es clara, usa la metadata;
- si hay varios sidecars posibles para el mismo media, marca conflicto y envia a revision;
- si no hay sidecar o no contiene metadata util, usa fallback clasico de fecha/naming del motor;
- genera reporte HTML/JSON con media, sidecars usados, sidecars huerfanos, media sin sidecar, confianza y conflictos;
- no elimina automaticamente la fuente.

## Digestión de proveedores

La fase futura de Import Providers deberia analizar:

```text
estructura de carpetas
metadatos auxiliares
JSON externos
albumes
relaciones entre recursos
sidecars y archivos asociados
```

El objetivo es construir un modelo interno de interpretacion por proveedor, por ejemplo:

```text
GoogleTakeout
ApplePhotos
Samsung Gallery
Immich
XmpSidecarLibrary
UnknownProvider
```

## Filosofía operativa

Cuando el usuario seleccione una exportacion de proveedor, UDMRS deberia intentar primero:

```text
ImportProvider
```

Motivo: una exportacion de proveedor puede contener mas contexto que una carpeta caotica normal:

```text
fecha original
albumes
ubicaciones
metadata externa
relaciones foto-video
estructura de exportacion
```

El motor clasico debe seguir existiendo como red de seguridad permanente.

## Degradación segura

Si el proveedor cambia su estructura o el formato no se reconoce con suficiente confianza:

```text
Intentar interpretar como proveedor conocido
↓
No reconocido / confianza insuficiente
↓
Registrar advertencia clara
↓
Degradar automaticamente al motor clasico
↓
Organize / RepairExif / flujo tradicional
```

Regla importante:

```text
Provider Import = primera opcion cuando hay evidencia suficiente
Motor clasico = red de seguridad permanente
```

UDMRS nunca debe perder la capacidad de procesar una biblioteca porque un proveedor haya cambiado su estructura, sus JSON o su forma de exportar.

## Reglas futuras

- No bloquear la importacion si el proveedor no se reconoce.
- No descartar archivos por no entender el formato del proveedor.
- No exigir intervencion manual para volver al flujo clasico.
- No mezclar la logica de proveedor con el Organize clasico.
- Mantener un unico destino final: `Fotos_Organizadas\Year\Quarter` o su equivalente localizado.
- Mantener un unico indice: `ProcessedFiles.json`.
- Compartir deduplicacion global, Reconcile, Purge y Recovery entre ambos flujos.
- Mantener DryRun obligatorio antes de Apply.

## Reporte esperado

Un futuro ImportProvider deberia generar un reporte claro con:

```text
proveedor detectado
confianza de deteccion
sidecars usados
archivos importables
archivos dudosos
duplicados
conflictos
elementos ignorados
motivos de fallback
```

## Pendiente futuro

Cuando existan exportaciones reales:

```text
iCloud Photos
Google Takeout
otros proveedores
```

se debe hacer primero una auditoria documental de sus estructuras reales y crear documentacion tecnica especifica por proveedor antes de implementar la segunda entrada de importacion.


