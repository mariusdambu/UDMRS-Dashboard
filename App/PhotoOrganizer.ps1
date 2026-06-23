<#
.SYNOPSIS
Safely organizes a Windows photo/video library with conservative duplicate
detection and optional EXIF repair.

.DESCRIPTION
Default source:
  %UserProfile%\OneDrive\Imagini

The script is intentionally conservative:
  - DryRun is the default.
  - Real file operations require -Apply.
  - EXIF writes require both -Apply and -RepairExif.
  - Internal work folders and configured user/vendor excluded folders are fully excluded.
  - Files are never overwritten. Existing identical targets are skipped.
  - Duplicates are moved/copied to the localized duplicates folder, not deleted.
  - Ambiguous files are moved/copied to the localized review folder only when -Apply is used.

External tool recommended:
  exiftool.exe in PATH, or pass -ExifToolPath.
#>

[CmdletBinding()]
param(
    [string]$SourcePath = (Join-Path $env:USERPROFILE 'OneDrive\Imagini'),
    [string]$DestinationPath = '',
    [switch]$Apply,
    [switch]$RepairExif,
    [switch]$CopyInsteadOfMove,
    [switch]$TestScan,
    [switch]$ReconcileProcessedDatabase,
    [switch]$PurgeMissingFromProcessedDatabase,
    [switch]$DedupeCleanup,
    [switch]$RetentionCleanup,
    [switch]$RecoverFromWrongDuplicateMove,
    [string]$RecoveryLogPath = '',
    [switch]$RepairOnlyExistingOrganizedLibrary,
    [switch]$MetadataAudit,
    [switch]$MetadataRepair,
    [switch]$SyncFileSystemDates,
    [switch]$Diagnostic,
    [switch]$KeepEmptyFolders,
    [ValidateSet('es', 'ro', 'en')]
    [string]$Language = 'es',
    [switch]$NormalizeExistingFolders,
    [switch]$RenameExistingFoldersToCurrentLanguage,
    [switch]$RenameInternalFoldersToCurrentLanguage,
    [string]$ImportProvider = '',
    [string]$ImportProviderPath = '',
    [switch]$DeleteImportProviderSourceAfterSuccess,
    [string]$LogPath = '',
    [string]$ProgressPath = '',
    [string]$ProcessedDbPath = '',
    [int]$BatchSize = 250,
    [ValidateSet('Safe', 'Balanced', 'HighPerformance')]
    [string]$PerformanceMode = 'Balanced',
    [int]$MaxParallelJobs = 0,
    [string]$ExifToolPath = 'exiftool.exe',
    [int]$ExifToolTimeoutSeconds = 30,
    [int]$ExifBatchTimeoutSeconds = 60,
    [int]$ExifRepairConfidence = 95,
    [int]$AutoActionConfidence = 99,
    [int]$MetadataBackupRetentionDays = 30,
    [int]$ConfirmedDuplicatesRetentionDays = 45,
    [double]$MetadataBackupMaxGB = 0,
    [double]$ConfirmedDuplicatesMaxGB = 0,
    [ValidateSet('QuarterlyFolders')]
    [string]$OrganizationProfile = 'QuarterlyFolders'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$ImageExtensions = @('.jpg', '.jpeg', '.png', '.heic', '.heif', '.tif', '.tiff', '.bmp', '.gif', '.webp', '.dng', '.cr2', '.cr3', '.nef', '.arw', '.rw2', '.orf')
$VideoExtensions = @('.mp4', '.mov', '.m4v', '.avi', '.mkv', '.3gp', '.mts', '.m2ts', '.wmv')
$MediaExtensions = @($ImageExtensions + $VideoExtensions)
$RawExtensions = @('.dng', '.cr2', '.cr3', '.nef', '.arw', '.rw2', '.orf')
$VisibleCaptureMetadataReadTags = @(
    '-DateTimeOriginal', '-CreateDate', '-MediaCreateDate',
    '-DateCreated', '-XMP:DateCreated', '-XMP:CreateDate', '-XMP:ModifyDate',
    '-PNG:CreationTime', '-TrackCreateDate', '-TrackModifyDate',
    '-QuickTime:CreateDate', '-QuickTime:ModifyDate'
)
$VisibleCaptureMetadataFields = @(
    'DateTimeOriginal', 'CreateDate', 'MediaCreateDate',
    'DateCreated', 'XMPDateCreated', 'XMPCreateDate', 'XMPModifyDate',
    'CreationTime', 'PNGCreationTime', 'TrackCreateDate', 'TrackModifyDate',
    'QuickTimeCreateDate', 'QuickTimeModifyDate'
)

$NoDescriptionByLanguage = @{
    es = 'Sin descripcion'
    ro = 'Fara descriere'
    en = 'No description'
}

$QuarterFolderNamesByLanguage = @{
    es = @{
        1 = 'Ene-Mar'; 2 = 'Abr-Jun'; 3 = 'Jul-Sep'; 4 = 'Oct-Dic'
    }
    ro = @{
        1 = 'Ian-Mar'; 2 = 'Apr-Iun'; 3 = 'Iul-Sep'; 4 = 'Oct-Dec'
    }
    en = @{
        1 = 'Jan-Mar'; 2 = 'Apr-Jun'; 3 = 'Jul-Sep'; 4 = 'Oct-Dec'
    }
}

$DefaultLanguageResources = @{
    es = @{
        OrganizedFolder = 'Fotos_Organizadas'
        NeedsReviewFolder = '_NecesitaRevision'
        DuplicatesFolder = '_Duplicados_Para_Revisar'
        ConfirmedDuplicatesQuarantineFolder = '_Cuarentena_Duplicados_Confirmados'
        MetadataBackupFolder = '_CopiaSeguridadMetadatos'
        LogsFolder = 'Logs'
        MediaMetadataIssuesFolder = 'MediaMetadataIssues'
    }
    ro = @{
        OrganizedFolder = 'Poze_Organizate'
        NeedsReviewFolder = '_De_Revizuit'
        DuplicatesFolder = '_Duplicate_De_Revizuit'
        ConfirmedDuplicatesQuarantineFolder = '_Carantina_Duplicate_Confirmate'
        MetadataBackupFolder = '_Backup_Metadate'
        LogsFolder = 'Logs'
        MediaMetadataIssuesFolder = 'MediaMetadataIssues'
    }
    en = @{
        OrganizedFolder = 'Organized_Photos'
        NeedsReviewFolder = '_NeedsReview'
        DuplicatesFolder = '_Duplicates_To_Review'
        ConfirmedDuplicatesQuarantineFolder = '_Confirmed_Duplicates_Quarantine'
        MetadataBackupFolder = '_MetadataBackup'
        LogsFolder = 'Logs'
        MediaMetadataIssuesFolder = 'MediaMetadataIssues'
    }
}

$SummaryTextByLanguage = @{
    es = @{
        FinalSummary = 'Resumen final'
        FilesFound = 'Archivos encontrados'
        LocalFilesDetected = 'Archivos locales detectados'
        CloudPlaceholdersSkipped = 'Placeholders cloud omitidos'
        CloudPlaceholdersInIndex = 'Placeholders cloud en indice'
        MissingReal = 'Missing reales'
        FilesAnalyzed = 'Archivos analizados'
        IncrementalSkipped = 'Omitidos por incremental'
        ExactDuplicatesFound = 'Duplicados exactos encontrados'
        NearDuplicatesFound = 'Duplicados similares encontrados'
        ExifRepaired = 'EXIF reparados'
        FilesMoved = 'Archivos movidos'
        FilesCopied = 'Archivos copiados'
        ExistingIdenticalSkipped = 'Identicos ya existentes'
        NeedsReview = 'Casos NeedsReview'
        EmptyFoldersRemoved = 'Carpetas vacias eliminadas'
        JunkOnlyFoldersRemoved = 'Carpetas solo basura eliminadas'
        JunkOnlySmallMarkerFoldersRemoved = 'Marcadores pequenos eliminados'
        ZombieNormalizeFoldersRemoved = 'Zombie Normalize eliminados'
        OrganizationProfile = 'Perfil de organizacion'
        FoldersReduced = 'Carpetas reducidas'
        JsonPathsUpdated = 'Rutas JSON actualizadas'
        SkippedUncertainNames = 'Nombres inciertos omitidos'
        SkippedOneDrive = 'Cloud/problematicos omitidos'
        Inaccessible = 'Inaccesibles'
        LockedFiles = 'Archivos bloqueados'
        RetryCount = 'Reintentos de transferencia'
        MetadataCorruptedMedia = 'Medios con metadata corrupta'
        SlowExifCandidates = 'Archivos EXIF lentos confirmados'
        SlowExifDetections = 'Detecciones EXIF lentas confirmadas'
        ExifBatchTimeouts = 'Timeouts de lotes EXIF'
        ExifBatchTimeoutAffectedFiles = 'Archivos afectados por timeouts de lotes EXIF'
        ExifBatchFallbacks = 'Fallbacks de lotes EXIF'
        ExifBatchFallbackAffectedFiles = 'Archivos afectados por fallbacks de lotes EXIF'
        MetadataBackupSizeGB = 'Tamano backup metadata GB'
        RetentionDeletedItems = 'Elementos eliminados por retencion'
        RetentionRecoveredGB = 'Espacio recuperado por retencion GB'
        JsonReconcileValid = 'JSON entradas validas'
        JsonReconcileStale = 'JSON entradas obsoletas'
        JsonReconcilePathsUpdated = 'JSON rutas actualizadas'
        JsonReconcileMissing = 'JSON entradas marcadas como desaparecidas'
        JsonReconcilePurged = 'JSON entradas desaparecidas purgadas'
        JsonReconcileEntriesRemoved = 'JSON entradas eliminadas'
        JsonReconcileConflicts = 'JSON conflictos'
        DryRunActions = 'Acciones simuladas DryRun'
        Errors = 'Errores'
        DryRunNoChanges = 'No se ha modificado nada. Ejecuta con -Apply para aplicar cambios reales.'
    }
    ro = @{
        FinalSummary = 'Rezumat final'
        FilesFound = 'Fisiere gasite'
        LocalFilesDetected = 'Fisiere locale detectate'
        CloudPlaceholdersSkipped = 'Placeholdere cloud omise'
        CloudPlaceholdersInIndex = 'Placeholdere cloud in index'
        MissingReal = 'Lipsa reala'
        FilesAnalyzed = 'Fisiere analizate'
        IncrementalSkipped = 'Omise incremental'
        ExactDuplicatesFound = 'Duplicate exacte gasite'
        NearDuplicatesFound = 'Duplicate similare gasite'
        ExifRepaired = 'EXIF reparate'
        FilesMoved = 'Fisiere mutate'
        FilesCopied = 'Fisiere copiate'
        ExistingIdenticalSkipped = 'Identice deja existente'
        NeedsReview = 'Cazuri NeedsReview'
        EmptyFoldersRemoved = 'Foldere goale eliminate'
        JunkOnlyFoldersRemoved = 'Foldere doar gunoi eliminate'
        JunkOnlySmallMarkerFoldersRemoved = 'Markere mici eliminate'
        ZombieNormalizeFoldersRemoved = 'Zombie Normalize eliminate'
        OrganizationProfile = 'Profil organizare'
        FoldersReduced = 'Foldere reduse'
        JsonPathsUpdated = 'Rute JSON actualizate'
        SkippedUncertainNames = 'Nume incerte omise'
        SkippedOneDrive = 'Cloud/problematice omise'
        Inaccessible = 'Inaccesibile'
        LockedFiles = 'Fisiere blocate'
        RetryCount = 'Reincercari transfer'
        MetadataCorruptedMedia = 'Media cu metadata corupta'
        SlowExifCandidates = 'Fisiere EXIF lente confirmate'
        SlowExifDetections = 'Detectii EXIF lente confirmate'
        ExifBatchTimeouts = 'Timeouturi loturi EXIF'
        ExifBatchTimeoutAffectedFiles = 'Fisiere afectate de timeouturi loturi EXIF'
        ExifBatchFallbacks = 'Fallbackuri loturi EXIF'
        ExifBatchFallbackAffectedFiles = 'Fisiere afectate de fallbackuri loturi EXIF'
        MetadataBackupSizeGB = 'Dimensiune backup metadata GB'
        RetentionDeletedItems = 'Elemente eliminate prin retentie'
        RetentionRecoveredGB = 'Spatiu recuperat prin retentie GB'
        JsonReconcileValid = 'Intrari JSON valide'
        JsonReconcileStale = 'Intrari JSON invechite'
        JsonReconcilePathsUpdated = 'Rute JSON actualizate'
        JsonReconcileMissing = 'Intrari JSON marcate ca lipsa'
        JsonReconcilePurged = 'Intrari JSON lipsa eliminate'
        JsonReconcileEntriesRemoved = 'Intrari JSON eliminate'
        JsonReconcileConflicts = 'Conflicte JSON'
        DryRunActions = 'Actiuni simulate DryRun'
        Errors = 'Erori'
        DryRunNoChanges = 'Nu s-a modificat nimic. Ruleaza cu -Apply pentru modificari reale.'
    }
    en = @{
        FinalSummary = 'Final summary'
        FilesFound = 'Files found'
        LocalFilesDetected = 'Local files detected'
        CloudPlaceholdersSkipped = 'Cloud placeholders skipped'
        CloudPlaceholdersInIndex = 'Cloud placeholders in index'
        MissingReal = 'Missing real'
        FilesAnalyzed = 'Files analyzed'
        IncrementalSkipped = 'Incremental skipped'
        ExactDuplicatesFound = 'Exact duplicates found'
        NearDuplicatesFound = 'Similar duplicates found'
        ExifRepaired = 'EXIF repaired'
        FilesMoved = 'Files moved'
        FilesCopied = 'Files copied'
        ExistingIdenticalSkipped = 'Existing identical skipped'
        NeedsReview = 'NeedsReview cases'
        EmptyFoldersRemoved = 'Empty folders removed'
        JunkOnlyFoldersRemoved = 'Junk-only folders removed'
        JunkOnlySmallMarkerFoldersRemoved = 'Small marker junk removed'
        ZombieNormalizeFoldersRemoved = 'Normalize zombie folders removed'
        OrganizationProfile = 'Organization profile'
        FoldersReduced = 'Folders reduced'
        JsonPathsUpdated = 'JSON paths updated'
        SkippedUncertainNames = 'Uncertain names skipped'
        SkippedOneDrive = 'Cloud/problematic skipped'
        Inaccessible = 'Inaccessible'
        LockedFiles = 'Locked files'
        RetryCount = 'Transfer retries'
        MetadataCorruptedMedia = 'Metadata-corrupted media'
        SlowExifCandidates = 'Confirmed slow EXIF files'
        SlowExifDetections = 'Confirmed slow EXIF detections'
        ExifBatchTimeouts = 'EXIF batch timeouts'
        ExifBatchTimeoutAffectedFiles = 'Files affected by EXIF batch timeouts'
        ExifBatchFallbacks = 'EXIF batch fallbacks'
        ExifBatchFallbackAffectedFiles = 'Files affected by EXIF batch fallbacks'
        MetadataBackupSizeGB = 'Metadata backup size GB'
        RetentionDeletedItems = 'Retention deleted items'
        RetentionRecoveredGB = 'Retention recovered space GB'
        JsonReconcileValid = 'Valid JSON entries'
        JsonReconcileStale = 'Stale JSON entries'
        JsonReconcilePathsUpdated = 'JSON paths updated'
        JsonReconcileMissing = 'JSON entries marked missing'
        JsonReconcilePurged = 'Missing JSON entries purged'
        JsonReconcileEntriesRemoved = 'JSON entries removed'
        JsonReconcileConflicts = 'JSON conflicts'
        DryRunActions = 'Simulated DryRun actions'
        Errors = 'Errors'
        DryRunNoChanges = 'No changes were made. Run with -Apply to apply real changes.'
    }
}

function Get-ScriptRootPath {
    if ($PSScriptRoot) { return $PSScriptRoot }
    return [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Path)
}

function Get-InstallRootPath {
    $scriptRoot = Get-ScriptRootPath
    if ([System.IO.Path]::GetFileName($scriptRoot) -ieq 'App') {
        return [System.IO.Path]::GetDirectoryName($scriptRoot)
    }
    return $scriptRoot
}

function Get-UserDataRootPath {
    $base = if (-not [string]::IsNullOrWhiteSpace($env:APPDATA)) { $env:APPDATA } else { [System.IO.Path]::GetTempPath() }
    return (Join-Path $base 'PhotoOrganizer')
}

function Get-LanguageResources {
    $resources = @{}
    foreach ($code in $DefaultLanguageResources.Keys) {
        $resources[$code] = @{}
        foreach ($key in $DefaultLanguageResources[$code].Keys) {
            $resources[$code][$key] = $DefaultLanguageResources[$code][$key]
        }
    }

    $resourcePath = Join-Path (Get-ScriptRootPath) 'LanguageResources.json'
    if (Test-Path -LiteralPath $resourcePath -PathType Leaf) {
        try {
            $loaded = Get-Content -LiteralPath $resourcePath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($languageProperty in $loaded.PSObject.Properties) {
                $code = [string]$languageProperty.Name
                if (-not $resources.ContainsKey($code)) { $resources[$code] = @{} }
                foreach ($folderProperty in $languageProperty.Value.PSObject.Properties) {
                    $resources[$code][[string]$folderProperty.Name] = [string]$folderProperty.Value
                }
            }
        }
        catch {
            Write-Warning "Could not read LanguageResources.json. Built-in folder names will be used. Error: $($_.Exception.Message)"
        }
    }

    foreach ($code in @('es', 'ro', 'en')) {
        if (-not $resources.ContainsKey($code)) { $resources[$code] = @{} }
        foreach ($key in $DefaultLanguageResources.es.Keys) {
            if (-not $resources[$code].ContainsKey($key) -or [string]::IsNullOrWhiteSpace($resources[$code][$key])) {
                $fallbackCode = if ($DefaultLanguageResources.ContainsKey($code)) { $code } else { 'es' }
                $resources[$code][$key] = $DefaultLanguageResources[$fallbackCode][$key]
            }
        }
    }

    return $resources
}

$LanguageResourcesByCode = Get-LanguageResources
$InternalFolderResources = if ($LanguageResourcesByCode.ContainsKey($Language)) { $LanguageResourcesByCode[$Language] } else { $LanguageResourcesByCode.es }

function Get-InternalFolderName {
    param([string]$Key)

    if ($InternalFolderResources.ContainsKey($Key) -and -not [string]::IsNullOrWhiteSpace($InternalFolderResources[$Key])) {
        return [string]$InternalFolderResources[$Key]
    }

    return [string]$DefaultLanguageResources.es[$Key]
}

function Get-AllInternalFolderNames {
    param([string]$Key)

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($code in $LanguageResourcesByCode.Keys) {
        if ($LanguageResourcesByCode[$code].ContainsKey($Key)) {
            $value = [string]$LanguageResourcesByCode[$code][$Key]
            if (-not [string]::IsNullOrWhiteSpace($value) -and -not $names.Contains($value)) {
                $names.Add($value)
            }
        }
    }

    return @($names)
}

function Resolve-LocalizedDestinationPath {
    param(
        [string]$RequestedDestinationPath,
        [string]$ResolvedSourcePath
    )

    $organizedFolder = Get-InternalFolderName -Key 'OrganizedFolder'
    if ([string]::IsNullOrWhiteSpace($RequestedDestinationPath)) {
        return Join-Path $ResolvedSourcePath $organizedFolder
    }

    $trimmed = $RequestedDestinationPath.TrimEnd('\')
    $leaf = [System.IO.Path]::GetFileName($trimmed)
    $knownOrganizedNames = @(Get-AllInternalFolderNames -Key 'OrganizedFolder')
    if ($knownOrganizedNames -contains $leaf) {
        $parent = [System.IO.Path]::GetDirectoryName($trimmed)
        if (-not [string]::IsNullOrWhiteSpace($parent)) {
            return Join-Path $parent $organizedFolder
        }
    }

    return $RequestedDestinationPath
}

$NoDescriptionText = $NoDescriptionByLanguage[$Language]
$SummaryText = if ($SummaryTextByLanguage.ContainsKey($Language)) { $SummaryTextByLanguage[$Language] } else { $SummaryTextByLanguage.es }

function Get-SummaryText {
    param([string]$Key)

    if ($SummaryText.ContainsKey($Key) -and -not [string]::IsNullOrWhiteSpace([string]$SummaryText[$Key])) {
        return [string]$SummaryText[$Key]
    }
    return [string]$SummaryTextByLanguage.es[$Key]
}

$ImportProviderTextByLanguage = @{
    es = @{
        Mode = 'Modo'
        MediaFiles = 'Medios'
        LogicalAssets = 'Assets lógicos'
        Occurrences = 'Ocurrencias'
        Albums = 'Álbumes'
        InternalDuplicateOccurrences = 'Duplicados físicos internos'
        TrashOccurrences = 'Ocurrencias en papelera'
        Videos = 'Vídeos'
        Conflicts = 'Conflictos'
        HighConfidence = 'Confianza alta'
        MediumConfidence = 'Confianza media'
        LowConfidence = 'Confianza baja'
        ExifReadsSkipped = 'Lecturas EXIF omitidas'
        ExifReadsSkippedVideo = 'Vídeos omitidos por confianza provider'
        TakeoutSourceDeletion = 'Eliminación fuente ImportProvider'
        AssetsFirst500 = 'Assets (primeros 500)'
        Hash = 'Hash'
        Status = 'Estado'
        Path = 'Ruta'
        OccurrenceCount = 'Ocurrencias'
        AlbumColumn = 'Álbumes'
        DateSource = 'Fuente de fecha'
        Confidence = 'Confianza'
        ExifVerification = 'Verificación EXIF'
        EmbeddedDateState = 'Estado fecha embebida'
        Warnings = 'Avisos'
        SidecarsFound = 'Sidecars encontrados'
        SidecarsUsed = 'Sidecars usados'
        OrphanSidecars = 'Sidecars huérfanos'
        MediaWithoutSidecar = 'Media sin sidecar'
        AmbiguousSidecars = 'Sidecars ambiguos'
        ClassicFallbackAssets = 'Assets con fallback clásico'
        Copied = 'Copiados'
        Sidecar = 'Sidecar'
        SidecarFields = 'Campos sidecar'
        ImportStarted = 'ImportProvider {0} iniciado. Modo: {1}. Raíz={2}'
        GoogleScan = 'ImportProvider Google Takeout scan: medios={0}; sidecars={1}; álbumes={2}; jsonRaíz={3}'
        AppleScan = 'ImportProvider Apple Photos / iCloud scan: medios={0}; detallesCsv={1}; albumesCsv={2}; papelera={3}; livePhotoCandidatos={4}'
        XmpScan = 'ImportProvider XMP / Sidecar Library scan: medios={0}; sidecars={1}; metadataCarpeta={2}'
        ExifPlan = 'Plan de verificación EXIF ImportProvider {0}: leer={1}; omitidosPorConfianzaProvider={2}; diagnóstico={3}'
        ExifPlanVideoTrusted = 'Plan de verificación EXIF ImportProvider {0}: vídeos omitidos por confianza provider={1}'
        ExifVerificationStart = 'Verificando metadata EXIF/QuickTime: 0 / {0} (0%). Velocidad: calculando. Transcurrido: 0m. Restante estimado: desconocido.'
        ExifVerificationProgress = 'Verificando metadata EXIF/QuickTime: {0} / {1} ({2}%). Velocidad: {3}/min. Transcurrido: {4}. Restante estimado: {5}. Archivo actual: {6}'
        ExifVerificationComplete = 'Verificación EXIF/QuickTime completada: {0} / {1} (100%). Transcurrido: {2}.'
        ReportWritten = 'Reporte ImportProvider escrito: {0}'
        JsonReportWritten = 'Reporte JSON ImportProvider escrito: {0}'
        SourceNotDeleted = 'Fuente ImportProvider no eliminada. Motivo=No solicitado. Ruta={0}'
        SourceDeletionSkippedDryRun = 'Eliminación de fuente ImportProvider omitida. Motivo=DryRun. Ruta={0}'
        SourceDeletionSkippedErrors = 'Eliminación de fuente ImportProvider omitida. Motivo=El import terminó con errores; Errores={0}. Ruta={1}'
        SourceDeletionSkippedMissing = 'Eliminación de fuente ImportProvider omitida. Motivo=La ruta seleccionada ya no existe. Ruta={0}'
        SourceDeletionBlockedRoot = 'Eliminación de fuente ImportProvider bloqueada. Motivo=La ruta seleccionada es una raíz de disco. Ruta={0}'
        SourceDeletionBlockedReportInsideSource = 'Eliminación de fuente ImportProvider bloqueada. Motivo=La carpeta de logs/reportes está dentro de la fuente seleccionada; se conserva la evidencia. Ruta={0}; LogRoot={1}'
        SourceDeleted = 'Fuente ImportProvider eliminada tras importación correcta: {0}'
        SourceDeletionFailed = 'Error eliminando fuente ImportProvider: {0}. Error={1}'
        SourceDeletionStatusLog = 'Estado de eliminación de fuente ImportProvider: {0}. Ruta={1}. Reporte={2}'
        WouldImportAsset = 'Se importaría asset de provider: {0} -> {1}'
        ImportProviderHashFailed = 'Falló hash de ImportProvider: {0}. Error={1}'
        ImportProviderInaccessibleFolder = 'Carpeta inaccesible de ImportProvider omitida: {0} - {1}'
        GoogleSummary = 'Resumen ImportProvider Google Takeout: medios={0}; assetsLogicos={1}; ocurrencias={2}; albumes={3}; duplicadosInternos={4}; papelera={5}; videos={6}; importables={7}; conflictos={8}; confianzaAlta={9}; confianzaMedia={10}; confianzaBaja={11}; exifLeidos={12}; exifOmitidosPorConfianzaProvider={13}; sidecarsUsados={14}; sidecarsAmbiguos={15}; jsonHuerfanos={16}; mediaSinJson={17}; copiados={18}; reporte={19}'
        AppleSummary = 'Resumen ImportProvider Apple Photos / iCloud: medios={0}; assetsLogicos={1}; albumesCsv={2}; referenciasAlbum={3}; papelera={4}; videos={5}; livePhotoCandidatos={6}; importables={7}; conflictos={8}; confianzaAlta={9}; confianzaMedia={10}; confianzaBaja={11}; exifLeidos={12}; exifOmitidosPorConfianzaProvider={13}; mediaSinDetalles={14}; copiados={15}; reporte={16}'
        XmpSummary = 'Resumen ImportProvider XMP / Sidecar Library: medios={0}; sidecars={1}; sidecarsUsados={2}; sidecarsHuerfanos={3}; mediaSinSidecar={4}; sidecarsAmbiguos={5}; fallbackClasico={6}; importables={7}; conflictos={8}; confianzaAlta={9}; confianzaMedia={10}; confianzaBaja={11}; exifLeidos={12}; exifOmitidosPorConfianzaProvider={13}; copiados={14}; reporte={15}'
    }
    ro = @{
        Mode = 'Mod'
        MediaFiles = 'Fișiere media'
        LogicalAssets = 'Asset-uri logice'
        Occurrences = 'Apariții'
        Albums = 'Albume'
        InternalDuplicateOccurrences = 'Duplicate fizice interne'
        TrashOccurrences = 'Apariții în coș'
        Videos = 'Video'
        Conflicts = 'Conflicte'
        HighConfidence = 'Încredere mare'
        MediumConfidence = 'Încredere medie'
        LowConfidence = 'Încredere mică'
        ExifReadsSkipped = 'Citiri EXIF omise'
        ExifReadsSkippedVideo = 'Video omise prin încredere provider'
        TakeoutSourceDeletion = 'Ștergere sursă ImportProvider'
        AssetsFirst500 = 'Asset-uri (primele 500)'
        Hash = 'Hash'
        Status = 'Stare'
        Path = 'Cale'
        OccurrenceCount = 'Apariții'
        AlbumColumn = 'Albume'
        DateSource = 'Sursa datei'
        Confidence = 'Încredere'
        ExifVerification = 'Verificare EXIF'
        EmbeddedDateState = 'Stare dată încorporată'
        Warnings = 'Avertismente'
        SidecarsFound = 'Sidecar-uri găsite'
        SidecarsUsed = 'Sidecar-uri folosite'
        OrphanSidecars = 'Sidecar-uri orfane'
        MediaWithoutSidecar = 'Media fără sidecar'
        AmbiguousSidecars = 'Sidecar-uri ambigue'
        ClassicFallbackAssets = 'Asset-uri cu fallback clasic'
        Copied = 'Copiate'
        Sidecar = 'Sidecar'
        SidecarFields = 'Câmpuri sidecar'
        ImportStarted = 'ImportProvider {0} pornit. Mod: {1}. Rădăcină={2}'
        GoogleScan = 'Scan ImportProvider Google Takeout: media={0}; sidecar-uri={1}; albume={2}; jsonRădăcină={3}'
        AppleScan = 'Scan ImportProvider Apple Photos / iCloud: media={0}; detaliiCsv={1}; albumeCsv={2}; coș={3}; candidațiLivePhoto={4}'
        XmpScan = 'Scan ImportProvider XMP / Sidecar Library: media={0}; sidecar-uri={1}; metadataFolder={2}'
        ExifPlan = 'Plan verificare EXIF ImportProvider {0}: citire={1}; omisePrinÎncredereProvider={2}; diagnostic={3}'
        ExifPlanVideoTrusted = 'Plan verificare EXIF ImportProvider {0}: video omise prin încredere provider={1}'
        ExifVerificationStart = 'Verificare metadata EXIF/QuickTime: 0 / {0} (0%). Viteză: se calculează. Trecut: 0m. Rămas estimat: necunoscut.'
        ExifVerificationProgress = 'Verificare metadata EXIF/QuickTime: {0} / {1} ({2}%). Viteză: {3}/min. Trecut: {4}. Rămas estimat: {5}. Fișier curent: {6}'
        ExifVerificationComplete = 'Verificare EXIF/QuickTime completă: {0} / {1} (100%). Trecut: {2}.'
        ReportWritten = 'Raport ImportProvider scris: {0}'
        JsonReportWritten = 'Raport JSON ImportProvider scris: {0}'
        SourceNotDeleted = 'Sursa ImportProvider nu a fost ștearsă. Motiv=Nesolicitat. Cale={0}'
        SourceDeletionSkippedDryRun = 'Ștergerea sursei ImportProvider omisă. Motiv=DryRun. Cale={0}'
        SourceDeletionSkippedErrors = 'Ștergerea sursei ImportProvider omisă. Motiv=Importul s-a terminat cu erori; Erori={0}. Cale={1}'
        SourceDeletionSkippedMissing = 'Ștergerea sursei ImportProvider omisă. Motiv=Calea selectată nu mai există. Cale={0}'
        SourceDeletionBlockedRoot = 'Ștergerea sursei ImportProvider blocată. Motiv=Calea selectată este rădăcina unui disc. Cale={0}'
        SourceDeletionBlockedReportInsideSource = 'Ștergerea sursei ImportProvider blocată. Motiv=Folderul de loguri/rapoarte este în sursa selectată; se păstrează evidența. Cale={0}; LogRoot={1}'
        SourceDeleted = 'Sursa ImportProvider ștearsă după import reușit: {0}'
        SourceDeletionFailed = 'Ștergerea sursei ImportProvider a eșuat: {0}. Eroare={1}'
        SourceDeletionStatusLog = 'Stare ștergere sursă ImportProvider: {0}. Cale={1}. Raport={2}'
        WouldImportAsset = 'Asset-ul providerului ar fi importat: {0} -> {1}'
        ImportProviderHashFailed = 'Hash ImportProvider eșuat: {0}. Eroare={1}'
        ImportProviderInaccessibleFolder = 'Folder inaccesibil ImportProvider omis: {0} - {1}'
        GoogleSummary = 'Rezumat ImportProvider Google Takeout: media={0}; asset-uriLogice={1}; apariții={2}; albume={3}; duplicateInterne={4}; coș={5}; video={6}; importabile={7}; conflicte={8}; încredereMare={9}; încredereMedie={10}; încredereMică={11}; exifCitite={12}; exifOmisePrinÎncredereProvider={13}; sidecar-uriFolosite={14}; sidecar-uriAmbigue={15}; jsonOrfane={16}; mediaFărăJson={17}; copiate={18}; raport={19}'
        AppleSummary = 'Rezumat ImportProvider Apple Photos / iCloud: media={0}; asset-uriLogice={1}; albumeCsv={2}; referințeAlbum={3}; coș={4}; video={5}; candidațiLivePhoto={6}; importabile={7}; conflicte={8}; încredereMare={9}; încredereMedie={10}; încredereMică={11}; exifCitite={12}; exifOmisePrinÎncredereProvider={13}; mediaFărăDetalii={14}; copiate={15}; raport={16}'
        XmpSummary = 'Rezumat ImportProvider XMP / Sidecar Library: media={0}; sidecar-uri={1}; sidecar-uriFolosite={2}; sidecar-uriOrfane={3}; mediaFărăSidecar={4}; sidecar-uriAmbigue={5}; fallbackClasic={6}; importabile={7}; conflicte={8}; încredereMare={9}; încredereMedie={10}; încredereMică={11}; exifCitite={12}; exifOmisePrinÎncredereProvider={13}; copiate={14}; raport={15}'
    }
    en = @{
        Mode = 'Mode'
        MediaFiles = 'Media files'
        LogicalAssets = 'Logical assets'
        Occurrences = 'Occurrences'
        Albums = 'Albums'
        InternalDuplicateOccurrences = 'Internal duplicate occurrences'
        TrashOccurrences = 'Trash occurrences'
        Videos = 'Videos'
        Conflicts = 'Conflicts'
        HighConfidence = 'High confidence'
        MediumConfidence = 'Medium confidence'
        LowConfidence = 'Low confidence'
        ExifReadsSkipped = 'EXIF reads skipped'
        ExifReadsSkippedVideo = 'Videos skipped by provider trust'
        TakeoutSourceDeletion = 'ImportProvider source deletion'
        AssetsFirst500 = 'Assets (first 500)'
        Hash = 'Hash'
        Status = 'Status'
        Path = 'Path'
        OccurrenceCount = 'Occurrences'
        AlbumColumn = 'Albums'
        DateSource = 'Date source'
        Confidence = 'Confidence'
        ExifVerification = 'EXIF verification'
        EmbeddedDateState = 'Embedded date state'
        Warnings = 'Warnings'
        SidecarsFound = 'Sidecars found'
        SidecarsUsed = 'Sidecars used'
        OrphanSidecars = 'Orphan sidecars'
        MediaWithoutSidecar = 'Media without sidecar'
        AmbiguousSidecars = 'Ambiguous sidecars'
        ClassicFallbackAssets = 'Classic fallback assets'
        Copied = 'Copied'
        Sidecar = 'Sidecar'
        SidecarFields = 'Sidecar fields'
        ImportStarted = 'ImportProvider {0} started. Mode: {1}. Root={2}'
        GoogleScan = 'ImportProvider Google Takeout scan: media={0}; sidecars={1}; albums={2}; rootJson={3}'
        AppleScan = 'ImportProvider Apple Photos / iCloud scan: media={0}; detailsCsv={1}; albumCsv={2}; trash={3}; livePhotoCandidates={4}'
        XmpScan = 'ImportProvider XMP / Sidecar Library scan: media={0}; sidecars={1}; folderMetadata={2}'
        ExifPlan = 'ImportProvider {0} EXIF verification plan: read={1}; skippedProviderTrusted={2}; diagnostic={3}'
        ExifPlanVideoTrusted = 'ImportProvider {0} EXIF verification plan: videosSkippedProviderTrusted={1}'
        ExifVerificationStart = 'Verifying EXIF/QuickTime metadata: 0 / {0} (0%). Speed: calculating. Elapsed: 0m. Estimated remaining: unknown.'
        ExifVerificationProgress = 'Verifying EXIF/QuickTime metadata: {0} / {1} ({2}%). Speed: {3}/min. Elapsed: {4}. Estimated remaining: {5}. Current file: {6}'
        ExifVerificationComplete = 'EXIF/QuickTime verification complete: {0} / {1} (100%). Elapsed: {2}.'
        ReportWritten = 'ImportProvider report written: {0}'
        JsonReportWritten = 'ImportProvider JSON report written: {0}'
        SourceNotDeleted = 'ImportProvider source not deleted. Reason=Not requested. Path={0}'
        SourceDeletionSkippedDryRun = 'ImportProvider source deletion skipped. Reason=DryRun. Path={0}'
        SourceDeletionSkippedErrors = 'ImportProvider source deletion skipped. Reason=Import completed with errors; Errors={0}. Path={1}'
        SourceDeletionSkippedMissing = 'ImportProvider source deletion skipped. Reason=Selected path no longer exists. Path={0}'
        SourceDeletionBlockedRoot = 'ImportProvider source deletion blocked. Reason=Selected path is a filesystem root. Path={0}'
        SourceDeletionBlockedReportInsideSource = 'ImportProvider source deletion blocked. Reason=Log/report folder is inside the selected source; preserving evidence. Path={0}; LogRoot={1}'
        SourceDeleted = 'ImportProvider source deleted after successful import: {0}'
        SourceDeletionFailed = 'ImportProvider source deletion failed: {0}. Error={1}'
        SourceDeletionStatusLog = 'ImportProvider source deletion status: {0}. Path={1}. Report={2}'
        WouldImportAsset = 'Would import provider asset: {0} -> {1}'
        ImportProviderHashFailed = 'ImportProvider hash failed: {0}. Error={1}'
        ImportProviderInaccessibleFolder = 'ImportProvider inaccessible folder skipped: {0} - {1}'
        GoogleSummary = 'ImportProvider Google Takeout summary: media={0}; logicalAssets={1}; occurrences={2}; albums={3}; internalDuplicateOccurrences={4}; trashOccurrences={5}; videos={6}; importable={7}; conflicts={8}; confidenceHigh={9}; confidenceMedium={10}; confidenceLow={11}; exifRead={12}; exifSkippedProviderTrusted={13}; sidecarsUsed={14}; sidecarsAmbiguous={15}; jsonOrphans={16}; mediaWithoutJson={17}; copied={18}; report={19}'
        AppleSummary = 'ImportProvider Apple Photos / iCloud summary: media={0}; logicalAssets={1}; albumCsv={2}; albumReferences={3}; trash={4}; videos={5}; livePhotoCandidates={6}; importable={7}; conflicts={8}; confidenceHigh={9}; confidenceMedium={10}; confidenceLow={11}; exifRead={12}; exifSkippedProviderTrusted={13}; mediaWithoutDetails={14}; copied={15}; report={16}'
        XmpSummary = 'ImportProvider XMP / Sidecar Library summary: media={0}; sidecars={1}; sidecarsUsed={2}; orphanSidecars={3}; mediaWithoutSidecar={4}; ambiguousSidecars={5}; fallbackClassic={6}; importable={7}; conflicts={8}; confidenceHigh={9}; confidenceMedium={10}; confidenceLow={11}; exifRead={12}; exifSkippedProviderTrusted={13}; copied={14}; report={15}'
    }
}

function Get-ImportProviderText {
    param([string]$Key)

    $resources = if ($ImportProviderTextByLanguage.ContainsKey($Language)) { $ImportProviderTextByLanguage[$Language] } else { $ImportProviderTextByLanguage.es }
    if ($resources.ContainsKey($Key)) { return [string]$resources[$Key] }
    if ($ImportProviderTextByLanguage.en.ContainsKey($Key)) { return [string]$ImportProviderTextByLanguage.en[$Key] }
    return $Key
}

function Write-ImportProviderExifVerificationProgress {
    param(
        [int]$Current,
        [int]$Total,
        [datetime]$StartedAt,
        [string]$CurrentFile = '',
        [switch]$Force
    )

    if ($Total -le 0) { return }
    $now = Get-Date
    if (-not $script:ImportProviderExifProgress) {
        $script:ImportProviderExifProgress = [ordered]@{
            LastLogAt = [datetime]::MinValue
            LastLoggedCurrent = -1
        }
    }

    $lastLogAt = [datetime]$script:ImportProviderExifProgress.LastLogAt
    $lastCurrent = [int]$script:ImportProviderExifProgress.LastLoggedCurrent
    $secondsSinceLog = if ($lastLogAt -eq [datetime]::MinValue) { [double]::PositiveInfinity } else { ($now - $lastLogAt).TotalSeconds }
    $itemsSinceLog = [math]::Abs($Current - $lastCurrent)
    if ($Force -and $Current -ge $Total -and $lastCurrent -eq $Current -and $Current -gt 0) {
        return
    }
    $shouldLog = $Force -or $lastCurrent -lt 0 -or $itemsSinceLog -ge 100 -or $secondsSinceLog -ge 15 -or $Current -ge $Total
    if (-not $shouldLog) { return }

    $elapsed = $now - $StartedAt
    $elapsedSeconds = [math]::Max(1, $elapsed.TotalSeconds)
    $speedPerMinute = [math]::Round(($Current / $elapsedSeconds) * 60, 1)
    $remainingText = 'unknown'
    if ($speedPerMinute -gt 0 -and $Current -gt 0) {
        $remainingSeconds = [int](([math]::Max(0, $Total - $Current) / [double]$speedPerMinute) * 60)
        $remainingText = Format-OperationalDuration -Duration ([TimeSpan]::FromSeconds($remainingSeconds))
    }

    $percent = [math]::Round(($Current / [double]$Total) * 100, 1)
    $leaf = if ([string]::IsNullOrWhiteSpace($CurrentFile)) { '' } else { Split-Path -Leaf $CurrentFile }
    $message = if ($Current -le 0) {
        (Get-ImportProviderText -Key 'ExifVerificationStart') -f $Total
    }
    elseif ($Current -ge $Total) {
        (Get-ImportProviderText -Key 'ExifVerificationComplete') -f $Current, $Total, (Format-OperationalDuration -Duration $elapsed)
    }
    else {
        (Get-ImportProviderText -Key 'ExifVerificationProgress') -f $Current, $Total, $percent, $speedPerMinute, (Format-OperationalDuration -Duration $elapsed), $remainingText, $leaf
    }

    $script:ImportProviderExifProgress.LastLogAt = $now
    $script:ImportProviderExifProgress.LastLoggedCurrent = $Current
    $script:OperationalProgress = [ordered]@{
        Name = 'ImportProvider EXIF verification'
        Total = $Total
        Current = $Current
        StartedAt = $StartedAt
        LastLogAt = $now
        LastLoggedCurrent = $Current
        WarmupItems = 0
        WarmupMinutes = 0
        WarmupCompletedAt = $StartedAt
        WarmupCompletedCurrent = 0
        CurrentStage = 'EXIF/QuickTime metadata verification'
    }
    Write-Log -Message $message -Phase 'ImportProvider EXIF verification'
}

function Write-SummaryLine {
    param(
        [string]$Key,
        [object]$Value
    )

    Write-Log -Message ('{0}: {1}' -f (Get-SummaryText -Key $Key), $Value) -Phase 'Complete'
}

function Write-SummaryHostLine {
    param(
        [string]$Key,
        [object]$Value
    )

    Write-Host ('{0}: {1}' -f (Get-SummaryText -Key $Key), $Value)
}

$Stats = [ordered]@{
    FilesAnalyzed = 0
    ExactDuplicatesFound = 0
    NearDuplicatesFound = 0
    ExifRepaired = 0
    CaptureDateMaterializationCandidates = 0
    CaptureMetadataWritten = 0
    FileSystemDatesSynced = 0
    DateKnownButMetadataNotWritten = 0
    FilesMoved = 0
    FilesCopied = 0
    ExistingIdenticalSkipped = 0
    NeedsReview = 0
    EmptyFoldersRemoved = 0
    JunkOnlyFoldersRemoved = 0
    JunkOnlySmallMarkerFoldersRemoved = 0
    ZombieNormalizeFoldersRemoved = 0
    FoldersReduced = 0
    JsonPathsUpdated = 0
    SkippedUncertainNames = 0
    FilesFound = 0
    LocalFilesDetected = 0
    CloudPlaceholdersSkipped = 0
    CloudPlaceholdersInIndex = 0
    MissingReal = 0
    IncrementalSkipped = 0
    SkippedOneDrive = 0
    Inaccessible = 0
    LockedFiles = 0
    RetryCount = 0
    MetadataCorruptedMedia = 0
    SlowExifCandidates = 0
    SlowExifDetections = 0
    ExifBatchTimeouts = 0
    ExifBatchTimeoutAffectedFiles = 0
    ExifBatchFallbacks = 0
    ExifBatchFallbackAffectedFiles = 0
    MetadataBackupSizeGB = 0
    RetentionDeletedItems = 0
    RetentionRecoveredGB = 0
    ConfirmedDuplicatesRevalidated = 0
    ConfirmedDuplicatesRetained = 0
    JsonReconcileValid = 0
    JsonReconcileStale = 0
    JsonReconcilePathsUpdated = 0
    JsonReconcileMissing = 0
    JsonReconcilePurged = 0
    JsonReconcileEntriesRemoved = 0
    JsonReconcileConflicts = 0
    DryRunActions = 0
    Errors = 0
}

$script:LogWriter = $null
$script:LastProgressWrite = Get-Date
$script:LastHeartbeat = Get-Date
$script:CurrentPhase = 'Starting'
$script:LastMessage = ''
$script:Cancelled = $false
$script:ConfirmedDuplicateQuarantineEntries = New-Object System.Collections.Generic.List[object]
$script:RunStartTime = Get-Date
$script:RunId = $script:RunStartTime.ToString('yyyyMMdd-HHmmss')
$script:CurrentBatch = 0
$script:TotalBatches = 0
$script:FilesPerSecond = 0.0
$script:EtaText = ''
$script:ProcessorCount = [Environment]::ProcessorCount
$script:WorkerCount = 1
$script:ActiveWorkers = 0
$script:QueueSize = 0
$script:CloudPlaceholdersInIndexKeys = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
$script:OperationalProgress = $null
$script:CpuPercent = 0
$script:RamPercent = 0
$script:LastResourceSample = [datetime]::MinValue
$script:LastProcessCpu = [TimeSpan]::Zero
$script:LastProcessCpuSample = Get-Date
$script:PowerShellRuntime = ('PowerShell {0}' -f $PSVersionTable.PSVersion.ToString())
$script:DriveKind = 'Unknown'
$script:ExifProblemFiles = @{}
$script:MetadataCorruptionFiles = @{}
$script:SlowExifCandidates = @{}
$script:HashHelperInitialized = $false

function Initialize-Logging {
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        $logDirectory = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($LogPath))
        if (-not [string]::IsNullOrWhiteSpace($logDirectory) -and -not (Test-Path -LiteralPath $logDirectory)) {
            New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
        }

        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $script:LogWriter = New-Object System.IO.StreamWriter -ArgumentList $LogPath, $true, $utf8NoBom
        $script:LogWriter.AutoFlush = $true
    }
}

function Close-Logging {
    if ($script:LogWriter) {
        $script:LogWriter.Flush()
        $script:LogWriter.Dispose()
        $script:LogWriter = $null
    }
}

function Write-ProgressState {
    param(
        [string]$Phase = $script:CurrentPhase,
        [string]$Message = $script:LastMessage,
        [string]$Status = 'Running'
    )

    $script:CurrentPhase = $Phase
    $script:LastMessage = $Message
    $script:LastProgressWrite = Get-Date

    if ([string]::IsNullOrWhiteSpace($ProgressPath)) {
        return
    }

    try {
        Update-ResourceMetrics

        $progressDirectory = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($ProgressPath))
        if (-not [string]::IsNullOrWhiteSpace($progressDirectory) -and -not (Test-Path -LiteralPath $progressDirectory)) {
            New-Item -ItemType Directory -Path $progressDirectory -Force | Out-Null
        }

        $state = [ordered]@{
            pid = $PID
            status = $Status
            phase = $Phase
            message = $Message
            timestamp = (Get-Date).ToString('o')
            filesFound = $Stats.FilesFound
            localFilesDetected = $Stats.LocalFilesDetected
            cloudPlaceholdersSkipped = $Stats.CloudPlaceholdersSkipped
            cloudPlaceholdersInIndex = $Stats.CloudPlaceholdersInIndex
            missingReal = $Stats.MissingReal
            filesAnalyzed = $Stats.FilesAnalyzed
            incrementalSkipped = $Stats.IncrementalSkipped
            filesPerSecond = $script:FilesPerSecond
            eta = $script:EtaText
            elapsedSeconds = [int]((Get-Date) - $script:RunStartTime).TotalSeconds
            currentBatch = $script:CurrentBatch
            totalBatches = $script:TotalBatches
            operationalName = if ($script:OperationalProgress) { $script:OperationalProgress.Name } else { '' }
            operationalCurrent = if ($script:OperationalProgress) { [int]$script:OperationalProgress.Current } else { 0 }
            operationalTotal = if ($script:OperationalProgress) { [int]$script:OperationalProgress.Total } else { 0 }
            operationalStage = if ($script:OperationalProgress) { [string]$script:OperationalProgress.CurrentStage } else { '' }
            performanceMode = $PerformanceMode
            processorCount = $script:ProcessorCount
            workerCount = $script:WorkerCount
            activeWorkers = $script:ActiveWorkers
            queueSize = $script:QueueSize
            cpuPercent = $script:CpuPercent
            ramPercent = $script:RamPercent
            powershellRuntime = $script:PowerShellRuntime
            driveKind = $script:DriveKind
            duplicatesFound = ($Stats.ExactDuplicatesFound + $Stats.NearDuplicatesFound)
            exactDuplicatesFound = $Stats.ExactDuplicatesFound
            nearDuplicatesFound = $Stats.NearDuplicatesFound
            exifRepaired = $Stats.ExifRepaired
            filesMoved = $Stats.FilesMoved
            filesCopied = $Stats.FilesCopied
            needsReview = $Stats.NeedsReview
            emptyFoldersRemoved = $Stats.EmptyFoldersRemoved
            skippedOneDrive = $Stats.SkippedOneDrive
            inaccessible = $Stats.Inaccessible
            lockedFiles = $Stats.LockedFiles
            retryCount = $Stats.RetryCount
            metadataCorruptedMedia = $Stats.MetadataCorruptedMedia
            slowExifCandidates = $Stats.SlowExifCandidates
            slowExifDetections = $Stats.SlowExifDetections
            exifBatchTimeouts = $Stats.ExifBatchTimeouts
            exifBatchTimeoutAffectedFiles = $Stats.ExifBatchTimeoutAffectedFiles
            exifBatchFallbacks = $Stats.ExifBatchFallbacks
            exifBatchFallbackAffectedFiles = $Stats.ExifBatchFallbackAffectedFiles
            metadataBackupSizeGb = $Stats.MetadataBackupSizeGB
            retentionDeletedItems = $Stats.RetentionDeletedItems
            retentionRecoveredGb = $Stats.RetentionRecoveredGB
            errors = $Stats.Errors
            logPath = $LogPath
        }

        $tmp = $ProgressPath + '.tmp'
        $state | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $tmp -Encoding UTF8
        Move-Item -LiteralPath $tmp -Destination $ProgressPath -Force
    }
    catch {
    }
}

function Write-Log {
    param(
        [string]$Message,
        [string]$Phase = $script:CurrentPhase,
        [string]$Status = 'Running'
    )

    $line = '{0} [{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Phase, $Message
    if ($script:LogWriter) {
        $script:LogWriter.WriteLine($line)
        $script:LogWriter.Flush()
    }
    Write-Host $Message
    Write-ProgressState -Phase $Phase -Message $Message -Status $Status
}

function Write-DiagnosticLog {
    param([string]$Message)
    if ($Diagnostic) {
        Write-Log -Message ("DIAGNOSTIC: " + $Message) -Phase $script:CurrentPhase
    }
}

function Write-Heartbeat {
    param(
        [string]$Phase,
        [string]$Message,
        [int]$EverySeconds = 10
    )

    if (((Get-Date) - $script:LastHeartbeat).TotalSeconds -ge $EverySeconds) {
        $script:LastHeartbeat = Get-Date
        Write-Log -Message $Message -Phase $Phase
    }
}

function Update-Throughput {
    $elapsed = [math]::Max(1, ((Get-Date) - $script:RunStartTime).TotalSeconds)
    $script:FilesPerSecond = [math]::Round($Stats.FilesAnalyzed / $elapsed, 2)
    $remaining = [math]::Max(0, $Stats.FilesFound - $Stats.FilesAnalyzed)
    if ($script:FilesPerSecond -gt 0) {
        $etaSeconds = [int]($remaining / $script:FilesPerSecond)
        $script:EtaText = ([TimeSpan]::FromSeconds($etaSeconds)).ToString()
    }
    else {
        $script:EtaText = ''
    }
}

function Format-OperationalDuration {
    param([TimeSpan]$Duration)

    $totalMinutes = [int][math]::Max(0, [math]::Round($Duration.TotalMinutes))
    $hours = [int][math]::Floor($totalMinutes / 60)
    $minutes = [int]($totalMinutes % 60)
    if ($hours -gt 0) {
        return ('{0}h {1}m' -f $hours, $minutes)
    }
    return ('{0}m' -f $minutes)
}

function Start-OperationalProgress {
    param(
        [string]$Name,
        [int]$Total,
        [string]$Phase,
        [string]$Message = '',
        [int]$WarmupItems = 100,
        [int]$WarmupMinutes = 2
    )

    $script:OperationalProgress = [ordered]@{
        Name = $Name
        Total = [math]::Max(0, $Total)
        Current = 0
        StartedAt = Get-Date
        LastLogAt = [datetime]::MinValue
        LastLoggedCurrent = -1
        CurrentStage = $Phase
        WarmupItems = [math]::Max(0, $WarmupItems)
        WarmupMinutes = [math]::Max(0, $WarmupMinutes)
        WarmupCompletedAt = $null
        WarmupCompletedCurrent = 0
    }

    $startMessage = if ([string]::IsNullOrWhiteSpace($Message)) {
        ('{0} progress started. Total items: {1}' -f $Name, $Total)
    }
    else {
        ('{0} progress started. Total items: {1}. {2}' -f $Name, $Total, $Message)
    }
    Write-Log -Message $startMessage -Phase $Phase
}

function Write-OperationalStage {
    param(
        [string]$Name,
        [string]$Stage,
        [string]$Phase
    )

    if ($script:OperationalProgress) {
        $script:OperationalProgress.CurrentStage = $Stage
    }
    Write-Log -Message ('{0} stage: {1}' -f $Name, $Stage) -Phase $Phase
}

function Update-OperationalProgress {
    param(
        [int]$Current,
        [int]$Total = -1,
        [string]$Phase = $script:CurrentPhase,
        [string]$Stage = '',
        [int]$EveryItems = 500,
        [int]$EveryMinutes = 5,
        [switch]$Force
    )

    if (-not $script:OperationalProgress) {
        return
    }

    if ($Total -ge 0) {
        $script:OperationalProgress.Total = [math]::Max(0, $Total)
    }
    $script:OperationalProgress.Current = [math]::Max(0, $Current)
    if (-not [string]::IsNullOrWhiteSpace($Stage)) {
        $script:OperationalProgress.CurrentStage = $Stage
    }

    $now = Get-Date
    $lastLogAt = [datetime]$script:OperationalProgress.LastLogAt
    $lastCurrent = [int]$script:OperationalProgress.LastLoggedCurrent
    $itemsSinceLog = [math]::Abs($Current - $lastCurrent)
    $minutesSinceLog = if ($lastLogAt -eq [datetime]::MinValue) { [double]::PositiveInfinity } else { ($now - $lastLogAt).TotalMinutes }
    $shouldLog = $Force -or $lastCurrent -lt 0 -or $itemsSinceLog -ge $EveryItems -or $minutesSinceLog -ge $EveryMinutes
    if (-not $shouldLog) {
        return
    }

    $totalItems = [int]$script:OperationalProgress.Total
    $elapsed = $now - ([datetime]$script:OperationalProgress.StartedAt)
    $elapsedSeconds = [math]::Max(1, $elapsed.TotalSeconds)
    $percent = if ($totalItems -gt 0) { [math]::Round(($Current / [double]$totalItems) * 100, 1) } else { 0 }
    $warmupItems = [int]$script:OperationalProgress.WarmupItems
    $warmupMinutes = [int]$script:OperationalProgress.WarmupMinutes
    $isWarmedUp = ($Current -ge $warmupItems) -and ($elapsed.TotalMinutes -ge $warmupMinutes)
    if ($Force -and $totalItems -gt 0 -and $Current -ge $totalItems) {
        $isWarmedUp = $true
    }

    if ($isWarmedUp -and $null -eq $script:OperationalProgress.WarmupCompletedAt) {
        $script:OperationalProgress.WarmupCompletedAt = $now
        $script:OperationalProgress.WarmupCompletedCurrent = $Current
    }

    $speedText = 'warming up'
    $remainingText = 'unknown'
    $itemsPerHour = 0
    if ($isWarmedUp) {
        $speedStart = if ($script:OperationalProgress.WarmupCompletedAt) { [datetime]$script:OperationalProgress.WarmupCompletedAt } else { [datetime]$script:OperationalProgress.StartedAt }
        $speedStartCurrent = if ($script:OperationalProgress.WarmupCompletedAt) { [int]$script:OperationalProgress.WarmupCompletedCurrent } else { 0 }
        $speedElapsedSeconds = [math]::Max(1, ($now - $speedStart).TotalSeconds)
        $effectiveItems = [math]::Max(0, $Current - $speedStartCurrent)
        if ($effectiveItems -le 0 -and $Force) {
            $effectiveItems = $Current
            $speedElapsedSeconds = $elapsedSeconds
        }
        $itemsPerHour = [math]::Round(($effectiveItems / $speedElapsedSeconds) * 3600, 0)
        if ($itemsPerHour -le 0 -and $Current -gt 0) {
            $itemsPerHour = [math]::Max(1, [math]::Round(($Current / $elapsedSeconds) * 3600, 0))
        }
        $speedText = ('{0} items/h' -f $itemsPerHour)
    }
    else {
        $remainingText = ('stabilizing after warmup ({0} items and {1} min)' -f $warmupItems, $warmupMinutes)
    }

    if ($itemsPerHour -gt 0 -and $totalItems -gt 0) {
        $remainingItems = [math]::Max(0, $totalItems - $Current)
        $remainingSeconds = [int](($remainingItems / [double]$itemsPerHour) * 3600)
        $remainingText = Format-OperationalDuration -Duration ([TimeSpan]::FromSeconds($remainingSeconds))
    }

    $stageText = if ($script:OperationalProgress.CurrentStage) { [string]$script:OperationalProgress.CurrentStage } else { $Phase }
    $message = ('{0} progress: {1} / {2} ({3}%). Current speed: {4}. Elapsed: {5}. Estimated remaining: {6}. Stage: {7}' -f `
        $script:OperationalProgress.Name,
        $Current,
        $totalItems,
        $percent,
        $speedText,
        (Format-OperationalDuration -Duration $elapsed),
        $remainingText,
        $stageText)

    $script:OperationalProgress.LastLogAt = $now
    $script:OperationalProgress.LastLoggedCurrent = $Current
    Write-Log -Message $message -Phase $Phase
}

function Complete-OperationalProgress {
    param(
        [string]$Phase = $script:CurrentPhase,
        [string]$Message = ''
    )

    if (-not $script:OperationalProgress) {
        return
    }

    $total = [int]$script:OperationalProgress.Total
    $current = if ($total -gt 0) { $total } else { [int]$script:OperationalProgress.Current }
    Update-OperationalProgress -Current $current -Total $total -Phase $Phase -Force
    if (-not [string]::IsNullOrWhiteSpace($Message)) {
        Write-Log -Message $Message -Phase $Phase
    }
    $script:OperationalProgress = $null
}

function Update-ResourceMetrics {
    if (((Get-Date) - $script:LastResourceSample).TotalSeconds -lt 5) {
        return
    }

    $sampleTime = Get-Date
    try {
        $process = [System.Diagnostics.Process]::GetCurrentProcess()
        $cpuDelta = ($process.TotalProcessorTime - $script:LastProcessCpu).TotalSeconds
        $timeDelta = [math]::Max(0.1, ($sampleTime - $script:LastProcessCpuSample).TotalSeconds)
        $script:CpuPercent = [int][math]::Min(100, [math]::Round(($cpuDelta / ($timeDelta * [math]::Max(1, $script:ProcessorCount))) * 100))
        $script:LastProcessCpu = $process.TotalProcessorTime
        $script:LastProcessCpuSample = $sampleTime
    }
    catch {
    }

    $script:LastResourceSample = $sampleTime
    try {
        $processor = Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop | Select-Object -First 1
        if ($processor -and $processor.LoadPercentage -ne $null) {
            $script:CpuPercent = [int]$processor.LoadPercentage
        }
    }
    catch {
        try {
            $counter = Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Stop
            $script:CpuPercent = [int][math]::Round($counter.CounterSamples[0].CookedValue)
        }
        catch {
            try {
                $processor = Get-WmiObject -Class Win32_Processor -ErrorAction Stop | Select-Object -First 1
                if ($processor -and $processor.LoadPercentage -ne $null) {
                    $script:CpuPercent = [int]$processor.LoadPercentage
                }
            }
            catch {
            }
        }
    }

    if ($script:CpuPercent -lt 0) {
        $script:CpuPercent = 0
    }

    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        if ($os.TotalVisibleMemorySize -gt 0) {
            $used = $os.TotalVisibleMemorySize - $os.FreePhysicalMemory
            $script:RamPercent = [int][math]::Round(($used / [double]$os.TotalVisibleMemorySize) * 100)
        }
    }
    catch {
        try {
            $os = Get-WmiObject -Class Win32_OperatingSystem -ErrorAction Stop
            if ($os.TotalVisibleMemorySize -gt 0) {
                $used = $os.TotalVisibleMemorySize - $os.FreePhysicalMemory
                $script:RamPercent = [int][math]::Round(($used / [double]$os.TotalVisibleMemorySize) * 100)
            }
        }
        catch {
        }
    }

    if ($script:RamPercent -le 0) {
        try {
            Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop
            $computerInfo = New-Object Microsoft.VisualBasic.Devices.ComputerInfo
            if ($computerInfo.TotalPhysicalMemory -gt 0) {
                $usedBytes = $computerInfo.TotalPhysicalMemory - $computerInfo.AvailablePhysicalMemory
                $script:RamPercent = [int][math]::Round(($usedBytes / [double]$computerInfo.TotalPhysicalMemory) * 100)
            }
        }
        catch {
        }
    }
}

function Resolve-WorkerCount {
    $threads = [math]::Max(1, [Environment]::ProcessorCount)
    $pct = switch ($PerformanceMode) {
        'Safe' { 0.25 }
        'HighPerformance' { 0.75 }
        default { 0.50 }
    }

    $workers = [int][math]::Floor($threads * $pct)
    $workers = [math]::Max(1, $workers)
    if ($MaxParallelJobs -gt 0) {
        $workers = [math]::Min($workers, $MaxParallelJobs)
    }

    return [math]::Min($workers, $threads)
}

function Resolve-DriveKind {
    param([string]$Path)

    try {
        $root = [System.IO.Path]::GetPathRoot((Resolve-FullPath $Path))
        if ([string]::IsNullOrWhiteSpace($root)) { return 'Unknown' }
        $driveLetter = $root.TrimEnd('\').TrimEnd(':')
        $partition = Get-CimInstance -ClassName Win32_LogicalDiskToPartition -ErrorAction Stop |
            Where-Object { $_.Dependent -match ('DeviceID="' + [regex]::Escape($driveLetter + ':') + '"') } |
            Select-Object -First 1
        if (-not $partition) { return 'Local' }
        $diskIndex = [regex]::Match($partition.Antecedent, 'Disk #(\d+)').Groups[1].Value
        if ([string]::IsNullOrWhiteSpace($diskIndex)) { return 'Local' }
        $physical = Get-PhysicalDisk -ErrorAction Stop | Select-Object -First 1
        if ($physical -and $physical.MediaType) {
            return [string]$physical.MediaType
        }
    }
    catch {
    }

    return 'Local/Unknown'
}

function Initialize-HashHelper {
    if ($script:HashHelperInitialized) { return }

    $source = @"
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;

public class PhotoOrganizerHashResult {
    public string Path;
    public string Hash;
    public string Error;
}

public static class PhotoOrganizerHasher {
    public static PhotoOrganizerHashResult[] Compute(string[] paths, int maxDegree) {
        if (maxDegree < 1) maxDegree = 1;
        var bag = new ConcurrentBag<PhotoOrganizerHashResult>();
        var options = new ParallelOptions { MaxDegreeOfParallelism = maxDegree };
        Parallel.ForEach(paths, options, path => {
            var result = new PhotoOrganizerHashResult();
            result.Path = path;
            for (int attempt = 0; attempt < 3; attempt++) {
                try {
                    using (var sha = SHA256.Create())
                    using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1024 * 1024, FileOptions.SequentialScan)) {
                        var hash = sha.ComputeHash(stream);
                        result.Hash = BitConverter.ToString(hash).Replace("-", "");
                        result.Error = null;
                        break;
                    }
                }
                catch (Exception ex) {
                    result.Error = ex.Message;
                    Thread.Sleep(250 * (attempt + 1));
                }
            }
            bag.Add(result);
        });
        return bag.ToArray();
    }
}
"@

    try {
        Add-Type -TypeDefinition $source -ErrorAction Stop
        $script:HashHelperInitialized = $true
    }
    catch {
        Write-Log -Message "Parallel hash helper unavailable, falling back to sequential hash: $($_.Exception.Message)" -Phase 'Validation'
        $script:HashHelperInitialized = $false
    }
}

function Get-Sha256Batch {
    param([object[]]$Files)

    $results = @{}
    if (-not $Files -or $Files.Count -eq 0) {
        return $results
    }

    $localFiles = New-Object System.Collections.Generic.List[object]
    foreach ($file in $Files) {
        $availability = Detect-StorageAvailability -Item $file
        if ($availability.State -eq 'CloudPlaceholder') {
            Register-CloudPlaceholderSkipped -Path $file.FullName -Phase 'Hash queue' -Availability $availability
            $results[$file.FullName] = [pscustomobject]@{
                Hash = $null
                Error = 'CloudPlaceholder'
            }
            continue
        }
        if ($availability.State -eq 'MissingReal') {
            $Stats.MissingReal++
            $results[$file.FullName] = [pscustomobject]@{
                Hash = $null
                Error = 'MissingReal'
            }
            continue
        }
        $localFiles.Add($file)
    }

    if ($localFiles.Count -eq 0) {
        return $results
    }

    $Files = @($localFiles.ToArray())

    Initialize-HashHelper
    if ($script:HashHelperInitialized) {
        $paths = [string[]]@($Files | ForEach-Object { $_.FullName })
        $script:ActiveWorkers = [math]::Min($script:WorkerCount, $paths.Count)
        try {
            $hashRows = [PhotoOrganizerHasher]::Compute($paths, $script:WorkerCount)
            foreach ($row in $hashRows) {
                $results[$row.Path] = [pscustomobject]@{
                    Hash = $row.Hash
                    Error = $row.Error
                }
            }
            return $results
        }
        finally {
            $script:ActiveWorkers = 0
        }
    }

    foreach ($file in $Files) {
        try {
            $results[$file.FullName] = [pscustomobject]@{
                Hash = Get-Sha256 -Path $file.FullName
                Error = $null
            }
        }
        catch {
            $results[$file.FullName] = [pscustomobject]@{
                Hash = $null
                Error = $_.Exception.Message
            }
        }
    }

    return $results
}

function Stop-WithError {
    param([string]$Message)
    Write-Log -Message $Message -Phase 'Error' -Status 'Error'
    throw $Message
}

function Write-Notice {
    param([string]$Message)
    Write-Log -Message $Message
}

trap {
    try {
        Write-Log -Message "Fatal error: $($_.Exception.Message)" -Phase 'Error' -Status 'Error'
        if ($Diagnostic) {
            Write-Log -Message "Fatal stack: $($_.ScriptStackTrace)" -Phase 'Error' -Status 'Error'
        }
    }
    catch {
    }
    Close-Logging
    exit 1
}

function ConvertTo-CommandArgument {
    param([string]$Value)
    if ($null -eq $Value) { return '""' }
    return '"' + ($Value -replace '"', '\"') + '"'
}

function Test-ExifProblemText {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $false
    }

    return ($Text -match '(?i)unknown trailer|truncated data|possible garbage at end of file|garbage at end|truncated|premature end|corrupt|invalid atom|bad atom|invalid chunk|missing.*moov|moov atom not found|end of file')
}

function Register-ExifProblemFile {
    param(
        [string]$Path,
        [string]$Reason,
        [string]$Detail = ''
    )

    $fullPath = Resolve-FullPath $Path
    $fileName = [System.IO.Path]::GetFileName($fullPath)
    $ext = [System.IO.Path]::GetExtension($fullPath)
    $size = 0
    try { $size = (Get-Item -LiteralPath $fullPath -ErrorAction Stop).Length } catch { }

    $script:ExifProblemFiles[$fullPath] = $Reason
    if ($Reason -eq 'Media corruption / WhatsApp metadata issue' -and -not $script:MetadataCorruptionFiles.ContainsKey($fullPath)) {
        $script:MetadataCorruptionFiles[$fullPath] = $true
        $Stats.MetadataCorruptedMedia++
    }
    if ([string]::IsNullOrWhiteSpace($Detail)) {
        Write-Log -Message ("EXIF problem detected: {0}. Extension={1}. Size={2} bytes. Path={3}" -f $fileName, $ext, $size, $fullPath) -Phase 'Processing'
    }
    else {
        Write-Log -Message ("EXIF problem detected: {0}. Extension={1}. Size={2} bytes. Path={3}. Detail={4}" -f $fileName, $ext, $size, $fullPath, $Detail) -Phase 'Processing'
    }
}

function Register-SlowExifCandidate {
    param(
        [string]$Path,
        [string]$Reason,
        [string]$Detail = '',
        [double]$Seconds = 0,
        [int]$BatchNumber = 0
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return }

    $fullPath = Resolve-FullPath $Path
    $extension = [System.IO.Path]::GetExtension($fullPath)
    $size = 0
    try { $size = (Get-Item -LiteralPath $fullPath -ErrorAction Stop).Length } catch { }
    $key = $fullPath.ToLowerInvariant()

    if (-not $script:SlowExifCandidates.ContainsKey($key)) {
        $script:SlowExifCandidates[$key] = [pscustomobject]@{
            Path = $fullPath
            Extension = $extension
            SizeBytes = $size
            Detections = 0
            LastSeconds = 0
            LastBatch = 0
            Reasons = New-Object System.Collections.Generic.List[string]
        }
        $Stats.SlowExifCandidates++
    }

    $candidate = $script:SlowExifCandidates[$key]
    $candidate.Detections++
    $candidate.LastSeconds = [math]::Round([math]::Max(0, $Seconds), 2)
    $candidate.LastBatch = $BatchNumber
    if (-not $candidate.Reasons.Contains($Reason)) {
        $candidate.Reasons.Add($Reason) | Out-Null
    }
    $Stats.SlowExifDetections++

    $timeText = if ($Seconds -gt 0) { ('{0:N2} sec' -f $Seconds) } else { 'unknown' }
    $batchText = if ($BatchNumber -gt 0) { [string]$BatchNumber } else { 'n/a' }
    $detailText = if ([string]::IsNullOrWhiteSpace($Detail)) { '' } else { "; Detail=$Detail" }
    Write-Log -Message ("Slow EXIF candidate: Path={0}; Extension={1}; Size={2} bytes; Time={3}; Batch={4}; Reason={5}; Count={6}{7}" -f $fullPath, $extension, $size, $timeText, $batchText, $Reason, $candidate.Detections, $detailText) -Phase 'Processing'
}

function Register-ExifBatchTimeout {
    param(
        [int]$AffectedFiles,
        [double]$Seconds = 0,
        [int]$BatchNumber = 0,
        [string]$Detail = ''
    )

    $Stats.ExifBatchTimeouts++
    $Stats.ExifBatchTimeoutAffectedFiles += [math]::Max(0, $AffectedFiles)
    $timeText = if ($Seconds -gt 0) { ('{0:N2} sec' -f $Seconds) } else { 'unknown' }
    $batchText = if ($BatchNumber -gt 0) { [string]$BatchNumber } else { 'n/a' }
    $detailText = if ([string]::IsNullOrWhiteSpace($Detail)) { '' } else { "; Detail=$Detail" }
    Write-Log -Message ("EXIF batch timeout recorded: Batch={0}; AffectedFiles={1}; Time={2}{3}" -f $batchText, $AffectedFiles, $timeText, $detailText) -Phase 'Processing'
}

function Register-ExifBatchFallback {
    param(
        [int]$AffectedFiles,
        [double]$Seconds = 0,
        [int]$BatchNumber = 0,
        [string]$Reason = 'Batch fallback'
    )

    $Stats.ExifBatchFallbacks++
    $Stats.ExifBatchFallbackAffectedFiles += [math]::Max(0, $AffectedFiles)
    $timeText = if ($Seconds -gt 0) { ('{0:N2} sec' -f $Seconds) } else { 'unknown' }
    $batchText = if ($BatchNumber -gt 0) { [string]$BatchNumber } else { 'n/a' }
    Write-Log -Message ("EXIF batch fallback recorded: Batch={0}; AffectedFiles={1}; Time={2}; Reason={3}" -f $batchText, $AffectedFiles, $timeText, $Reason) -Phase 'Processing'
}

function Invoke-ExifTool {
    param(
        [string]$Path,
        [string[]]$Arguments,
        [int]$TimeoutSeconds = $ExifToolTimeoutSeconds
    )

    $result = [ordered]@{
        Success = $false
        TimedOut = $false
        ExitCode = $null
        Output = ''
        Error = ''
        DurationSeconds = 0
    }

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $result.Error = 'ExifTool path is empty.'
        return [pscustomobject]$result
    }

    $argFile = $null
    try {
        $argFile = [System.IO.Path]::GetTempFileName()
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllLines($argFile, @($Arguments), $utf8NoBom)

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $Path
        $effectiveArguments = @('-charset', 'filename=UTF8', '-@', $argFile)
        $psi.Arguments = (($effectiveArguments | ForEach-Object { ConvertTo-CommandArgument $_ }) -join ' ')
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        try {
            $psi.StandardOutputEncoding = [System.Text.Encoding]::UTF8
            $psi.StandardErrorEncoding = [System.Text.Encoding]::UTF8
        }
        catch {
        }
        $psi.CreateNoWindow = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi

        if (-not $process.Start()) {
            $result.Error = 'ExifTool failed to start.'
            return [pscustomobject]$result
        }

        $startedAt = Get-Date
        if (-not $process.WaitForExit([math]::Max(1, $TimeoutSeconds) * 1000)) {
            $result.DurationSeconds = [math]::Round(((Get-Date) - $startedAt).TotalSeconds, 2)
            $result.TimedOut = $true
            try { $process.Kill() } catch { }
            $result.Error = "ExifTool timeout after $TimeoutSeconds seconds."
            return [pscustomobject]$result
        }

        $result.DurationSeconds = [math]::Round(((Get-Date) - $startedAt).TotalSeconds, 2)
        $result.ExitCode = $process.ExitCode
        $result.Output = $process.StandardOutput.ReadToEnd()
        $result.Error = $process.StandardError.ReadToEnd()
        $result.Success = ($process.ExitCode -eq 0)
        return [pscustomobject]$result
    }
    catch {
        $result.Error = $_.Exception.Message
        return [pscustomobject]$result
    }
    finally {
        if ($process) {
            $process.Dispose()
        }
        if ($argFile -and (Test-Path -LiteralPath $argFile)) {
            try { Remove-Item -LiteralPath $argFile -Force } catch { }
        }
    }
}

function Test-ExifTool {
    param([string]$Path)
    $probe = Invoke-ExifTool -Path $Path -Arguments @('-ver') -TimeoutSeconds 5
    return ($probe.Success -and -not [string]::IsNullOrWhiteSpace($probe.Output))
}

function Resolve-ExifToolPath {
    param([string]$PreferredPath)

    if (-not [string]::IsNullOrWhiteSpace($PreferredPath) -and (Test-ExifTool -Path $PreferredPath)) {
        return $PreferredPath
    }

    $scriptRoot = Get-ScriptRootPath
    $installRoot = Get-InstallRootPath
    $localCandidates = @(
        (Join-Path $scriptRoot 'exiftool.exe'),
        (Join-Path $scriptRoot 'exiftool(-k).exe'),
        (Join-Path $installRoot 'Tools\ExifTool\exiftool.exe'),
        (Join-Path $installRoot 'Tools\ExifTool\exiftool(-k).exe')
    )

    try {
        $searchRoots = @($scriptRoot, (Join-Path $installRoot 'Tools\ExifTool')) | Select-Object -Unique
        foreach ($searchRoot in $searchRoots) {
            if (-not (Test-Path -LiteralPath $searchRoot -PathType Container)) { continue }
            $nested = Get-ChildItem -LiteralPath $searchRoot -Recurse -File -Filter 'exiftool*.exe' -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty FullName
            $localCandidates += @($nested)
        }
    }
    catch {
    }

    $orderedCandidates = @($localCandidates |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique |
        Sort-Object @{ Expression = { if ($_ -match '\(-k\)') { 1 } else { 0 } } }, @{ Expression = { $_ } })

    foreach ($candidate in $orderedCandidates) {
        if (Test-ExifTool -Path $candidate) {
            return $candidate
        }
    }

    return $PreferredPath
}

function Resolve-FullPath {
    param([string]$Path)
    $expanded = [Environment]::ExpandEnvironmentVariables($Path)
    return [System.IO.Path]::GetFullPath($expanded)
}

function ConvertTo-RelativePath {
    param(
        [string]$Path,
        [string]$BasePath
    )

    try {
        $fullPath = (Resolve-FullPath $Path).TrimEnd('\')
        $fullBase = (Resolve-FullPath $BasePath).TrimEnd('\')
        if ($fullPath.Equals($fullBase, [StringComparison]::OrdinalIgnoreCase)) {
            return ''
        }
        if ($fullPath.StartsWith($fullBase + '\', [StringComparison]::OrdinalIgnoreCase)) {
            return $fullPath.Substring($fullBase.Length + 1)
        }
        return $fullPath
    }
    catch {
        return $Path
    }
}

function Test-IsChildPath {
    param(
        [string]$Path,
        [string]$ParentPath
    )
    $fullPath = (Resolve-FullPath $Path).TrimEnd('\')
    $fullParent = (Resolve-FullPath $ParentPath).TrimEnd('\')
    return $fullPath.Equals($fullParent, [StringComparison]::OrdinalIgnoreCase) -or
        $fullPath.StartsWith($fullParent + '\', [StringComparison]::OrdinalIgnoreCase)
}

$script:FolderProtectionRules = @()
$script:UserExcludedRoots = @()
$script:VendorManagedRoots = @()
$script:FolderProtectionSkipCounts = @{
    UserExcluded = 0
    VendorManaged = 0
    InternalProtected = 0
}

function Expand-ConfiguredFolderPath {
    param(
        [string]$Path,
        [string]$ResolvedSourcePath,
        [string]$ResolvedDestinationPath,
        [string]$ResolvedScriptRoot
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ''
    }

    $expanded = [Environment]::ExpandEnvironmentVariables($Path)
    $destinationBasePath = ''
    try {
        if (-not [string]::IsNullOrWhiteSpace($ResolvedDestinationPath)) {
            $destinationBasePath = [System.IO.Path]::GetDirectoryName($ResolvedDestinationPath.TrimEnd('\'))
        }
    }
    catch {
        $destinationBasePath = ''
    }

    $expanded = $expanded.Replace('<SourcePath>', $ResolvedSourcePath)
    $expanded = $expanded.Replace('<DestinationPath>', $ResolvedDestinationPath)
    $expanded = $expanded.Replace('<DestinationBase>', $destinationBasePath)
    $expanded = $expanded.Replace('<AppRoot>', $ResolvedScriptRoot)
    $expanded = $expanded.Replace('<ScriptRoot>', $ResolvedScriptRoot)

    return Resolve-FullPath $expanded
}

function New-FolderProtectionRule {
    param(
        [string]$Path,
        [string]$Role,
        [string]$Label,
        [string]$Reason
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    $normalizedRole = if ([string]::IsNullOrWhiteSpace($Role)) { 'UserExcluded' } else { $Role }
    if ($normalizedRole -notin @('UserExcluded', 'VendorManaged')) {
        $normalizedRole = 'UserExcluded'
    }

    [pscustomobject]@{
        Path = (Resolve-FullPath $Path).TrimEnd('\')
        Role = $normalizedRole
        Label = if ([string]::IsNullOrWhiteSpace($Label)) { [System.IO.Path]::GetFileName($Path.TrimEnd('\')) } else { $Label }
        Reason = if ([string]::IsNullOrWhiteSpace($Reason)) { 'User protected folder' } else { $Reason }
    }
}

function Initialize-FolderProtectionRules {
    param(
        [string]$ResolvedSourcePath,
        [string]$ResolvedDestinationPath,
        [string]$ResolvedScriptRoot
    )

    $rules = New-Object System.Collections.Generic.List[object]
    $userConfigRoot = Join-Path (Get-UserDataRootPath) 'Config'
    $userConfigPath = Join-Path $userConfigRoot 'UserExcludedFolders.json'
    if (-not (Test-Path -LiteralPath $userConfigPath -PathType Leaf)) {
        try {
            if (-not (Test-Path -LiteralPath $userConfigRoot -PathType Container)) {
                New-Item -ItemType Directory -Path $userConfigRoot -Force | Out-Null
            }
            [pscustomobject]@{
                version             = 1
                userExcludedFolders = @()
                vendorPresets       = @(
                    [pscustomobject]@{
                        path    = '<SourcePath>\Samsung Gallery'
                        label   = 'Samsung Gallery'
                        reason  = 'External app managed folder'
                        role    = 'VendorManaged'
                    },
                    [pscustomobject]@{
                        path    = '<SourcePath>\Camera Roll'
                        label   = 'Camera Roll'
                        reason  = 'External camera upload folder'
                        role    = 'VendorManaged'
                    }
                )
            } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $userConfigPath -Encoding UTF8
        }
        catch {
            Write-Log -Message "User excluded folders config could not be created: $userConfigPath. Error=$($_.Exception.Message)" -Phase 'Validation'
        }
    }
    $entries = @()
    if (Test-Path -LiteralPath $userConfigPath -PathType Leaf) {
        try {
            $config = Get-Content -LiteralPath $userConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($config -and $config.PSObject.Properties.Name -contains 'userExcludedFolders') {
                $entries = @($config.userExcludedFolders)
            }
        }
        catch {
            Write-Log -Message "User excluded folders config could not be read: $userConfigPath. Error=$($_.Exception.Message)" -Phase 'Validation'
        }
    }

    foreach ($entry in $entries) {
        if ($null -eq $entry) { continue }
        $enabled = $true
        if ($entry.PSObject.Properties.Name -contains 'enabled') {
            $enabled = [bool]$entry.enabled
        }
        if (-not $enabled) { continue }

        $configuredPath = if ($entry.PSObject.Properties.Name -contains 'path') { [string]$entry.path } else { '' }
        $expandedPath = Expand-ConfiguredFolderPath -Path $configuredPath -ResolvedSourcePath $ResolvedSourcePath -ResolvedDestinationPath $ResolvedDestinationPath -ResolvedScriptRoot $ResolvedScriptRoot
        if ([string]::IsNullOrWhiteSpace($expandedPath)) { continue }

        $role = if ($entry.PSObject.Properties.Name -contains 'role') { [string]$entry.role } else { 'UserExcluded' }
        $label = if ($entry.PSObject.Properties.Name -contains 'label') { [string]$entry.label } else { '' }
        $reason = if ($entry.PSObject.Properties.Name -contains 'reason') { [string]$entry.reason } else { '' }
        $rule = New-FolderProtectionRule -Path $expandedPath -Role $role -Label $label -Reason $reason
        if ($null -ne $rule) {
            $rules.Add($rule)
        }
    }

    $script:FolderProtectionRules = @($rules.ToArray() | Sort-Object Path -Unique)
    $script:UserExcludedRoots = @($script:FolderProtectionRules | Where-Object { $_.Role -eq 'UserExcluded' } | ForEach-Object { $_.Path })
    $script:VendorManagedRoots = @($script:FolderProtectionRules | Where-Object { $_.Role -eq 'VendorManaged' } | ForEach-Object { $_.Path })
}

function Get-FolderProtectionRole {
    param([string]$Path)

    foreach ($rule in @($script:FolderProtectionRules)) {
        if (Test-IsChildPath -Path $Path -ParentPath ([string]$rule.Path)) {
            return $rule
        }
    }

    if (Test-IsExcludedPath -Path $Path) {
        return [pscustomobject]@{
            Path = $Path
            Role = 'InternalProtected'
            Label = 'Internal protected folder'
            Reason = 'Application managed/protected folder'
        }
    }

    return [pscustomobject]@{
        Path = $Path
        Role = 'None'
        Label = ''
        Reason = ''
    }
}

function Write-FolderProtectionSkipLog {
    param(
        [string]$Path,
        [string]$Phase
    )

    $role = Get-FolderProtectionRole -Path $Path
    if ($role.Role -eq 'UserExcluded') {
        $script:FolderProtectionSkipCounts.UserExcluded++
        Write-Log -Message "User excluded folder skipped: $Path. Reason: $($role.Reason)" -Phase $Phase
        return
    }
    if ($role.Role -eq 'VendorManaged') {
        $script:FolderProtectionSkipCounts.VendorManaged++
        Write-Log -Message "Vendor-managed folder skipped: $Path. Provider: $($role.Label). Reason: $($role.Reason)" -Phase $Phase
        return
    }
    if ($role.Role -eq 'InternalProtected') {
        $script:FolderProtectionSkipCounts.InternalProtected++
    }
    Write-DiagnosticLog "Excluded directory: $Path"
}

function Test-IsExcludedPath {
    param([string]$Path)
    return [bool]($ExcludedRoots | Where-Object { Test-IsChildPath -Path $Path -ParentPath $_ })
}

function Test-IsProtectedInternalFolderSegment {
    param(
        [string]$Path,
        [string]$RootPath
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($RootPath)) {
        return $false
    }

    $relative = ConvertTo-RelativePath -Path $Path -BasePath $RootPath
    if ([string]::IsNullOrWhiteSpace($relative)) {
        return $false
    }

    $protectedNames = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
    foreach ($rule in @($script:FolderProtectionRules)) {
        try {
            [void]$protectedNames.Add([System.IO.Path]::GetFileName(([string]$rule.Path).TrimEnd('\')))
        }
        catch {
        }
    }
    foreach ($folderKey in @('OrganizedFolder', 'NeedsReviewFolder', 'DuplicatesFolder', 'ConfirmedDuplicatesQuarantineFolder', 'MetadataBackupFolder', 'LogsFolder', 'MediaMetadataIssuesFolder')) {
        foreach ($folderName in (Get-AllInternalFolderNames -Key $folderKey)) {
            if (-not [string]::IsNullOrWhiteSpace($folderName)) {
                [void]$protectedNames.Add($folderName)
            }
        }
    }

    foreach ($part in @($relative -split '[\\/]' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        if ($protectedNames.Contains($part)) {
            return $true
        }
    }

    return $false
}

$StorageAttributeFlags = @{
    Offline = 0x00001000
    ReparsePoint = 0x00000400
    RecallOnOpen = 0x00040000
    Pinned = 0x00080000
    Unpinned = 0x00100000
    RecallOnDataAccess = 0x00400000
}

function Get-CloudProviderHint {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return 'Unknown' }
    if ($Path -match '(?i)[\\/]OneDrive([\\/]|$)') { return 'OneDrive' }
    if ($Path -match '(?i)[\\/]Dropbox([\\/]|$)') { return 'Dropbox' }
    if ($Path -match '(?i)[\\/]iCloudDrive([\\/]|$)|[\\/]iCloud Photos([\\/]|$)|[\\/]iCloud([\\/]|$)') { return 'iCloud' }
    if ($Path -match '(?i)[\\/]Google Drive([\\/]|$)|[\\/]My Drive([\\/]|$)') { return 'GoogleDrive' }
    if ($Path -match '(?i)[\\/]SynologyDrive([\\/]|$)|[\\/]Synology Drive([\\/]|$)') { return 'SynologyDrive' }
    if ($Path -match '(?i)[\\/]Box([\\/]|$)|[\\/]Box Drive([\\/]|$)') { return 'Box' }
    return 'Unknown'
}

function Detect-StorageAvailability {
    param(
        [System.IO.FileSystemInfo]$Item,
        [string]$Path = '',
        [switch]$Directory
    )

    $targetPath = if ($Item) { [string]$Item.FullName } else { [string]$Path }
    if ([string]::IsNullOrWhiteSpace($targetPath)) {
        return [pscustomobject]@{ State = 'MissingReal'; Reason = 'Path is empty'; ProviderHint = 'Unknown'; Attributes = 0; Path = $targetPath }
    }

    if (-not $Item) {
        try {
            $exists = if ($Directory) { Test-Path -LiteralPath $targetPath -PathType Container } else { Test-Path -LiteralPath $targetPath }
            if (-not $exists) {
                return [pscustomobject]@{ State = 'MissingReal'; Reason = 'Path does not exist'; ProviderHint = (Get-CloudProviderHint -Path $targetPath); Attributes = 0; Path = $targetPath }
            }
            $Item = Get-Item -LiteralPath $targetPath -Force -ErrorAction Stop
        }
        catch {
            return [pscustomobject]@{ State = 'MissingReal'; Reason = $_.Exception.Message; ProviderHint = (Get-CloudProviderHint -Path $targetPath); Attributes = 0; Path = $targetPath }
        }
    }

    $provider = Get-CloudProviderHint -Path $targetPath
    try {
        $attributeValue = [int64]$Item.Attributes
    }
    catch {
        return [pscustomobject]@{ State = 'CloudPlaceholder'; Reason = 'Attributes unavailable'; ProviderHint = $provider; Attributes = 0; Path = $targetPath }
    }

    $reasons = New-Object System.Collections.Generic.List[string]
    foreach ($flagName in @('Offline', 'RecallOnOpen', 'RecallOnDataAccess')) {
        if (($attributeValue -band [int64]$StorageAttributeFlags[$flagName]) -ne 0) {
            [void]$reasons.Add($flagName)
        }
    }

    $isReparse = (($attributeValue -band [int64]$StorageAttributeFlags.ReparsePoint) -ne 0)
    $isPinned = (($attributeValue -band [int64]$StorageAttributeFlags.Pinned) -ne 0)
    $isUnpinned = (($attributeValue -band [int64]$StorageAttributeFlags.Unpinned) -ne 0)

    if ($isUnpinned -and $reasons.Count -gt 0) {
        [void]$reasons.Add('Unpinned')
    }

    if (-not $Directory -and $isReparse -and -not $isPinned) {
        [void]$reasons.Add('ReparsePoint')
    }
    elseif ($isReparse -and $Diagnostic) {
        Write-DiagnosticLog "Reparse directory will be scanned if no cloud recall/offline attributes are present: $targetPath"
    }

    if ($reasons.Count -gt 0) {
        return [pscustomobject]@{
            State = 'CloudPlaceholder'
            Reason = ($reasons.ToArray() -join ',')
            ProviderHint = $provider
            Attributes = $attributeValue
            Path = $targetPath
        }
    }

    return [pscustomobject]@{
        State = 'LocalVerified'
        Reason = 'Local attributes'
        ProviderHint = $provider
        Attributes = $attributeValue
        Path = $targetPath
    }
}

function Register-CloudPlaceholderSkipped {
    param(
        [string]$Path,
        [string]$Phase,
        [object]$Availability
    )

    $Stats.CloudPlaceholdersSkipped++
    $Stats.SkippedOneDrive++
    $provider = if ($Availability -and $Availability.ProviderHint) { [string]$Availability.ProviderHint } else { Get-CloudProviderHint -Path $Path }
    $reason = if ($Availability -and $Availability.Reason) { [string]$Availability.Reason } else { 'Cloud placeholder' }
    if ($Stats.CloudPlaceholdersSkipped -le 200 -or $Stats.CloudPlaceholdersSkipped % 1000 -eq 0 -or $Diagnostic) {
        Write-Log -Message "CloudPlaceholderSkipped: Path=$Path; Provider=$provider; Reason=$reason" -Phase $Phase
    }
}

function Test-IsCloudOnlyOrProblematic {
    param([System.IO.FileSystemInfo]$Item)

    $availability = Detect-StorageAvailability -Item $Item -Directory:($Item -is [System.IO.DirectoryInfo])
    if ($availability.State -eq 'CloudPlaceholder') {
        return ('CloudPlaceholder/{0}/{1}' -f $availability.ProviderHint, $availability.Reason)
    }
    if ($availability.State -eq 'MissingReal') {
        return ('MissingReal/{0}' -f $availability.Reason)
    }
    return $null
}

function Test-ReadableFile {
    param([string]$Path)

    for ($attempt = 0; $attempt -lt 3; $attempt++) {
        try {
            $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
            $stream.Dispose()
            return $true
        }
        catch {
            Start-Sleep -Milliseconds (250 * ($attempt + 1))
        }
    }

    return $false
}

function Get-CompatibleFiles {
    param([string]$RootPath)

    Write-Log -Message "Scanning files..." -Phase 'Scan'
    $files = New-Object System.Collections.Generic.List[object]
    $directories = New-Object System.Collections.Generic.Stack[string]
    $directories.Push($RootPath)
    $lastScanLog = Get-Date

    while ($directories.Count -gt 0) {
        $directory = $directories.Pop()

        if (Test-IsExcludedPath -Path $directory) {
            Write-FolderProtectionSkipLog -Path $directory -Phase 'Scan'
            continue
        }

        Write-Heartbeat -Phase 'Scan' -Message "Scanning folder: $directory" -EverySeconds 10

        try {
            $childDirectories = @(Get-ChildItem -LiteralPath $directory -Directory -Force -ErrorAction Stop)
        }
        catch {
            $Stats.Inaccessible++
            Write-Log -Message "Inaccessible folder skipped: $directory - $($_.Exception.Message)" -Phase 'Scan'
            continue
        }

        foreach ($childDirectory in $childDirectories) {
            if (Test-IsExcludedPath -Path $childDirectory.FullName) {
                Write-FolderProtectionSkipLog -Path $childDirectory.FullName -Phase 'Scan'
                continue
            }

            $dirAvailability = Detect-StorageAvailability -Item $childDirectory -Directory
            if ($dirAvailability.State -eq 'CloudPlaceholder') {
                Register-CloudPlaceholderSkipped -Path $childDirectory.FullName -Phase 'Scan' -Availability $dirAvailability
                continue
            }
            if ($dirAvailability.State -eq 'MissingReal') {
                $Stats.MissingReal++
                Write-Log -Message "Directory missing during scan: $($childDirectory.FullName). Reason=$($dirAvailability.Reason)" -Phase 'Scan'
                continue
            }

            $directories.Push($childDirectory.FullName)
        }

        try {
            $childFiles = @(Get-ChildItem -LiteralPath $directory -File -Force -ErrorAction Stop)
        }
        catch {
            $Stats.Inaccessible++
            Write-Log -Message "Inaccessible files in folder skipped: $directory - $($_.Exception.Message)" -Phase 'Scan'
            continue
        }

        foreach ($file in $childFiles) {
            if (Test-IsExcludedPath -Path $file.FullName) { continue }
            $ext = $file.Extension.ToLowerInvariant()
            if ($MediaExtensions -notcontains $ext) { continue }

            $availability = Detect-StorageAvailability -Item $file
            if ($availability.State -eq 'CloudPlaceholder') {
                Register-CloudPlaceholderSkipped -Path $file.FullName -Phase 'Scan' -Availability $availability
                continue
            }
            if ($availability.State -eq 'MissingReal') {
                $Stats.MissingReal++
                Write-Log -Message "File missing during scan: $($file.FullName). Reason=$($availability.Reason)" -Phase 'Scan'
                continue
            }

            if (-not (Test-ReadableFile -Path $file.FullName)) {
                $Stats.Inaccessible++
                Write-Log -Message "Inaccessible file skipped: $($file.FullName)" -Phase 'Scan'
                continue
            }

            $files.Add($file)
            $Stats.LocalFilesDetected++

            if ($files.Count -eq 1) {
                Write-Log -Message "First compatible file detected: $($file.FullName)" -Phase 'Scan'
            }
            if ($files.Count % 1000 -eq 0) {
                Write-Log -Message "Scanning files... found $($files.Count) compatible files so far." -Phase 'Scan'
            }
            elseif (((Get-Date) - $lastScanLog).TotalSeconds -ge 10) {
                $lastScanLog = Get-Date
                Write-Log -Message "Scanning files... found $($files.Count) compatible files so far. Current folder: $directory" -Phase 'Scan'
            }
        }
    }

    $Stats.FilesFound = $files.Count
    Write-Log -Message "Found $($files.Count) files." -Phase 'Scan'
    $protectedSkipTotal = [int]$script:FolderProtectionSkipCounts.UserExcluded + [int]$script:FolderProtectionSkipCounts.VendorManaged + [int]$script:FolderProtectionSkipCounts.InternalProtected
    if ($protectedSkipTotal -gt 0) {
        Write-Log -Message ("Protected/excluded folders skipped during scan: userExcluded={0}; vendorManaged={1}; internalProtected={2}" -f $script:FolderProtectionSkipCounts.UserExcluded, $script:FolderProtectionSkipCounts.VendorManaged, $script:FolderProtectionSkipCounts.InternalProtected) -Phase 'Scan'
    }
    if ($files.Count -eq 0 -and $protectedSkipTotal -gt 0) {
        Write-Log -Message "No processable media files were found after applying protected/excluded folders. If you expected files, review SourcePath and the dashboard excluded folders list." -Phase 'Scan' -Status 'Warning'
    }
    Write-Log -Message "Storage availability: Local files detected: $($Stats.LocalFilesDetected); Cloud placeholders skipped: $($Stats.CloudPlaceholdersSkipped); Missing real: $($Stats.MissingReal)" -Phase 'Scan'
    return $files.ToArray()
}

function Get-MediaFilesWithoutProtectedTraversal {
    param(
        [string]$RootPath,
        [string]$Phase,
        [switch]$CountLocalFiles
    )

    if ([string]::IsNullOrWhiteSpace($RootPath) -or -not (Test-Path -LiteralPath $RootPath -PathType Container)) {
        return @()
    }

    $rootFullPath = (Resolve-FullPath $RootPath).TrimEnd('\')
    $rootProtection = Get-FolderProtectionRole -Path $rootFullPath
    $rootIsAllowedInternalProtectedScanRoot = (Test-IsExcludedPath -Path $rootFullPath) -and $rootProtection.Role -eq 'InternalProtected'
    $files = New-Object System.Collections.Generic.List[object]
    $directories = New-Object System.Collections.Generic.Stack[string]
    $directories.Push($rootFullPath)

    while ($directories.Count -gt 0) {
        $directory = $directories.Pop()
        if (Test-IsExcludedPath -Path $directory) {
            $protection = Get-FolderProtectionRole -Path $directory
            $isCoveredOnlyByAllowedRoot = $false
            if ($rootIsAllowedInternalProtectedScanRoot -and (Test-IsChildPath -Path $directory -ParentPath $rootFullPath) -and $protection.Role -eq 'InternalProtected') {
                $nestedExcludedRoot = $false
                foreach ($excludedRoot in @($ExcludedRoots)) {
                    $resolvedExcludedRoot = (Resolve-FullPath $excludedRoot).TrimEnd('\')
                    if ($resolvedExcludedRoot.Equals($rootFullPath, [StringComparison]::OrdinalIgnoreCase)) { continue }
                    if (Test-IsChildPath -Path $directory -ParentPath $resolvedExcludedRoot) {
                        $nestedExcludedRoot = $true
                        break
                    }
                }
                $isCoveredOnlyByAllowedRoot = -not $nestedExcludedRoot
            }
            if (-not $isCoveredOnlyByAllowedRoot) {
                Write-FolderProtectionSkipLog -Path $directory -Phase $Phase
                continue
            }
        }

        try {
            $childDirectories = @(Get-ChildItem -LiteralPath $directory -Directory -Force -ErrorAction Stop)
        }
        catch {
            $Stats.Inaccessible++
            Write-Log -Message "Inaccessible folder skipped: $directory - $($_.Exception.Message)" -Phase $Phase
            continue
        }

        foreach ($childDirectory in $childDirectories) {
            if (Test-IsExcludedPath -Path $childDirectory.FullName) {
                $protection = Get-FolderProtectionRole -Path $childDirectory.FullName
                $isCoveredOnlyByAllowedRoot = $false
                if ($rootIsAllowedInternalProtectedScanRoot -and (Test-IsChildPath -Path $childDirectory.FullName -ParentPath $rootFullPath) -and $protection.Role -eq 'InternalProtected') {
                    $nestedExcludedRoot = $false
                    foreach ($excludedRoot in @($ExcludedRoots)) {
                        $resolvedExcludedRoot = (Resolve-FullPath $excludedRoot).TrimEnd('\')
                        if ($resolvedExcludedRoot.Equals($rootFullPath, [StringComparison]::OrdinalIgnoreCase)) { continue }
                        if (Test-IsChildPath -Path $childDirectory.FullName -ParentPath $resolvedExcludedRoot) {
                            $nestedExcludedRoot = $true
                            break
                        }
                    }
                    $isCoveredOnlyByAllowedRoot = -not $nestedExcludedRoot
                }
                if (-not $isCoveredOnlyByAllowedRoot) {
                    Write-FolderProtectionSkipLog -Path $childDirectory.FullName -Phase $Phase
                    continue
                }
            }
            $dirAvailability = Detect-StorageAvailability -Item $childDirectory -Directory
            if ($dirAvailability.State -eq 'CloudPlaceholder') {
                Register-CloudPlaceholderSkipped -Path $childDirectory.FullName -Phase $Phase -Availability $dirAvailability
                continue
            }
            if ($dirAvailability.State -eq 'MissingReal') {
                $Stats.MissingReal++
                continue
            }
            $directories.Push($childDirectory.FullName)
        }

        try {
            $childFiles = @(Get-ChildItem -LiteralPath $directory -File -Force -ErrorAction Stop)
        }
        catch {
            $Stats.Inaccessible++
            Write-Log -Message "Inaccessible files in folder skipped: $directory - $($_.Exception.Message)" -Phase $Phase
            continue
        }

        foreach ($file in $childFiles) {
            if (Test-IsExcludedPath -Path $file.FullName) {
                $protection = Get-FolderProtectionRole -Path $file.FullName
                $isCoveredOnlyByAllowedRoot = $false
                if ($rootIsAllowedInternalProtectedScanRoot -and (Test-IsChildPath -Path $file.FullName -ParentPath $rootFullPath) -and $protection.Role -eq 'InternalProtected') {
                    $nestedExcludedRoot = $false
                    foreach ($excludedRoot in @($ExcludedRoots)) {
                        $resolvedExcludedRoot = (Resolve-FullPath $excludedRoot).TrimEnd('\')
                        if ($resolvedExcludedRoot.Equals($rootFullPath, [StringComparison]::OrdinalIgnoreCase)) { continue }
                        if (Test-IsChildPath -Path $file.FullName -ParentPath $resolvedExcludedRoot) {
                            $nestedExcludedRoot = $true
                            break
                        }
                    }
                    $isCoveredOnlyByAllowedRoot = -not $nestedExcludedRoot
                }
                if (-not $isCoveredOnlyByAllowedRoot) { continue }
            }
            if ($MediaExtensions -notcontains $file.Extension.ToLowerInvariant()) { continue }

            $availability = Detect-StorageAvailability -Item $file
            if ($availability.State -eq 'CloudPlaceholder') {
                Register-CloudPlaceholderSkipped -Path $file.FullName -Phase $Phase -Availability $availability
                continue
            }
            if ($availability.State -eq 'MissingReal') {
                $Stats.MissingReal++
                continue
            }

            if ($CountLocalFiles) {
                $Stats.LocalFilesDetected++
            }
            $files.Add($file)
        }
    }

    return @($files.ToArray())
}

function Get-Sha256 {
    param([string]$Path)
    $availability = Detect-StorageAvailability -Path $Path
    if ($availability.State -eq 'CloudPlaceholder') {
        Register-CloudPlaceholderSkipped -Path $Path -Phase 'Hash queue' -Availability $availability
        throw "CloudPlaceholder: $Path"
    }
    if ($availability.State -eq 'MissingReal') {
        $Stats.MissingReal++
        throw "MissingReal: $Path"
    }
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

$script:ProcessedByHash = @{}
$script:ProcessedRecords = New-Object System.Collections.Generic.List[object]
$script:StaleProcessedRecords = New-Object System.Collections.Generic.List[object]
$script:StaleProcessedByHash = @{}
$script:ProcessedDirtyCount = 0
$script:JsonBackupDoneByPath = @{}
$script:IndexBackupRetentionDays = 7
$script:IndexBackupMaxFiles = 10

function Get-FallbackProcessedDbPath {
    return (Join-Path (Join-Path (Get-UserDataRootPath) 'Config') 'ProcessedFiles.json')
}

function Get-IndexBackupRootPath {
    return (Join-Path (Get-UserDataRootPath) 'IndexBackups')
}

function Get-InternalJsonFiles {
    $paths = New-Object System.Collections.Generic.List[string]
    foreach ($path in @($ProcessedDbPath, (Get-FallbackProcessedDbPath), $ProgressPath)) {
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            $resolved = Resolve-FullPath $path
            if (-not $paths.Contains($resolved)) {
                $paths.Add($resolved)
            }
        }
    }

    foreach ($candidate in @(
            (Join-Path (Join-Path (Get-UserDataRootPath) 'Config') 'ProcessedFiles.json'),
            (Join-Path (Get-UserDataRootPath) 'settings.json'),
            (Join-Path (Get-UserDataRootPath) 'dashboard-settings.json'),
            (Join-Path (Get-ScriptRootPath) 'LanguageResources.json')
        )) {
        if (-not [string]::IsNullOrWhiteSpace($candidate)) {
            $resolved = Resolve-FullPath $candidate
            if (-not $paths.Contains($resolved)) {
                $paths.Add($resolved)
            }
        }
    }

    return @($paths.ToArray())
}

function Backup-InternalJsonFile {
    param(
        [string]$Path,
        [string]$Reason
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $null
    }

    $resolvedPath = Resolve-FullPath $Path
    if (-not (Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
        return $null
    }

    $key = $resolvedPath.ToLowerInvariant()
    if ($script:JsonBackupDoneByPath.ContainsKey($key)) {
        return $script:JsonBackupDoneByPath[$key]
    }

    Move-LegacyJsonBackupsToIndexBackups

    $backupRoot = Join-Path (Get-IndexBackupRootPath) $script:RunId
    if (-not (Test-Path -LiteralPath $backupRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
    }

    $safeName = ([System.IO.Path]::GetFileName($resolvedPath) -replace '[^\w.\-]+', '_')
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backupPath = Join-Path $backupRoot ("{0}-{1}" -f $stamp, $safeName)
    Copy-Item -LiteralPath $resolvedPath -Destination $backupPath -Force
    $script:JsonBackupDoneByPath[$key] = $backupPath
    Write-Log -Message "JSON backup created before mutation: $resolvedPath -> $backupPath. Reason: $Reason" -Phase 'Safety'
    Invoke-IndexBackupRetentionCleanup
    return $backupPath
}

function Backup-InternalJsonFilesForMutation {
    param(
        [string[]]$Paths,
        [string]$Reason
    )

    foreach ($path in @($Paths)) {
        Backup-InternalJsonFile -Path $path -Reason $Reason | Out-Null
    }
}

function Get-UniqueBackupTargetPath {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return $Path
    }

    $directory = [System.IO.Path]::GetDirectoryName($Path)
    $nameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $extension = [System.IO.Path]::GetExtension($Path)
    for ($i = 1; $i -lt 1000; $i++) {
        $candidate = Join-Path $directory ("{0}-migrated-{1}{2}" -f $nameWithoutExtension, $i, $extension)
        if (-not (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    return (Join-Path $directory ("{0}-migrated-{1}{2}" -f $nameWithoutExtension, ([guid]::NewGuid().ToString('N')), $extension))
}

function Move-LegacyJsonBackupsToIndexBackups {
    $legacyRoot = Join-Path $LogRoot 'JsonBackups'
    $targetRoot = Get-IndexBackupRootPath
    if ([string]::IsNullOrWhiteSpace($legacyRoot) -or
        -not (Test-Path -LiteralPath $legacyRoot -PathType Container)) {
        return
    }

    try {
        if (-not (Test-Path -LiteralPath $targetRoot -PathType Container)) {
            New-Item -ItemType Directory -Path $targetRoot -Force | Out-Null
        }

        $moved = 0
        foreach ($file in @(Get-ChildItem -LiteralPath $legacyRoot -File -Recurse -Force -ErrorAction SilentlyContinue)) {
            try {
                $relative = ConvertTo-RelativePath -Path $file.FullName -BasePath $legacyRoot
                $target = Join-Path $targetRoot $relative
                $targetDirectory = [System.IO.Path]::GetDirectoryName($target)
                if (-not (Test-Path -LiteralPath $targetDirectory -PathType Container)) {
                    New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
                }
                $target = Get-UniqueBackupTargetPath -Path $target
                Move-Item -LiteralPath $file.FullName -Destination $target -Force
                $moved++
            }
            catch {
                Write-DiagnosticLog "Legacy JsonBackup migration skipped file: $($file.FullName). Error=$($_.Exception.Message)"
            }
        }

        foreach ($directory in @(Get-ChildItem -LiteralPath $legacyRoot -Directory -Recurse -Force -ErrorAction SilentlyContinue | Sort-Object FullName -Descending)) {
            try {
                if (@(Get-ChildItem -LiteralPath $directory.FullName -Force -ErrorAction Stop).Count -eq 0) {
                    Remove-Item -LiteralPath $directory.FullName -Force -ErrorAction Stop
                }
            }
            catch {
            }
        }

        try {
            if (@(Get-ChildItem -LiteralPath $legacyRoot -Force -ErrorAction Stop).Count -eq 0) {
                Remove-Item -LiteralPath $legacyRoot -Force -ErrorAction Stop
            }
        }
        catch {
        }

        if ($moved -gt 0) {
            Write-Log -Message "Migrated legacy JSON index backups from Logs\\JsonBackups to IndexBackups. Files moved: $moved" -Phase 'Safety'
        }
    }
    catch {
        Write-Log -Message "Legacy JSON index backup migration could not be completed: $($_.Exception.Message)" -Phase 'Safety'
    }
}

function Invoke-IndexBackupRetentionCleanup {
    $root = Get-IndexBackupRootPath
    if ([string]::IsNullOrWhiteSpace($root) -or -not (Test-Path -LiteralPath $root -PathType Container)) {
        return
    }

    try {
        $files = @(Get-ChildItem -LiteralPath $root -File -Recurse -Force -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
        if ($files.Count -le 1) {
            return
        }

        $cutoff = (Get-Date).AddDays(-1 * [math]::Max(1, $script:IndexBackupRetentionDays))
        $protected = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
        [void]$protected.Add($files[0].FullName)

        $recentKept = 0
        foreach ($file in $files) {
            if ($recentKept -ge [math]::Max(1, $script:IndexBackupMaxFiles)) {
                break
            }
            if ($file.LastWriteTime -ge $cutoff) {
                [void]$protected.Add($file.FullName)
                $recentKept++
            }
        }

        $deleted = 0
        $recovered = [int64]0
        foreach ($file in $files) {
            if ($protected.Contains($file.FullName)) { continue }
            try {
                $bytes = [int64]$file.Length
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                $deleted++
                $recovered += $bytes
            }
            catch {
                Write-DiagnosticLog "Index backup retention skipped locked file: $($file.FullName). Error=$($_.Exception.Message)"
            }
        }

        foreach ($directory in @(Get-ChildItem -LiteralPath $root -Directory -Recurse -Force -ErrorAction SilentlyContinue | Sort-Object FullName -Descending)) {
            try {
                if (@(Get-ChildItem -LiteralPath $directory.FullName -Force -ErrorAction Stop).Count -eq 0) {
                    Remove-Item -LiteralPath $directory.FullName -Force -ErrorAction Stop
                }
            }
            catch {
            }
        }

        if ($deleted -gt 0) {
            Write-Log -Message ("Index backup retention cleanup completed. Deleted {0} files. Recovered {1} MB. Root={2}" -f $deleted, ([math]::Round($recovered / 1MB, 2)), $root) -Phase 'Safety'
        }
    }
    catch {
        Write-Log -Message "Index backup retention cleanup could not be completed: $($_.Exception.Message)" -Phase 'Safety'
    }
}

function Add-StaleProcessedRecord {
    param([object]$Record)

    if ($null -eq $Record -or [string]::IsNullOrWhiteSpace([string]$Record.hash)) {
        return
    }

    $hashKey = ([string]$Record.hash).ToUpperInvariant()
    $script:StaleProcessedRecords.Add($Record)
    if (-not $script:StaleProcessedByHash.ContainsKey($hashKey)) {
        $script:StaleProcessedByHash[$hashKey] = New-Object System.Collections.Generic.List[object]
    }
    $script:StaleProcessedByHash[$hashKey].Add($Record)
}

function Get-ProcessedRecordRegisteredPath {
    param([object]$Record)

    if ($null -eq $Record) {
        return ''
    }

    $rawPath = ''
    $basePath = ''
    if ($Record.PSObject.Properties.Name -contains 'newRelativePath' -and
        -not [string]::IsNullOrWhiteSpace([string]$Record.newRelativePath)) {
        $rawPath = [string]$Record.newRelativePath
        $basePath = $DestinationBase
    }
    elseif ($Record.PSObject.Properties.Name -contains 'originalRelativePath' -and
        -not [string]::IsNullOrWhiteSpace([string]$Record.originalRelativePath)) {
        $rawPath = [string]$Record.originalRelativePath
        $basePath = $SourcePath
    }
    else {
        return ''
    }

    $cleanPath = $rawPath.Replace('/', '\').Trim()
    if ([System.IO.Path]::IsPathRooted($cleanPath)) {
        return (Resolve-FullPath $cleanPath)
    }

    return (Resolve-FullPath (Join-Path $basePath $cleanPath.TrimStart('\')))
}

function Resolve-InternalFolderAliasPathCandidate {
    param([string]$RegisteredPath)

    if ([string]::IsNullOrWhiteSpace($RegisteredPath)) {
        return [pscustomobject]@{ Status = 'None'; Path = ''; Paths = @(); RoleKey = ''; OldName = ''; NewName = '' }
    }

    $resolvedPath = Resolve-FullPath $RegisteredPath
    $folderKeys = @(
        'OrganizedFolder',
        'NeedsReviewFolder',
        'DuplicatesFolder',
        'ConfirmedDuplicatesQuarantineFolder',
        'MetadataBackupFolder',
        'LogsFolder',
        'MediaMetadataIssuesFolder'
    )

    $candidateMap = @{}
    foreach ($folderKey in $folderKeys) {
        $aliases = @(
            @((Get-AllInternalFolderNames -Key $folderKey)) +
            @((Get-InternalFolderName -Key $folderKey))
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

        foreach ($oldName in $aliases) {
            $pattern = '(^|[\\/])' + [regex]::Escape($oldName) + '($|[\\/])'
            $match = [regex]::Match($resolvedPath, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if (-not $match.Success) { continue }

            foreach ($newName in $aliases) {
                if ($oldName.Equals($newName, [StringComparison]::OrdinalIgnoreCase)) { continue }

                $candidate = $resolvedPath.Substring(0, $match.Index) +
                    $match.Groups[1].Value +
                    $newName +
                    $match.Groups[2].Value +
                    $resolvedPath.Substring($match.Index + $match.Length)
                $candidate = Resolve-FullPath $candidate

                if ($candidate.TrimEnd('\').Equals($resolvedPath.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
                    continue
                }

                if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                    $key = $candidate.TrimEnd('\').ToLowerInvariant()
                    if (-not $candidateMap.ContainsKey($key)) {
                        $candidateMap[$key] = [pscustomobject]@{
                            Path = $candidate
                            RoleKey = $folderKey
                            OldName = $oldName
                            NewName = $newName
                        }
                    }
                }
            }
        }
    }

    $candidates = @($candidateMap.Values)
    if ($candidates.Count -eq 1) {
        return [pscustomobject]@{
            Status = 'Unique'
            Path = [string]$candidates[0].Path
            Paths = @([string]$candidates[0].Path)
            RoleKey = [string]$candidates[0].RoleKey
            OldName = [string]$candidates[0].OldName
            NewName = [string]$candidates[0].NewName
        }
    }
    if ($candidates.Count -gt 1) {
        return [pscustomobject]@{
            Status = 'Multiple'
            Path = ''
            Paths = @($candidates | ForEach-Object { [string]$_.Path })
            RoleKey = ''
            OldName = ''
            NewName = ''
        }
    }

    return [pscustomobject]@{ Status = 'None'; Path = ''; Paths = @(); RoleKey = ''; OldName = ''; NewName = '' }
}

function Set-ProcessedRecordMissing {
    param([object]$Record)

    if ($null -eq $Record) {
        return
    }

    $now = (Get-Date).ToString('o')
    $Record.status = 'Missing/Stale'
    $Record.lastSeen = if (-not [string]::IsNullOrWhiteSpace([string]$Record.lastSeen)) { [string]$Record.lastSeen } else { [string]$Record.date }
    if ([string]::IsNullOrWhiteSpace([string]$Record.missingSince)) {
        $Record.missingSince = $now
    }
    $Record.date = $now
}

function Set-ProcessedRecordProperty {
    param(
        [object]$Record,
        [string]$Name,
        [object]$Value
    )

    if ($Record.PSObject.Properties.Name -contains $Name) {
        $Record.$Name = $Value
    }
    else {
        $Record | Add-Member -NotePropertyName $Name -NotePropertyValue $Value -Force
    }
}

function Set-ProcessedRecordCloudPlaceholder {
    param(
        [object]$Record,
        [string]$CurrentPath,
        [object]$Availability
    )

    if ($null -eq $Record) {
        return
    }

    $now = (Get-Date).ToString('o')
    Set-ProcessedRecordProperty -Record $Record -Name 'status' -Value 'CloudPlaceholderKnown'
    Set-ProcessedRecordProperty -Record $Record -Name 'lastSeen' -Value $now
    Set-ProcessedRecordProperty -Record $Record -Name 'missingSince' -Value ''
    Set-ProcessedRecordProperty -Record $Record -Name 'date' -Value $now
    Set-ProcessedRecordProperty -Record $Record -Name 'verifiedLocal' -Value $false
    Set-ProcessedRecordProperty -Record $Record -Name 'storageState' -Value 'CloudPlaceholder'
    Set-ProcessedRecordProperty -Record $Record -Name 'providerHint' -Value $(if ($Availability -and $Availability.ProviderHint) { [string]$Availability.ProviderHint } else { Get-CloudProviderHint -Path $CurrentPath })
}

function Register-CloudPlaceholderInIndex {
    param(
        [string]$Hash,
        [string]$Path
    )

    $hashKey = if ([string]::IsNullOrWhiteSpace($Hash)) { '<no-hash>' } else { ([string]$Hash).ToUpperInvariant() }
    $pathKey = if ([string]::IsNullOrWhiteSpace($Path)) { '<no-path>' } else { (Resolve-FullPath $Path).TrimEnd('\').ToLowerInvariant() }
    $key = '{0}|{1}' -f $hashKey, $pathKey
    if ($script:CloudPlaceholdersInIndexKeys.Add($key)) {
        $Stats.CloudPlaceholdersInIndex = $script:CloudPlaceholdersInIndexKeys.Count
        return $true
    }

    $Stats.CloudPlaceholdersInIndex = $script:CloudPlaceholdersInIndexKeys.Count
    return $false
}

function Set-ProcessedRecordCurrentPath {
    param(
        [object]$Record,
        [string]$CurrentPath
    )

    if ($null -eq $Record -or [string]::IsNullOrWhiteSpace($CurrentPath)) {
        return
    }

    if (Test-IsChildPath -Path $CurrentPath -ParentPath $DestinationBase) {
        $Record.newRelativePath = ConvertTo-RelativePath -Path $CurrentPath -BasePath $DestinationBase
        if ($Record.PSObject.Properties.Name -contains 'originalRelativePath') {
            $Record.originalRelativePath = ''
        }
    }
    elseif (Test-IsChildPath -Path $CurrentPath -ParentPath $SourcePath) {
        $Record.originalRelativePath = ConvertTo-RelativePath -Path $CurrentPath -BasePath $SourcePath
        if ($Record.PSObject.Properties.Name -contains 'newRelativePath') {
            $Record.newRelativePath = ''
        }
    }
    else {
        $Record.newRelativePath = Resolve-FullPath $CurrentPath
        if ($Record.PSObject.Properties.Name -contains 'originalRelativePath') {
            $Record.originalRelativePath = ''
        }
    }

    if (Test-IsChildPath -Path $CurrentPath -ParentPath $OrganizedRoot) {
        $Record.status = 'Reconciled - registered path updated'
    }
    else {
        $Record.status = 'ManualMoved/ExternalLocation'
    }
    $Record.lastSeen = (Get-Date).ToString('o')
    $Record.missingSince = ''
    $Record.date = $Record.lastSeen
    Set-ProcessedRecordProperty -Record $Record -Name 'verifiedLocal' -Value $true
    Set-ProcessedRecordProperty -Record $Record -Name 'storageState' -Value 'LocalVerified'
    Set-ProcessedRecordProperty -Record $Record -Name 'providerHint' -Value (Get-CloudProviderHint -Path $CurrentPath)
}

function Resolve-ProcessedRecordRegisteredPathWithInternalAlias {
    param(
        [object]$Record,
        [string]$RegisteredPath,
        [string]$HashKey,
        [string]$Context
    )

    $aliasResolution = Resolve-InternalFolderAliasPathCandidate -RegisteredPath $RegisteredPath
    if ($aliasResolution.Status -eq 'Unique') {
        Set-ProcessedRecordCurrentPath -Record $Record -CurrentPath ([string]$aliasResolution.Path)
        $Stats.JsonReconcilePathsUpdated++
        $script:ProcessedDirtyCount++
        Write-Log -Message ("JSON internal folder alias path resolved. Hash={0}. OldPath={1}. NewPath={2}. Role={3}. Alias={4}->{5}. Context={6}" -f $HashKey, $RegisteredPath, $aliasResolution.Path, $aliasResolution.RoleKey, $aliasResolution.OldName, $aliasResolution.NewName, $Context) -Phase 'JSON reconciliation'
        return [string]$aliasResolution.Path
    }

    if ($aliasResolution.Status -eq 'Multiple') {
        Set-ProcessedRecordProperty -Record $Record -Name 'status' -Value 'AliasConflict - Multiple internal alias paths'
        Set-ProcessedRecordProperty -Record $Record -Name 'missingSince' -Value ''
        Set-ProcessedRecordProperty -Record $Record -Name 'verifiedLocal' -Value $false
        Set-ProcessedRecordProperty -Record $Record -Name 'storageState' -Value 'AliasConflict'
        Set-ProcessedRecordProperty -Record $Record -Name 'internalAliasCandidates' -Value (@($aliasResolution.Paths) -join ' | ')
        Set-ProcessedRecordProperty -Record $Record -Name 'lastSeen' -Value (Get-Date).ToString('o')
        Set-ProcessedRecordProperty -Record $Record -Name 'date' -Value $Record.lastSeen
        $Stats.JsonReconcileConflicts++
        $script:ProcessedDirtyCount++
        Write-Log -Message ("JSON internal folder alias conflict. Hash={0}. OldPath={1}. Candidates={2}. Context={3}" -f $HashKey, $RegisteredPath, (@($aliasResolution.Paths) -join ' | '), $Context) -Phase 'JSON reconciliation'
    }

    return ''
}

function Find-ProcessedRecordByPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not $script:ProcessedRecords) {
        return $null
    }

    $resolvedPath = (Resolve-FullPath $Path).TrimEnd('\')
    foreach ($record in @($script:ProcessedRecords.ToArray())) {
        $registeredPath = Get-ProcessedRecordRegisteredPath -Record $record
        if (-not [string]::IsNullOrWhiteSpace($registeredPath) -and
            (Resolve-FullPath $registeredPath).TrimEnd('\').Equals($resolvedPath, [StringComparison]::OrdinalIgnoreCase)) {
            return $record
        }
    }

    return $null
}

function Update-ProcessedIndexHashAfterExifRepair {
    param(
        [string]$Path,
        [string]$OldHash,
        [string]$NewHash
    )

    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($NewHash)) {
        return $false
    }

    $record = Find-ProcessedRecordByPath -Path $Path
    if ($null -eq $record) {
        Write-DiagnosticLog "No ProcessedFiles entry found by path after EXIF repair; new files will be indexed when registered. Path=$Path"
        return $false
    }

    $newHashKey = $NewHash.ToUpperInvariant()
    $oldHashKey = if ([string]::IsNullOrWhiteSpace($OldHash)) {
        if ($record.hash) { ([string]$record.hash).ToUpperInvariant() } else { '' }
    }
    else {
        $OldHash.ToUpperInvariant()
    }
    if ($oldHashKey -eq $newHashKey) {
        return $false
    }

    if ($script:ProcessedByHash.ContainsKey($newHashKey)) {
        $existing = $script:ProcessedByHash[$newHashKey]
        $existingPath = Get-ProcessedRecordRegisteredPath -Record $existing
        if (-not [string]::IsNullOrWhiteSpace($existingPath) -and
            -not (Resolve-FullPath $existingPath).TrimEnd('\').Equals((Resolve-FullPath $Path).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
            $Stats.JsonReconcileConflicts++
            Write-Log -Message "Processed index hash update after EXIF repair skipped due to hash conflict: oldHash=$oldHashKey newHash=$newHashKey path=$Path existingPath=$existingPath" -Phase 'JSON reconciliation'
            return $false
        }
    }

    if ($script:ProcessedByHash.ContainsKey($oldHashKey)) {
        $existingOld = $script:ProcessedByHash[$oldHashKey]
        $existingOldPath = Get-ProcessedRecordRegisteredPath -Record $existingOld
        if (-not [string]::IsNullOrWhiteSpace($existingOldPath) -and
            (Resolve-FullPath $existingOldPath).TrimEnd('\').Equals((Resolve-FullPath $Path).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
            $script:ProcessedByHash.Remove($oldHashKey)
        }
    }

    Set-ProcessedRecordProperty -Record $record -Name 'hash' -Value $newHashKey
    Set-ProcessedRecordProperty -Record $record -Name 'lastSeen' -Value (Get-Date).ToString('o')
    Set-ProcessedRecordProperty -Record $record -Name 'date' -Value $record.lastSeen
    Set-ProcessedRecordProperty -Record $record -Name 'missingSince' -Value ''
    Set-ProcessedRecordProperty -Record $record -Name 'verifiedLocal' -Value $true
    Set-ProcessedRecordProperty -Record $record -Name 'storageState' -Value 'LocalVerified'
    Set-ProcessedRecordProperty -Record $record -Name 'providerHint' -Value (Get-CloudProviderHint -Path $Path)
    if (Test-ProcessedRecordFeedsDuplicateIndex -Record $record) {
        $script:ProcessedByHash[$newHashKey] = $record
    }
    $script:ProcessedDirtyCount++
    Write-Log -Message "Processed index hash updated after EXIF repair: oldHash=$oldHashKey newHash=$newHashKey path=$Path" -Phase 'JSON reconciliation'
    return $true
}

function Test-ProcessedRecordFeedsDuplicateIndex {
    param([object]$Record)

    if ($null -eq $Record) {
        return $false
    }

    $status = if ($Record.PSObject.Properties.Name -contains 'status') { [string]$Record.status } else { '' }
    if ($status -like 'Missing/Stale*') { return $false }
    if ($status -like 'ManualMoved/ExternalLocation*') { return $false }
    if ($status -like 'CloudPlaceholderKnown*') { return $false }
    if ($status -like 'AliasConflict*') { return $false }
    if ($Record.PSObject.Properties.Name -contains 'storageState' -and [string]$Record.storageState -eq 'CloudPlaceholder') { return $false }
    if ($Record.PSObject.Properties.Name -contains 'storageState' -and [string]$Record.storageState -eq 'AliasConflict') { return $false }
    if ($Record.PSObject.Properties.Name -contains 'verifiedLocal' -and [string]$Record.verifiedLocal -eq 'False') { return $false }
    return $true
}

function Load-ProcessedDatabase {
    if ([string]::IsNullOrWhiteSpace($ProcessedDbPath)) {
        return
    }

    if (-not (Test-Path -LiteralPath $ProcessedDbPath -PathType Leaf)) {
        Write-Log -Message "Processed database not found. A new one will be created: $ProcessedDbPath" -Phase 'Incremental'
        return
    }

    try {
        $script:ProcessedByHash = @{}
        $script:ProcessedRecords = New-Object System.Collections.Generic.List[object]
        $script:StaleProcessedRecords = New-Object System.Collections.Generic.List[object]
        $script:StaleProcessedByHash = @{}

        $json = Get-Content -LiteralPath $ProcessedDbPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $jsonProperties = @($json.PSObject.Properties.Name)
        if (($jsonProperties -contains 'schemaVersion') -and $json.schemaVersion -eq 2 -and ($jsonProperties -contains 'files') -and $json.files) {
            $records = @($json.files)
        }
        else {
            Write-Log -Message "Migrating old ProcessedFiles schema to schemaVersion 2." -Phase 'Incremental'
            $records = @($json)
        }
        foreach ($record in $records) {
            $recordProperties = @($record.PSObject.Properties.Name)
            if (($recordProperties -contains 'hash') -and $record.hash) {
                $hashKey = ([string]$record.hash).ToUpperInvariant()
                $normalized = [pscustomobject]@{
                    hash = $hashKey
                    originalRelativePath = if (($recordProperties -contains 'originalRelativePath') -and $record.originalRelativePath) { [string]$record.originalRelativePath } elseif (($recordProperties -contains 'originalPath') -and $record.originalPath) { ConvertTo-RelativePath -Path ([string]$record.originalPath) -BasePath $SourcePath } else { '' }
                    newRelativePath = if (($recordProperties -contains 'newRelativePath') -and $record.newRelativePath) { [string]$record.newRelativePath } elseif (($recordProperties -contains 'newPath') -and $record.newPath) { ConvertTo-RelativePath -Path ([string]$record.newPath) -BasePath $DestinationBase } else { '' }
                    status = if (($recordProperties -contains 'status') -and $record.status) { [string]$record.status } else { 'Migrated' }
                    date = if (($recordProperties -contains 'date') -and $record.date) { [string]$record.date } else { (Get-Date).ToString('o') }
                    lastSeen = if (($recordProperties -contains 'lastSeen') -and $record.lastSeen) { [string]$record.lastSeen } elseif (($recordProperties -contains 'date') -and $record.date) { [string]$record.date } else { (Get-Date).ToString('o') }
                    missingSince = if (($recordProperties -contains 'missingSince') -and $record.missingSince) { [string]$record.missingSince } else { '' }
                    verifiedLocal = if ($recordProperties -contains 'verifiedLocal') { [bool]$record.verifiedLocal } else { $true }
                    storageState = if (($recordProperties -contains 'storageState') -and $record.storageState) { [string]$record.storageState } else { 'LocalVerified' }
                    providerHint = if (($recordProperties -contains 'providerHint') -and $record.providerHint) { [string]$record.providerHint } else { 'Unknown' }
                    toolVersion = 'PhotoOrganizer 2'
                }

                $registeredPath = Get-ProcessedRecordRegisteredPath -Record $normalized
                if ([string]::IsNullOrWhiteSpace($registeredPath) -or -not (Test-Path -LiteralPath $registeredPath -PathType Leaf)) {
                    $aliasPath = Resolve-ProcessedRecordRegisteredPathWithInternalAlias -Record $normalized -RegisteredPath $registeredPath -HashKey $hashKey -Context 'Load'
                    if (-not [string]::IsNullOrWhiteSpace($aliasPath)) {
                        $registeredPath = $aliasPath
                    }
                    elseif (($normalized.PSObject.Properties.Name -contains 'status') -and ([string]$normalized.status -like 'AliasConflict*')) {
                        $script:ProcessedRecords.Add($normalized)
                        continue
                    }
                }

                if ([string]::IsNullOrWhiteSpace($registeredPath) -or -not (Test-Path -LiteralPath $registeredPath -PathType Leaf)) {
                    Set-ProcessedRecordMissing -Record $normalized
                    Add-StaleProcessedRecord -Record $normalized
                    $script:ProcessedRecords.Add($normalized)
                    $Stats.JsonReconcileStale++
                    $Stats.JsonReconcileMissing++
                    $Stats.MissingReal++
                    Write-Log -Message "Stale JSON entry detected. Hash=$hashKey. Path=$registeredPath. Reason=Registered path no longer exists" -Phase 'JSON reconciliation'
                    continue
                }

                $availability = Detect-StorageAvailability -Path $registeredPath
                if ($availability.State -eq 'CloudPlaceholder') {
                    Set-ProcessedRecordCloudPlaceholder -Record $normalized -CurrentPath $registeredPath -Availability $availability
                    $isNewCloudIndexEntry = Register-CloudPlaceholderInIndex -Hash $hashKey -Path $registeredPath
                    $script:ProcessedRecords.Add($normalized)
                    $script:ProcessedDirtyCount++
                    if ($isNewCloudIndexEntry) {
                        Write-Log -Message "Cloud placeholder JSON entry kept without hash verification. Hash=$hashKey. Path=$registeredPath. Provider=$($availability.ProviderHint). Reason=$($availability.Reason)" -Phase 'JSON reconciliation'
                    }
                    else {
                        Write-DiagnosticLog "Cloud placeholder JSON entry already counted. Hash=$hashKey. Path=$registeredPath"
                    }
                    continue
                }
                if ($availability.State -eq 'MissingReal') {
                    Set-ProcessedRecordMissing -Record $normalized
                    Add-StaleProcessedRecord -Record $normalized
                    $script:ProcessedRecords.Add($normalized)
                    $Stats.JsonReconcileStale++
                    $Stats.JsonReconcileMissing++
                    $Stats.MissingReal++
                    Write-Log -Message "Stale JSON entry detected. Hash=$hashKey. Path=$registeredPath. Reason=$($availability.Reason)" -Phase 'JSON reconciliation'
                    continue
                }

                $normalizedStatus = if ($normalized.PSObject.Properties.Name -contains 'status') { [string]$normalized.status } else { '' }
                $normalizedMissingSince = if ($normalized.PSObject.Properties.Name -contains 'missingSince') { [string]$normalized.missingSince } else { '' }
                $normalizedStorageState = if ($normalized.PSObject.Properties.Name -contains 'storageState') { [string]$normalized.storageState } else { '' }
                if ($normalizedStatus -like 'Missing/Stale*' -or $normalizedStatus -like 'CloudPlaceholderKnown*' -or $normalizedStorageState -eq 'CloudPlaceholder' -or -not [string]::IsNullOrWhiteSpace($normalizedMissingSince)) {
                    Set-ProcessedRecordCurrentPath -Record $normalized -CurrentPath $registeredPath
                    $Stats.JsonReconcilePathsUpdated++
                    $script:ProcessedDirtyCount++
                    Write-Log -Message "Previously missing JSON entry exists on disk again. Hash=$hashKey. Path=$registeredPath. Status=$($normalized.status)" -Phase 'JSON reconciliation'
                }

                $Stats.JsonReconcileValid++
                $script:ProcessedRecords.Add($normalized)
                if (Test-ProcessedRecordFeedsDuplicateIndex -Record $normalized) {
                    $script:ProcessedByHash[$hashKey] = $normalized
                }
            }
        }
        if ($script:StaleProcessedRecords.Count -gt 0) {
            $script:ProcessedDirtyCount += $script:StaleProcessedRecords.Count
            Write-Log -Message "JSON reconciliation: valid entries: $($Stats.JsonReconcileValid); stale entries: $($Stats.JsonReconcileStale); cloud placeholders in index: $($Stats.CloudPlaceholdersInIndex); paths updated: $($Stats.JsonReconcilePathsUpdated); missing marked: $($Stats.JsonReconcileMissing); entries removed: 0; conflicts: $($Stats.JsonReconcileConflicts)" -Phase 'JSON reconciliation'
        }
        Write-Log -Message "Processed database loaded: $($script:ProcessedByHash.Count) active hashes." -Phase 'Incremental'
        $jsonProperties = @($json.PSObject.Properties.Name)
        if (($jsonProperties -notcontains 'schemaVersion') -or $json.schemaVersion -ne 2) {
            $script:ProcessedDirtyCount = 1
            Save-ProcessedDatabase
        }
    }
    catch {
        Write-Log -Message "Processed database could not be read and will be ignored: $($_.Exception.Message)" -Phase 'Incremental'
        $script:ProcessedByHash = @{}
        $script:ProcessedRecords = New-Object System.Collections.Generic.List[object]
        $script:StaleProcessedRecords = New-Object System.Collections.Generic.List[object]
        $script:StaleProcessedByHash = @{}
    }
}

function Load-ProcessedIndexLight {
    if ([string]::IsNullOrWhiteSpace($ProcessedDbPath)) {
        return
    }

    if (-not (Test-Path -LiteralPath $ProcessedDbPath -PathType Leaf)) {
        Write-Log -Message "Processed database not found. A new one will be created: $ProcessedDbPath" -Phase 'Incremental'
        return
    }

    try {
        $script:ProcessedByHash = @{}
        $script:ProcessedRecords = New-Object System.Collections.Generic.List[object]
        $script:StaleProcessedRecords = New-Object System.Collections.Generic.List[object]
        $script:StaleProcessedByHash = @{}

        $json = Get-Content -LiteralPath $ProcessedDbPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $jsonProperties = @($json.PSObject.Properties.Name)
        if (($jsonProperties -contains 'schemaVersion') -and $json.schemaVersion -eq 2 -and ($jsonProperties -contains 'files') -and $json.files) {
            $records = @($json.files)
        }
        else {
            Write-Log -Message "Processed database legacy schema read in lightweight mode. Deep migration is deferred to ReconcileProcessedDatabase or another mutating maintenance run." -Phase 'Incremental'
            $records = @($json)
        }

        foreach ($record in $records) {
            if ($null -eq $record) { continue }
            $recordProperties = @($record.PSObject.Properties.Name)
            if (($recordProperties -notcontains 'hash') -or [string]::IsNullOrWhiteSpace([string]$record.hash)) {
                continue
            }

            $hashKey = ([string]$record.hash).ToUpperInvariant()
            $normalized = [pscustomobject]@{
                hash = $hashKey
                originalRelativePath = if (($recordProperties -contains 'originalRelativePath') -and $record.originalRelativePath) { [string]$record.originalRelativePath } elseif (($recordProperties -contains 'originalPath') -and $record.originalPath) { ConvertTo-RelativePath -Path ([string]$record.originalPath) -BasePath $SourcePath } else { '' }
                newRelativePath = if (($recordProperties -contains 'newRelativePath') -and $record.newRelativePath) { [string]$record.newRelativePath } elseif (($recordProperties -contains 'newPath') -and $record.newPath) { ConvertTo-RelativePath -Path ([string]$record.newPath) -BasePath $DestinationBase } else { '' }
                status = if (($recordProperties -contains 'status') -and $record.status) { [string]$record.status } else { 'LoadedLight' }
                date = if (($recordProperties -contains 'date') -and $record.date) { [string]$record.date } else { (Get-Date).ToString('o') }
                lastSeen = if (($recordProperties -contains 'lastSeen') -and $record.lastSeen) { [string]$record.lastSeen } elseif (($recordProperties -contains 'date') -and $record.date) { [string]$record.date } else { (Get-Date).ToString('o') }
                missingSince = if (($recordProperties -contains 'missingSince') -and $record.missingSince) { [string]$record.missingSince } else { '' }
                verifiedLocal = if ($recordProperties -contains 'verifiedLocal') { [bool]$record.verifiedLocal } else { $true }
                storageState = if (($recordProperties -contains 'storageState') -and $record.storageState) { [string]$record.storageState } else { 'LocalVerified' }
                providerHint = if (($recordProperties -contains 'providerHint') -and $record.providerHint) { [string]$record.providerHint } else { 'Unknown' }
                toolVersion = 'PhotoOrganizer 2'
            }

            $script:ProcessedRecords.Add($normalized)
            if (Test-ProcessedRecordFeedsDuplicateIndex -Record $normalized) {
                $script:ProcessedByHash[$hashKey] = $normalized
            }
        }

        Write-Log -Message "Processed index loaded in lightweight mode: $($script:ProcessedByHash.Count) active hashes; records: $($script:ProcessedRecords.Count). Path validation and JSON reconciliation were not performed." -Phase 'Incremental'
    }
    catch {
        Write-Log -Message "Processed database could not be read and will be ignored: $($_.Exception.Message)" -Phase 'Incremental'
        $script:ProcessedByHash = @{}
        $script:ProcessedRecords = New-Object System.Collections.Generic.List[object]
        $script:StaleProcessedRecords = New-Object System.Collections.Generic.List[object]
        $script:StaleProcessedByHash = @{}
    }
}

function Resolve-ProcessedRecordForOrganizeHash {
    param(
        [string]$HashKey,
        [object]$Record,
        [System.IO.FileInfo]$File
    )

    if ([string]::IsNullOrWhiteSpace($HashKey) -or $null -eq $Record) {
        return $null
    }

    if (-not (Test-ProcessedRecordFeedsDuplicateIndex -Record $Record)) {
        return $null
    }

    $registeredPath = Get-ProcessedRecordRegisteredPath -Record $Record
    if ([string]::IsNullOrWhiteSpace($registeredPath)) {
        Write-Log -Message "Organize ignored processed index hash match because the registered path is empty. Hash=$HashKey. CurrentFile=$($File.FullName)" -Phase 'Incremental' -Status 'Warning'
        $script:ProcessedByHash.Remove($HashKey) | Out-Null
        return $null
    }

    if (-not (Test-Path -LiteralPath $registeredPath -PathType Leaf)) {
        $aliasResolution = Resolve-InternalFolderAliasPathCandidate -RegisteredPath $registeredPath
        if ($aliasResolution.Status -eq 'Unique' -and (Test-Path -LiteralPath ([string]$aliasResolution.Path) -PathType Leaf)) {
            $registeredPath = [string]$aliasResolution.Path
            Write-Log -Message "Organize resolved one processed index alias on demand. Hash=$HashKey. OldPath=$(Get-ProcessedRecordRegisteredPath -Record $Record). CandidatePath=$registeredPath" -Phase 'Incremental'
        }
        elseif ($aliasResolution.Status -eq 'Multiple') {
            Write-Log -Message "Organize ignored ambiguous processed index hash match. Hash=$HashKey. RegisteredPath=$registeredPath. Candidates=$(@($aliasResolution.Paths) -join ' | '). Run ReconcileProcessedDatabase for a full index repair." -Phase 'Incremental' -Status 'Warning'
            $script:ProcessedByHash.Remove($HashKey) | Out-Null
            return $null
        }
        else {
            Write-Log -Message "Organize ignored stale processed index hash match. Hash=$HashKey. RegisteredPath=$registeredPath. CurrentFile=$($File.FullName). Run ReconcileProcessedDatabase to repair the index." -Phase 'Incremental' -Status 'Warning'
            $script:ProcessedByHash.Remove($HashKey) | Out-Null
            return $null
        }
    }

    $availability = Detect-StorageAvailability -Path $registeredPath
    if ($availability.State -eq 'MissingReal') {
        Write-Log -Message "Organize ignored processed index hash match because the registered path is not locally available. Hash=$HashKey. Path=$registeredPath. Reason=$($availability.Reason)" -Phase 'Incremental' -Status 'Warning'
        $script:ProcessedByHash.Remove($HashKey) | Out-Null
        return $null
    }

    if ($availability.State -eq 'CloudPlaceholder') {
        Write-DiagnosticLog "Organize kept historical hash match pointing to cloud placeholder without hydrating it. Hash=$HashKey. Path=$registeredPath. Provider=$($availability.ProviderHint). Reason=$($availability.Reason)"
    }

    return $Record
}

function Save-ProcessedDatabase {
    if ([string]::IsNullOrWhiteSpace($ProcessedDbPath)) {
        return
    }
    if ($script:ProcessedDirtyCount -le 0) {
        return
    }
    if (-not $Apply) {
        Write-DiagnosticLog "Processed database has pending changes, but -Apply is not set. Save skipped: $ProcessedDbPath"
        return
    }

    try {
        $dbDirectory = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($ProcessedDbPath))
        if (-not [string]::IsNullOrWhiteSpace($dbDirectory) -and -not (Test-Path -LiteralPath $dbDirectory)) {
            New-Item -ItemType Directory -Path $dbDirectory -Force | Out-Null
        }

        Backup-InternalJsonFilesForMutation -Paths @($ProcessedDbPath) -Reason 'ProcessedFiles save'
        $tmp = $ProcessedDbPath + '.tmp'
        $recordsArray = @($script:ProcessedRecords.ToArray())
        [pscustomobject]@{
            schemaVersion = 2
            toolVersion = 'PhotoOrganizer 2'
            updatedAt = (Get-Date).ToString('o')
            files = $recordsArray
        } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $tmp -Encoding UTF8
        Move-Item -LiteralPath $tmp -Destination $ProcessedDbPath -Force
        $script:ProcessedDirtyCount = 0
        Write-DiagnosticLog "Processed database saved: $ProcessedDbPath"
    }
    catch {
        $originalError = $_.Exception.Message
        $fallbackPath = Get-FallbackProcessedDbPath
        if ((Resolve-FullPath $ProcessedDbPath) -ne (Resolve-FullPath $fallbackPath)) {
            Write-Log -Message "Processed database save failed at $ProcessedDbPath. Falling back to $fallbackPath. Error: $originalError" -Phase 'Incremental'
            $script:ProcessedDirtyCount = [math]::Max(1, $script:ProcessedDirtyCount)
            $script:ProcessedDbPath = $fallbackPath
            Set-Variable -Name ProcessedDbPath -Scope Script -Value $fallbackPath
            try {
                $fallbackDirectory = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($ProcessedDbPath))
                if (-not (Test-Path -LiteralPath $fallbackDirectory)) {
                    New-Item -ItemType Directory -Path $fallbackDirectory -Force | Out-Null
                }
                Backup-InternalJsonFilesForMutation -Paths @($ProcessedDbPath) -Reason 'ProcessedFiles fallback save'
                $tmp = $ProcessedDbPath + '.tmp'
                [pscustomobject]@{
                    schemaVersion = 2
                    toolVersion = 'PhotoOrganizer 2'
                    updatedAt = (Get-Date).ToString('o')
                    files = @($script:ProcessedRecords.ToArray())
                } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $tmp -Encoding UTF8
                Move-Item -LiteralPath $tmp -Destination $ProcessedDbPath -Force
                $script:ProcessedDirtyCount = 0
                Write-Log -Message "Processed database saved to fallback: $ProcessedDbPath" -Phase 'Incremental'
                return
            }
            catch {
                $Stats.Errors++
                Write-Log -Message "Processed database fallback save failed: $($_.Exception.Message)" -Phase 'Incremental'
                return
            }
        }
        $Stats.Errors++
        Write-Log -Message "Processed database save failed: $originalError" -Phase 'Incremental'
    }
}

function Test-ProcessedHash {
    param([string]$Hash)
    if ([string]::IsNullOrWhiteSpace($Hash)) {
        return $false
    }
    return $script:ProcessedByHash.ContainsKey($Hash.ToUpperInvariant())
}

function Invoke-ProcessedDatabaseSmartReconciliation {
    param([object[]]$Files)

    Write-Log -Message ("ReconcileProcessedDatabase started. Mode: {0}" -f $(if ($Apply) { 'APPLY' } else { 'DRY RUN' })) -Phase 'JSON reconciliation'
    $preflightPathUpdates = $Stats.JsonReconcilePathsUpdated
    $preflightConflicts = $Stats.JsonReconcileConflicts
    $Stats.JsonReconcileMissing = 0
    $Stats.JsonReconcileEntriesRemoved = 0
    $Stats.JsonReconcilePathsUpdated = $preflightPathUpdates
    $Stats.JsonReconcileConflicts = $preflightConflicts

    if (-not $script:StaleProcessedRecords -or $script:StaleProcessedRecords.Count -eq 0) {
        if ($Apply -and $script:ProcessedDirtyCount -gt 0) {
            Save-ProcessedDatabase
        }
        elseif (-not $Apply -and $script:ProcessedDirtyCount -gt 0) {
            Write-Log -Message "DryRun only. ProcessedFiles.json has pending path/status updates but was not modified." -Phase 'JSON reconciliation'
        }
        Write-Log -Message "JSON reconciliation: valid entries: $($Stats.JsonReconcileValid); stale entries: 0; cloud placeholders in index: $($Stats.CloudPlaceholdersInIndex); paths updated: $($Stats.JsonReconcilePathsUpdated); missing marked: 0; entries removed: 0; conflicts: $($Stats.JsonReconcileConflicts)" -Phase 'JSON reconciliation' -Status 'Completed'
        return
    }

    $matchesByHash = @{}
    foreach ($hashKey in $script:StaleProcessedByHash.Keys) {
        $matchesByHash[$hashKey] = New-Object System.Collections.Generic.List[string]
    }

    $batchSizeForReconcile = [math]::Max(1, $BatchSize)
    $totalBatchesForReconcile = [int][math]::Ceiling($Files.Count / [double]$batchSizeForReconcile)
    Start-OperationalProgress -Name 'JSON reconciliation' -Total $Files.Count -Phase 'JSON reconciliation' -Message 'Hashing current files to rescue stale ProcessedFiles entries.'
    $batchNumber = 0
    for ($batchStart = 0; $batchStart -lt $Files.Count; $batchStart += $batchSizeForReconcile) {
        $batchNumber++
        $batchEnd = [math]::Min($batchStart + $batchSizeForReconcile - 1, $Files.Count - 1)
        $batch = @($Files[$batchStart..$batchEnd])
        Write-Log -Message "JSON reconciliation hashing batch $batchNumber/$totalBatchesForReconcile ($($batch.Count) files)." -Phase 'JSON reconciliation'
        Update-OperationalProgress -Current $batchStart -Total $Files.Count -Phase 'JSON reconciliation' -Stage 'Hashing current files for stale JSON reconciliation' -EveryItems ([math]::Max(1, $BatchSize * 10)) -EveryMinutes 5
        $hashResults = Get-Sha256Batch -Files $batch
        foreach ($file in $batch) {
            if (-not $hashResults.ContainsKey($file.FullName)) {
                continue
            }

            $hashResult = $hashResults[$file.FullName]
            if ($null -eq $hashResult -or [string]::IsNullOrWhiteSpace([string]$hashResult.Hash)) {
                continue
            }

            $hashKey = ([string]$hashResult.Hash).ToUpperInvariant()
            if ($matchesByHash.ContainsKey($hashKey)) {
                $matchesByHash[$hashKey].Add($file.FullName)
            }
        }
        Update-OperationalProgress -Current ($batchEnd + 1) -Total $Files.Count -Phase 'JSON reconciliation' -Stage 'Hashing current files for stale JSON reconciliation' -EveryItems ([math]::Max(1, $BatchSize * 10)) -EveryMinutes 5
    }
    Complete-OperationalProgress -Phase 'JSON reconciliation' -Message 'JSON reconciliation scan completed.'

    $newProcessedByHash = @{}
    foreach ($hashKey in $script:ProcessedByHash.Keys) {
        $newProcessedByHash[$hashKey] = $script:ProcessedByHash[$hashKey]
    }

    foreach ($hashKey in $script:StaleProcessedByHash.Keys) {
        $recordsForHash = @($script:StaleProcessedByHash[$hashKey].ToArray())
        $matches = @($matchesByHash[$hashKey].ToArray() | Select-Object -Unique)
        $oldPath = Get-ProcessedRecordRegisteredPath -Record $recordsForHash[0]

        if ($matches.Count -eq 1) {
            $record = $recordsForHash[0]
            Set-ProcessedRecordCurrentPath -Record $record -CurrentPath $matches[0]
            if (Test-ProcessedRecordFeedsDuplicateIndex -Record $record) {
                $newProcessedByHash[$hashKey] = $record
            }
            $Stats.JsonReconcilePathsUpdated++
            Write-Log -Message "JSON path would be updated. Hash=$hashKey. OldPath=$oldPath. NewPath=$($matches[0]). Status=$($record.status)" -Phase 'JSON reconciliation'
            continue
        }

        if ($matches.Count -eq 0) {
            foreach ($record in $recordsForHash) {
                Set-ProcessedRecordMissing -Record $record
            }
            $Stats.JsonReconcileMissing += $recordsForHash.Count
            Write-Log -Message "JSON entry would be kept as Missing/Stale. Hash=$hashKey. Path=$oldPath. Reason=Registered path no longer exists and hash was not found in current SourcePath" -Phase 'JSON reconciliation'
            continue
        }

        foreach ($record in $recordsForHash) {
            Set-ProcessedRecordMissing -Record $record
            $record.status = 'Missing/Stale - Multiple current paths for same hash'
        }
        $Stats.JsonReconcileConflicts++
        Write-Log -Message "Multiple current paths for same hash. Hash=$hashKey. RegisteredPath=$oldPath. Matches=$($matches -join ' | ')" -Phase 'JSON reconciliation'
    }

    if ($Apply -and ($Stats.JsonReconcilePathsUpdated + $Stats.JsonReconcileMissing + $Stats.JsonReconcileConflicts) -gt 0) {
        $script:ProcessedByHash = $newProcessedByHash
        $script:ProcessedDirtyCount = ($Stats.JsonReconcilePathsUpdated + $Stats.JsonReconcileMissing + $Stats.JsonReconcileConflicts)
        Save-ProcessedDatabase
    }
    elseif (-not $Apply) {
        $Stats.DryRunActions += ($Stats.JsonReconcilePathsUpdated + $Stats.JsonReconcileMissing)
        Write-Log -Message "DryRun only. ProcessedFiles.json was not modified." -Phase 'JSON reconciliation'
    }

    Write-Log -Message "JSON reconciliation: valid entries: $($Stats.JsonReconcileValid); stale entries: $($Stats.JsonReconcileStale); cloud placeholders in index: $($Stats.CloudPlaceholdersInIndex); paths updated: $($Stats.JsonReconcilePathsUpdated); missing marked: $($Stats.JsonReconcileMissing); entries removed: $($Stats.JsonReconcileEntriesRemoved); conflicts: $($Stats.JsonReconcileConflicts)" -Phase 'JSON reconciliation' -Status 'Completed'
}

function Get-MediaFilesForIndexMaintenanceRoots {
    param([string[]]$Roots)

    $fileByPath = @{}
    foreach ($root in @($Roots)) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        $resolvedRoot = Resolve-FullPath $root
        if (-not (Test-Path -LiteralPath $resolvedRoot -PathType Container)) { continue }
        Write-Log -Message "Index maintenance scan root: $resolvedRoot" -Phase 'JSON reconciliation'
        foreach ($file in @(Get-MediaFilesWithoutProtectedTraversal -RootPath $resolvedRoot -Phase 'JSON reconciliation')) {
            $fileByPath[(Resolve-FullPath $file.FullName).ToLowerInvariant()] = $file
        }
    }

    return @($fileByPath.Values)
}

function Invoke-ProcessedDatabasePathValidation {
    param([string]$Reason)

    Write-Log -Message "Processed database path validation started. Reason: $Reason" -Phase 'JSON reconciliation'
    $script:StaleProcessedRecords = New-Object System.Collections.Generic.List[object]
    $script:StaleProcessedByHash = @{}
    $script:ProcessedByHash = @{}
    $valid = 0
    $stale = 0
    $rescued = 0

    foreach ($record in @($script:ProcessedRecords.ToArray())) {
        if ($null -eq $record -or [string]::IsNullOrWhiteSpace([string]$record.hash)) { continue }
        $hashKey = ([string]$record.hash).ToUpperInvariant()
        $registeredPath = Get-ProcessedRecordRegisteredPath -Record $record

        if ([string]::IsNullOrWhiteSpace($registeredPath) -or -not (Test-Path -LiteralPath $registeredPath -PathType Leaf)) {
            $aliasPath = Resolve-ProcessedRecordRegisteredPathWithInternalAlias -Record $record -RegisteredPath $registeredPath -HashKey $hashKey -Context $Reason
            if (-not [string]::IsNullOrWhiteSpace($aliasPath)) {
                $registeredPath = $aliasPath
                $rescued++
            }
            elseif (($record.PSObject.Properties.Name -contains 'status') -and ([string]$record.status -like 'AliasConflict*')) {
                continue
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($registeredPath) -and (Test-Path -LiteralPath $registeredPath -PathType Leaf)) {
            $availability = Detect-StorageAvailability -Path $registeredPath
            if ($availability.State -eq 'CloudPlaceholder') {
                Set-ProcessedRecordCloudPlaceholder -Record $record -CurrentPath $registeredPath -Availability $availability
                $script:ProcessedDirtyCount++
                $isNewCloudIndexEntry = Register-CloudPlaceholderInIndex -Hash $hashKey -Path $registeredPath
                $valid++
                if ($isNewCloudIndexEntry) {
                    Write-Log -Message "Cloud placeholder JSON entry kept during path validation. Hash=$hashKey. Path=$registeredPath. Provider=$($availability.ProviderHint). Reason=$($availability.Reason)" -Phase 'JSON reconciliation'
                }
                else {
                    Write-DiagnosticLog "Cloud placeholder JSON entry already counted during path validation. Hash=$hashKey. Path=$registeredPath"
                }
                continue
            }
            if ($availability.State -eq 'MissingReal') {
                Set-ProcessedRecordMissing -Record $record
                Add-StaleProcessedRecord -Record $record
                $script:ProcessedDirtyCount++
                $stale++
                $Stats.JsonReconcileStale++
                $Stats.JsonReconcileMissing++
                $Stats.MissingReal++
                Write-Log -Message "Stale JSON entry detected after path changes. Hash=$hashKey. Path=$registeredPath. Reason=$($availability.Reason)" -Phase 'JSON reconciliation'
                continue
            }

            $status = if ($record.PSObject.Properties.Name -contains 'status') { [string]$record.status } else { '' }
            $missingSince = if ($record.PSObject.Properties.Name -contains 'missingSince') { [string]$record.missingSince } else { '' }
            $storageState = if ($record.PSObject.Properties.Name -contains 'storageState') { [string]$record.storageState } else { '' }
            if ($status -like 'Missing/Stale*' -or $status -like 'CloudPlaceholderKnown*' -or $storageState -eq 'CloudPlaceholder' -or -not [string]::IsNullOrWhiteSpace($missingSince)) {
                Set-ProcessedRecordCurrentPath -Record $record -CurrentPath $registeredPath
                $script:ProcessedDirtyCount++
                $rescued++
                $Stats.JsonReconcilePathsUpdated++
                Write-Log -Message "Previously missing JSON entry exists after path changes. Hash=$hashKey. Path=$registeredPath. Status=$($record.status)" -Phase 'JSON reconciliation'
            }
            if (Test-ProcessedRecordFeedsDuplicateIndex -Record $record) {
                $script:ProcessedByHash[$hashKey] = $record
            }
            $valid++
            continue
        }

        Set-ProcessedRecordMissing -Record $record
        Add-StaleProcessedRecord -Record $record
        $script:ProcessedDirtyCount++
        $stale++
        $Stats.JsonReconcileStale++
        $Stats.JsonReconcileMissing++
        $Stats.MissingReal++
        Write-Log -Message "Stale JSON entry detected after path changes. Hash=$hashKey. Path=$registeredPath. Reason=Registered path no longer exists" -Phase 'JSON reconciliation'
    }

    Write-Log -Message "Processed database path validation: valid entries: $valid; stale entries: $stale; rescued entries: $rescued" -Phase 'JSON reconciliation'
}

function Invoke-PostMutationProcessedDatabaseValidation {
    param(
        [string[]]$Roots,
        [string]$Reason
    )

    if (-not $Apply) { return }
    Write-Log -Message "Post-operation index validation started. Reason: $Reason" -Phase 'JSON reconciliation'
    Invoke-ProcessedDatabasePathValidation -Reason $Reason
    if ($script:StaleProcessedRecords.Count -gt 0) {
        $files = Get-MediaFilesForIndexMaintenanceRoots -Roots $Roots
        Invoke-ProcessedDatabaseSmartReconciliation -Files $files
    }
    else {
        if ($script:ProcessedDirtyCount -gt 0) {
            Save-ProcessedDatabase
        }
        Write-Log -Message "Post-operation index validation found no stale entries." -Phase 'JSON reconciliation'
    }
}
function Invoke-PurgeMissingFromProcessedDatabase {
    Write-Log -Message ("PurgeMissingFromProcessedDatabase started. Mode: {0}" -f $(if ($Apply) { 'APPLY' } else { 'DRY RUN' })) -Phase 'JSON reconciliation'

    $keptRecords = New-Object System.Collections.Generic.List[object]
    $purged = 0
    $rescued = 0
    foreach ($record in @($script:ProcessedRecords.ToArray())) {
        $status = if ($record.PSObject.Properties.Name -contains 'status') { [string]$record.status } else { '' }
        $missingSince = if ($record.PSObject.Properties.Name -contains 'missingSince') { [string]$record.missingSince } else { '' }
        if ($status -like 'CloudPlaceholderKnown*' -or ($record.PSObject.Properties.Name -contains 'storageState' -and [string]$record.storageState -eq 'CloudPlaceholder')) {
            $keptRecords.Add($record)
            $registeredPath = Get-ProcessedRecordRegisteredPath -Record $record
            $isNewCloudIndexEntry = Register-CloudPlaceholderInIndex -Hash ([string]$record.hash) -Path $registeredPath
            if ($isNewCloudIndexEntry) {
                Write-Log -Message "Cloud placeholder JSON entry kept during purge. Hash=$($record.hash). Path=$registeredPath" -Phase 'JSON reconciliation'
            }
            else {
                Write-DiagnosticLog "Cloud placeholder JSON entry already counted during purge. Hash=$($record.hash). Path=$registeredPath"
            }
            continue
        }
        $isMissing = ($status -like 'Missing/Stale*' -or -not [string]::IsNullOrWhiteSpace($missingSince))
        if ($isMissing) {
            $registeredPath = Get-ProcessedRecordRegisteredPath -Record $record
            if (-not [string]::IsNullOrWhiteSpace($registeredPath)) {
                $protection = Get-FolderProtectionRole -Path $registeredPath
                if ($protection.Role -in @('UserExcluded', 'VendorManaged')) {
                    $keptRecords.Add($record)
                    Write-Log -Message "Missing JSON entry kept because it belongs to an external protected folder. Hash=$($record.hash). Path=$registeredPath. Role=$($protection.Role). Reason=$($protection.Reason)" -Phase 'JSON reconciliation'
                    continue
                }
            }
            if (-not [string]::IsNullOrWhiteSpace($registeredPath) -and (Test-Path -LiteralPath $registeredPath -PathType Leaf)) {
                $availability = Detect-StorageAvailability -Path $registeredPath
                if ($availability.State -eq 'CloudPlaceholder') {
                    Set-ProcessedRecordCloudPlaceholder -Record $record -CurrentPath $registeredPath -Availability $availability
                    $keptRecords.Add($record)
                    $rescued++
                    $isNewCloudIndexEntry = Register-CloudPlaceholderInIndex -Hash ([string]$record.hash) -Path $registeredPath
                    if ($isNewCloudIndexEntry) {
                        Write-Log -Message "Missing JSON entry kept as cloud placeholder, not purged. Hash=$($record.hash). Path=$registeredPath. Provider=$($availability.ProviderHint). Reason=$($availability.Reason)" -Phase 'JSON reconciliation'
                    }
                    else {
                        Write-DiagnosticLog "Cloud placeholder JSON entry already counted during purge rescue. Hash=$($record.hash). Path=$registeredPath"
                    }
                    continue
                }
                Set-ProcessedRecordCurrentPath -Record $record -CurrentPath $registeredPath
                $keptRecords.Add($record)
                $rescued++
                Write-Log -Message "Missing JSON entry kept because registered path exists. Hash=$($record.hash). Path=$registeredPath. Status=$($record.status)" -Phase 'JSON reconciliation'
                continue
            }
            if ($status -like 'Missing/Stale - Multiple current paths*') {
                $keptRecords.Add($record)
                $Stats.JsonReconcileConflicts++
                Write-Log -Message "Missing JSON entry kept because hash exists in multiple current paths and requires review. Hash=$($record.hash). Path=$registeredPath. Status=$status" -Phase 'JSON reconciliation'
                continue
            }
            $purged++
            Write-Log -Message "Missing JSON entry would be purged. Hash=$($record.hash). Path=$registeredPath. MissingSince=$missingSince" -Phase 'JSON reconciliation'
            continue
        }
        $keptRecords.Add($record)
    }

    $Stats.JsonReconcilePathsUpdated += $rescued
    $Stats.JsonReconcilePurged = $purged
    $Stats.JsonReconcileEntriesRemoved = $purged

    if ($Apply -and ($purged + $rescued) -gt 0) {
        $script:ProcessedRecords = $keptRecords
        $script:ProcessedByHash = @{}
        foreach ($record in @($script:ProcessedRecords.ToArray())) {
            if ($record.hash) {
                if (Test-ProcessedRecordFeedsDuplicateIndex -Record $record) {
                    $script:ProcessedByHash[([string]$record.hash).ToUpperInvariant()] = $record
                }
            }
        }
        $script:ProcessedDirtyCount = ($purged + $rescued)
        Save-ProcessedDatabase
    }
    elseif (-not $Apply) {
        $Stats.DryRunActions += ($purged + $rescued)
        Write-Log -Message "DryRun only. Missing entries were not purged from ProcessedFiles.json." -Phase 'JSON reconciliation'
    }

    Write-Log -Message "JSON purge missing: existing entries rescued: $rescued; missing entries found: $purged; purged: $(if ($Apply) { $purged } else { 0 })" -Phase 'JSON reconciliation' -Status 'Completed'
}

function Test-ProcessedRecordMatchesCurrentPath {
    param(
        [object]$Record,
        [System.IO.FileInfo]$File
    )

    if ($null -eq $Record -or $null -eq $File) {
        return $false
    }

    $sourceRelative = ConvertTo-RelativePath -Path $File.FullName -BasePath $SourcePath
    $destinationRelative = ConvertTo-RelativePath -Path $File.FullName -BasePath $DestinationBase

    if ($Record.PSObject.Properties.Name -contains 'originalRelativePath' -and
        -not [string]::IsNullOrWhiteSpace([string]$Record.originalRelativePath) -and
        ([string]$Record.originalRelativePath).Equals($sourceRelative, [StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    if ($Record.PSObject.Properties.Name -contains 'newRelativePath' -and
        -not [string]::IsNullOrWhiteSpace([string]$Record.newRelativePath) -and
        ([string]$Record.newRelativePath).Equals($destinationRelative, [StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    return $false
}

function Register-ProcessedFile {
    param(
        [pscustomobject]$Item,
        [string]$NewPath,
        [string]$Status
    )

    if ([string]::IsNullOrWhiteSpace($Item.Sha256)) {
        return
    }

    $hashKey = $Item.Sha256.ToUpperInvariant()
    if ($script:ProcessedByHash.ContainsKey($hashKey)) {
        return
    }

    $record = [pscustomobject]@{
        hash = $hashKey
        originalRelativePath = ConvertTo-RelativePath -Path $Item.File.FullName -BasePath $SourcePath
        newRelativePath = ConvertTo-RelativePath -Path $NewPath -BasePath $DestinationBase
        date = (Get-Date).ToString('o')
        lastSeen = (Get-Date).ToString('o')
        missingSince = ''
        status = $Status
        verifiedLocal = $true
        storageState = 'LocalVerified'
        providerHint = Get-CloudProviderHint -Path $NewPath
        captureDate = if ($Item.DateInfo -and $Item.DateInfo.Date) { $Item.DateInfo.Date.ToString('o') } else { '' }
        captureDateSource = if ($Item.DateInfo) { [string]$Item.DateInfo.Source } else { '' }
        captureDateConfidence = if ($Item.DateInfo) { [int]$Item.DateInfo.Confidence } else { 0 }
        embeddedCaptureDateState = if ($Item.PSObject.Properties.Name -contains 'EmbeddedCaptureDateState') { [string]$Item.EmbeddedCaptureDateState } else { 'NotChecked' }
        captureDateMaterializationStatus = if ($Item.PSObject.Properties.Name -contains 'CaptureDateMaterializationStatus') { [string]$Item.CaptureDateMaterializationStatus } else { '' }
        embeddedCaptureMetadataWritten = if ($Item.PSObject.Properties.Name -contains 'EmbeddedCaptureMetadataWritten') { [bool]$Item.EmbeddedCaptureMetadataWritten } else { $false }
        fileSystemDatesSynced = if ($Item.PSObject.Properties.Name -contains 'FileSystemDatesSynced') { [bool]$Item.FileSystemDatesSynced } else { $false }
        dateKnownButMetadataNotWritten = if ($Item.PSObject.Properties.Name -contains 'DateKnownButMetadataNotWritten') { [bool]$Item.DateKnownButMetadataNotWritten } else { $false }
        toolVersion = 'PhotoOrganizer 2'
    }

    $script:ProcessedByHash[$hashKey] = $record
    $script:ProcessedRecords.Add($record)
    $script:ProcessedDirtyCount++

    if ($script:ProcessedDirtyCount -ge 100) {
        Save-ProcessedDatabase
    }
}

function Register-ImportedProviderFile {
    param(
        [pscustomobject]$Item,
        [string]$NewPath,
        [string]$Status,
        [string]$ProviderName,
        [string]$ProviderRootPath
    )

    if ([string]::IsNullOrWhiteSpace($Item.Sha256)) {
        return
    }

    $hashKey = $Item.Sha256.ToUpperInvariant()
    if ($script:ProcessedByHash.ContainsKey($hashKey)) {
        return
    }

    $record = [pscustomobject]@{
        hash = $hashKey
        originalRelativePath = ConvertTo-RelativePath -Path $Item.File.FullName -BasePath $ProviderRootPath
        newRelativePath = ConvertTo-RelativePath -Path $NewPath -BasePath $DestinationBase
        date = (Get-Date).ToString('o')
        lastSeen = (Get-Date).ToString('o')
        missingSince = ''
        status = $Status
        verifiedLocal = $true
        storageState = 'LocalVerified'
        providerHint = $ProviderName
        importProvider = $ProviderName
        occurrenceCount = if ($Item.PSObject.Properties.Name -contains 'ProviderOccurrenceCount') { [int]$Item.ProviderOccurrenceCount } else { 1 }
        albumNames = if ($Item.PSObject.Properties.Name -contains 'ProviderAlbumNames') { @($Item.ProviderAlbumNames) } else { @() }
        metadataConfidence = if ($Item.PSObject.Properties.Name -contains 'ProviderMetadataConfidence') { [string]$Item.ProviderMetadataConfidence } else { '' }
        exifVerification = if ($Item.PSObject.Properties.Name -contains 'ProviderExifVerification') { [string]$Item.ProviderExifVerification } else { '' }
        captureDate = if ($Item.DateInfo -and $Item.DateInfo.Date) { $Item.DateInfo.Date.ToString('o') } else { '' }
        captureDateSource = if ($Item.DateInfo) { [string]$Item.DateInfo.Source } else { '' }
        captureDateConfidence = if ($Item.DateInfo) { [int]$Item.DateInfo.Confidence } else { 0 }
        embeddedCaptureDateState = if ($Item.PSObject.Properties.Name -contains 'EmbeddedCaptureDateState') { [string]$Item.EmbeddedCaptureDateState } else { 'NotChecked' }
        captureDateMaterializationStatus = if ($Item.PSObject.Properties.Name -contains 'CaptureDateMaterializationStatus') { [string]$Item.CaptureDateMaterializationStatus } else { '' }
        embeddedCaptureMetadataWritten = if ($Item.PSObject.Properties.Name -contains 'EmbeddedCaptureMetadataWritten') { [bool]$Item.EmbeddedCaptureMetadataWritten } else { $false }
        fileSystemDatesSynced = if ($Item.PSObject.Properties.Name -contains 'FileSystemDatesSynced') { [bool]$Item.FileSystemDatesSynced } else { $false }
        dateKnownButMetadataNotWritten = if ($Item.PSObject.Properties.Name -contains 'DateKnownButMetadataNotWritten') { [bool]$Item.DateKnownButMetadataNotWritten } else { $false }
        toolVersion = 'PhotoOrganizer 2'
    }

    $script:ProcessedByHash[$hashKey] = $record
    $script:ProcessedRecords.Add($record)
    $script:ProcessedDirtyCount++

    if ($script:ProcessedDirtyCount -ge 100) {
        Save-ProcessedDatabase
    }
}

function Initialize-DestinationStructure {
    if (-not $Apply) {
        return
    }

    Write-Log -Message "Creating destination structure..." -Phase 'Moving'
    $directories = @($OrganizedRoot, $NeedsReviewRoot, $MediaMetadataIssuesRoot, $DuplicatesRoot, $ConfirmedDuplicatesQuarantineRoot)
    if ($RepairExif -or $MetadataRepair) {
        $directories += $MetadataBackupRoot
    }

    foreach ($directory in $directories) {
        if (-not (Test-Path -LiteralPath $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }
    }
}

$JunkOnlyFileNames = @('desktop.ini', 'Thumbs.db', 'ehthumbs.db', '.DS_Store', '.picasa.ini', '.tmp', '.dropbox')
$JunkOnlyDirectoryNames = @('.Trashes', '.fseventsd', '.Spotlight-V100')
$JunkProtectedExtensions = @(
    '.jpg', '.jpeg', '.png', '.heic', '.webp', '.gif', '.bmp', '.tif', '.tiff',
    '.mp4', '.mov', '.avi', '.mkv', '.3gp',
    '.dng', '.raw', '.cr2', '.cr3', '.nef', '.arw', '.rw2', '.orf',
    '.json', '.xmp', '.txt', '.pdf', '.doc', '.docx', '.url', '.lnk',
    '.html', '.htm', '.xml', '.csv'
)

function Test-JunkOnlyFile {
    param([System.IO.FileInfo]$File)

    if ($null -eq $File) { return $false }
    $extension = [string]$File.Extension
    if ($JunkProtectedExtensions -contains $extension.ToLowerInvariant()) {
        return $false
    }

    if (($JunkOnlyFileNames -contains $File.Name) -or $File.Name -like '._*' -or $File.Name -like 'Icon?' -or $extension.Equals('.tmp', [StringComparison]::OrdinalIgnoreCase)) {
        return $true
    }

    if ($File.Length -ge 16KB) {
        return $false
    }

    $isHiddenOrSystem = (($File.Attributes -band [System.IO.FileAttributes]::Hidden) -eq [System.IO.FileAttributes]::Hidden) -or
        (($File.Attributes -band [System.IO.FileAttributes]::System) -eq [System.IO.FileAttributes]::System)
    $hasNoExtension = [string]::IsNullOrWhiteSpace($extension)
    $looksLikeMarker = $File.Name.StartsWith('.', [StringComparison]::OrdinalIgnoreCase)

    return ($isHiddenOrSystem -or $hasNoExtension -or $looksLikeMarker)
}

function Get-SafeFolderCleanupState {
    param([string]$Path)

    try {
        $children = @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction Stop)
    }
    catch {
        return [pscustomobject]@{
            CanRemove = $false
            IsJunkOnly = $false
            IsSmallMarkerOnly = $false
            Error = $_.Exception.Message
        }
    }

    if ($children.Count -eq 0) {
        return [pscustomobject]@{
            CanRemove = $true
            IsJunkOnly = $false
            IsSmallMarkerOnly = $false
            Error = ''
        }
    }

    $files = @($children | Where-Object { -not $_.PSIsContainer })
    $directories = @($children | Where-Object { $_.PSIsContainer })
    foreach ($file in $files) {
        if ($JunkProtectedExtensions -contains ([string]$file.Extension).ToLowerInvariant()) {
            return [pscustomobject]@{
                CanRemove = $false
                IsJunkOnly = $false
                IsSmallMarkerOnly = $false
                Error = ''
            }
        }
    }

    $hasSmallMarker = $false
    foreach ($file in $files) {
        if (-not (Test-JunkOnlyFile -File $file)) {
            return [pscustomobject]@{
                CanRemove = $false
                IsJunkOnly = $false
                IsSmallMarkerOnly = $false
                Error = ''
            }
        }

        $extension = [string]$file.Extension
        $isKnownJunk = (($JunkOnlyFileNames -contains $file.Name) -or $file.Name -like '._*' -or $file.Name -like 'Icon?' -or $extension.Equals('.tmp', [StringComparison]::OrdinalIgnoreCase))
        if (-not $isKnownJunk) {
            $hasSmallMarker = $true
        }
    }

    foreach ($directory in $directories) {
        if (-not ($JunkOnlyDirectoryNames -contains $directory.Name)) {
            return [pscustomobject]@{
                CanRemove = $false
                IsJunkOnly = $false
                IsSmallMarkerOnly = $false
                Error = ''
            }
        }

        $childState = Get-SafeFolderCleanupState -Path $directory.FullName
        if (-not $childState.CanRemove) {
            return [pscustomobject]@{
                CanRemove = $false
                IsJunkOnly = $false
                IsSmallMarkerOnly = $false
                Error = ''
            }
        }
        if ($childState.IsSmallMarkerOnly) {
            $hasSmallMarker = $true
        }
    }

    if (($files.Count + $directories.Count) -gt 0) {
        return [pscustomobject]@{
            CanRemove = $true
            IsJunkOnly = $true
            IsSmallMarkerOnly = $hasSmallMarker
            Error = ''
        }
    }

    return [pscustomobject]@{
        CanRemove = $false
        IsJunkOnly = $false
        IsSmallMarkerOnly = $false
        Error = ''
    }
}

function Test-IsOrganizedCleanupRoot {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $resolvedPath = (Resolve-FullPath $Path).TrimEnd('\')
    $organizedRoots = New-Object System.Collections.Generic.List[string]
    foreach ($root in @($OrganizedRoot)) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        $resolvedRoot = (Resolve-FullPath $root).TrimEnd('\')
        if (-not $organizedRoots.Contains($resolvedRoot)) {
            $organizedRoots.Add($resolvedRoot)
        }
    }
    if ($script:DedupeOrganizedRoots) {
        foreach ($root in @($script:DedupeOrganizedRoots)) {
            if ([string]::IsNullOrWhiteSpace($root)) { continue }
            $resolvedRoot = (Resolve-FullPath $root).TrimEnd('\')
            if (-not $organizedRoots.Contains($resolvedRoot)) {
                $organizedRoots.Add($resolvedRoot)
            }
        }
    }

    return [bool](@($organizedRoots.ToArray()) | Where-Object { $resolvedPath.Equals($_, [StringComparison]::OrdinalIgnoreCase) })
}

function Test-IsCleanupProtectedPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $true }
    $resolvedPath = (Resolve-FullPath $Path).TrimEnd('\')

    foreach ($root in @($SourcePath, $OrganizedRoot)) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        if ($resolvedPath.Equals((Resolve-FullPath $root).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    foreach ($excludedRoot in @($ExcludedRoots)) {
        if ([string]::IsNullOrWhiteSpace($excludedRoot)) { continue }
        $resolvedExcluded = (Resolve-FullPath $excludedRoot).TrimEnd('\')
        if (Test-IsOrganizedCleanupRoot -Path $resolvedExcluded) {
            if ($resolvedPath.Equals($resolvedExcluded, [StringComparison]::OrdinalIgnoreCase)) {
                return $true
            }
            continue
        }
        if (Test-IsChildPath -Path $resolvedPath -ParentPath $resolvedExcluded) {
            return $true
        }
    }

    return $false
}

function Remove-SafeEmptyOrJunkOnlyFolder {
    param(
        [string]$Path,
        [string]$Phase = 'Cleaning folders',
        [switch]$NormalizeZombie
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return $false }
    if (Test-IsCleanupProtectedPath -Path $Path) { return $false }

    $resolvedPath = (Resolve-FullPath $Path).TrimEnd('\')
    $state = Get-SafeFolderCleanupState -Path $Path
    if (-not $state.CanRemove) {
        if (-not [string]::IsNullOrWhiteSpace($state.Error)) {
            Write-DiagnosticLog "Folder cleanup skipped for ${Path}: $($state.Error)"
        }
        return $false
    }

    $kind = if ($state.IsJunkOnly) { 'junk-only folder' } else { 'empty folder' }
    if (-not $Apply) {
        $Stats.DryRunActions++
        Write-Log -Message "Would remove ${kind}: $Path" -Phase $Phase
        return $true
    }

    try {
        Remove-Item -LiteralPath $Path -Recurse -Force
        $Stats.EmptyFoldersRemoved++
        if ($state.IsJunkOnly) {
            $Stats.JunkOnlyFoldersRemoved++
        }
        if ($state.IsSmallMarkerOnly) {
            $Stats.JunkOnlySmallMarkerFoldersRemoved++
        }
        if ($NormalizeZombie) {
            $Stats.ZombieNormalizeFoldersRemoved++
        }
        Write-Log -Message "Removed ${kind}: $Path" -Phase $Phase
        return $true
    }
    catch {
        Write-DiagnosticLog "Folder cleanup skipped for ${Path}: $($_.Exception.Message)"
        return $false
    }
}

function Remove-EmptySourceFolders {
    if ($KeepEmptyFolders) {
        Write-Log -Message "KeepEmptyFolders enabled. Empty folder cleanup skipped." -Phase 'Cleaning folders'
        return
    }

    Write-Log -Message "Cleaning empty folders..." -Phase 'Cleaning folders'
    $cleanupRoots = New-Object System.Collections.Generic.List[string]
    foreach ($root in @($SourcePath, $OrganizedRoot)) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        $resolvedRoot = Resolve-FullPath $root
        if (-not (Test-Path -LiteralPath $resolvedRoot -PathType Container)) { continue }

        $coveredByExistingRoot = $false
        foreach ($existingRoot in @($cleanupRoots.ToArray())) {
            if (Test-IsChildPath -Path $resolvedRoot -ParentPath $existingRoot) {
                $coveredByExistingRoot = $true
                break
            }
        }
        if ($coveredByExistingRoot) {
            continue
        }

        foreach ($existingRoot in @($cleanupRoots.ToArray())) {
            if (Test-IsChildPath -Path $existingRoot -ParentPath $resolvedRoot) {
                $cleanupRoots.Remove($existingRoot) | Out-Null
            }
        }
        if (-not $cleanupRoots.Contains($resolvedRoot)) {
            $cleanupRoots.Add($resolvedRoot)
        }
    }

    $directories = [ordered]@{}
    foreach ($root in @($cleanupRoots.ToArray())) {
        Get-ChildItem -LiteralPath $root -Directory -Recurse -Force -ErrorAction SilentlyContinue |
            ForEach-Object {
                $key = (Resolve-FullPath $_.FullName).TrimEnd('\').ToLowerInvariant()
                if (-not $directories.Contains($key)) {
                    $directories[$key] = $_
                }
            }
    }

    Start-OperationalProgress -Name 'Folder cleanup' -Total $directories.Count -Phase 'Cleaning folders' -Message 'Removing empty, junk-only and zombie branches.'
    $cleanupIndex = 0
    foreach ($directory in @($directories.Values | Sort-Object FullName -Descending)) {
        $cleanupIndex++
        Update-OperationalProgress -Current $cleanupIndex -Total $directories.Count -Phase 'Cleaning folders' -Stage 'Cleaning empty/junk-only folders' -EveryItems 1000 -EveryMinutes 5
        Remove-SafeEmptyOrJunkOnlyFolder -Path $directory.FullName -Phase 'Cleaning folders' | Out-Null
    }
    Complete-OperationalProgress -Phase 'Cleaning folders' -Message 'Folder cleanup phase completed.'
}

function Remove-EmptyFolderChain {
    param(
        [string]$StartPath,
        [string]$StopRoot,
        [string]$Phase = 'NormalizeExistingFolders'
    )

    if ([string]::IsNullOrWhiteSpace($StartPath)) { return }
    $current = Resolve-FullPath $StartPath
    $stop = if ([string]::IsNullOrWhiteSpace($StopRoot)) { '' } else { (Resolve-FullPath $StopRoot).TrimEnd('\') }

    while (-not [string]::IsNullOrWhiteSpace($current) -and (Test-Path -LiteralPath $current -PathType Container)) {
        $resolvedCurrent = (Resolve-FullPath $current).TrimEnd('\')
        if (-not [string]::IsNullOrWhiteSpace($stop) -and $resolvedCurrent.Equals($stop, [StringComparison]::OrdinalIgnoreCase)) {
            break
        }

        $removed = Remove-SafeEmptyOrJunkOnlyFolder -Path $current -Phase $Phase -NormalizeZombie
        if (-not $removed -or -not $Apply) {
            break
        }

        $parent = [System.IO.Path]::GetDirectoryName($resolvedCurrent)
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent.Equals($resolvedCurrent, [StringComparison]::OrdinalIgnoreCase)) {
            break
        }
        $current = $parent
    }
}

function ConvertTo-SafeName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $NoDescriptionText
    }
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $safe = $Name.Trim()
    foreach ($char in $invalid) {
        $safe = $safe.Replace([string]$char, ' ')
    }
    $safe = ($safe -replace '\s+', ' ').Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return $NoDescriptionText
    }
    if ($safe.Length -gt 90) {
        return $safe.Substring(0, 90).Trim()
    }
    return $safe
}

function ConvertTo-MediaDate {
    param([object]$Value)
    if ($null -eq $Value) {
        return $null
    }

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $null
    }

    $normalized = $text.Trim()
    $normalized = $normalized -replace '^(\d{4}):(\d{2}):(\d{2})', '$1-$2-$3'
    $normalized = $normalized -replace 'T', ' '
    $normalized = $normalized -replace 'Z$', ''
    $normalized = $normalized -replace '([+-]\d{2}):?(\d{2})$', ''

    $formats = @(
        'yyyy-MM-dd HH:mm:ss',
        'yyyy-MM-dd HH:mm',
        'yyyy-MM-dd',
        'yyyyMMdd HHmmss',
        'yyyyMMdd',
        'dd-MM-yyyy HH:mm:ss',
        'dd-MM-yyyy',
        'dd.MM.yyyy HH:mm:ss',
        'dd.MM.yyyy'
    )

    foreach ($format in $formats) {
        try {
            return [datetime]::ParseExact(
                $normalized,
                $format,
                [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::AssumeLocal
            )
        }
        catch {
        }
    }

    try {
        return [datetime]::Parse($normalized, [Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        return $null
    }
}

function New-FilenameDateInfo {
    param(
        [System.Text.RegularExpressions.Match]$Match,
        [string]$Kind,
        [int]$Confidence,
        [string]$Pattern,
        [bool]$HasReliableTime,
        [bool]$SyntheticTime
    )

    try {
        $year = [int]$Match.Groups['y'].Value
        $month = [int]$Match.Groups['m'].Value
        $day = [int]$Match.Groups['d'].Value
        $hour = if ($Match.Groups['h'].Success -and $Match.Groups['h'].Value) { [int]$Match.Groups['h'].Value } else { 12 }
        $minute = if ($Match.Groups['min'].Success -and $Match.Groups['min'].Value) { [int]$Match.Groups['min'].Value } else { 0 }
        $second = if ($Match.Groups['s'].Success -and $Match.Groups['s'].Value) { [int]$Match.Groups['s'].Value } else { 0 }
        $date = Get-Date -Year $year -Month $month -Day $day -Hour $hour -Minute $minute -Second $second
        return [pscustomobject]@{
            Date = $date
            Kind = $Kind
            Confidence = $Confidence
            Pattern = $Pattern
            HasReliableTime = $HasReliableTime
            SyntheticTime = $SyntheticTime
        }
    }
    catch {
        return $null
    }
}

function Get-DateInfoFromFileName {
    param([string]$FileName)

    $name = [System.IO.Path]::GetFileNameWithoutExtension($FileName)

    $reliablePatterns = @(
        @{ Regex = '^(?:(?:20\d{2}|19\d{2})-[01]\d-[0-3]\d(?: [0-2]\d[0-5]\d)? - )?(?:IMG|VID)-(?<y>20\d{2}|19\d{2})(?<m>0[1-9]|1[0-2])(?<d>0[1-9]|[12]\d|3[01])-WA\d+(?:[-_][\p{L}\p{N}]+)*(?: \(\d+\))?$'; Kind = 'ReliableDateOnly'; Confidence = 97; Pattern = 'StructuredDateOnly'; HasReliableTime = $false; SyntheticTime = $true },
        @{ Regex = '^(?:IMG|VID)_(?<y>20\d{2}|19\d{2})(?<m>0[1-9]|1[0-2])(?<d>0[1-9]|[12]\d|3[01])[_-](?<h>[0-2]\d)(?<min>[0-5]\d)(?<s>[0-5]\d)(?:\D|$)'; Kind = 'ReliableDateTime'; Confidence = 98; Pattern = 'StructuredDateTime'; HasReliableTime = $true; SyntheticTime = $false },
        @{ Regex = '^PXL_(?<y>20\d{2}|19\d{2})(?<m>0[1-9]|1[0-2])(?<d>0[1-9]|[12]\d|3[01])[_-](?<h>[0-2]\d)(?<min>[0-5]\d)(?<s>[0-5]\d)\d*(?:\D|$)'; Kind = 'ReliableDateTime'; Confidence = 98; Pattern = 'StructuredDateTime'; HasReliableTime = $true; SyntheticTime = $false },
        @{ Regex = '^(?<y>20\d{2}|19\d{2})(?<m>0[1-9]|1[0-2])(?<d>0[1-9]|[12]\d|3[01])[_-](?<h>[0-2]\d)(?<min>[0-5]\d)(?<s>[0-5]\d)(?:\D|$)'; Kind = 'ReliableDateTime'; Confidence = 98; Pattern = 'StructuredDateTime'; HasReliableTime = $true; SyntheticTime = $false }
    )

    foreach ($patternInfo in $reliablePatterns) {
        $m = [regex]::Match($name, [string]$patternInfo.Regex, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if ($m.Success) {
            return New-FilenameDateInfo -Match $m -Kind ([string]$patternInfo.Kind) -Confidence ([int]$patternInfo.Confidence) -Pattern ([string]$patternInfo.Pattern) -HasReliableTime ([bool]$patternInfo.HasReliableTime) -SyntheticTime ([bool]$patternInfo.SyntheticTime)
        }
    }

    $genericPatterns = @(
        '(?<y>20\d{2}|19\d{2})(?<m>0[1-9]|1[0-2])(?<d>0[1-9]|[12]\d|3[01])[_\-\s]?(?<h>[0-2]\d)?(?<min>[0-5]\d)?(?<s>[0-5]\d)?',
        '(?<y>20\d{2}|19\d{2})[-_.](?<m>0?[1-9]|1[0-2])[-_.](?<d>0?[1-9]|[12]\d|3[01])(?!\d)\D+(?<h>[0-2]?\d)(?<min>[0-5]\d)(?<s>[0-5]\d)?',
        '(?<y>20\d{2}|19\d{2})[-_.](?<m>0?[1-9]|1[0-2])[-_.](?<d>0?[1-9]|[12]\d|3[01])(?!\d)(?:\D+(?<h>[0-2]?\d)[-.hH_](?<min>[0-5]\d)(?:[-.mM_](?<s>[0-5]\d))?)?',
        '(?<d>0?[1-9]|[12]\d|3[01])(?!\d)[-_.](?<m>0?[1-9]|1[0-2])[-_.](?<y>20\d{2}|19\d{2})'
    )

    foreach ($pattern in $genericPatterns) {
        $m = [regex]::Match($name, $pattern)
        if ($m.Success) {
            $hasTime = ($m.Groups['h'].Success -and $m.Groups['h'].Value)
            return New-FilenameDateInfo -Match $m -Kind 'GenericDate' -Confidence 96 -Pattern 'GenericDate' -HasReliableTime ([bool]$hasTime) -SyntheticTime (-not [bool]$hasTime)
        }
    }

    return $null
}

function Get-DateFromFileName {
    param([string]$FileName)

    $dateInfo = Get-DateInfoFromFileName -FileName $FileName
    if ($dateInfo) {
        return $dateInfo.Date
    }

    return $null
}

function Get-DateFromFolderName {
    param([string]$Path)

    $parts = @($Path -split '[\\/]')
    for ($i = $parts.Count - 1; $i -ge 0; $i--) {
        $candidate = $parts[$i]
        $dt = Get-DateFromFileName -FileName $candidate
        if ($null -ne $dt) {
            return $dt
        }

        $m = [regex]::Match($candidate, '(?<y>20\d{2}|19\d{2})\D+(?<m>0?[1-9]|1[0-2])')
        if ($m.Success) {
            try {
                return Get-Date -Year ([int]$m.Groups['y'].Value) -Month ([int]$m.Groups['m'].Value) -Day 1 -Hour 12
            }
            catch {
            }
        }
    }
    return $null
}

function Get-ExifMetadata {
    param(
        [string]$Path,
        [bool]$ExifToolAvailable,
        [int]$TimeoutSeconds = 30
    )

    $metadata = [ordered]@{
        DateTimeOriginal = $null
        CreateDate = $null
        MediaCreateDate = $null
        XMPDateCreated = $null
        XMPCreateDate = $null
        XMPModifyDate = $null
        PNGCreationTime = $null
        TrackCreateDate = $null
        TrackModifyDate = $null
        QuickTimeCreateDate = $null
        QuickTimeModifyDate = $null
        GPSLatitude = $null
        GPSLongitude = $null
        ImageWidth = $null
        ImageHeight = $null
        Warning = $null
        ExifScore = 0
        Raw = $null
        ReadStatus = 'NotChecked'
    }

    if (-not $ExifToolAvailable) {
        $metadata.ReadStatus = 'Unreadable'
        return [pscustomobject]$metadata
    }

    $availability = Detect-StorageAvailability -Path $Path
    if ($availability.State -eq 'CloudPlaceholder') {
        Register-CloudPlaceholderSkipped -Path $Path -Phase 'EXIF' -Availability $availability
        $metadata.ReadStatus = 'Unreadable'
        return [pscustomobject]$metadata
    }
    if ($availability.State -eq 'MissingReal') {
        $Stats.MissingReal++
        Write-Log -Message "EXIF skipped missing file: $Path. Reason=$($availability.Reason)" -Phase 'EXIF'
        $metadata.ReadStatus = 'Unreadable'
        return [pscustomobject]$metadata
    }

    try {
        $readArguments = @('-j', '-n') + $VisibleCaptureMetadataReadTags + @(
            '-GPSLatitude', '-GPSLongitude', '-ImageWidth', '-ImageHeight',
            '-FileType', '-MIMEType', '-Warning', $Path
        )
        $exif = Invoke-ExifTool -Path $ExifToolPath -Arguments $readArguments -TimeoutSeconds $TimeoutSeconds
        if ($exif.TimedOut) {
            Register-SlowExifCandidate -Path $Path -Reason 'timeout' -Detail "Per-file ExifTool timeout after $TimeoutSeconds seconds" -Seconds $exif.DurationSeconds -BatchNumber $script:CurrentBatch
            Register-ExifProblemFile -Path $Path -Reason "Slow EXIF file detected" -Detail "ExifTool timeout after $TimeoutSeconds seconds"
            $metadata.ReadStatus = 'Unreadable'
            return [pscustomobject]$metadata
        }
        if ($exif.DurationSeconds -ge 10) {
            Register-SlowExifCandidate -Path $Path -Reason 'slow per-file read' -Detail 'Per-file EXIF read exceeded 10 seconds' -Seconds $exif.DurationSeconds -BatchNumber $script:CurrentBatch
        }
        if (-not $exif.Success) {
            if (Test-ExifProblemText -Text $exif.Error) {
                Register-SlowExifCandidate -Path $Path -Reason 'metadata warning' -Detail $exif.Error -Seconds $exif.DurationSeconds -BatchNumber $script:CurrentBatch
                Register-ExifProblemFile -Path $Path -Reason "Media corruption / WhatsApp metadata issue" -Detail $exif.Error
            }
            Write-DiagnosticLog "ExifTool metadata read failed for ${Path}: $($exif.Error)"
            $metadata.ReadStatus = 'Unreadable'
            return [pscustomobject]$metadata
        }

        if (Test-ExifProblemText -Text $exif.Error) {
            Register-SlowExifCandidate -Path $Path -Reason 'metadata warning' -Detail $exif.Error -Seconds $exif.DurationSeconds -BatchNumber $script:CurrentBatch
            Register-ExifProblemFile -Path $Path -Reason "Media corruption / WhatsApp metadata issue" -Detail $exif.Error
            $metadata.ReadStatus = 'Unreadable'
            return [pscustomobject]$metadata
        }

        $json = $exif.Output
        if ([string]::IsNullOrWhiteSpace($json)) {
            $metadata.ReadStatus = 'Unreadable'
            return [pscustomobject]$metadata
        }

        $raw = ($json | ConvertFrom-Json)[0]
        $metadata.Raw = $raw
        $metadata.ReadStatus = 'Read'
        foreach ($field in @($VisibleCaptureMetadataFields + @('GPSLatitude', 'GPSLongitude', 'ImageWidth', 'ImageHeight', 'Warning'))) {
            if ($raw.PSObject.Properties.Name -contains $field) {
                $metadata[$field] = $raw.$field
            }
        }
        if ($raw.PSObject.Properties.Name -contains 'DateCreated') { $metadata['XMPDateCreated'] = $raw.DateCreated }
        if ($raw.PSObject.Properties.Name -contains 'CreationTime') { $metadata['PNGCreationTime'] = $raw.CreationTime }

        if (Test-ExifProblemText -Text ([string]$metadata.Warning)) {
            Register-SlowExifCandidate -Path $Path -Reason 'metadata warning' -Detail ([string]$metadata.Warning) -Seconds $exif.DurationSeconds -BatchNumber $script:CurrentBatch
            Register-ExifProblemFile -Path $Path -Reason "Media corruption / WhatsApp metadata issue" -Detail ([string]$metadata.Warning)
            $metadata.ReadStatus = 'Unreadable'
            return [pscustomobject]$metadata
        }

        if ($metadata.DateTimeOriginal) { $metadata.ExifScore += 4 }
        if ($metadata.CreateDate) { $metadata.ExifScore += 2 }
        if ($metadata.MediaCreateDate) { $metadata.ExifScore += 2 }
        if ($metadata['XMPDateCreated'] -or $metadata['XMPCreateDate'] -or $metadata['PNGCreationTime'] -or $metadata['TrackCreateDate']) { $metadata.ExifScore += 2 }
        if ($metadata.GPSLatitude -and $metadata.GPSLongitude) { $metadata.ExifScore += 2 }
        if ($metadata.ImageWidth -and $metadata.ImageHeight) { $metadata.ExifScore += 1 }
    }
    catch {
        $metadata.ReadStatus = 'Unreadable'
    }

    return [pscustomobject]$metadata
}

function New-EmptyMetadata {
    param([string]$ReadStatus = 'NotChecked')

    return [pscustomobject]([ordered]@{
        DateTimeOriginal = $null
        CreateDate = $null
        MediaCreateDate = $null
        XMPDateCreated = $null
        XMPCreateDate = $null
        XMPModifyDate = $null
        PNGCreationTime = $null
        TrackCreateDate = $null
        TrackModifyDate = $null
        QuickTimeCreateDate = $null
        QuickTimeModifyDate = $null
        GPSLatitude = $null
        GPSLongitude = $null
        ImageWidth = $null
        ImageHeight = $null
        Warning = $null
        ExifScore = 0
        Raw = $null
        ReadStatus = $ReadStatus
    })
}

function Convert-ExifRawToMetadata {
    param([object]$Raw)

    $metadata = [ordered]@{
        DateTimeOriginal = $null
        CreateDate = $null
        MediaCreateDate = $null
        XMPDateCreated = $null
        XMPCreateDate = $null
        XMPModifyDate = $null
        PNGCreationTime = $null
        TrackCreateDate = $null
        TrackModifyDate = $null
        QuickTimeCreateDate = $null
        QuickTimeModifyDate = $null
        GPSLatitude = $null
        GPSLongitude = $null
        ImageWidth = $null
        ImageHeight = $null
        Warning = $null
        ExifScore = 0
        Raw = $Raw
        ReadStatus = 'Read'
    }

    foreach ($field in @($VisibleCaptureMetadataFields + @('GPSLatitude', 'GPSLongitude', 'ImageWidth', 'ImageHeight', 'Warning'))) {
        if ($Raw.PSObject.Properties.Name -contains $field) {
            $metadata[$field] = $Raw.$field
        }
    }
    if ($Raw.PSObject.Properties.Name -contains 'DateCreated') { $metadata['XMPDateCreated'] = $Raw.DateCreated }
    if ($Raw.PSObject.Properties.Name -contains 'CreationTime') { $metadata['PNGCreationTime'] = $Raw.CreationTime }

    if ($metadata.DateTimeOriginal) { $metadata.ExifScore += 4 }
    if ($metadata.CreateDate) { $metadata.ExifScore += 2 }
    if ($metadata.MediaCreateDate) { $metadata.ExifScore += 2 }
    if ($metadata['XMPDateCreated'] -or $metadata['XMPCreateDate'] -or $metadata['PNGCreationTime'] -or $metadata['TrackCreateDate']) { $metadata.ExifScore += 2 }
    if ($metadata.GPSLatitude -and $metadata.GPSLongitude) { $metadata.ExifScore += 2 }
    if ($metadata.ImageWidth -and $metadata.ImageHeight) { $metadata.ExifScore += 1 }

    return [pscustomobject]$metadata
}

function Get-ExifMetadataBatch {
    param(
        [object[]]$Files,
        [bool]$ExifToolAvailable,
        [switch]$ShowImportProviderProgress
    )

    $result = @{}
    foreach ($file in $Files) {
        $result[$file.FullName] = New-EmptyMetadata
    }

    if (-not $ExifToolAvailable -or $Files.Count -eq 0) {
        if (-not $ExifToolAvailable) {
            foreach ($file in $Files) {
                $result[$file.FullName].ReadStatus = 'Unreadable'
            }
        }
        return $result
    }

    $localFiles = New-Object System.Collections.Generic.List[object]
    foreach ($file in $Files) {
        $availability = Detect-StorageAvailability -Item $file
        if ($availability.State -eq 'CloudPlaceholder') {
            Register-CloudPlaceholderSkipped -Path $file.FullName -Phase 'EXIF' -Availability $availability
            $result[$file.FullName].ReadStatus = 'Unreadable'
            continue
        }
        if ($availability.State -eq 'MissingReal') {
            $Stats.MissingReal++
            Write-Log -Message "EXIF batch skipped missing file: $($file.FullName). Reason=$($availability.Reason)" -Phase 'EXIF'
            $result[$file.FullName].ReadStatus = 'Unreadable'
            continue
        }
        $localFiles.Add($file)
    }

    if ($localFiles.Count -eq 0) {
        return $result
    }

    $Files = @($localFiles.ToArray())
    $progressStartedAt = Get-Date
    $script:ImportProviderExifProgress = $null
    if ($ShowImportProviderProgress) {
        Write-ImportProviderExifVerificationProgress -Current 0 -Total $Files.Count -StartedAt $progressStartedAt -Force
    }

    $arguments = @('-j', '-n') + $VisibleCaptureMetadataReadTags + @(
        '-GPSLatitude', '-GPSLongitude', '-ImageWidth', '-ImageHeight',
        '-FileType', '-MIMEType', '-Warning'
    )
    $arguments += @($Files | ForEach-Object { $_.FullName })

    $timeout = [math]::Max(1, $ExifBatchTimeoutSeconds)
    $exif = Invoke-ExifTool -Path $ExifToolPath -Arguments $arguments -TimeoutSeconds $timeout
    if ($exif.TimedOut) {
        Write-Log -Message "EXIF batch timeout after $timeout seconds. Switching immediately to per-file mode with 30-second timeout." -Phase 'Processing'
        Register-ExifBatchTimeout -AffectedFiles $Files.Count -Seconds $exif.DurationSeconds -BatchNumber $script:CurrentBatch -Detail 'Falling back to per-file metadata reads'
        $processed = 0
        $progressStartedAt = Get-Date
        $script:ImportProviderExifProgress = $null
        foreach ($file in $Files) {
            $result[$file.FullName] = Get-ExifMetadata -Path $file.FullName -ExifToolAvailable $ExifToolAvailable -TimeoutSeconds 30
            $processed++
            if ($ShowImportProviderProgress) {
                Write-ImportProviderExifVerificationProgress -Current $processed -Total $Files.Count -StartedAt $progressStartedAt -CurrentFile $file.FullName
            }
        }
        if ($ShowImportProviderProgress) {
            Write-ImportProviderExifVerificationProgress -Current $Files.Count -Total $Files.Count -StartedAt $progressStartedAt -Force
            $script:OperationalProgress = $null
        }
        return $result
    }
    if (-not $exif.Success) {
        Write-Log -Message "ExifTool batch failed. Falling back to per-file metadata reads: $($exif.Error)" -Phase 'Processing'
        Register-ExifBatchFallback -AffectedFiles $Files.Count -Seconds $exif.DurationSeconds -BatchNumber $script:CurrentBatch -Reason 'Batch EXIF read failed'
        $processed = 0
        $progressStartedAt = Get-Date
        $script:ImportProviderExifProgress = $null
        foreach ($file in $Files) {
            $result[$file.FullName] = Get-ExifMetadata -Path $file.FullName -ExifToolAvailable $ExifToolAvailable -TimeoutSeconds 30
            $processed++
            if ($ShowImportProviderProgress) {
                Write-ImportProviderExifVerificationProgress -Current $processed -Total $Files.Count -StartedAt $progressStartedAt -CurrentFile $file.FullName
            }
        }
        if ($ShowImportProviderProgress) {
            Write-ImportProviderExifVerificationProgress -Current $Files.Count -Total $Files.Count -StartedAt $progressStartedAt -Force
            $script:OperationalProgress = $null
        }
        return $result
    }

    if (Test-ExifProblemText -Text $exif.Error) {
        $mappedProblem = $false
        foreach ($file in $Files) {
            if ($exif.Error.IndexOf($file.FullName, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                $exif.Error.IndexOf($file.Name, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                Register-SlowExifCandidate -Path $file.FullName -Reason 'metadata warning' -Detail $exif.Error -Seconds $exif.DurationSeconds -BatchNumber $script:CurrentBatch
                Register-ExifProblemFile -Path $file.FullName -Reason "Media corruption / WhatsApp metadata issue" -Detail $exif.Error
                $mappedProblem = $true
            }
        }

        if (-not $mappedProblem) {
            Write-Log -Message "ExifTool batch reported corrupt/partial media but did not identify a file clearly. Falling back to per-file metadata reads." -Phase 'Processing'
            Register-ExifBatchFallback -AffectedFiles $Files.Count -Seconds $exif.DurationSeconds -BatchNumber $script:CurrentBatch -Reason 'Batch warning did not identify a single file'
            $processed = 0
            $progressStartedAt = Get-Date
            $script:ImportProviderExifProgress = $null
            foreach ($file in $Files) {
                $result[$file.FullName] = Get-ExifMetadata -Path $file.FullName -ExifToolAvailable $ExifToolAvailable -TimeoutSeconds 30
                $processed++
                if ($ShowImportProviderProgress) {
                    Write-ImportProviderExifVerificationProgress -Current $processed -Total $Files.Count -StartedAt $progressStartedAt -CurrentFile $file.FullName
                }
            }
            if ($ShowImportProviderProgress) {
                Write-ImportProviderExifVerificationProgress -Current $Files.Count -Total $Files.Count -StartedAt $progressStartedAt -Force
                $script:OperationalProgress = $null
            }
            return $result
        }
    }

    try {
        $rows = @($exif.Output | ConvertFrom-Json)
        foreach ($row in $rows) {
            if ($row.SourceFile) {
                $full = Resolve-FullPath ([string]$row.SourceFile)
                $metadata = Convert-ExifRawToMetadata -Raw $row
                $result[$full] = $metadata
                if (Test-ExifProblemText -Text ([string]$metadata.Warning)) {
                    Register-ExifProblemFile -Path $full -Reason "Media corruption / WhatsApp metadata issue" -Detail ([string]$metadata.Warning)
                }
            }
        }
        if ($ShowImportProviderProgress) {
            Write-ImportProviderExifVerificationProgress -Current $Files.Count -Total $Files.Count -StartedAt $progressStartedAt -Force
            $script:OperationalProgress = $null
        }
    }
    catch {
        Write-Log -Message "ExifTool batch JSON parse failed: $($_.Exception.Message)" -Phase 'Processing'
        if ($ShowImportProviderProgress) {
            Write-ImportProviderExifVerificationProgress -Current $Files.Count -Total $Files.Count -StartedAt $progressStartedAt -Force
            $script:OperationalProgress = $null
        }
    }

    foreach ($file in $Files) {
        if ([string]$result[$file.FullName].ReadStatus -eq 'NotChecked') {
            $result[$file.FullName].ReadStatus = 'Unreadable'
        }
    }

    return $result
}

function Get-PrimaryDate {
    param(
        [pscustomobject]$Item,
        [bool]$IsVideo
    )

    $sources = New-Object System.Collections.Generic.List[object]

    $dateTimeOriginal = ConvertTo-MediaDate $Item.Metadata.DateTimeOriginal
    if ($dateTimeOriginal) {
        $sources.Add([pscustomobject]@{ Source = 'EXIF DateTimeOriginal'; Date = $dateTimeOriginal; Confidence = 99 })
    }

    $createDate = ConvertTo-MediaDate $Item.Metadata.CreateDate
    if ($createDate) {
        $sources.Add([pscustomobject]@{ Source = 'EXIF CreateDate'; Date = $createDate; Confidence = 98 })
    }

    if ($IsVideo) {
        $mediaCreateDate = ConvertTo-MediaDate $Item.Metadata.MediaCreateDate
        if ($mediaCreateDate) {
            $sources.Add([pscustomobject]@{ Source = 'MediaCreateDate'; Date = $mediaCreateDate; Confidence = 98 })
        }
        $trackCreateDate = ConvertTo-MediaDate $Item.Metadata.TrackCreateDate
        if ($trackCreateDate) {
            $sources.Add([pscustomobject]@{ Source = 'TrackCreateDate'; Date = $trackCreateDate; Confidence = 98 })
        }
    }

    foreach ($visibleField in @(
        @{ Name = 'XMPDateCreated'; Source = 'XMP DateCreated'; Confidence = 98 },
        @{ Name = 'XMPCreateDate'; Source = 'XMP CreateDate'; Confidence = 98 },
        @{ Name = 'PNGCreationTime'; Source = 'PNG CreationTime'; Confidence = 98 }
    )) {
        if ($Item.Metadata.PSObject.Properties.Name -contains $visibleField.Name) {
            $visibleDate = ConvertTo-MediaDate $Item.Metadata.($visibleField.Name)
            if ($visibleDate) {
                $sources.Add([pscustomobject]@{ Source = $visibleField.Source; Date = $visibleDate; Confidence = [int]$visibleField.Confidence })
            }
        }
    }

    $fileNameDateInfo = Get-DateInfoFromFileName -FileName $Item.File.Name
    $fileNameDate = if ($fileNameDateInfo) { $fileNameDateInfo.Date } else { $null }
    if ($fileNameDateInfo) {
        $sources.Add([pscustomobject]@{
            Source = 'FileName'
            Date = $fileNameDateInfo.Date
            Confidence = $fileNameDateInfo.Confidence
            FilenameDateKind = $fileNameDateInfo.Kind
            FilenameDatePattern = $fileNameDateInfo.Pattern
            FilenameDateHasReliableTime = $fileNameDateInfo.HasReliableTime
            FilenameDateSyntheticTime = $fileNameDateInfo.SyntheticTime
        })
    }

    $filenameConflictsWithExif = $false
    if ($dateTimeOriginal -and $fileNameDate -and $dateTimeOriginal.Date -ne $fileNameDate.Date) {
        $filenameConflictsWithExif = $true
        Write-Log -Message ("Filename date conflicts with EXIF. EXIF used. File={0}; EXIF={1:yyyy-MM-dd}; Filename={2:yyyy-MM-dd}" -f $Item.File.FullName, $dateTimeOriginal, $fileNameDate) -Phase 'Decision queue'
    }

    $folderDate = Get-DateFromFolderName -Path $Item.File.DirectoryName
    if ($folderDate) {
        $sources.Add([pscustomobject]@{ Source = 'OriginalFolder'; Date = $folderDate; Confidence = 90 })
    }

    if ($sources.Count -eq 0) {
        return [pscustomobject]@{
            Date = $Item.File.LastWriteTime
            Confidence = 50
            Source = 'WindowsLastWriteTime'
            Sources = @('WindowsLastWriteTime')
            FilenameDateConflictsWithExif = $false
        }
    }

    $best = $sources | Sort-Object Confidence -Descending | Select-Object -First 1
    $supporting = @($sources | Where-Object {
        $_.Source -ne $best.Source -and [math]::Abs(($_.Date.Date - $best.Date.Date).TotalDays) -le 1
    })

    $confidence = [int]$best.Confidence
    if ($supporting.Count -gt 0 -and $confidence -lt 99) {
        $confidence = [math]::Min(99, $confidence + 2)
    }

    return [pscustomobject]@{
        Date = $best.Date
        Confidence = $confidence
        Source = $best.Source
        Sources = @($sources | ForEach-Object { $_.Source })
        FilenameDateKind = if ($fileNameDateInfo) { $fileNameDateInfo.Kind } else { $null }
        FilenameDatePattern = if ($fileNameDateInfo) { $fileNameDateInfo.Pattern } else { $null }
        FilenameDateHasReliableTime = if ($fileNameDateInfo) { [bool]$fileNameDateInfo.HasReliableTime } else { $false }
        FilenameDateSyntheticTime = if ($fileNameDateInfo) { [bool]$fileNameDateInfo.SyntheticTime } else { $false }
        FilenameDateConflictsWithExif = $filenameConflictsWithExif
    }
}

function Convert-DateInfoSourceForDiagnostics {
    param([string]$Source)

    if ([string]::IsNullOrWhiteSpace($Source)) { return 'Fallback' }

    switch -Regex ($Source) {
        '^(EXIF DateTimeOriginal|EXIF CreateDate|MediaCreateDate)$' { return 'EXIF' }
        '^EXIF fallback:' { return 'EXIF' }
        '^FileName$|^File name$' { return 'Filename' }
        '^OriginalFolder$|^Current organized path$|^Current path \+ file name$|^Current path conflicts with file name$' { return 'FolderInference' }
        '^ProcessedFiles path$' { return 'HistoricalProcessedRecord' }
        '^GoogleTakeout$' { return 'GoogleTakeout' }
        '^XmpSidecarLibrary$' { return 'ImportProvider' }
        '^MigrationReconcile$' { return 'MigrationReconcile' }
        '^WindowsLastWriteTime$' { return 'Fallback' }
        default { return $Source }
    }
}

function Write-DateInfoDiagnostic {
    param(
        [System.IO.FileInfo]$File,
        [pscustomobject]$DateInfo,
        [string]$Context
    )

    if (-not $Diagnostic -or $null -eq $DateInfo) {
        return
    }

    $finalDate = if ($DateInfo.Date) { $DateInfo.Date.ToString('yyyy-MM-dd HH:mm:ss') } else { 'null' }
    $rawSource = if ($DateInfo.Source) { [string]$DateInfo.Source } else { 'Unknown' }
    $source = Convert-DateInfoSourceForDiagnostics -Source $rawSource
    $confidence = if ($DateInfo.PSObject.Properties.Name -contains 'Confidence') { [string]$DateInfo.Confidence } else { 'Unknown' }
    $sources = if ($DateInfo.PSObject.Properties.Name -contains 'Sources' -and $DateInfo.Sources) { (@($DateInfo.Sources) -join ',') } else { $rawSource }
    $detail = if ($DateInfo.PSObject.Properties.Name -contains 'Detail' -and $DateInfo.Detail) { [string]$DateInfo.Detail } else { '' }
    $conflict = if ($DateInfo.PSObject.Properties.Name -contains 'FilenameDateConflictsWithExif') { [bool]$DateInfo.FilenameDateConflictsWithExif } else { $false }

    Write-DiagnosticLog ("DateInfo resolved: Context={0}; Path={1}; FinalDate={2}; Source={3}; RawSource={4}; Confidence={5}; Sources={6}; FilenameConflict={7}; Detail={8}" -f $Context, $File.FullName, $finalDate, $source, $rawSource, $confidence, $sources, $conflict, $detail)
}

function Get-ContentTypeName {
    param(
        [string]$Path,
        [string]$Extension,
        [bool]$IsVideo
    )

    $lower = $Path.ToLowerInvariant()
    if ($lower -match 'whatsapp') { return 'WhatsApp' }
    if ($lower -match 'screenshot|captura|screen shot') { return 'Screenshots' }
    if ($IsVideo) { return 'Videos' }
    if ($lower -match 'camera roll|dcim|img_|pxl_|dsc_|photo') { return 'Camera Roll' }
    return $null
}

function Get-DestinationPath {
    param(
        [pscustomobject]$Item,
        [string]$RootPath,
        [string]$OrganizedRoot
    )

    if ($Item.DateInfo.Confidence -lt $ExifRepairConfidence) {
        return $NeedsReviewRoot
    }

    $date = $Item.DateInfo.Date
    $targetFileName = Get-QuarterlyTargetFileName -Item $Item
    $Item | Add-Member -NotePropertyName TargetFileName -NotePropertyValue $targetFileName -Force
    return Get-QuarterlyDestinationDirectory -Date $date -OrganizedRoot $OrganizedRoot
}

function Get-QuarterFolderName {
    param([datetime]$Date)

    $quarter = [int][math]::Ceiling($Date.Month / 3.0)
    $names = if ($QuarterFolderNamesByLanguage.ContainsKey($Language)) { $QuarterFolderNamesByLanguage[$Language] } else { $QuarterFolderNamesByLanguage['es'] }
    return $names[$quarter]
}

function Get-QuarterlyDestinationDirectory {
    param(
        [datetime]$Date,
        [string]$OrganizedRoot
    )

    return Join-Path (Join-Path $OrganizedRoot ('{0:yyyy}' -f $Date)) (Get-QuarterFolderName -Date $Date)
}

function Test-DateInfoHasReliableTime {
    param([pscustomobject]$DateInfo)

    if ($null -eq $DateInfo -or $null -eq $DateInfo.Date) {
        return $false
    }

    if ($DateInfo.Date.TimeOfDay.TotalSeconds -le 0) {
        return $false
    }

    if (@('EXIF DateTimeOriginal', 'EXIF CreateDate', 'MediaCreateDate', 'TrackCreateDate', 'XMP DateCreated', 'XMP CreateDate', 'PNG CreationTime', 'GoogleTakeout', 'ApplePhotos', 'XmpSidecarLibrary') -contains [string]$DateInfo.Source) {
        return $true
    }

    if ([string]$DateInfo.Source -eq 'FileName' -and
        $DateInfo.PSObject.Properties.Name -contains 'FilenameDateKind' -and
        [string]$DateInfo.FilenameDateKind -eq 'ReliableDateTime') {
        return $true
    }

    if ([string]$DateInfo.Source -eq 'FileName' -and ($DateInfo.Date.Hour -ne 12 -or $DateInfo.Date.Minute -ne 0 -or $DateInfo.Date.Second -ne 0)) {
        return $true
    }

    return $false
}

function Get-QuarterlyOriginalFileName {
    param([string]$FileName)

    $extension = [System.IO.Path]::GetExtension($FileName)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $cleanBase = ConvertTo-SafeName $baseName
    $clean = if ([string]::IsNullOrWhiteSpace($extension)) { $cleanBase } else { $cleanBase + $extension }
    $clean = [regex]::Replace($clean, '^\d{4}-\d{2}-\d{2}(?:\s+\d{4})?\s+-\s+', '')
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $FileName
    }
    return $clean
}

function Remove-RedundantCaptureDatePrefixFromFileName {
    param([string]$FileName)

    $extension = [System.IO.Path]::GetExtension($FileName)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    if ([string]::IsNullOrWhiteSpace($baseName)) {
        return $FileName
    }

    $cleanBase = ConvertTo-SafeName $baseName
    $patterns = @(
        '^(?<date>\d{4}-\d{2}-\d{2})\s+\d{4}\s+-\s+\k<date>(?=\s+\d{4,6}(?:\b|$)|\s|$)',
        '^(?<date>\d{4}-\d{2}-\d{2})\s+-\s+\k<date>(?=\s+\d{4,6}(?:\b|$)|\s|$)',
        '^(?<date>\d{4}-\d{2}-\d{2})\s+\k<date>(?=\s+\d{4,6}(?:\b|$)|\s|$)'
    )

    foreach ($pattern in $patterns) {
        $updated = ([regex]$pattern).Replace($cleanBase, '${date}', 1)
        if (-not $updated.Equals($cleanBase, [StringComparison]::Ordinal)) {
            $cleanBase = ($updated -replace '\s{2,}', ' ').Trim()
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($cleanBase)) {
        return $FileName
    }

    return $cleanBase + $extension
}

function Test-FileNameStartsWithSameCaptureDate {
    param(
        [string]$FileName,
        [datetime]$Date
    )

    $baseName = ConvertTo-SafeName ([System.IO.Path]::GetFileNameWithoutExtension($FileName))
    if ([string]::IsNullOrWhiteSpace($baseName)) {
        return $false
    }

    $dateText = [regex]::Escape(('{0:yyyy-MM-dd}' -f $Date))
    if ($baseName -match ("^$dateText(?:\s+\d{4,6})?(?:\b|\s|$)")) {
        return $true
    }

    $filenameDateInfo = Get-DateInfoFromFileName -FileName $FileName
    if ($filenameDateInfo -and
        @('ReliableDateTime', 'ReliableDateOnly') -contains [string]$filenameDateInfo.Kind -and
        $filenameDateInfo.Date.Date -eq $Date.Date) {
        return $true
    }

    return $false
}

function New-QuarterlyFileName {
    param(
        [datetime]$Date,
        [string]$OriginalName,
        [bool]$IncludeTime
    )

    $extension = [System.IO.Path]::GetExtension($OriginalName)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($OriginalName)
    $safeBase = ConvertTo-SafeName $baseName
    $cleanOriginalName = Remove-RedundantCaptureDatePrefixFromFileName -FileName ($safeBase + $extension)
    if (Test-FileNameStartsWithSameCaptureDate -FileName $cleanOriginalName -Date $Date) {
        return $cleanOriginalName
    }

    $prefix = if ($IncludeTime) { '{0:yyyy-MM-dd HHmm}' -f $Date } else { '{0:yyyy-MM-dd}' -f $Date }
    return ('{0} - {1}' -f $prefix, $cleanOriginalName)
}

function Get-QuarterlyTargetFileName {
    param([pscustomobject]$Item)

    $date = $Item.DateInfo.Date
    $originalName = Get-QuarterlyOriginalFileName -FileName $Item.File.Name
    if (Test-DateInfoHasReliableTime -DateInfo $Item.DateInfo) {
        return New-QuarterlyFileName -Date $date -OriginalName $originalName -IncludeTime $true
    }

    $Stats.SkippedUncertainNames++
    return New-QuarterlyFileName -Date $date -OriginalName $originalName -IncludeTime $false
}

function Test-KnownFolderContainerExists {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }

    try {
        if ([System.IO.Directory]::Exists($Path)) {
            return $true
        }
    }
    catch {
    }

    try {
        if (Test-Path -LiteralPath $Path) {
            $item = Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
            return [bool]($item -and $item.PSIsContainer)
        }
    }
    catch {
    }

    return $false
}

function Invoke-KnownFolderContainerRename {
    param(
        [string]$OldPath,
        [string]$TargetPath,
        [string]$Label
    )

    $oldResolved = Resolve-FullPath $OldPath
    $targetResolved = Resolve-FullPath $TargetPath
    if ($oldResolved.TrimEnd('\').Equals($targetResolved.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
        return 0
    }

    if (-not (Test-KnownFolderContainerExists -Path $oldResolved)) {
        return 0
    }

    $availability = Detect-StorageAvailability -Path $oldResolved -Directory
    $cloudNote = if ($availability.State -eq 'CloudPlaceholder') {
        " Cloud-backed/container rename allowed; file content will not be read or hydrated. Provider=$($availability.ProviderHint); Reason=$($availability.Reason)."
    }
    else {
        ''
    }

    if (Test-Path -LiteralPath $targetResolved) {
        if (-not (Test-KnownFolderContainerExists -Path $targetResolved)) {
            Write-Log -Message "$Label rename skipped because target path exists and is not a folder: $targetResolved" -Phase 'Validation'
            return 0
        }

        $children = @(Get-ChildItem -LiteralPath $oldResolved -Force -ErrorAction SilentlyContinue)
        $conflicts = @(
            foreach ($child in $children) {
                $childTarget = Join-Path $targetResolved $child.Name
                if (Test-Path -LiteralPath $childTarget) {
                    $childTarget
                }
            }
        )

        if ($conflicts.Count -gt 0) {
            Write-Log -Message "$Label consolidation skipped because target exists and child name conflicts were found: $oldResolved -> $targetResolved. Conflicts: $($conflicts -join '; ')" -Phase 'Validation'
            return 0
        }

        if (-not $Apply) {
            Write-Log -Message "Would consolidate $Label container into existing target: $oldResolved -> $targetResolved. Immediate children to move: $($children.Count). File content will not be read or hydrated.$cloudNote" -Phase 'Validation'
            return 1
        }

        try {
            foreach ($child in $children) {
                Move-Item -LiteralPath $child.FullName -Destination (Join-Path $targetResolved $child.Name)
            }
            Remove-Item -LiteralPath $oldResolved -Force
            Write-Log -Message "Consolidated $Label container into existing target: $oldResolved -> $targetResolved. Immediate children moved: $($children.Count).$cloudNote" -Phase 'Validation'
            return 1
        }
        catch {
            $Stats.Errors++
            Write-Log -Message "$Label consolidation failed: $oldResolved -> $targetResolved. Error: $($_.Exception.Message)" -Phase 'Validation'
            return 0
        }
    }

    if (-not $Apply) {
        Write-Log -Message "Would rename $Label container: $oldResolved -> $targetResolved.$cloudNote" -Phase 'Validation'
        return 1
    }

    try {
        Move-Item -LiteralPath $oldResolved -Destination $targetResolved
        Write-Log -Message "Renamed $Label container: $oldResolved -> $targetResolved.$cloudNote" -Phase 'Validation'
        return 1
    }
    catch {
        $Stats.Errors++
        Write-Log -Message "$Label rename failed: $oldResolved -> $targetResolved. Error: $($_.Exception.Message)" -Phase 'Validation'
        return 0
    }
}

function Rename-InternalFoldersToCurrentLanguage {
    if (-not $RenameInternalFoldersToCurrentLanguage) {
        return
    }

    $modeText = if ($Apply) { 'APPLY' } else { 'DRY RUN' }
    Write-Log -Message "Renaming internal folders to current language: $Language. Mode=$modeText. Known folder container rename only; file content is not scanned." -Phase 'Validation'
    $renamed = 0
    $aliasesScanned = 0
    $aliasesDetected = 0
    $renameKeys = @('OrganizedFolder', 'NeedsReviewFolder', 'DuplicatesFolder', 'ConfirmedDuplicatesQuarantineFolder', 'MetadataBackupFolder', 'LogsFolder')
    $localAppDataBase = if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { Join-Path $env:LOCALAPPDATA 'PhotoOrganizer' } else { '' }
    $bases = @($DestinationBase, $SourcePath, $OrganizedRoot, $localAppDataBase) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { Resolve-FullPath $_ } | Select-Object -Unique

    foreach ($base in $bases) {
        if (-not (Test-KnownFolderContainerExists -Path $base)) {
            Write-Log -Message "Internal folder alias scan skipped because base folder does not exist or is not visible: $base" -Phase 'Validation'
            continue
        }

        $immediateFolders = @()
        try {
            $immediateFolders = @(Get-ChildItem -LiteralPath $base -Directory -Force -ErrorAction Stop)
            Write-Log -Message "Internal folder alias scan base: $base. Immediate folders: $($immediateFolders.Count)." -Phase 'Validation'
        }
        catch {
            Write-Log -Message "Internal folder alias scan could not enumerate immediate folders under base: $base. Error: $($_.Exception.Message)" -Phase 'Validation'
        }

        foreach ($key in $renameKeys) {
            if ($key -eq 'OrganizedFolder' -and (Resolve-FullPath $base).TrimEnd('\').Equals((Resolve-FullPath $OrganizedRoot).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
                continue
            }

            $targetName = Get-InternalFolderName -Key $key
            foreach ($oldName in (Get-AllInternalFolderNames -Key $key)) {
                if ([string]::IsNullOrWhiteSpace($oldName) -or $oldName -eq $targetName) { continue }
                $aliasesScanned++

                $oldPath = Resolve-FullPath (Join-Path $base $oldName)
                $targetPath = Resolve-FullPath (Join-Path $base $targetName)

                if (-not (Test-KnownFolderContainerExists -Path $oldPath)) { continue }
                $aliasesDetected++
                Write-Log -Message "Detected internal folder alias: $oldPath -> $targetPath" -Phase 'Validation'
                if ($oldPath.TrimEnd('\').Equals((Resolve-FullPath $SourcePath).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
                    Write-Log -Message "Internal folder rename skipped because it is the selected source: $oldPath" -Phase 'Validation'
                    continue
                }

                $renamed += Invoke-KnownFolderContainerRename -OldPath $oldPath -TargetPath $targetPath -Label 'internal folder'
            }
        }

        $targetNeedsReviewName = Get-InternalFolderName -Key 'NeedsReviewFolder'
        foreach ($needsReviewName in (Get-AllInternalFolderNames -Key 'NeedsReviewFolder')) {
            foreach ($mediaName in (Get-AllInternalFolderNames -Key 'MediaMetadataIssuesFolder')) {
                $targetMediaName = Get-InternalFolderName -Key 'MediaMetadataIssuesFolder'
                if ([string]::IsNullOrWhiteSpace($mediaName) -or $mediaName -eq $targetMediaName) { continue }

                foreach ($needsContainerName in @($needsReviewName, $targetNeedsReviewName) | Select-Object -Unique) {
                    $aliasesScanned++
                    $oldMediaPath = Resolve-FullPath (Join-Path (Join-Path $base $needsContainerName) $mediaName)
                    $targetMediaPath = Resolve-FullPath (Join-Path (Join-Path $base $needsContainerName) $targetMediaName)
                    if (Test-KnownFolderContainerExists -Path $oldMediaPath) {
                        $aliasesDetected++
                        Write-Log -Message "Detected internal media metadata folder alias: $oldMediaPath -> $targetMediaPath" -Phase 'Validation'
                    }
                    $renamed += Invoke-KnownFolderContainerRename -OldPath $oldMediaPath -TargetPath $targetMediaPath -Label 'internal media metadata folder'
                }
            }
        }
    }

    $verb = if ($Apply) { 'Renamed folders' } else { 'Planned folder renames' }
    Write-Log -Message ("Internal folder aliases scanned: {0}; detected aliases: {1}" -f $aliasesScanned, $aliasesDetected) -Phase 'Validation'
    if ($aliasesDetected -eq 0) {
        Write-Log -Message "No alias folders found under scanned roots. Checked SourcePath, DestinationBase, OrganizedRoot and local metadata backup base." -Phase 'Validation'
    }
    Write-Log -Message ("Internal folder language rename complete. {0}: {1}" -f $verb, $renamed) -Phase 'Validation'
}

function Rename-ExistingFoldersToCurrentLanguage {
    if (-not $RenameExistingFoldersToCurrentLanguage) {
        return
    }

    if (-not (Test-Path -LiteralPath $OrganizedRoot -PathType Container)) {
        Write-Log -Message "RenameExistingFoldersToCurrentLanguage requested, but organized root does not exist yet." -Phase 'Validation'
        return
    }

    $modeText = if ($Apply) { 'APPLY' } else { 'DRY RUN' }
    Write-Log -Message "Renaming existing quarterly folders to current language: $Language. Mode=$modeText. Directory-level rename only; file content is not scanned." -Phase 'Validation'
    $renamed = 0

    try {
        $yearFolders = @(Get-ChildItem -LiteralPath $OrganizedRoot -Directory -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^\d{4}$' })
        foreach ($yearFolder in $yearFolders) {
            $quarterFolders = @(Get-ChildItem -LiteralPath $yearFolder.FullName -Directory -Force -ErrorAction SilentlyContinue)
            foreach ($quarterFolder in $quarterFolders) {
                $startMonth = Get-QuarterStartMonthFromFolderName -FolderName $quarterFolder.Name
                if (-not $startMonth) { continue }
                $targetName = Get-QuarterFolderName -Date (Get-Date -Year ([int]$yearFolder.Name) -Month $startMonth -Day 1)
                if ($quarterFolder.Name -eq $targetName) { continue }
                $targetPath = Join-Path $yearFolder.FullName $targetName
                $renamed += Invoke-KnownFolderContainerRename -OldPath $quarterFolder.FullName -TargetPath $targetPath -Label 'quarter folder'
            }
        }
    }
    catch {
        $Stats.Errors++
        Write-Log -Message "Existing folder rename failed: $($_.Exception.Message)" -Phase 'Validation'
    }

    $verb = if ($Apply) { 'Renamed folders' } else { 'Planned folder renames' }
    Write-Log -Message ("Existing folder language rename complete. {0}: {1}" -f $verb, $renamed) -Phase 'Validation'
}

function Test-FolderLanguageRenameOnlyMode {
    if (-not ($RenameExistingFoldersToCurrentLanguage -or $RenameInternalFoldersToCurrentLanguage)) {
        return $false
    }

    return -not (
        $RetentionCleanup -or
        -not [string]::IsNullOrWhiteSpace($ImportProvider) -or
        $DedupeCleanup -or
        $RecoverFromWrongDuplicateMove -or
        $NormalizeExistingFolders -or
        $RepairOnlyExistingOrganizedLibrary -or
        $ReconcileProcessedDatabase -or
        $PurgeMissingFromProcessedDatabase -or
        $TestScan
    )
}

function Resolve-UniquePath {
    param(
        [string]$Directory,
        [string]$FileName,
        [string]$SourceHash
    )

    if (-not (Test-Path -LiteralPath $Directory)) {
        return [pscustomobject]@{
            Path = (Join-Path $Directory $FileName)
            SkipExistingIdentical = $false
        }
    }

    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $ext = [System.IO.Path]::GetExtension($FileName)
    $candidate = Join-Path $Directory $FileName
    $index = 1

    while (Test-Path -LiteralPath $candidate) {
        try {
            $existingHash = Get-Sha256 -Path $candidate
            if ($existingHash -eq $SourceHash) {
                return [pscustomobject]@{
                    Path = $candidate
                    SkipExistingIdentical = $true
                }
            }
        }
        catch {
        }

        $candidate = Join-Path $Directory ('{0} ({1}){2}' -f $base, $index, $ext)
        $index++
    }

    return [pscustomobject]@{
        Path = $candidate
        SkipExistingIdentical = $false
    }
}

function Test-LockedFileError {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)

    if ($null -eq $ErrorRecord) { return $false }
    $message = [string]$ErrorRecord.Exception.Message
    $hresult = 0
    try { $hresult = $ErrorRecord.Exception.HResult } catch { }

    if ($hresult -eq -2147024864) { return $true } # Win32 sharing violation
    return ($message -match '(?i)being used by another process|used by another process|cannot access the file|process cannot access|sharing violation|esta siendo usado por otro proceso|está siendo usado por otro proceso|otro proceso')
}

function Release-OwnHandles {
    try {
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
    }
    catch {
    }
}

function Write-ExifToolProcessState {
    param([string]$FilePath)

    try {
        $running = @(Get-Process -Name 'exiftool*' -ErrorAction SilentlyContinue)
        if ($running.Count -gt 0) {
            Write-DiagnosticLog ("ExifTool still running while file is locked ({0} processes). File: {1}" -f $running.Count, $FilePath)
        }
    }
    catch {
    }
}

function Invoke-TransferOperation {
    param(
        [pscustomobject]$Item,
        [string]$DestinationPath
    )

    if ($CopyInsteadOfMove) {
        Copy-Item -LiteralPath $Item.File.FullName -Destination $DestinationPath -ErrorAction Stop
    }
    else {
        Move-Item -LiteralPath $Item.File.FullName -Destination $DestinationPath -ErrorAction Stop
    }
}

function Invoke-TransferWithLockRetry {
    param(
        [pscustomobject]$Item,
        [string]$DestinationPath,
        [string]$Reason
    )

    $delays = @(500, 2000, 5000, 10000)
    for ($attempt = 0; $attempt -lt ($delays.Count + 1); $attempt++) {
        try {
            Invoke-TransferOperation -Item $Item -DestinationPath $DestinationPath
            return $true
        }
        catch {
            if (-not (Test-LockedFileError -ErrorRecord $_)) {
                throw
            }

            if ($attempt -ge $delays.Count) {
                $Stats.LockedFiles++
                Write-Log -Message ("Locked file persisted after {0} attempts. File will be sent to NeedsReview when possible. Path={1}. Error={2}" -f ($delays.Count + 1), $Item.File.FullName, $_.Exception.Message) -Phase 'Moving'
                return $false
            }

            $Stats.RetryCount++
            Write-ExifToolProcessState -FilePath $Item.File.FullName
            Release-OwnHandles
            $delay = $delays[$attempt]
            Write-Log -Message ("Locked file retry {0}/4 after {1}ms. File={2}. Reason={3}" -f ($attempt + 1), $delay, $Item.File.FullName, $Reason) -Phase 'Moving'
            Start-Sleep -Milliseconds $delay
        }
    }

    return $false
}

function Remove-SourceAfterVerifiedIdenticalTarget {
    param(
        [pscustomobject]$Item,
        [string]$TargetPath,
        [string]$Reason
    )

    if (-not $Apply -or $CopyInsteadOfMove) {
        return $false
    }

    $sourcePath = Resolve-FullPath $Item.File.FullName
    $targetFullPath = Resolve-FullPath $TargetPath
    if ($sourcePath.TrimEnd('\').Equals($targetFullPath.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) {
        Write-Log -Message "Source cleanup skipped because source and target are the same path: $sourcePath" -Phase 'Moving'
        return $false
    }

    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf) -or -not (Test-Path -LiteralPath $targetFullPath -PathType Leaf)) {
        return $false
    }

    try {
        $sourceHash = if (-not [string]::IsNullOrWhiteSpace($Item.Sha256)) { $Item.Sha256 } else { Get-Sha256 -Path $sourcePath }
        $targetHash = Get-Sha256 -Path $targetFullPath
        if (-not $sourceHash.Equals($targetHash, [StringComparison]::OrdinalIgnoreCase)) {
            Write-Log -Message "Source cleanup skipped because identical-target verification failed: $sourcePath -> $targetFullPath" -Phase 'Moving'
            return $false
        }

        Remove-Item -LiteralPath $sourcePath -Force
        $Stats.FilesMoved++
        Write-Log -Message "Removed source because identical target already exists: $sourcePath -> $targetFullPath. Reason: $Reason" -Phase 'Moving'
        return $true
    }
    catch {
        $Stats.Errors++
        Write-Log -Message "Could not remove source after identical target verification: $sourcePath. Error: $($_.Exception.Message)" -Phase 'Moving'
        return $false
    }
}

function Move-LockedFileToNeedsReview {
    param(
        [pscustomobject]$Item,
        [string]$OriginalDestinationPath,
        [string]$Reason
    )

    try {
        if (-not (Test-Path -LiteralPath $NeedsReviewRoot)) {
            New-Item -ItemType Directory -Path $NeedsReviewRoot -Force | Out-Null
        }

        if (-not (Test-IsChildPath -Path $OriginalDestinationPath -ParentPath $NeedsReviewRoot)) {
            $Stats.NeedsReview++
        }

        $needsReviewTarget = Resolve-UniquePath -Directory $NeedsReviewRoot -FileName $Item.File.Name -SourceHash $Item.Sha256
        if ($needsReviewTarget.SkipExistingIdentical) {
            Register-ProcessedFile -Item $Item -NewPath $needsReviewTarget.Path -Status "$Reason - Locked, existing identical in NeedsReview"
            return
        }

        Release-OwnHandles
        Invoke-TransferOperation -Item $Item -DestinationPath $needsReviewTarget.Path
        if ($CopyInsteadOfMove) {
            $Stats.FilesCopied++
            Register-ProcessedFile -Item $Item -NewPath $needsReviewTarget.Path -Status "$Reason - Locked - Copied to NeedsReview"
        }
        else {
            $Stats.FilesMoved++
            Register-ProcessedFile -Item $Item -NewPath $needsReviewTarget.Path -Status "$Reason - Locked - Moved to NeedsReview"
        }
        Write-Log -Message "Locked file moved to NeedsReview: $($Item.File.FullName)" -Phase 'Moving'
    }
    catch {
        $Stats.Errors++
        Write-Log -Message "Locked file could not be moved to NeedsReview and was left in place: $($Item.File.FullName). Intended destination was: $OriginalDestinationPath. Error: $($_.Exception.Message)" -Phase 'Moving'
    }
}

function Get-FolderSizeBytes {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return [int64]0
    }

    $total = [int64]0
    try {
        Get-ChildItem -LiteralPath $Path -File -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
            $total += [int64]$_.Length
        }
    }
    catch {
    }

    return $total
}

function Get-RetentionItemSizeBytes {
    param([System.IO.FileSystemInfo]$Item)

    if ($null -eq $Item) { return [int64]0 }
    if ($Item -is [System.IO.FileInfo]) { return [int64]$Item.Length }
    return (Get-FolderSizeBytes -Path $Item.FullName)
}

function Get-ConfirmedDuplicateManifestRootPath {
    $stateRoot = $null
    if (-not [string]::IsNullOrWhiteSpace($ProcessedDbPath)) {
        try {
            $stateRoot = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($ProcessedDbPath))
        }
        catch {
        }
    }
    if ([string]::IsNullOrWhiteSpace($stateRoot)) {
        $stateRoot = Get-UserDataRootPath
    }
    return (Join-Path $stateRoot 'QuarantineManifests')
}

function ConvertTo-ConfirmedDuplicateManifestPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $resolved = Resolve-FullPath $Path
    try {
        $relative = ConvertTo-RelativePath -Path $resolved -BasePath $DestinationBase
        if (-not [string]::IsNullOrWhiteSpace($relative) -and
            -not [System.IO.Path]::IsPathRooted($relative) -and
            -not $relative.StartsWith('..')) {
            return $relative
        }
    }
    catch {
    }
    return $resolved
}

function Resolve-ConfirmedDuplicateManifestPath {
    param([string]$PathValue)

    if ([string]::IsNullOrWhiteSpace($PathValue)) { return '' }
    if ([System.IO.Path]::IsPathRooted($PathValue)) {
        return (Resolve-FullPath $PathValue)
    }
    return (Resolve-FullPath (Join-Path $DestinationBase $PathValue))
}

function Add-ConfirmedDuplicateQuarantineEntry {
    param(
        [string]$QuarantinePath,
        [string]$CanonicalPath,
        [string]$Hash
    )

    if (-not $Apply -or [string]::IsNullOrWhiteSpace($QuarantinePath) -or
        [string]::IsNullOrWhiteSpace($CanonicalPath) -or [string]::IsNullOrWhiteSpace($Hash)) {
        return
    }
    if ($null -eq $script:ConfirmedDuplicateQuarantineEntries) {
        $script:ConfirmedDuplicateQuarantineEntries = New-Object System.Collections.Generic.List[object]
    }
    $script:ConfirmedDuplicateQuarantineEntries.Add([pscustomobject]@{
        hash = $Hash.ToUpperInvariant()
        quarantinePath = ConvertTo-ConfirmedDuplicateManifestPath -Path $QuarantinePath
        canonicalPath = ConvertTo-ConfirmedDuplicateManifestPath -Path $CanonicalPath
        quarantinedAt = (Get-Date).ToString('o')
    })
}

function Update-ConfirmedDuplicateQuarantineCanonicalPath {
    param(
        [string]$Hash,
        [string]$OldCanonicalPath,
        [string]$NewCanonicalPath
    )

    if ($null -eq $script:ConfirmedDuplicateQuarantineEntries -or
        [string]::IsNullOrWhiteSpace($Hash) -or [string]::IsNullOrWhiteSpace($NewCanonicalPath)) {
        return
    }
    $oldValue = ConvertTo-ConfirmedDuplicateManifestPath -Path $OldCanonicalPath
    $newValue = ConvertTo-ConfirmedDuplicateManifestPath -Path $NewCanonicalPath
    foreach ($entry in @($script:ConfirmedDuplicateQuarantineEntries.ToArray())) {
        if ([string]$entry.hash -eq $Hash.ToUpperInvariant() -and
            ([string]$entry.canonicalPath).Equals($oldValue, [StringComparison]::OrdinalIgnoreCase)) {
            $entry.canonicalPath = $newValue
        }
    }
}

function Save-ConfirmedDuplicateQuarantineRunManifest {
    param([bool]$Successful)

    if (-not $Apply -or $null -eq $script:ConfirmedDuplicateQuarantineEntries -or
        $script:ConfirmedDuplicateQuarantineEntries.Count -eq 0) {
        return $null
    }

    try {
        $root = Get-ConfirmedDuplicateManifestRootPath
        if (-not (Test-Path -LiteralPath $root -PathType Container)) {
            New-Item -ItemType Directory -Path $root -Force | Out-Null
        }
        $safeRunId = ([string]$script:RunId -replace '[^A-Za-z0-9_.-]', '_')
        $manifestPath = Join-Path $root ("ConfirmedDuplicates-{0}-{1}.json" -f $safeRunId, $PID)
        $payload = [pscustomobject]@{
            schemaVersion = 1
            runId = [string]$script:RunId
            completedAt = (Get-Date).ToString('o')
            status = if ($Successful) { 'Completed' } else { 'CompletedWithErrors' }
            entries = @($script:ConfirmedDuplicateQuarantineEntries.ToArray())
        }
        $tmp = $manifestPath + '.tmp'
        $payload | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $tmp -Encoding UTF8
        Move-Item -LiteralPath $tmp -Destination $manifestPath -Force
        Write-Log -Message ("Confirmed duplicate quarantine manifest saved: {0}. Entries={1}; Status={2}" -f $manifestPath, $payload.entries.Count, $payload.status) -Phase 'DedupeCleanup'
        return $manifestPath
    }
    catch {
        $Stats.Errors++
        Write-Log -Message "Confirmed duplicate quarantine manifest could not be saved. Quarantined files will remain protected from automatic retention. Error: $($_.Exception.Message)" -Phase 'DedupeCleanup' -Status 'Warning'
        return $null
    }
}

function ConvertFrom-ConfirmedDuplicateManifestTimestamp {
    param(
        [object]$Value,
        [datetime]$Fallback = [datetime]::MinValue
    )

    if ($null -eq $Value) { return $Fallback }
    if ($Value -is [datetime]) { return [datetime]$Value }

    $parsedOffset = [datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse(
            [string]$Value,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [System.Globalization.DateTimeStyles]::AllowWhiteSpaces,
            [ref]$parsedOffset)) {
        return $parsedOffset.LocalDateTime
    }
    return $Fallback
}

function Get-ConfirmedDuplicateQuarantineManifests {
    $root = Get-ConfirmedDuplicateManifestRootPath
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
        return @()
    }

    $result = New-Object System.Collections.Generic.List[object]
    foreach ($file in @(Get-ChildItem -LiteralPath $root -File -Filter 'ConfirmedDuplicates-*.json' -Force -ErrorAction SilentlyContinue)) {
        try {
            $json = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
            if ($null -eq $json -or -not ($json.PSObject.Properties.Name -contains 'entries')) { continue }
            $completedAt = ConvertFrom-ConfirmedDuplicateManifestTimestamp -Value $json.completedAt
            $result.Add([pscustomobject]@{
                Path = $file.FullName
                RunId = [string]$json.runId
                CompletedAt = $completedAt
                Status = [string]$json.status
                Entries = @($json.entries)
            })
        }
        catch {
            Write-Log -Message "Confirmed duplicate quarantine manifest ignored because it could not be read: $($file.FullName). Error: $($_.Exception.Message)" -Phase 'Retention cleanup' -Status 'Warning'
        }
    }
    return @($result.ToArray())
}

function New-RetentionCandidate {
    param(
        [System.IO.FileSystemInfo]$Item,
        [string]$Kind,
        [string]$Reason
    )

    [pscustomobject]@{
        Path = $Item.FullName
        Kind = $Kind
        Reason = $Reason
        LastWriteTime = $Item.LastWriteTime
        Bytes = Get-RetentionItemSizeBytes -Item $Item
        IsDirectory = ($Item -is [System.IO.DirectoryInfo])
    }
}

function Select-RetentionCandidatesByPolicy {
    param(
        [object[]]$Candidates,
        [string]$RootPath,
        [int]$RetentionDays,
        [double]$MaxGB
    )

    if (-not $Candidates -or $Candidates.Count -eq 0) {
        return @()
    }

    $cutoff = (Get-Date).AddDays(-1 * [math]::Max(0, $RetentionDays))
    $ageEligible = @($Candidates | Where-Object { $_.LastWriteTime -lt $cutoff } | Sort-Object LastWriteTime)

    if ($MaxGB -le 0) {
        return $ageEligible
    }

    $currentBytes = Get-FolderSizeBytes -Path $RootPath
    $limitBytes = [int64]($MaxGB * 1GB)
    if ($currentBytes -le $limitBytes) {
        return @()
    }

    $selected = New-Object System.Collections.Generic.List[object]
    $projectedBytes = [int64]$currentBytes
    foreach ($candidate in $ageEligible) {
        if ($projectedBytes -le $limitBytes) { break }
        $selected.Add($candidate)
        $projectedBytes -= [int64]$candidate.Bytes
    }

    return @($selected.ToArray())
}

function Get-RetentionCandidatesTotalBytes {
    param([object[]]$Candidates)

    $total = [int64]0
    foreach ($candidate in @($Candidates)) {
        if ($null -eq $candidate) { continue }
        try {
            $total += [int64]$candidate.Bytes
        }
        catch {
        }
    }

    return $total
}

function Invoke-RetentionCandidateDeletion {
    param(
        [object[]]$Candidates,
        [string]$Phase
    )

    $deleted = 0
    $recoveredBytes = [int64]0
    foreach ($candidate in @($Candidates)) {
        $actionText = if ($Apply) { 'Deleting' } else { 'Would delete' }
        Write-Log -Message ("{0} retention candidate: {1}; SizeGB={2}; Reason={3}" -f $actionText, $candidate.Path, ([math]::Round(([int64]$candidate.Bytes / 1GB), 3)), $candidate.Reason) -Phase $Phase
        if (-not $Apply) {
            continue
        }

        try {
            if ($candidate.IsDirectory) {
                Remove-Item -LiteralPath $candidate.Path -Recurse -Force -ErrorAction Stop
            }
            else {
                Remove-Item -LiteralPath $candidate.Path -Force -ErrorAction Stop
            }
            $deleted++
            $recoveredBytes += [int64]$candidate.Bytes
        }
        catch {
            $Stats.Errors++
            Write-Log -Message "Retention cleanup could not delete $($candidate.Path): $($_.Exception.Message)" -Phase $Phase
        }
    }

    $Stats.RetentionDeletedItems += $deleted
    $Stats.RetentionRecoveredGB = [math]::Round((($Stats.RetentionRecoveredGB * 1GB) + $recoveredBytes) / 1GB, 3)
    return [pscustomobject]@{
        Deleted = $deleted
        RecoveredBytes = $recoveredBytes
    }
}

function Remove-EmptyRetentionDirectories {
    param(
        [string]$RootPath,
        [string]$Phase
    )

    if (-not $Apply -or -not (Test-Path -LiteralPath $RootPath -PathType Container)) {
        return
    }

    $directories = @(Get-ChildItem -LiteralPath $RootPath -Directory -Recurse -Force -ErrorAction SilentlyContinue | Sort-Object FullName -Descending)
    foreach ($directory in $directories) {
        try {
            $children = @(Get-ChildItem -LiteralPath $directory.FullName -Force -ErrorAction Stop)
            if ($children.Count -eq 0) {
                Remove-Item -LiteralPath $directory.FullName -Force -ErrorAction Stop
                Write-DiagnosticLog "Retention cleanup removed empty directory: $($directory.FullName)"
            }
        }
        catch {
            Write-DiagnosticLog "Retention cleanup could not inspect/remove empty directory: $($directory.FullName). Error: $($_.Exception.Message)"
        }
    }
}

function Invoke-MetadataBackupRetentionCleanup {
    param([string]$BackupBasePath)

    $phase = 'Retention cleanup'
    Write-Log -Message ("Metadata retention cleanup started. Mode={0}; Root={1}; RetentionDays={2}; MaxGB={3}" -f $(if ($Apply) { 'APPLY' } else { 'DRY RUN' }), $BackupBasePath, $MetadataBackupRetentionDays, $MetadataBackupMaxGB) -Phase $phase
    if ([string]::IsNullOrWhiteSpace($BackupBasePath) -or -not (Test-Path -LiteralPath $BackupBasePath -PathType Container)) {
        Write-Log -Message "Metadata retention cleanup skipped. Root not found: $BackupBasePath" -Phase $phase
        return
    }

    $runFolders = @(Get-ChildItem -LiteralPath $BackupBasePath -Directory -Force -ErrorAction SilentlyContinue)
    if ($runFolders.Count -eq 0) {
        Write-Log -Message "Metadata retention cleanup found no backup folders." -Phase $phase
        return
    }

    $latestFolder = @($runFolders | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
    $protected = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
    [void]$protected.Add((Resolve-FullPath $MetadataBackupRoot).TrimEnd('\'))
    if ($latestFolder.Count -gt 0) {
        [void]$protected.Add((Resolve-FullPath $latestFolder[0].FullName).TrimEnd('\'))
    }

    $candidates = New-Object System.Collections.Generic.List[object]
    foreach ($folder in $runFolders) {
        $resolved = (Resolve-FullPath $folder.FullName).TrimEnd('\')
        if ($protected.Contains($resolved)) {
            Write-DiagnosticLog "Metadata retention protected folder: $resolved"
            continue
        }
        $candidates.Add((New-RetentionCandidate -Item $folder -Kind 'MetadataBackupFolder' -Reason 'Old metadata backup run folder'))
    }

    $selected = @(Select-RetentionCandidatesByPolicy -Candidates @($candidates.ToArray()) -RootPath $BackupBasePath -RetentionDays $MetadataBackupRetentionDays -MaxGB $MetadataBackupMaxGB)
    $result = Invoke-RetentionCandidateDeletion -Candidates $selected -Phase $phase
    $wouldDelete = if ($Apply) { $result.Deleted } else { $selected.Count }
    $space = if ($Apply) { $result.RecoveredBytes } else { Get-RetentionCandidatesTotalBytes -Candidates $selected }
    $verb = if ($Apply) { 'Deleted' } else { 'Would delete' }
    $spaceVerb = if ($Apply) { 'Recovered space' } else { 'Potential space' }
    Write-Log -Message ("Metadata retention cleanup: {0} {1} old backup folders. {2}: {3} GB" -f $verb, $wouldDelete, $spaceVerb, ([math]::Round($space / 1GB, 3))) -Phase $phase
}

function Get-ConfirmedDuplicateIndexPathsByHash {
    param([string]$QuarantineRootPath)

    $result = @{}
    foreach ($record in @($script:ProcessedRecords.ToArray())) {
        if ($null -eq $record -or [string]::IsNullOrWhiteSpace([string]$record.hash)) { continue }
        $hashKey = ([string]$record.hash).ToUpperInvariant()
        $registeredPath = Get-ProcessedRecordRegisteredPath -Record $record
        if ([string]::IsNullOrWhiteSpace($registeredPath) -or
            (Test-IsChildPath -Path $registeredPath -ParentPath $QuarantineRootPath)) {
            continue
        }
        if (-not $result.ContainsKey($hashKey)) {
            $result[$hashKey] = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
        }
        [void]$result[$hashKey].Add((Resolve-FullPath $registeredPath))
    }
    return $result
}

function Test-ConfirmedDuplicateRetentionSafety {
    param(
        [object]$Candidate,
        [hashtable]$IndexPathsByHash,
        [string]$QuarantineRootPath
    )

    $result = [ordered]@{
        Safe = $false
        Reason = 'Unknown'
        CanonicalPath = ''
        ActualHash = ''
    }
    $quarantinePath = Resolve-FullPath $Candidate.Path
    if (-not (Test-IsChildPath -Path $quarantinePath -ParentPath $QuarantineRootPath)) {
        $result.Reason = 'QuarantinePathOutsideExpectedRoot'
        return [pscustomobject]$result
    }
    if (-not (Test-Path -LiteralPath $quarantinePath -PathType Leaf)) {
        $result.Reason = 'QuarantineFileMissing'
        return [pscustomobject]$result
    }

    $extension = [System.IO.Path]::GetExtension($quarantinePath).ToLowerInvariant()
    if ($RawExtensions -contains $extension) {
        $result.Reason = 'RawOrDngProtected'
        return [pscustomobject]$result
    }

    $quarantineAvailability = Detect-StorageAvailability -Path $quarantinePath
    if ($quarantineAvailability.State -ne 'LocalVerified') {
        $result.Reason = 'QuarantineNotLocallyVerified:' + [string]$quarantineAvailability.State
        return [pscustomobject]$result
    }

    try {
        $actualHash = (Get-Sha256 -Path $quarantinePath).ToUpperInvariant()
        $result.ActualHash = $actualHash
    }
    catch {
        $result.Reason = 'QuarantineHashFailed:' + $_.Exception.Message
        return [pscustomobject]$result
    }

    $expectedHash = ([string]$Candidate.ExpectedHash).ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($expectedHash) -or $actualHash -ne $expectedHash) {
        $result.Reason = 'QuarantineHashChanged'
        return [pscustomobject]$result
    }

    $canonicalPaths = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
    $manifestCanonicalPath = Resolve-ConfirmedDuplicateManifestPath -PathValue ([string]$Candidate.CanonicalPathValue)
    if (-not [string]::IsNullOrWhiteSpace($manifestCanonicalPath)) {
        [void]$canonicalPaths.Add($manifestCanonicalPath)
    }
    if ($IndexPathsByHash.ContainsKey($expectedHash)) {
        foreach ($indexPath in @($IndexPathsByHash[$expectedHash])) {
            [void]$canonicalPaths.Add((Resolve-FullPath $indexPath))
        }
    }
    [void]$canonicalPaths.Remove($quarantinePath)
    foreach ($path in @($canonicalPaths)) {
        if (Test-IsChildPath -Path $path -ParentPath $QuarantineRootPath) {
            [void]$canonicalPaths.Remove($path)
        }
    }

    if ($canonicalPaths.Count -eq 0) {
        $result.Reason = 'CanonicalNotFound'
        return [pscustomobject]$result
    }
    if ($canonicalPaths.Count -gt 1) {
        $result.Reason = 'CanonicalIndexConflict'
        return [pscustomobject]$result
    }

    $canonicalPath = @($canonicalPaths | Select-Object -First 1)[0]
    $result.CanonicalPath = $canonicalPath
    if (-not (Test-Path -LiteralPath $canonicalPath -PathType Leaf)) {
        $result.Reason = 'CanonicalMissing'
        return [pscustomobject]$result
    }
    $canonicalAvailability = Detect-StorageAvailability -Path $canonicalPath
    if ($canonicalAvailability.State -ne 'LocalVerified') {
        $result.Reason = 'CanonicalNotLocallyVerified:' + [string]$canonicalAvailability.State
        return [pscustomobject]$result
    }

    try {
        $canonicalHash = (Get-Sha256 -Path $canonicalPath).ToUpperInvariant()
    }
    catch {
        $result.Reason = 'CanonicalHashFailed:' + $_.Exception.Message
        return [pscustomobject]$result
    }
    if ($canonicalHash -ne $actualHash) {
        $result.Reason = 'CanonicalHashMismatch'
        return [pscustomobject]$result
    }

    $result.Safe = $true
    $result.Reason = 'CurrentSha256Revalidated'
    return [pscustomobject]$result
}

function Invoke-ConfirmedDuplicatesRetentionCleanup {
    param([string]$QuarantineRootPath)

    $phase = 'Retention cleanup'
    Write-Log -Message ("Confirmed duplicates retention cleanup started. Mode={0}; Root={1}; RetentionDays={2}; MaxGB={3}; Revalidation=CurrentSHA256" -f $(if ($Apply) { 'APPLY' } else { 'DRY RUN' }), $QuarantineRootPath, $ConfirmedDuplicatesRetentionDays, $ConfirmedDuplicatesMaxGB) -Phase $phase
    if ([string]::IsNullOrWhiteSpace($QuarantineRootPath) -or -not (Test-Path -LiteralPath $QuarantineRootPath -PathType Container)) {
        Write-Log -Message "Confirmed duplicates retention cleanup skipped. Root not found: $QuarantineRootPath" -Phase $phase
        return
    }

    $files = @(Get-ChildItem -LiteralPath $QuarantineRootPath -File -Recurse -Force -ErrorAction SilentlyContinue)
    if ($files.Count -eq 0) {
        Write-Log -Message "Confirmed duplicates retention cleanup found no files." -Phase $phase
        return
    }

    $manifests = @(Get-ConfirmedDuplicateQuarantineManifests)
    $successfulManifests = @($manifests | Where-Object { $_.Status -eq 'Completed' } | Sort-Object CompletedAt -Descending)
    $latestSuccessfulManifestPath = if ($successfulManifests.Count -gt 0) { [string]$successfulManifests[0].Path } else { '' }
    $referencedPaths = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
    $candidates = New-Object System.Collections.Generic.List[object]
    $protectedByRun = 0
    $protectedByErrors = 0

    foreach ($manifest in $manifests) {
        $protectManifest = ([string]$manifest.Path).Equals($latestSuccessfulManifestPath, [StringComparison]::OrdinalIgnoreCase)
        $manifestHasErrors = $manifest.Status -ne 'Completed'
        foreach ($entry in @($manifest.Entries)) {
            $quarantinePath = Resolve-ConfirmedDuplicateManifestPath -PathValue ([string]$entry.quarantinePath)
            if ([string]::IsNullOrWhiteSpace($quarantinePath)) { continue }
            [void]$referencedPaths.Add($quarantinePath)
            if (-not (Test-Path -LiteralPath $quarantinePath -PathType Leaf)) { continue }
            if ($protectManifest) {
                $protectedByRun++
                continue
            }
            if ($manifestHasErrors) {
                $protectedByErrors++
                continue
            }

            $quarantinedAt = ConvertFrom-ConfirmedDuplicateManifestTimestamp -Value $entry.quarantinedAt -Fallback $manifest.CompletedAt
            $file = Get-Item -LiteralPath $quarantinePath -Force -ErrorAction SilentlyContinue
            if ($null -eq $file) { continue }
            $candidates.Add([pscustomobject]@{
                Path = $file.FullName
                Kind = 'ConfirmedDuplicateFile'
                Reason = 'Old confirmed duplicate pending current SHA256 revalidation'
                LastWriteTime = $quarantinedAt
                Bytes = [int64]$file.Length
                IsDirectory = $false
                ExpectedHash = [string]$entry.hash
                CanonicalPathValue = [string]$entry.canonicalPath
                RunId = [string]$manifest.RunId
                ManifestPath = [string]$manifest.Path
            })
        }
    }

    $historicalUntracked = 0
    foreach ($file in $files) {
        if (-not $referencedPaths.Contains((Resolve-FullPath $file.FullName))) {
            $historicalUntracked++
        }
    }
    if ($historicalUntracked -gt 0) {
        Write-Log -Message "Historical confirmed duplicate quarantine files retained because no successful validation manifest exists: $historicalUntracked" -Phase $phase -Status 'Warning'
    }

    $selected = @(Select-RetentionCandidatesByPolicy -Candidates @($candidates.ToArray()) -RootPath $QuarantineRootPath -RetentionDays $ConfirmedDuplicatesRetentionDays -MaxGB $ConfirmedDuplicatesMaxGB)
    $safeCandidates = New-Object System.Collections.Generic.List[object]
    $retainedReasons = @{}
    if ($selected.Count -gt 0) {
        Load-ProcessedIndexLight
        $indexPathsByHash = Get-ConfirmedDuplicateIndexPathsByHash -QuarantineRootPath $QuarantineRootPath
        foreach ($candidate in $selected) {
            $validation = Test-ConfirmedDuplicateRetentionSafety -Candidate $candidate -IndexPathsByHash $indexPathsByHash -QuarantineRootPath $QuarantineRootPath
            if ($validation.Safe) {
                $candidate.Reason = "Current SHA256 revalidated against canonical: $($validation.CanonicalPath)"
                $safeCandidates.Add($candidate)
                $Stats.ConfirmedDuplicatesRevalidated++
            }
            else {
                $Stats.ConfirmedDuplicatesRetained++
                if (-not $retainedReasons.ContainsKey($validation.Reason)) { $retainedReasons[$validation.Reason] = 0 }
                $retainedReasons[$validation.Reason]++
                Write-Log -Message "Confirmed duplicate retained after revalidation: Path=$($candidate.Path); Reason=$($validation.Reason); Canonical=$($validation.CanonicalPath)" -Phase $phase -Status 'Warning'
            }
        }
    }

    $result = Invoke-RetentionCandidateDeletion -Candidates @($safeCandidates.ToArray()) -Phase $phase
    if ($Apply) {
        Remove-EmptyRetentionDirectories -RootPath $QuarantineRootPath -Phase $phase
    }

    $wouldDelete = if ($Apply) { $result.Deleted } else { $safeCandidates.Count }
    $space = if ($Apply) { $result.RecoveredBytes } else { Get-RetentionCandidatesTotalBytes -Candidates @($safeCandidates.ToArray()) }
    $verb = if ($Apply) { 'Deleted' } else { 'Would delete' }
    $spaceVerb = if ($Apply) { 'Recovered space' } else { 'Potential space' }
    $reasonSummary = if ($retainedReasons.Count -gt 0) { @($retainedReasons.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join '; ' } else { 'none' }
    Write-Log -Message ("Confirmed duplicates retention cleanup: {0} {1} revalidated duplicates older than {2} days. {3}: {4} GB. Eligible={5}; retainedAfterValidation={6}; historicalWithoutManifest={7}; latestRunProtected={8}; errorRunsProtected={9}; retainedReasons={10}" -f $verb, $wouldDelete, $ConfirmedDuplicatesRetentionDays, $spaceVerb, ([math]::Round($space / 1GB, 3)), $selected.Count, $Stats.ConfirmedDuplicatesRetained, $historicalUntracked, $protectedByRun, $protectedByErrors, $reasonSummary) -Phase $phase
}

function Invoke-RetentionCleanup {
    Write-Log -Message ("RetentionCleanup started. Mode: {0}" -f $(if ($Apply) { 'APPLY' } else { 'DRY RUN' })) -Phase 'Retention cleanup'
    Write-Log -Message "Retention cleanup never touches Duplicates_To_Review / _Duplicados_Para_Revisar." -Phase 'Retention cleanup'

    $metadataBackupBasePath = [System.IO.Path]::GetDirectoryName($MetadataBackupRoot.TrimEnd('\'))
    Invoke-MetadataBackupRetentionCleanup -BackupBasePath $metadataBackupBasePath
    Invoke-ConfirmedDuplicatesRetentionCleanup -QuarantineRootPath $ConfirmedDuplicatesQuarantineRoot

    Write-Log -Message ("RetentionCleanup summary: deletedItems={0}; recoveredSpaceGB={1}; metadataRetentionDays={2}; confirmedDuplicatesRetentionDays={3}; confirmedDuplicatesRevalidated={4}; confirmedDuplicatesRetained={5}" -f $Stats.RetentionDeletedItems, $Stats.RetentionRecoveredGB, $MetadataBackupRetentionDays, $ConfirmedDuplicatesRetentionDays, $Stats.ConfirmedDuplicatesRevalidated, $Stats.ConfirmedDuplicatesRetained) -Phase 'Complete' -Status 'Completed'
}

function Update-MetadataBackupSize {
    $bytes = Get-FolderSizeBytes -Path $MetadataBackupRoot
    $Stats.MetadataBackupSizeGB = [math]::Round(($bytes / 1GB), 3)
    return $Stats.MetadataBackupSizeGB
}

function Test-MetadataBackupsHaveVerifiedTargets {
    if (-not (Test-Path -LiteralPath $MetadataBackupRoot)) {
        return $true
    }

    $backupFiles = @(Get-ChildItem -LiteralPath $MetadataBackupRoot -File -Recurse -Force -ErrorAction SilentlyContinue)
    if ($backupFiles.Count -eq 0) {
        return $true
    }

    $recordsByOriginal = @{}
    foreach ($record in @($script:ProcessedRecords.ToArray())) {
        if ($record.PSObject.Properties.Name -notcontains 'originalRelativePath' -or [string]::IsNullOrWhiteSpace([string]$record.originalRelativePath)) {
            continue
        }

        $key = ([string]$record.originalRelativePath).Replace('/', '\').TrimStart('\')
        if (-not $recordsByOriginal.ContainsKey($key)) {
            $recordsByOriginal[$key] = New-Object System.Collections.Generic.List[object]
        }
        $recordsByOriginal[$key].Add($record)
    }

    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($backupFile in $backupFiles) {
        $relative = (ConvertTo-RelativePath -Path $backupFile.FullName -BasePath $MetadataBackupRoot).Replace('/', '\').TrimStart('\')
        if (-not $recordsByOriginal.ContainsKey($relative)) {
            if ($missing.Count -lt 20) { $missing.Add("No processed record: $relative") }
            continue
        }

        $targetFound = $false
        foreach ($record in $recordsByOriginal[$relative]) {
            if ($record.PSObject.Properties.Name -notcontains 'newRelativePath' -or [string]::IsNullOrWhiteSpace([string]$record.newRelativePath)) {
                continue
            }

            $targetPath = Join-Path $DestinationBase (([string]$record.newRelativePath).Replace('/', '\').TrimStart('\'))
            if (Test-Path -LiteralPath $targetPath -PathType Leaf) {
                $targetFound = $true
                break
            }
        }

        if (-not $targetFound -and $missing.Count -lt 20) {
            $missing.Add("Missing final target: $relative")
        }
    }

    if ($missing.Count -gt 0) {
        foreach ($item in $missing) {
            Write-Log -Message "Metadata backup verification failed: $item" -Phase 'Complete'
        }
        Write-Log -Message "Metadata backup verification failed. Checked backups: $($backupFiles.Count)." -Phase 'Complete'
        return $false
    }

    Write-Log -Message "Metadata backup verification passed. Checked backups: $($backupFiles.Count). All final targets exist." -Phase 'Complete'
    return $true
}

function Invoke-SafeTransfer {
    param(
        [pscustomobject]$Item,
        [string]$DestinationDirectory,
        [string]$Reason
    )

    if ((Resolve-FullPath $DestinationDirectory).TrimEnd('\').Equals((Resolve-FullPath $DuplicatesRoot).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase) -and
        ($script:DedupeOrganizedRoots | Where-Object { Test-IsChildPath -Path $Item.File.FullName -ParentPath $_ })) {
        Stop-WithError "CRITICAL: Attempted to move organized library item to duplicates during combined mode. Operation blocked. File=$($Item.File.FullName)"
    }

    $availability = Detect-StorageAvailability -Item $Item.File
    if ($availability.State -eq 'CloudPlaceholder') {
        Register-CloudPlaceholderSkipped -Path $Item.File.FullName -Phase 'Moving' -Availability $availability
        return
    }
    if ($availability.State -eq 'MissingReal') {
        $Stats.MissingReal++
        Write-Log -Message "Transfer skipped missing file: $($Item.File.FullName). Reason=$($availability.Reason)" -Phase 'Moving'
        return
    }

    $targetFileName = $Item.File.Name
    if ($Item.PSObject.Properties.Name -contains 'TargetFileName' -and -not [string]::IsNullOrWhiteSpace([string]$Item.TargetFileName)) {
        $targetFileName = [string]$Item.TargetFileName
    }

    $resolved = Resolve-UniquePath -Directory $DestinationDirectory -FileName $targetFileName -SourceHash $Item.Sha256
    if ($resolved.SkipExistingIdentical) {
        $Stats.ExistingIdenticalSkipped++
        Write-Verbose "Skipped existing identical target: $($resolved.Path)"
        if ($Apply) {
            Remove-SourceAfterVerifiedIdenticalTarget -Item $Item -TargetPath $resolved.Path -Reason $Reason | Out-Null
            Register-ProcessedFile -Item $Item -NewPath $resolved.Path -Status "$Reason - Existing identical target"
        }
        return
    }

    if (-not $Apply) {
        $Stats.DryRunActions++
        Write-Verbose "DryRun: $Reason -> $($resolved.Path)"
        return
    }

    if (-not (Test-Path -LiteralPath $DestinationDirectory)) {
        New-Item -ItemType Directory -Path $DestinationDirectory -Force | Out-Null
    }

    $transferred = Invoke-TransferWithLockRetry -Item $Item -DestinationPath $resolved.Path -Reason $Reason
    if (-not $transferred) {
        Move-LockedFileToNeedsReview -Item $Item -OriginalDestinationPath $resolved.Path -Reason $Reason
        return
    }

    if ($CopyInsteadOfMove) {
        $Stats.FilesCopied++
        Register-ProcessedFile -Item $Item -NewPath $resolved.Path -Status "$Reason - Copied"
    }
    else {
        $Stats.FilesMoved++
        Register-ProcessedFile -Item $Item -NewPath $resolved.Path -Status "$Reason - Moved"
    }
}

function Get-ImportProviderRegistry {
    return @(
        [pscustomobject]@{
            Id = 'GoogleTakeout'
            Aliases = @('GoogleTakeout', 'GooglePhotos', 'Google')
            DisplayName = 'Google Photos / Google Takeout'
            Status = 'Available'
            Gate = 'Implemented'
            SemanticSignals = @('supplemental-metadata.json', 'albums', 'trash', 'provider timestamps', 'physical album duplicates')
            Notes = 'First functional second-head provider.'
        },
        [pscustomobject]@{
            Id = 'ApplePhotos'
            Aliases = @('ApplePhotos', 'ICloudPhotos', 'iCloudPhotos', 'iCloud', 'Apple')
            DisplayName = 'Apple Photos / iCloud Photos'
            Status = 'Available'
            Gate = 'Implemented'
            SemanticSignals = @('Photo Details CSV', 'album CSV references', 'deleted flag', 'provider timestamps', 'Live Photo pair candidates')
            Notes = 'Conservative importer based on real iCloud Photos export structure.'
        },
        [pscustomobject]@{
            Id = 'SamsungGallery'
            Aliases = @('SamsungGallery', 'Samsung')
            DisplayName = 'Samsung Gallery'
            Status = 'Planned'
            Gate = 'SampleGated'
            SemanticSignals = @('albums', 'stories', 'trash', 'device/cloud gallery context')
            Notes = 'Officially planned because of adoption, but disabled until a real export proves recoverable semantics.'
        },
        [pscustomobject]@{
            Id = 'Immich'
            Aliases = @('Immich')
            DisplayName = 'Immich'
            Status = 'Planned'
            Gate = 'SampleRequired'
            SemanticSignals = @('XMP sidecars', 'albums', 'tags', 'ratings', 'descriptions', 'external library metadata')
            Notes = 'Provider should prefer exported metadata/API/database evidence when available.'
        },
        [pscustomobject]@{
            Id = 'XmpSidecarLibrary'
            Aliases = @('XmpSidecarLibrary', 'XMP', 'GenericXMP', 'SidecarLibrary')
            DisplayName = 'Generic XMP / Sidecar Library'
            Status = 'Available'
            Gate = 'Implemented'
            SemanticSignals = @('xmp sidecars', 'json sidecars', 'yaml sidecars', 'ratings', 'tags', 'descriptions', 'dates', 'locations')
            Notes = 'Generic conservative provider for sidecar-based or unknown exported libraries.'
        }
    )
}

function Resolve-ImportProviderId {
    param([string]$Provider)

    if ([string]::IsNullOrWhiteSpace($Provider)) { return '' }
    $normalized = ([string]$Provider).Trim()
    foreach ($entry in Get-ImportProviderRegistry) {
        if (@($entry.Aliases) -contains $normalized) {
            return [string]$entry.Id
        }
    }
    return ''
}

function Get-ImportProviderSpec {
    param([string]$Provider)

    $providerId = Resolve-ImportProviderId -Provider $Provider
    if ([string]::IsNullOrWhiteSpace($providerId)) { return $null }
    return @(Get-ImportProviderRegistry | Where-Object { $_.Id -eq $providerId } | Select-Object -First 1)[0]
}

function Get-ImportProviderDisplayName {
    param([string]$Provider)

    $spec = Get-ImportProviderSpec -Provider $Provider
    if ($spec) { return [string]$spec.DisplayName }
    return $Provider
}

function Assert-ImportProviderAvailable {
    param([string]$Provider)

    $spec = Get-ImportProviderSpec -Provider $Provider
    if (-not $spec) {
        $known = (@(Get-ImportProviderRegistry | ForEach-Object { $_.Id }) -join ', ')
        Stop-WithError "Unsupported ImportProvider: $Provider. Known providers: $known"
    }

    if ([string]$spec.Status -ne 'Available') {
        Stop-WithError ("ImportProvider {0} is {1}/{2}. {3}" -f $spec.DisplayName, $spec.Status, $spec.Gate, $spec.Notes)
    }

    return $spec
}

function Resolve-GoogleTakeoutPhotosRoot {
    param([string]$RootPath)

    $resolved = Resolve-FullPath $RootPath
    if (-not (Test-Path -LiteralPath $resolved -PathType Container)) {
        Stop-WithError "ImportProviderPath does not exist: $resolved"
    }

    $leaf = Split-Path -Leaf $resolved.TrimEnd('\')
    if ($leaf -match '^(Google Fotos|Google Photos|Google Foto)$') {
        return $resolved
    }

    foreach ($candidateName in @('Google Fotos', 'Google Photos', 'Google Foto')) {
        $candidate = Join-Path $resolved $candidateName
        if (Test-Path -LiteralPath $candidate -PathType Container) {
            return (Resolve-FullPath $candidate)
        }
    }

    $sidecar = Get-ChildItem -LiteralPath $resolved -File -Filter '*.json' -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '\.supp.*\.json$' } | Select-Object -First 1
    if ($sidecar) {
        return $resolved
    }

    Stop-WithError "Google Takeout photos folder was not detected under: $resolved"
}

function Read-ProviderJsonFile {
    param([string]$Path)

    try {
        return Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        Write-Log -Message "Provider JSON parse failed: $Path - $($_.Exception.Message)" -Phase 'ImportProvider' -Status 'Warning'
        return $null
    }
}

function Convert-GoogleTimestampToDateTime {
    param([object]$Timestamp)

    if ($null -eq $Timestamp) { return $null }
    $text = [string]$Timestamp
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $seconds = 0L
    if (-not [Int64]::TryParse($text, [ref]$seconds)) { return $null }
    try {
        return [DateTimeOffset]::FromUnixTimeSeconds($seconds).LocalDateTime
    }
    catch {
        return $null
    }
}

function Get-GoogleTakeoutFolderRole {
    param(
        [string]$DirectoryPath,
        [string]$GoogleRoot
    )

    $leaf = Split-Path -Leaf $DirectoryPath.TrimEnd('\')
    if ($leaf -match '(?i)^(Papelera|Trash|Bin|Coș de gunoi|Cos de gunoi)$') { return 'Trash' }
    if ($leaf -match '(?i)^(Archive|Archived|Arhivează|Arhiveaza)$') { return 'Archive' }
    if ((Test-Path -LiteralPath (Join-Path $DirectoryPath 'metadatos.json') -PathType Leaf) -or
        (Test-Path -LiteralPath (Join-Path $DirectoryPath 'metadata.json') -PathType Leaf)) {
        return 'Album'
    }
    if ($leaf -match '(?i)^(Fotos del|Photos from|Photos of|Photos in|Fotografii din)\s+\d{4}$') { return 'YearFolder' }
    if ((Resolve-FullPath $DirectoryPath).TrimEnd('\').Equals((Resolve-FullPath $GoogleRoot).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) { return 'Root' }
    return 'Unknown'
}

function Get-GoogleTakeoutAlbumMetadata {
    param([string]$DirectoryPath)

    foreach ($name in @('metadatos.json', 'metadata.json')) {
        $path = Join-Path $DirectoryPath $name
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            $json = Read-ProviderJsonFile -Path $path
            if ($json) {
                return [pscustomobject]@{
                    Path = $path
                    Title = if ($json.PSObject.Properties.Name -contains 'title') { [string]$json.title } else { Split-Path -Leaf $DirectoryPath }
                    Description = if ($json.PSObject.Properties.Name -contains 'description') { [string]$json.description } else { '' }
                    Access = if ($json.PSObject.Properties.Name -contains 'access') { [string]$json.access } else { '' }
                    Date = if ($json.PSObject.Properties.Name -contains 'date') { Convert-GoogleTimestampToDateTime -Timestamp $json.date.timestamp } else { $null }
                }
            }
        }
    }

    return $null
}

function Get-GoogleSidecarMediaTitle {
    param([System.IO.FileInfo]$JsonFile)

    $name = $JsonFile.Name -replace '\.supp.*\.json$', ''
    $name = $name -replace '\.sup\.json$', ''
    return $name
}

function Get-GoogleTakeoutMediaCopyInfo {
    param([System.IO.FileInfo]$MediaFile)

    if ($MediaFile.Name -match '^(?<base>.+)\((?<index>\d+)\)(?<extension>\.[^.]+)$') {
        return [pscustomobject]@{
            HasCopyIndex = $true
            CopyIndex = [int]$Matches.index
            BaseMediaName = $Matches.base + $Matches.extension
        }
    }

    return [pscustomobject]@{
        HasCopyIndex = $false
        CopyIndex = 0
        BaseMediaName = $MediaFile.Name
    }
}

function Get-GoogleTakeoutSidecarCopyInfo {
    param([System.IO.FileInfo]$JsonFile)

    $patterns = @(
        '^(?<base>.+?)\.supplemental-metadata\((?<index>\d+)\)\.json$',
        '^(?<base>.+?)\.suppl\((?<index>\d+)\)\.json$',
        '^(?<base>.+?)\.sup\((?<index>\d+)\)\.json$'
    )

    foreach ($pattern in $patterns) {
        if ($JsonFile.Name -match $pattern) {
            return [pscustomobject]@{
                HasCopyIndex = $true
                CopyIndex = [int]$Matches.index
                BaseMediaName = $Matches.base
            }
        }
    }

    return [pscustomobject]@{
        HasCopyIndex = $false
        CopyIndex = 0
        BaseMediaName = ''
    }
}

function Test-GoogleTakeoutMediaSidecarJson {
    param([object]$Json)

    if ($null -eq $Json) { return $false }
    $properties = @($Json.PSObject.Properties.Name)
    if ($properties -notcontains 'title' -or [string]::IsNullOrWhiteSpace([string]$Json.title)) {
        return $false
    }

    foreach ($signal in @('photoTakenTime', 'creationTime', 'geoData', 'geoDataExif', 'googlePhotosOrigin', 'appSource', 'url')) {
        if ($properties -contains $signal) {
            return $true
        }
    }

    return $false
}

function New-GoogleTakeoutSidecarRecord {
    param([System.IO.FileInfo]$JsonFile)

    $json = Read-ProviderJsonFile -Path $JsonFile.FullName
    if (-not $json) { return $null }
    if (-not (Test-GoogleTakeoutMediaSidecarJson -Json $json)) { return $null }
    $title = if ($json.PSObject.Properties.Name -contains 'title' -and -not [string]::IsNullOrWhiteSpace([string]$json.title)) {
        [string]$json.title
    }
    else {
        Get-GoogleSidecarMediaTitle -JsonFile $JsonFile
    }

    return [pscustomobject]@{
        File = $JsonFile
        Path = $JsonFile.FullName
        Directory = $JsonFile.DirectoryName
        Title = $title
        Json = $json
        PhotoTakenDate = if ($json.PSObject.Properties.Name -contains 'photoTakenTime') { Convert-GoogleTimestampToDateTime -Timestamp $json.photoTakenTime.timestamp } else { $null }
        CreationDate = if ($json.PSObject.Properties.Name -contains 'creationTime') { Convert-GoogleTimestampToDateTime -Timestamp $json.creationTime.timestamp } else { $null }
        HasGeoData = ($json.PSObject.Properties.Name -contains 'geoData') -and ([double]$json.geoData.latitude -ne 0 -or [double]$json.geoData.longitude -ne 0)
        HasGeoDataExif = ($json.PSObject.Properties.Name -contains 'geoDataExif') -and ([double]$json.geoDataExif.latitude -ne 0 -or [double]$json.geoDataExif.longitude -ne 0)
        IsTrashed = $json.PSObject.Properties.Name -contains 'trashed'
        Origin = if ($json.PSObject.Properties.Name -contains 'googlePhotosOrigin') { (@($json.googlePhotosOrigin.PSObject.Properties.Name) -join ',') } else { '' }
        AppSource = if ($json.PSObject.Properties.Name -contains 'appSource') { (@($json.appSource.PSObject.Properties | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ';') } else { '' }
        CopyInfo = Get-GoogleTakeoutSidecarCopyInfo -JsonFile $JsonFile
        Used = $false
    }
}

function Get-GoogleTakeoutImportDateInfo {
    param(
        [pscustomobject]$Asset,
        [pscustomobject]$PrimaryOccurrence,
        [pscustomobject]$Metadata,
        [bool]$EmbeddedMetadataRead = $true
    )

    $providerDates = @($Asset.Sidecars | Where-Object { $_.PhotoTakenDate } | ForEach-Object { $_.PhotoTakenDate })
    $providerDate = if ($providerDates.Count -gt 0) { $providerDates | Sort-Object | Select-Object -First 1 } else { $null }
    $hasWeakSidecarEvidence = @($Asset.Warnings | Where-Object { $_ -eq 'Ambiguous sidecar' -or $_ -eq 'Media without exact sidecar' }).Count -gt 0

    $exifDate = ConvertTo-MediaDate $Metadata.DateTimeOriginal
    if (-not $exifDate) { $exifDate = ConvertTo-MediaDate $Metadata.CreateDate }
    if (-not $exifDate -and $PrimaryOccurrence.IsVideo) { $exifDate = ConvertTo-MediaDate $Metadata.MediaCreateDate }

    if ($providerDate -and $exifDate) {
        if ([math]::Abs(($providerDate.Date - $exifDate.Date).TotalDays) -le 1) {
            return [pscustomobject]@{
                Date = $exifDate
                Confidence = 99
                Source = 'GoogleTakeout'
                Sources = @('GoogleTakeout', 'EXIF')
                Detail = 'Provider photoTakenTime agrees with embedded metadata'
                ProviderExifConflict = $false
                MetadataConfidence = if ($hasWeakSidecarEvidence) { 'LowConfidence' } else { 'HighConfidence' }
                ExifVerification = 'Read'
            }
        }

        return [pscustomobject]@{
            Date = $providerDate
            Confidence = 20
            Source = 'GoogleTakeout'
            Sources = @('GoogleTakeout', 'EXIF')
            Detail = ("Provider/EXIF date conflict. Provider={0:yyyy-MM-dd}; EXIF={1:yyyy-MM-dd}" -f $providerDate, $exifDate)
            ProviderExifConflict = $true
            MetadataConfidence = 'LowConfidence'
            ExifVerification = 'Read'
        }
    }

    if ($providerDate) {
        $trustedSkipState = if ($PrimaryOccurrence.IsVideo) { 'SkippedProviderTrustedVideo' } else { 'SkippedProviderTrusted' }
        $detail = if ($EmbeddedMetadataRead) {
            'Provider photoTakenTime used because embedded metadata is absent or incomplete'
        }
        elseif ($PrimaryOccurrence.IsVideo) {
            'Provider photoTakenTime used; embedded video metadata verification skipped for trusted exact sidecar'
        }
        else {
            'Provider photoTakenTime used; embedded metadata verification skipped for trusted exact sidecar'
        }
        return [pscustomobject]@{
            Date = $providerDate
            Confidence = 98
            Source = 'GoogleTakeout'
            Sources = @('GoogleTakeout')
            Detail = $detail
            ProviderExifConflict = $false
            MetadataConfidence = if ($hasWeakSidecarEvidence) { 'LowConfidence' } else { 'MediumConfidence' }
            ExifVerification = if ($EmbeddedMetadataRead) { 'Read' } else { $trustedSkipState }
        }
    }

    $item = [pscustomobject]@{
        File = $PrimaryOccurrence.File
        Metadata = $Metadata
    }
    $dateInfo = Get-PrimaryDate -Item $item -IsVideo $PrimaryOccurrence.IsVideo
    $dateInfo | Add-Member -NotePropertyName MetadataConfidence -NotePropertyValue 'LowConfidence' -Force
    $dateInfo | Add-Member -NotePropertyName ExifVerification -NotePropertyValue $(if ($EmbeddedMetadataRead) { 'Read' } else { 'SkippedNoProviderDate' }) -Force
    return $dateInfo
}

function Test-GoogleTakeoutAssetNeedsExifVerification {
    param([pscustomobject]$Asset)

    if ($Diagnostic) { return $true }
    if (-not $Asset -or -not $Asset.PrimaryOccurrence) { return $true }

    $sidecarStatuses = @($Asset.Occurrences | ForEach-Object { $_.SidecarStatus } | Select-Object -Unique)
    if ($sidecarStatuses -contains 'Ambiguous' -or $sidecarStatuses -contains 'Missing') { return $true }

    $providerDates = @($Asset.Sidecars | Where-Object { $_.PhotoTakenDate } | ForEach-Object { $_.PhotoTakenDate.Date } | Select-Object -Unique)
    if ($providerDates.Count -eq 0) { return $true }
    if ($providerDates.Count -gt 1) { return $true }

    return $false
}

function Select-GoogleTakeoutPrimaryOccurrence {
    param([object[]]$Occurrences)

    $ordered = @($Occurrences | Sort-Object `
        @{ Expression = { if ($_.FolderRole -eq 'YearFolder') { 0 } elseif ($_.FolderRole -eq 'Unknown' -or $_.FolderRole -eq 'Root' -or $_.FolderRole -eq 'Archive') { 1 } elseif ($_.FolderRole -eq 'Album') { 2 } else { 3 } } }, `
        @{ Expression = { $_.File.FullName.Length } })
    return $ordered | Select-Object -First 1
}

function ConvertTo-HtmlText {
    param([object]$Value)
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Write-GoogleTakeoutImportReport {
    param(
        [pscustomobject]$Summary,
        [object[]]$Assets,
        [object[]]$OrphanJson,
        [object[]]$MediaWithoutJson
    )

    if (-not (Test-Path -LiteralPath $LogRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null
    }

    $htmlPath = Join-Path $LogRoot ("ImportProvider-GoogleTakeout-{0}.html" -f $script:RunId)
    $jsonPath = Join-Path $LogRoot ("ImportProvider-GoogleTakeout-{0}.json" -f $script:RunId)
    $modeLabel = Get-ImportProviderText -Key 'Mode'
    $mediaFilesLabel = Get-ImportProviderText -Key 'MediaFiles'
    $logicalAssetsLabel = Get-ImportProviderText -Key 'LogicalAssets'
    $occurrencesLabel = Get-ImportProviderText -Key 'Occurrences'
    $albumsLabel = Get-ImportProviderText -Key 'Albums'
    $internalDuplicatesLabel = Get-ImportProviderText -Key 'InternalDuplicateOccurrences'
    $trashOccurrencesLabel = Get-ImportProviderText -Key 'TrashOccurrences'
    $videosLabel = Get-ImportProviderText -Key 'Videos'
    $conflictsLabel = Get-ImportProviderText -Key 'Conflicts'
    $highConfidenceLabel = Get-ImportProviderText -Key 'HighConfidence'
    $mediumConfidenceLabel = Get-ImportProviderText -Key 'MediumConfidence'
    $lowConfidenceLabel = Get-ImportProviderText -Key 'LowConfidence'
    $exifSkippedLabel = Get-ImportProviderText -Key 'ExifReadsSkipped'
    $exifSkippedVideoLabel = Get-ImportProviderText -Key 'ExifReadsSkippedVideo'
    $sourceDeletionLabel = Get-ImportProviderText -Key 'TakeoutSourceDeletion'
    $assetsLabel = Get-ImportProviderText -Key 'AssetsFirst500'
    $hashLabel = Get-ImportProviderText -Key 'Hash'
    $statusLabel = Get-ImportProviderText -Key 'Status'
    $pathLabel = Get-ImportProviderText -Key 'Path'
    $dateSourceLabel = Get-ImportProviderText -Key 'DateSource'
    $confidenceLabel = Get-ImportProviderText -Key 'Confidence'
    $exifVerificationLabel = Get-ImportProviderText -Key 'ExifVerification'
    $embeddedDateStateLabel = Get-ImportProviderText -Key 'EmbeddedDateState'
    $warningsLabel = Get-ImportProviderText -Key 'Warnings'
    $assetRows = New-Object System.Text.StringBuilder
    foreach ($asset in @($Assets | Select-Object -First 500)) {
        $albums = (@($asset.AlbumNames) -join ', ')
        $warnings = (@($asset.Warnings) -join '; ')
        [void]$assetRows.AppendLine(("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td></tr>" -f
            (ConvertTo-HtmlText $asset.Hash.Substring(0, [math]::Min(12, $asset.Hash.Length))),
            (ConvertTo-HtmlText $asset.Status),
            (ConvertTo-HtmlText $asset.PrimaryPath),
            (ConvertTo-HtmlText $asset.OccurrenceCount),
            (ConvertTo-HtmlText $albums),
            (ConvertTo-HtmlText $asset.DateSource),
            (ConvertTo-HtmlText $asset.MetadataConfidence),
            (ConvertTo-HtmlText $asset.ExifVerification),
            (ConvertTo-HtmlText $asset.EmbeddedCaptureDateState),
            (ConvertTo-HtmlText $warnings)))
    }

    $html = @"
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>ImportProvider Google Takeout</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #1f2937; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #d1d5db; padding: 6px 8px; text-align: left; vertical-align: top; }
    th { background: #f3f4f6; }
    .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 10px; margin: 16px 0; }
    .card { background: #f8fafc; border: 1px solid #dbe3ef; border-radius: 8px; padding: 10px; }
    .num { font-size: 24px; font-weight: 700; }
  </style>
</head>
<body>
  <h1>ImportProvider Google Takeout</h1>
  <p>${modeLabel}: $($Summary.Mode)</p>
  <div class="grid">
    <div class="card"><div class="num">$($Summary.MediaFiles)</div><div>$mediaFilesLabel</div></div>
    <div class="card"><div class="num">$($Summary.LogicalAssets)</div><div>$logicalAssetsLabel</div></div>
    <div class="card"><div class="num">$($Summary.Occurrences)</div><div>$occurrencesLabel</div></div>
    <div class="card"><div class="num">$($Summary.Albums)</div><div>$albumsLabel</div></div>
    <div class="card"><div class="num">$($Summary.InternalDuplicateOccurrences)</div><div>$internalDuplicatesLabel</div></div>
    <div class="card"><div class="num">$($Summary.TrashOccurrences)</div><div>$trashOccurrencesLabel</div></div>
    <div class="card"><div class="num">$($Summary.Videos)</div><div>$videosLabel</div></div>
    <div class="card"><div class="num">$($Summary.Conflicts)</div><div>$conflictsLabel</div></div>
    <div class="card"><div class="num">$($Summary.HighConfidenceAssets)</div><div>$highConfidenceLabel</div></div>
    <div class="card"><div class="num">$($Summary.MediumConfidenceAssets)</div><div>$mediumConfidenceLabel</div></div>
    <div class="card"><div class="num">$($Summary.LowConfidenceAssets)</div><div>$lowConfidenceLabel</div></div>
    <div class="card"><div class="num">$($Summary.ExifVerificationSkippedProviderTrusted)</div><div>$exifSkippedLabel</div></div>
    <div class="card"><div class="num">$($Summary.ExifVerificationSkippedProviderTrustedVideo)</div><div>$exifSkippedVideoLabel</div></div>
  </div>
  <p><strong>${sourceDeletionLabel}:</strong> $($Summary.SourceDeletionStatus) · $([System.Net.WebUtility]::HtmlEncode([string]$Summary.SourceDeletionPath))</p>
  <h2>$assetsLabel</h2>
  <table>
    <tr><th>$hashLabel</th><th>$statusLabel</th><th>$pathLabel</th><th>$occurrencesLabel</th><th>$albumsLabel</th><th>$dateSourceLabel</th><th>$confidenceLabel</th><th>$exifVerificationLabel</th><th>$embeddedDateStateLabel</th><th>$warningsLabel</th></tr>
    $assetRows
  </table>
</body>
</html>
"@
    $html | Set-Content -LiteralPath $htmlPath -Encoding UTF8

    $assetReportRows = @($Assets | ForEach-Object {
        [pscustomobject]@{
            hash = $_.Hash
            status = $_.Status
            primaryPath = $_.PrimaryPath
            targetPath = $_.TargetPath
            occurrenceCount = $_.OccurrenceCount
            albumNames = @($_.AlbumNames)
            warnings = @($_.Warnings)
            dateSource = $_.DateSource
            metadataConfidence = $_.MetadataConfidence
            exifVerification = $_.ExifVerification
            embeddedCaptureDateState = $_.EmbeddedCaptureDateState
            occurrences = @($_.Occurrences | ForEach-Object {
                [pscustomobject]@{
                    path = $_.File.FullName
                    folderRole = $_.FolderRole
                    albumTitle = $_.AlbumTitle
                    isTrash = $_.IsTrash
                    isVideo = $_.IsVideo
                    sidecarStatus = $_.SidecarStatus
                }
            })
        }
    })

    [pscustomobject]@{
        language = $Language
        labels = [pscustomobject]@{
            mode = $modeLabel
            mediaFiles = $mediaFilesLabel
            logicalAssets = $logicalAssetsLabel
            occurrences = $occurrencesLabel
            albums = $albumsLabel
            conflicts = $conflictsLabel
            confidence = $confidenceLabel
            exifVerification = $exifVerificationLabel
            warnings = $warningsLabel
        }
        summary = $Summary
        assets = $assetReportRows
        orphanJson = @($OrphanJson | ForEach-Object { $_.FullName })
        mediaWithoutJson = @($MediaWithoutJson | ForEach-Object { $_.FullName })
    } | ConvertTo-Json -Depth 16 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

    Write-Log -Message ((Get-ImportProviderText -Key 'ReportWritten') -f $htmlPath) -Phase 'ImportProvider'
    Write-Log -Message ((Get-ImportProviderText -Key 'JsonReportWritten') -f $jsonPath) -Phase 'ImportProvider'
    return [pscustomobject]@{ HtmlPath = $htmlPath; JsonPath = $jsonPath }
}

function Invoke-ImportProviderSourceDeletion {
    param(
        [string]$SelectedProviderPath,
        [pscustomobject]$Summary
    )

    $deletePath = Resolve-FullPath $SelectedProviderPath
    $Summary | Add-Member -NotePropertyName SourceDeletionPath -NotePropertyValue $deletePath -Force

    if (-not $DeleteImportProviderSourceAfterSuccess) {
        $Summary | Add-Member -NotePropertyName SourceDeletionStatus -NotePropertyValue 'NotRequested' -Force
        Write-Log -Message ((Get-ImportProviderText -Key 'SourceNotDeleted') -f $deletePath) -Phase 'ImportProvider'
        return
    }

    if (-not $Apply) {
        $Summary | Add-Member -NotePropertyName SourceDeletionStatus -NotePropertyValue 'SkippedDryRun' -Force
        Write-Log -Message ((Get-ImportProviderText -Key 'SourceDeletionSkippedDryRun') -f $deletePath) -Phase 'ImportProvider' -Status 'Warning'
        return
    }

    if ($Stats.Errors -gt 0) {
        $Summary | Add-Member -NotePropertyName SourceDeletionStatus -NotePropertyValue 'SkippedDueToErrors' -Force
        Write-Log -Message ((Get-ImportProviderText -Key 'SourceDeletionSkippedErrors') -f $Stats.Errors, $deletePath) -Phase 'ImportProvider' -Status 'Warning'
        return
    }

    if (-not (Test-Path -LiteralPath $deletePath -PathType Container)) {
        $Summary | Add-Member -NotePropertyName SourceDeletionStatus -NotePropertyValue 'SkippedPathNotFound' -Force
        Write-Log -Message ((Get-ImportProviderText -Key 'SourceDeletionSkippedMissing') -f $deletePath) -Phase 'ImportProvider' -Status 'Warning'
        return
    }

    $root = [System.IO.Path]::GetPathRoot($deletePath)
    if ($deletePath.TrimEnd('\') -eq $root.TrimEnd('\')) {
        $Summary | Add-Member -NotePropertyName SourceDeletionStatus -NotePropertyValue 'BlockedUnsafeRootPath' -Force
        Write-Log -Message ((Get-ImportProviderText -Key 'SourceDeletionBlockedRoot') -f $deletePath) -Phase 'ImportProvider' -Status 'Warning'
        return
    }

    if (Test-IsChildPath -Path $LogRoot -ParentPath $deletePath) {
        $Summary | Add-Member -NotePropertyName SourceDeletionStatus -NotePropertyValue 'BlockedReportInsideSource' -Force
        Write-Log -Message ((Get-ImportProviderText -Key 'SourceDeletionBlockedReportInsideSource') -f $deletePath, $LogRoot) -Phase 'ImportProvider' -Status 'Warning'
        return
    }

    try {
        Remove-Item -LiteralPath $deletePath -Recurse -Force -ErrorAction Stop
        $Summary | Add-Member -NotePropertyName SourceDeletionStatus -NotePropertyValue 'Deleted' -Force
        Write-Log -Message ((Get-ImportProviderText -Key 'SourceDeleted') -f $deletePath) -Phase 'ImportProvider' -Status 'Completed'
    }
    catch {
        $Summary | Add-Member -NotePropertyName SourceDeletionStatus -NotePropertyValue 'DeleteFailed' -Force
        Write-Log -Message ((Get-ImportProviderText -Key 'SourceDeletionFailed') -f $deletePath, $_.Exception.Message) -Phase 'ImportProvider' -Status 'Warning'
    }
}

function Copy-ProviderAssetToDestination {
    param(
        [pscustomobject]$Item,
        [string]$DestinationDirectory,
        [string]$ProviderName,
        [string]$ProviderRootPath,
        [string]$Reason
    )

    $targetFileName = Get-QuarterlyTargetFileName -Item $Item
    $Item | Add-Member -NotePropertyName TargetFileName -NotePropertyValue $targetFileName -Force
    $resolved = Resolve-UniquePath -Directory $DestinationDirectory -FileName $targetFileName -SourceHash $Item.Sha256
    if ($resolved.SkipExistingIdentical) {
        $Stats.ExistingIdenticalSkipped++
        if ($Apply) {
            Register-ImportedProviderFile -Item $Item -NewPath $resolved.Path -Status "$Reason - Existing identical target" -ProviderName $ProviderName -ProviderRootPath $ProviderRootPath
        }
        return $resolved.Path
    }

    if (-not $Apply) {
        $Stats.DryRunActions++
        Write-Log -Message ((Get-ImportProviderText -Key 'WouldImportAsset') -f $Item.File.FullName, $resolved.Path) -Phase 'ImportProvider'
        return $resolved.Path
    }

    if (-not (Test-Path -LiteralPath $DestinationDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $DestinationDirectory -Force | Out-Null
    }
    $sourceFile = $Item.File
    Copy-Item -LiteralPath $sourceFile.FullName -Destination $resolved.Path -ErrorAction Stop
    $Stats.FilesCopied++
    try {
        $destinationFile = Get-Item -LiteralPath $resolved.Path -ErrorAction Stop
        $Item.File = $destinationFile
        $materializationResult = Invoke-CaptureDateMaterialization -Item $Item -RootPath $DestinationBase -MetadataBackupRoot $MetadataBackupRoot -ExifToolAvailable $ExifToolAvailable -AllowMetadataWriteWithoutRepairExif -Reason "ImportProvider $ProviderName"
        if ($materializationResult.HashChanged -and -not [string]::IsNullOrWhiteSpace($materializationResult.NewHash)) {
            Write-Log -Message "Provider import hash updated after capture date materialization: oldHash=$($materializationResult.OldHash) newHash=$($materializationResult.NewHash) path=$($resolved.Path)" -Phase 'ImportProvider'
        }
    }
    finally {
        $Item.File = $sourceFile
    }
    Register-ImportedProviderFile -Item $Item -NewPath $resolved.Path -Status "$Reason - Copied" -ProviderName $ProviderName -ProviderRootPath $ProviderRootPath
    return $resolved.Path
}

function Convert-AppleDateTimeToDateTime {
    param([object]$Value)

    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }

    $styles = [System.Globalization.DateTimeStyles]::AllowWhiteSpaces -bor [System.Globalization.DateTimeStyles]::AssumeUniversal
    $result = [DateTime]::MinValue
    $candidates = @(
        $text,
        ($text -replace ',(?=\d{4}\b)', ', '),
        ($text -replace '^[A-Za-z]+\s+', ''),
        (($text -replace '^[A-Za-z]+\s+', '') -replace ',(?=\d{4}\b)', ', ')
    ) | Select-Object -Unique
    foreach ($candidate in $candidates) {
        foreach ($culture in @([System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.CultureInfo]::CurrentCulture)) {
            if ([DateTime]::TryParse($candidate, $culture, $styles, [ref]$result)) {
                return $result
            }
        }
    }
    return $null
}

function Find-ApplePhotoDetailsCsvFiles {
    param([string]$RootPath)

    $detailMatches = New-Object System.Collections.Generic.List[object]
    $csvFiles = @(Get-ChildItem -LiteralPath $RootPath -File -Filter '*.csv' -Recurse -Force -ErrorAction SilentlyContinue)
    foreach ($csv in $csvFiles) {
        try {
            $header = Get-Content -LiteralPath $csv.FullName -TotalCount 1 -ErrorAction Stop
            $required = @('imgName', 'fileChecksum', 'favorite', 'hidden', 'deleted', 'originalCreationDate', 'importDate')
            $matchCount = 0
            foreach ($name in $required) {
                if ($header -match "(^|,)$([regex]::Escape($name))($|,)") { $matchCount++ }
            }
            if ($matchCount -ge 5) { $detailMatches.Add($csv) }
        }
        catch {
            Write-Log -Message ((Get-ImportProviderText -Key 'ImportProviderInaccessibleFolder') -f $csv.FullName, $_.Exception.Message) -Phase 'ImportProvider' -Status 'Warning'
        }
    }
    return @($detailMatches.ToArray() | Sort-Object FullName)
}

function Find-ApplePhotoDetailsCsv {
    param([string]$RootPath)

    $detailMatches = @(Find-ApplePhotoDetailsCsvFiles -RootPath $RootPath)
    if ($detailMatches.Count -gt 0) { return $detailMatches[0] }
    return $null
}

function Resolve-ApplePhotosRoot {
    param([string]$RootPath)

    $resolved = Resolve-FullPath $RootPath
    if (-not (Test-Path -LiteralPath $resolved -PathType Container)) {
        Stop-WithError "ImportProviderPath does not exist: $resolved"
    }

    $detailsCsvFiles = @(Find-ApplePhotoDetailsCsvFiles -RootPath $resolved)
    if ($detailsCsvFiles.Count -eq 0) {
        Stop-WithError "Apple Photos / iCloud export was not detected under: $resolved. Expected CSV files with iCloud photo detail columns."
    }

    $photoDirectories = @($detailsCsvFiles | ForEach-Object { Resolve-FullPath $_.DirectoryName } | Select-Object -Unique)
    $candidateRoots = @($photoDirectories | ForEach-Object {
        $parent = Split-Path -Parent $_
        if ([string]::IsNullOrWhiteSpace($parent)) { $resolved } else { Resolve-FullPath $parent }
    } | Select-Object -Unique)

    $root = $resolved
    if ($candidateRoots.Count -eq 1) {
        $candidateRoot = [string]$candidateRoots[0]
        if ((Test-IsChildPath -Path $candidateRoot -ParentPath $resolved) -or $candidateRoot.Equals($resolved, [StringComparison]::OrdinalIgnoreCase)) {
            $root = $candidateRoot
        }
    }

    return [pscustomobject]@{
        Root = (Resolve-FullPath $root)
        PhotosDirectory = (Resolve-FullPath $photoDirectories[0])
        PhotosDirectories = @($photoDirectories)
        DetailsCsv = $detailsCsvFiles[0].FullName
        DetailsCsvs = @($detailsCsvFiles | ForEach-Object { $_.FullName })
        ExportPartRoots = @($candidateRoots)
        IsMultiPart = ($candidateRoots.Count -gt 1)
    }
}

function Split-AppleCsvLineCompat {
    param([string]$Line)

    try { Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue } catch { }
    $reader = New-Object System.IO.StringReader($Line)
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($reader)
    try {
        $parser.SetDelimiters(',')
        $parser.HasFieldsEnclosedInQuotes = $true
        return @($parser.ReadFields())
    }
    finally {
        $parser.Close()
        $reader.Dispose()
    }
}

function Read-ApplePhotoDetailsCsvRows {
    param([string]$CsvPath)

    $rows = New-Object System.Collections.Generic.List[object]
    $lines = @(Get-Content -LiteralPath $CsvPath -ErrorAction Stop)
    if ($lines.Count -le 1) { return @() }

    for ($i = 1; $i -lt $lines.Count; $i++) {
        $line = [string]$lines[$i]
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $fields = @(Split-AppleCsvLineCompat -Line $line)
        if ($fields.Count -lt 8) { continue }

        $originalCreationDate = [string]$fields[5]
        $index = 6
        if ($fields.Count -gt $index -and ([string]$fields[$index]) -match '^\d{4}\b') {
            $originalCreationDate = $originalCreationDate + ',' + [string]$fields[$index]
            $index++
        }

        $viewCount = if ($fields.Count -gt $index) { [string]$fields[$index] } else { '' }
        $index++

        $importDate = if ($fields.Count -gt $index) { [string]$fields[$index] } else { '' }
        $index++
        if ($fields.Count -gt $index -and ([string]$fields[$index]) -match '^\d{4}\b') {
            $importDate = $importDate + ',' + [string]$fields[$index]
        }

        $rows.Add([pscustomobject]@{
            imgName = [string]$fields[0]
            fileChecksum = [string]$fields[1]
            favorite = [string]$fields[2]
            hidden = [string]$fields[3]
            deleted = [string]$fields[4]
            originalCreationDate = $originalCreationDate
            viewCount = $viewCount
            importDate = $importDate
        })
    }

    return @($rows.ToArray())
}
function Read-ApplePhotoDetails {
    param([string[]]$CsvPaths)

    $records = @{}
    foreach ($csvPath in @($CsvPaths)) {
        if ([string]::IsNullOrWhiteSpace($csvPath)) { continue }
        if (-not (Test-Path -LiteralPath $csvPath -PathType Leaf)) { continue }

        $photosDirectory = Resolve-FullPath ([System.IO.Path]::GetDirectoryName($csvPath))
        $exportRoot = Split-Path -Parent $photosDirectory
        if ([string]::IsNullOrWhiteSpace($exportRoot)) { $exportRoot = $photosDirectory }
        $exportRoot = Resolve-FullPath $exportRoot

        $rows = @(Read-ApplePhotoDetailsCsvRows -CsvPath $csvPath)
        foreach ($row in $rows) {
            $name = [string]$row.imgName
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            if (-not $records.ContainsKey($name)) {
                $records[$name] = New-Object System.Collections.Generic.List[object]
            }
            $records[$name].Add([pscustomobject]@{
                Name = $name
                FileChecksum = [string]$row.fileChecksum
                Favorite = ([string]$row.favorite).Equals('yes', [StringComparison]::OrdinalIgnoreCase)
                Hidden = ([string]$row.hidden).Equals('yes', [StringComparison]::OrdinalIgnoreCase)
                Deleted = ([string]$row.deleted).Equals('yes', [StringComparison]::OrdinalIgnoreCase)
                OriginalCreationDate = Convert-AppleDateTimeToDateTime $row.originalCreationDate
                ImportDate = Convert-AppleDateTimeToDateTime $row.importDate
                DetailsCsv = $csvPath
                PhotosDirectory = $photosDirectory
                ExportRoot = $exportRoot
                Raw = $row
                Used = $false
            })
        }
    }
    return $records
}

function Get-ApplePhotosFolderRole {
    param(
        [string]$Path,
        [string]$AppleRoot
    )

    $relative = ConvertTo-RelativePath -Path $Path -BasePath $AppleRoot
    if ([string]::IsNullOrWhiteSpace($relative)) { return 'Root' }
    $segments = @($relative -split '[\\/]+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $trashAliases = @(
        'Recently Deleted', 'RecentlyDeleted',
        'Eliminado recientemente', 'Eliminados recientemente', 'Eliminadas recientemente', 'Eliminados', 'Papelera',
        'Sterse recent', 'Șterse recent', 'Cos', 'Coș'
    )
    foreach ($segment in $segments) {
        foreach ($alias in $trashAliases) {
            if ($segment.Equals($alias, [StringComparison]::OrdinalIgnoreCase)) { return 'Trash' }
        }
    }
    return 'Library'
}

function Get-AppleAlbumReferences {
    param(
        [string]$RootPath,
        [string[]]$DetailsCsvPaths,
        [hashtable]$MediaNames
    )

    $detailsPathSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($detailsPath in @($DetailsCsvPaths)) {
        if (-not [string]::IsNullOrWhiteSpace($detailsPath)) { [void]$detailsPathSet.Add((Resolve-FullPath $detailsPath)) }
    }

    $albumFiles = @(Get-ChildItem -LiteralPath $RootPath -File -Filter '*.csv' -Recurse -Force -ErrorAction SilentlyContinue | Where-Object {
        -not $detailsPathSet.Contains((Resolve-FullPath $_.FullName))
    })
    $referencesByName = @{}
    $usedAlbumFiles = New-Object System.Collections.Generic.List[object]
    $referenceCount = 0
    $memoryCsvCount = 0
    $albumCsvCount = 0

    foreach ($albumFile in $albumFiles) {
        try {
            $header = Get-Content -LiteralPath $albumFile.FullName -TotalCount 1 -ErrorAction Stop
            if ($header -match '(^|,)imageName($|,)') { $memoryCsvCount++ }
            elseif ($header -match '(^|,)Images($|,)' -or $header -match '(^|,)imgName($|,)') { $albumCsvCount++ }

            $rows = @(Import-Csv -LiteralPath $albumFile.FullName)
            $matched = 0
            foreach ($row in $rows) {
                $values = @($row.PSObject.Properties | ForEach-Object { [string]$_.Value })
                foreach ($value in $values) {
                    $name = $value.Trim('"').Trim()
                    if ([string]::IsNullOrWhiteSpace($name)) { continue }
                    if (-not $MediaNames.ContainsKey($name)) { continue }
                    if (-not $referencesByName.ContainsKey($name)) {
                        $referencesByName[$name] = New-Object System.Collections.Generic.List[string]
                    }
                    $albumName = $albumFile.BaseName
                    if (-not @($referencesByName[$name].ToArray()).Contains($albumName)) {
                        $referencesByName[$name].Add($albumName)
                        $referenceCount++
                    }
                    $matched++
                }
            }
            if ($matched -gt 0) { $usedAlbumFiles.Add($albumFile) }
        }
        catch {
            Write-Log -Message ((Get-ImportProviderText -Key 'ImportProviderInaccessibleFolder') -f $albumFile.FullName, $_.Exception.Message) -Phase 'ImportProvider' -Status 'Warning'
        }
    }

    return [pscustomobject]@{
        AlbumFiles = @($usedAlbumFiles.ToArray())
        ReferencesByName = $referencesByName
        ReferenceCount = $referenceCount
        CandidateCsvCount = @($albumFiles).Count
        AlbumCsvCount = $albumCsvCount
        MemoryCsvCount = $memoryCsvCount
    }
}
function Get-ApplePhotosImportDateInfo {
    param(
        [pscustomobject]$Asset,
        [pscustomobject]$Metadata,
        [bool]$EmbeddedMetadataRead
    )

    $providerDate = $Asset.ProviderDate
    $exifDate = ConvertTo-MediaDate $Metadata.DateTimeOriginal
    if (-not $exifDate) { $exifDate = ConvertTo-MediaDate $Metadata.CreateDate }
    if (-not $exifDate -and $Asset.IsVideo) { $exifDate = ConvertTo-MediaDate $Metadata.MediaCreateDate }

    if ($providerDate -and $exifDate) {
        if ([math]::Abs(($providerDate.Date - $exifDate.Date).TotalDays) -le 1) {
            return [pscustomobject]@{
                Date = $exifDate
                Confidence = 99
                Source = 'ApplePhotos'
                Sources = @('ApplePhotos', 'EXIF')
                Detail = 'Provider originalCreationDate agrees with embedded metadata'
                ProviderExifConflict = $false
                MetadataConfidence = 'HighConfidence'
                ExifVerification = 'Read'
            }
        }
        return [pscustomobject]@{
            Date = $providerDate
            Confidence = 20
            Source = 'ApplePhotos'
            Sources = @('ApplePhotos', 'EXIF')
            Detail = ("Provider/EXIF date conflict. Provider={0:yyyy-MM-dd}; EXIF={1:yyyy-MM-dd}" -f $providerDate, $exifDate)
            ProviderExifConflict = $true
            MetadataConfidence = 'LowConfidence'
            ExifVerification = 'Read'
        }
    }

    if ($providerDate) {
        return [pscustomobject]@{
            Date = $providerDate
            Confidence = 98
            Source = 'ApplePhotos'
            Sources = @('ApplePhotos')
            Detail = if ($EmbeddedMetadataRead) { 'Provider originalCreationDate used because embedded metadata is absent or incomplete' } else { 'Provider originalCreationDate used; embedded metadata verification skipped for trusted iCloud CSV row' }
            ProviderExifConflict = $false
            MetadataConfidence = 'MediumConfidence'
            ExifVerification = if ($EmbeddedMetadataRead) { 'Read' } else { 'SkippedProviderTrusted' }
        }
    }

    $item = [pscustomobject]@{
        File = $Asset.PrimaryOccurrence.File
        Metadata = $Metadata
    }
    $dateInfo = Get-PrimaryDate -Item $item -IsVideo $Asset.IsVideo
    $dateInfo | Add-Member -NotePropertyName MetadataConfidence -NotePropertyValue 'LowConfidence' -Force
    $dateInfo | Add-Member -NotePropertyName ExifVerification -NotePropertyValue $(if ($EmbeddedMetadataRead) { 'Read' } else { 'SkippedNoProviderDate' }) -Force
    return $dateInfo
}

function Test-ApplePhotosAssetNeedsExifVerification {
    param([pscustomobject]$Asset)

    if ($Diagnostic) { return $true }
    if (-not $Asset -or -not $Asset.ProviderDate) { return $true }
    if ($Asset.IsVideo) { return $true }
    if (@($Asset.Warnings | Where-Object { $_ -eq 'Media without Photo Details row' -or $_ -eq 'Ambiguous Photo Details row' }).Count -gt 0) { return $true }
    return $false
}

function Write-ApplePhotosImportReport {
    param(
        [pscustomobject]$Summary,
        [object[]]$Assets,
        [object[]]$AlbumFiles
    )

    if (-not (Test-Path -LiteralPath $LogRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null
    }

    $htmlPath = Join-Path $LogRoot ("ImportProvider-ApplePhotos-{0}.html" -f $script:RunId)
    $jsonPath = Join-Path $LogRoot ("ImportProvider-ApplePhotos-{0}.json" -f $script:RunId)
    $modeLabel = Get-ImportProviderText -Key 'Mode'
    $mediaFilesLabel = Get-ImportProviderText -Key 'MediaFiles'
    $logicalAssetsLabel = Get-ImportProviderText -Key 'LogicalAssets'
    $albumsLabel = Get-ImportProviderText -Key 'Albums'
    $trashOccurrencesLabel = Get-ImportProviderText -Key 'TrashOccurrences'
    $videosLabel = Get-ImportProviderText -Key 'Videos'
    $conflictsLabel = Get-ImportProviderText -Key 'Conflicts'
    $highConfidenceLabel = Get-ImportProviderText -Key 'HighConfidence'
    $mediumConfidenceLabel = Get-ImportProviderText -Key 'MediumConfidence'
    $lowConfidenceLabel = Get-ImportProviderText -Key 'LowConfidence'
    $exifSkippedLabel = Get-ImportProviderText -Key 'ExifReadsSkipped'
    $copiedLabel = Get-ImportProviderText -Key 'Copied'
    $assetsLabel = Get-ImportProviderText -Key 'AssetsFirst500'
    $hashLabel = Get-ImportProviderText -Key 'Hash'
    $statusLabel = Get-ImportProviderText -Key 'Status'
    $pathLabel = Get-ImportProviderText -Key 'Path'
    $dateSourceLabel = Get-ImportProviderText -Key 'DateSource'
    $confidenceLabel = Get-ImportProviderText -Key 'Confidence'
    $exifVerificationLabel = Get-ImportProviderText -Key 'ExifVerification'
    $embeddedDateStateLabel = Get-ImportProviderText -Key 'EmbeddedDateState'
    $warningsLabel = Get-ImportProviderText -Key 'Warnings'

    $assetRows = New-Object System.Text.StringBuilder
    foreach ($asset in @($Assets | Select-Object -First 500)) {
        $albums = (@($asset.AlbumNames) -join ', ')
        $warnings = (@($asset.Warnings) -join '; ')
        [void]$assetRows.AppendLine(("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td><td>{9}</td></tr>" -f
            (ConvertTo-HtmlText $asset.Hash.Substring(0, [math]::Min(12, $asset.Hash.Length))),
            (ConvertTo-HtmlText $asset.Status),
            (ConvertTo-HtmlText $asset.PrimaryPath),
            (ConvertTo-HtmlText $asset.ProviderChecksum),
            (ConvertTo-HtmlText $albums),
            (ConvertTo-HtmlText $asset.DateSource),
            (ConvertTo-HtmlText $asset.MetadataConfidence),
            (ConvertTo-HtmlText $asset.ExifVerification),
            (ConvertTo-HtmlText $asset.EmbeddedCaptureDateState),
            (ConvertTo-HtmlText $warnings)))
    }

    $html = @"
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>ImportProvider Apple Photos / iCloud</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #1f2937; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #d1d5db; padding: 6px 8px; text-align: left; vertical-align: top; }
    th { background: #f3f4f6; }
    .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 10px; margin: 16px 0; }
    .card { background: #f8fafc; border: 1px solid #dbe3ef; border-radius: 8px; padding: 10px; }
    .num { font-size: 24px; font-weight: 700; }
  </style>
</head>
<body>
  <h1>ImportProvider Apple Photos / iCloud</h1>
  <p>${modeLabel}: $($Summary.Mode)</p>
  <div class="grid">
    <div class="card"><div class="num">$($Summary.MediaFiles)</div><div>$mediaFilesLabel</div></div>
    <div class="card"><div class="num">$($Summary.LogicalAssets)</div><div>$logicalAssetsLabel</div></div>
    <div class="card"><div class="num">$($Summary.AlbumCsv)</div><div>$albumsLabel CSV</div></div>
    <div class="card"><div class="num">$($Summary.TrashOccurrences)</div><div>$trashOccurrencesLabel</div></div>
    <div class="card"><div class="num">$($Summary.Videos)</div><div>$videosLabel</div></div>
    <div class="card"><div class="num">$($Summary.LivePhotoCandidates)</div><div>Live Photo candidates</div></div>
    <div class="card"><div class="num">$($Summary.Conflicts)</div><div>$conflictsLabel</div></div>
    <div class="card"><div class="num">$($Summary.HighConfidenceAssets)</div><div>$highConfidenceLabel</div></div>
    <div class="card"><div class="num">$($Summary.MediumConfidenceAssets)</div><div>$mediumConfidenceLabel</div></div>
    <div class="card"><div class="num">$($Summary.LowConfidenceAssets)</div><div>$lowConfidenceLabel</div></div>
    <div class="card"><div class="num">$($Summary.ExifVerificationSkippedProviderTrusted)</div><div>$exifSkippedLabel</div></div>
    <div class="card"><div class="num">$($Summary.FilesCopied)</div><div>$copiedLabel</div></div>
  </div>
  <p><strong>$(Get-ImportProviderText -Key 'TakeoutSourceDeletion'):</strong> $($Summary.SourceDeletionStatus) · $([System.Net.WebUtility]::HtmlEncode([string]$Summary.SourceDeletionPath))</p>
  <h2>$assetsLabel</h2>
  <table>
    <tr><th>$hashLabel</th><th>$statusLabel</th><th>$pathLabel</th><th>Apple checksum</th><th>$albumsLabel</th><th>$dateSourceLabel</th><th>$confidenceLabel</th><th>$exifVerificationLabel</th><th>$embeddedDateStateLabel</th><th>$warningsLabel</th></tr>
    $assetRows
  </table>
</body>
</html>
"@
    $html | Set-Content -LiteralPath $htmlPath -Encoding UTF8

    [pscustomobject]@{
        Language = $Language
        Summary = $Summary
        Assets = @($Assets | ForEach-Object {
            [pscustomobject]@{
                Hash = [string]$_.Hash
                ProviderChecksum = [string]$_.ProviderChecksum
                Status = [string]$_.Status
                PrimaryPath = [string]$_.PrimaryPath
                TargetPath = [string]$_.TargetPath
                AlbumNames = @($_.AlbumNames)
                DateSource = [string]$_.DateSource
                MetadataConfidence = [string]$_.MetadataConfidence
                ExifVerification = [string]$_.ExifVerification
                EmbeddedCaptureDateState = [string]$_.EmbeddedCaptureDateState
                Warnings = @($_.Warnings)
            }
        })
        AlbumFiles = @($AlbumFiles | ForEach-Object { $_.FullName })
    } | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

    Write-Log -Message ((Get-ImportProviderText -Key 'ReportWritten') -f $htmlPath) -Phase 'ImportProvider'
    Write-Log -Message ((Get-ImportProviderText -Key 'JsonReportWritten') -f $jsonPath) -Phase 'ImportProvider'
    return [pscustomobject]@{ HtmlPath = $htmlPath; JsonPath = $jsonPath }
}

function Resolve-XmpSidecarLibraryRoot {
    param([string]$RootPath)

    $resolved = Resolve-FullPath $RootPath
    if (-not (Test-Path -LiteralPath $resolved -PathType Container)) {
        Stop-WithError "ImportProviderPath does not exist: $resolved"
    }
    return $resolved
}

function Get-InsensitivePropertyValue {
    param(
        [object]$Object,
        [string[]]$Names
    )

    if ($null -eq $Object) { return $null }
    $properties = @($Object.PSObject.Properties)
    foreach ($name in @($Names)) {
        $property = @($properties | Where-Object { $_.Name.Equals($name, [StringComparison]::OrdinalIgnoreCase) } | Select-Object -First 1)
        if ($property.Count -gt 0) {
            return $property[0].Value
        }
    }
    return $null
}

function Convert-ProviderTimestampToDateTime {
    param([object]$Value)

    if ($null -eq $Value) { return $null }
    if ($Value.PSObject.Properties.Name -contains 'timestamp') {
        return Convert-ProviderTimestampToDateTime -Value $Value.timestamp
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $seconds = 0L
    if ([Int64]::TryParse($text, [ref]$seconds) -and $seconds -gt 946684800 -and $seconds -lt 4102444800) {
        try { return [DateTimeOffset]::FromUnixTimeSeconds($seconds).LocalDateTime } catch { }
    }

    return ConvertTo-MediaDate $text
}

function Convert-SidecarValueToStringArray {
    param([object]$Value)

    if ($null -eq $Value) { return @() }
    if ($Value -is [array]) {
        return @($Value | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return @() }
    if ($text -match ',') {
        return @($text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    return @($text)
}

function Read-SimpleYamlSidecar {
    param([string]$Path)

    $map = [ordered]@{}
    try {
        foreach ($line in @(Get-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop)) {
            if ($line -match '^\s*#') { continue }
            if ($line -match '^\s*([A-Za-z0-9_.-]+)\s*:\s*(.*?)\s*$') {
                $key = $matches[1]
                $value = $matches[2].Trim()
                $value = $value.Trim('"').Trim("'")
                $map[$key] = $value
            }
        }
    }
    catch {
        Write-Log -Message "Provider YAML parse failed: $Path - $($_.Exception.Message)" -Phase 'ImportProvider' -Status 'Warning'
    }
    return [pscustomobject]$map
}

function Get-XmpSidecarAssociationKeys {
    param([System.IO.FileInfo]$Sidecar)

    $ext = $Sidecar.Extension.ToLowerInvariant()
    $name = $Sidecar.Name
    $directoryKey = (Resolve-FullPath $Sidecar.DirectoryName).ToLowerInvariant()
    $mediaExtPattern = (($MediaExtensions | ForEach-Object { [regex]::Escape($_.TrimStart('.')) }) -join '|')
    $sidecarExtPattern = [regex]::Escape($ext.TrimStart('.'))
    if ($name -match "^(?<media>.+\.($mediaExtPattern))\.$sidecarExtPattern$") {
        return @($directoryKey + '|' + $matches['media'].ToLowerInvariant())
    }

    $stem = [System.IO.Path]::GetFileNameWithoutExtension($name)
    if ($stem -match '^(metadata|metadatos|album|folder)$') {
        return @()
    }

    return @($directoryKey + '|' + $stem.ToLowerInvariant())
}

function Read-XmpSidecarMetadata {
    param([System.IO.FileInfo]$Sidecar)

    $record = [pscustomobject]@{
        File = $Sidecar
        Path = $Sidecar.FullName
        Directory = $Sidecar.DirectoryName
        SidecarKind = $Sidecar.Extension.TrimStart('.').ToUpperInvariant()
        Date = $null
        Title = ''
        Description = ''
        Tags = @()
        Rating = ''
        GPSLatitude = $null
        GPSLongitude = $null
        ParseStatus = 'Parsed'
        Used = $false
        FieldsFound = @()
    }

    $availability = Detect-StorageAvailability -Item $Sidecar
    if ($availability.State -eq 'CloudPlaceholder') {
        Register-CloudPlaceholderSkipped -Path $Sidecar.FullName -Phase 'ImportProvider' -Availability $availability
        $record.ParseStatus = 'CloudPlaceholder'
        return $record
    }
    if ($availability.State -eq 'MissingReal') {
        $Stats.MissingReal++
        $record.ParseStatus = 'MissingReal'
        return $record
    }

    try {
        $ext = $Sidecar.Extension.ToLowerInvariant()
        $raw = $null
        if ($ext -eq '.xmp') {
            if (-not $ExifToolAvailable) {
                $record.ParseStatus = 'ExifTool unavailable'
                return $record
            }
            $exif = Invoke-ExifTool -Path $ExifToolPath -Arguments @(
                '-j', '-n',
                '-DateTimeOriginal', '-CreateDate', '-MediaCreateDate', '-DateCreated', '-ModifyDate',
                '-Title', '-ObjectName', '-Headline', '-Description', '-Caption-Abstract',
                '-Subject', '-Keywords', '-HierarchicalSubject', '-Rating',
                '-GPSLatitude', '-GPSLongitude',
                $Sidecar.FullName
            ) -TimeoutSeconds 20
            if (-not $exif.Success -or [string]::IsNullOrWhiteSpace($exif.Output)) {
                $record.ParseStatus = if ($exif.TimedOut) { 'ExifTool timeout' } else { 'ExifTool read failed' }
                return $record
            }
            $raw = @($exif.Output | ConvertFrom-Json)[0]
        }
        elseif ($ext -eq '.json') {
            $raw = Read-ProviderJsonFile -Path $Sidecar.FullName
        }
        elseif ($ext -eq '.yaml' -or $ext -eq '.yml') {
            $raw = Read-SimpleYamlSidecar -Path $Sidecar.FullName
        }

        if ($null -eq $raw) {
            $record.ParseStatus = 'Parse failed'
            return $record
        }

        $dateValue = Get-InsensitivePropertyValue -Object $raw -Names @('DateTimeOriginal', 'DateCreated', 'dateCreated', 'photoTakenTime', 'dateTaken', 'dateTimeOriginal', 'CreateDate', 'createDate', 'createdAt', 'creationDate', 'captureDate', 'capturedAt', 'takenAt', 'timestamp', 'MediaCreateDate', 'ModifyDate')
        $record.Date = Convert-ProviderTimestampToDateTime -Value $dateValue
        $record.Title = [string](Get-InsensitivePropertyValue -Object $raw -Names @('Title', 'title', 'ObjectName', 'name', 'Headline', 'headline'))
        $record.Description = [string](Get-InsensitivePropertyValue -Object $raw -Names @('Description', 'description', 'Caption-Abstract', 'caption', 'comment', 'ImageDescription'))
        $record.Tags = @(Convert-SidecarValueToStringArray -Value (Get-InsensitivePropertyValue -Object $raw -Names @('Subject', 'subject', 'Keywords', 'keywords', 'tags', 'Tags', 'HierarchicalSubject')))
        $record.Rating = [string](Get-InsensitivePropertyValue -Object $raw -Names @('Rating', 'rating', 'score'))
        $record.GPSLatitude = Get-InsensitivePropertyValue -Object $raw -Names @('GPSLatitude', 'latitude', 'lat')
        $record.GPSLongitude = Get-InsensitivePropertyValue -Object $raw -Names @('GPSLongitude', 'longitude', 'lon', 'lng')

        $fields = New-Object System.Collections.Generic.List[string]
        if ($record.Date) { $fields.Add('date') }
        if (-not [string]::IsNullOrWhiteSpace($record.Title)) { $fields.Add('title') }
        if (-not [string]::IsNullOrWhiteSpace($record.Description)) { $fields.Add('description') }
        if (@($record.Tags).Count -gt 0) { $fields.Add('tags') }
        if (-not [string]::IsNullOrWhiteSpace($record.Rating)) { $fields.Add('rating') }
        if ($record.GPSLatitude -and $record.GPSLongitude) { $fields.Add('location') }
        $record.FieldsFound = @($fields.ToArray())
    }
    catch {
        $record.ParseStatus = "Parse failed: $($_.Exception.Message)"
    }

    return $record
}

function Get-XmpSidecarLibraryDateInfo {
    param(
        [pscustomobject]$Asset,
        [pscustomobject]$Metadata,
        [bool]$EmbeddedMetadataRead
    )

    $sidecar = $Asset.Sidecar
    $sidecarDate = if ($sidecar) { $sidecar.Date } else { $null }
    $exifDate = ConvertTo-MediaDate $Metadata.DateTimeOriginal
    if (-not $exifDate) { $exifDate = ConvertTo-MediaDate $Metadata.CreateDate }
    if (-not $exifDate -and $Asset.IsVideo) { $exifDate = ConvertTo-MediaDate $Metadata.MediaCreateDate }

    if ($sidecarDate -and $exifDate) {
        if ([math]::Abs(($sidecarDate.Date - $exifDate.Date).TotalDays) -le 1) {
            return [pscustomobject]@{
                Date = $exifDate
                Confidence = 99
                Source = 'XmpSidecarLibrary'
                Sources = @('XmpSidecarLibrary', 'EXIF')
                Detail = 'Sidecar capture date agrees with embedded metadata'
                ProviderExifConflict = $false
                MetadataConfidence = 'HighConfidence'
                ExifVerification = 'Read'
            }
        }
        return [pscustomobject]@{
            Date = $sidecarDate
            Confidence = 20
            Source = 'XmpSidecarLibrary'
            Sources = @('XmpSidecarLibrary', 'EXIF')
            Detail = ("Sidecar/EXIF date conflict. Sidecar={0:yyyy-MM-dd}; EXIF={1:yyyy-MM-dd}" -f $sidecarDate, $exifDate)
            ProviderExifConflict = $true
            MetadataConfidence = 'LowConfidence'
            ExifVerification = 'Read'
        }
    }

    if ($sidecarDate) {
        return [pscustomobject]@{
            Date = $sidecarDate
            Confidence = 97
            Source = 'XmpSidecarLibrary'
            Sources = @('XmpSidecarLibrary')
            Detail = 'Sidecar capture date used; embedded metadata verification skipped for a clear sidecar association'
            ProviderExifConflict = $false
            MetadataConfidence = 'MediumConfidence'
            ExifVerification = if ($EmbeddedMetadataRead) { 'Read' } else { 'SkippedProviderTrusted' }
        }
    }

    $item = [pscustomobject]@{
        File = $Asset.File
        Metadata = $Metadata
    }
    $dateInfo = Get-PrimaryDate -Item $item -IsVideo $Asset.IsVideo
    $dateInfo | Add-Member -NotePropertyName MetadataConfidence -NotePropertyValue 'LowConfidence' -Force
    $dateInfo | Add-Member -NotePropertyName ExifVerification -NotePropertyValue $(if ($EmbeddedMetadataRead) { 'Read' } else { 'SkippedNoSidecarDate' }) -Force
    return $dateInfo
}

function Write-XmpSidecarLibraryImportReport {
    param(
        [pscustomobject]$Summary,
        [object[]]$Assets,
        [object[]]$OrphanSidecars,
        [object[]]$FolderMetadata
    )

    if (-not (Test-Path -LiteralPath $LogRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null
    }

    $htmlPath = Join-Path $LogRoot ("ImportProvider-XmpSidecarLibrary-{0}.html" -f $script:RunId)
    $jsonPath = Join-Path $LogRoot ("ImportProvider-XmpSidecarLibrary-{0}.json" -f $script:RunId)
    $modeLabel = Get-ImportProviderText -Key 'Mode'
    $mediaFilesLabel = Get-ImportProviderText -Key 'MediaFiles'
    $sidecarsFoundLabel = Get-ImportProviderText -Key 'SidecarsFound'
    $sidecarsUsedLabel = Get-ImportProviderText -Key 'SidecarsUsed'
    $orphanSidecarsLabel = Get-ImportProviderText -Key 'OrphanSidecars'
    $mediaWithoutSidecarLabel = Get-ImportProviderText -Key 'MediaWithoutSidecar'
    $ambiguousSidecarsLabel = Get-ImportProviderText -Key 'AmbiguousSidecars'
    $highConfidenceLabel = Get-ImportProviderText -Key 'HighConfidence'
    $mediumConfidenceLabel = Get-ImportProviderText -Key 'MediumConfidence'
    $lowConfidenceLabel = Get-ImportProviderText -Key 'LowConfidence'
    $classicFallbackLabel = Get-ImportProviderText -Key 'ClassicFallbackAssets'
    $conflictsLabel = Get-ImportProviderText -Key 'Conflicts'
    $copiedLabel = Get-ImportProviderText -Key 'Copied'
    $assetsLabel = Get-ImportProviderText -Key 'AssetsFirst500'
    $statusLabel = Get-ImportProviderText -Key 'Status'
    $pathLabel = Get-ImportProviderText -Key 'Path'
    $sidecarLabel = Get-ImportProviderText -Key 'Sidecar'
    $dateSourceLabel = Get-ImportProviderText -Key 'DateSource'
    $confidenceLabel = Get-ImportProviderText -Key 'Confidence'
    $exifVerificationLabel = Get-ImportProviderText -Key 'ExifVerification'
    $embeddedDateStateLabel = Get-ImportProviderText -Key 'EmbeddedDateState'
    $sidecarFieldsLabel = Get-ImportProviderText -Key 'SidecarFields'
    $warningsLabel = Get-ImportProviderText -Key 'Warnings'
    $assetRows = New-Object System.Text.StringBuilder
    foreach ($asset in @($Assets | Select-Object -First 500)) {
        $warnings = (@($asset.Warnings) -join '; ')
        $fields = if ($asset.Sidecar) { (@($asset.Sidecar.FieldsFound) -join ', ') } else { '' }
        [void]$assetRows.AppendLine(("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td><td>{8}</td></tr>" -f
            (ConvertTo-HtmlText $asset.Status),
            (ConvertTo-HtmlText $asset.PrimaryPath),
            (ConvertTo-HtmlText $asset.SidecarStatus),
            (ConvertTo-HtmlText $asset.DateSource),
            (ConvertTo-HtmlText $asset.MetadataConfidence),
            (ConvertTo-HtmlText $asset.ExifVerification),
            (ConvertTo-HtmlText $asset.EmbeddedCaptureDateState),
            (ConvertTo-HtmlText $fields),
            (ConvertTo-HtmlText $warnings)))
    }

    $html = @"
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>ImportProvider XMP / Sidecar Library</title>
  <style>
    body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #1f2937; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #d1d5db; padding: 6px 8px; text-align: left; vertical-align: top; }
    th { background: #f3f4f6; }
    .grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 10px; margin: 16px 0; }
    .card { background: #f8fafc; border: 1px solid #dbe3ef; border-radius: 8px; padding: 10px; }
    .num { font-size: 24px; font-weight: 700; }
  </style>
</head>
<body>
  <h1>ImportProvider XMP / Sidecar Library</h1>
  <p>${modeLabel}: $($Summary.Mode)</p>
  <div class="grid">
    <div class="card"><div class="num">$($Summary.MediaFiles)</div><div>$mediaFilesLabel</div></div>
    <div class="card"><div class="num">$($Summary.SidecarsFound)</div><div>$sidecarsFoundLabel</div></div>
    <div class="card"><div class="num">$($Summary.SidecarsUsed)</div><div>$sidecarsUsedLabel</div></div>
    <div class="card"><div class="num">$($Summary.OrphanSidecars)</div><div>$orphanSidecarsLabel</div></div>
    <div class="card"><div class="num">$($Summary.MediaWithoutSidecar)</div><div>$mediaWithoutSidecarLabel</div></div>
    <div class="card"><div class="num">$($Summary.AmbiguousSidecars)</div><div>$ambiguousSidecarsLabel</div></div>
    <div class="card"><div class="num">$($Summary.HighConfidenceAssets)</div><div>$highConfidenceLabel</div></div>
    <div class="card"><div class="num">$($Summary.MediumConfidenceAssets)</div><div>$mediumConfidenceLabel</div></div>
    <div class="card"><div class="num">$($Summary.LowConfidenceAssets)</div><div>$lowConfidenceLabel</div></div>
    <div class="card"><div class="num">$($Summary.FallbackClassicAssets)</div><div>$classicFallbackLabel</div></div>
    <div class="card"><div class="num">$($Summary.Conflicts)</div><div>$conflictsLabel</div></div>
    <div class="card"><div class="num">$($Summary.FilesCopied)</div><div>$copiedLabel</div></div>
  </div>
  <p><strong>$(Get-ImportProviderText -Key 'TakeoutSourceDeletion'):</strong> $($Summary.SourceDeletionStatus) · $([System.Net.WebUtility]::HtmlEncode([string]$Summary.SourceDeletionPath))</p>
  <h2>$assetsLabel</h2>
  <table>
    <tr><th>$statusLabel</th><th>$pathLabel</th><th>$sidecarLabel</th><th>$dateSourceLabel</th><th>$confidenceLabel</th><th>$exifVerificationLabel</th><th>$embeddedDateStateLabel</th><th>$sidecarFieldsLabel</th><th>$warningsLabel</th></tr>
    $assetRows
  </table>
</body>
</html>
"@

    $html | Set-Content -LiteralPath $htmlPath -Encoding UTF8
    $reportAssets = @($Assets | ForEach-Object {
        [pscustomobject]@{
            Hash = if ($_.Hash) { [string]$_.Hash } else { '' }
            Status = [string]$_.Status
            PrimaryPath = [string]$_.PrimaryPath
            TargetPath = [string]$_.TargetPath
            SidecarStatus = [string]$_.SidecarStatus
            SidecarPath = if ($_.Sidecar) { [string]$_.Sidecar.Path } else { '' }
            DateSource = [string]$_.DateSource
            MetadataConfidence = [string]$_.MetadataConfidence
            ExifVerification = [string]$_.ExifVerification
            EmbeddedCaptureDateState = [string]$_.EmbeddedCaptureDateState
            AlbumNames = @($_.AlbumNames)
            Warnings = @($_.Warnings)
            SidecarFields = if ($_.Sidecar) { @($_.Sidecar.FieldsFound) } else { @() }
        }
    })
    [pscustomobject]@{
        Language = $Language
        Labels = [pscustomobject]@{
            Mode = $modeLabel
            MediaFiles = $mediaFilesLabel
            SidecarsFound = $sidecarsFoundLabel
            SidecarsUsed = $sidecarsUsedLabel
            OrphanSidecars = $orphanSidecarsLabel
            MediaWithoutSidecar = $mediaWithoutSidecarLabel
            AmbiguousSidecars = $ambiguousSidecarsLabel
            Confidence = $confidenceLabel
            ExifVerification = $exifVerificationLabel
            Warnings = $warningsLabel
        }
        Summary = $Summary
        Assets = $reportAssets
        OrphanSidecars = @($OrphanSidecars | ForEach-Object { $_.FullName })
        FolderMetadata = @($FolderMetadata | ForEach-Object { $_.FullName })
    } | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonPath -Encoding UTF8

    Write-Log -Message ((Get-ImportProviderText -Key 'ReportWritten') -f $htmlPath) -Phase 'ImportProvider'
    Write-Log -Message ((Get-ImportProviderText -Key 'JsonReportWritten') -f $jsonPath) -Phase 'ImportProvider'
    return [pscustomobject]@{ HtmlPath = $htmlPath; JsonPath = $jsonPath }
}

function Invoke-ImportProviderApplePhotos {
    param([string]$ProviderRootPath)

    $providerName = 'ApplePhotos'
    $displayName = Get-ImportProviderDisplayName -Provider $providerName
    $resolved = Resolve-ApplePhotosRoot -RootPath $ProviderRootPath
    Write-Log -Message ((Get-ImportProviderText -Key 'ImportStarted') -f $displayName, $(if ($Apply) { 'APPLY' } else { 'DRY RUN' }), $resolved.Root) -Phase 'ImportProvider'

    $mediaExts = @($ImageExtensions + $VideoExtensions)
    $allFiles = @(Get-ChildItem -LiteralPath $resolved.Root -File -Recurse -Force -ErrorAction SilentlyContinue)
    $mediaFiles = @($allFiles | Where-Object { $mediaExts -contains $_.Extension.ToLowerInvariant() })
    $detailsByName = Read-ApplePhotoDetails -CsvPaths @($resolved.DetailsCsvs)
    $mediaNames = @{}
    foreach ($file in $mediaFiles) {
        if (-not $mediaNames.ContainsKey($file.Name)) { $mediaNames[$file.Name] = $true }
    }
    $albumInfo = Get-AppleAlbumReferences -RootPath $resolved.Root -DetailsCsvPaths @($resolved.DetailsCsvs) -MediaNames $mediaNames

    $livePhotoBaseNames = @{}
    foreach ($group in @($mediaFiles | Group-Object @{ Expression = { (Resolve-FullPath $_.DirectoryName).ToLowerInvariant() + '|' + $_.BaseName.ToLowerInvariant() } })) {
        $images = @($group.Group | Where-Object { $ImageExtensions -contains $_.Extension.ToLowerInvariant() })
        $videos = @($group.Group | Where-Object { $VideoExtensions -contains $_.Extension.ToLowerInvariant() })
        if ($images.Count -gt 0 -and $videos.Count -gt 0) {
            $livePhotoBaseNames[[string]$group.Name] = $true
        }
    }

    $deletedCount = 0
    foreach ($list in $detailsByName.Values) {
        foreach ($entry in @($list.ToArray())) {
            if ($entry.Deleted) { $deletedCount++ }
        }
    }

    $Stats.FilesFound = $mediaFiles.Count
    $Stats.LocalFilesDetected = $mediaFiles.Count
    $Stats.FilesAnalyzed = 0
    Write-Log -Message ((Get-ImportProviderText -Key 'AppleScan') -f $mediaFiles.Count, @($resolved.DetailsCsvs).Count, $albumInfo.AlbumFiles.Count, $deletedCount, $livePhotoBaseNames.Count) -Phase 'ImportProvider'

    $hashMap = Get-Sha256Batch -Files $mediaFiles
    $occurrences = New-Object System.Collections.Generic.List[object]
    $mediaWithoutDetails = New-Object System.Collections.Generic.List[object]

    foreach ($file in $mediaFiles) {
        $Stats.FilesAnalyzed++
        $hashResult = if ($hashMap.ContainsKey($file.FullName)) { $hashMap[$file.FullName] } else { $null }
        if (-not $hashResult -or [string]::IsNullOrWhiteSpace($hashResult.Hash)) {
            $Stats.Errors++
            Write-Log -Message ((Get-ImportProviderText -Key 'ImportProviderHashFailed') -f $file.FullName, $hashResult.Error) -Phase 'ImportProvider' -Status 'Warning'
            continue
        }

        $details = @()
        if ($detailsByName.ContainsKey($file.Name)) {
            $details = @($detailsByName[$file.Name].ToArray())
            $samePartDetails = @($details | Where-Object { $_.ExportRoot -and ((Test-IsChildPath -Path $file.FullName -ParentPath $_.ExportRoot) -or (Resolve-FullPath $file.FullName).Equals((Resolve-FullPath $_.ExportRoot), [StringComparison]::OrdinalIgnoreCase)) })
            if ($samePartDetails.Count -gt 0) { $details = @($samePartDetails) }
            foreach ($detail in $details) { $detail.Used = $true }
        }
        else {
            $mediaWithoutDetails.Add($file)
        }

        $detail = if ($details.Count -gt 0) { $details[0] } else { $null }
        $albums = if ($albumInfo.ReferencesByName.ContainsKey($file.Name)) { @($albumInfo.ReferencesByName[$file.Name].ToArray()) } else { @() }
        $isVideo = $VideoExtensions -contains $file.Extension.ToLowerInvariant()
        $folderRole = Get-ApplePhotosFolderRole -Path $file.FullName -AppleRoot $resolved.Root
        $isTrash = $false
        if ($detail) { $isTrash = [bool]$detail.Deleted }
        if ($folderRole -eq 'Trash') { $isTrash = $true }
        $warnings = New-Object System.Collections.Generic.List[string]
        if (@($details).Count -eq 0) { $warnings.Add('Media without Photo Details row') }
        if (@($details).Count -gt 1) { $warnings.Add('Ambiguous Photo Details row') }
        if (@($albums).Count -gt 0) { $warnings.Add('Album reference') }
        $livePhotoKey = (Resolve-FullPath $file.DirectoryName).ToLowerInvariant() + '|' + $file.BaseName.ToLowerInvariant()
        if ($livePhotoBaseNames.ContainsKey($livePhotoKey)) { $warnings.Add('Live Photo pair candidate') }

        $occurrences.Add([pscustomobject]@{
            File = $file
            Hash = $hashResult.Hash.ToUpperInvariant()
            ProviderChecksum = if ($detail) { [string]$detail.FileChecksum } else { '' }
            Detail = $detail
            ProviderDate = if ($detail) { $detail.OriginalCreationDate } else { $null }
            IsTrash = $isTrash
            IsHidden = if ($detail) { [bool]$detail.Hidden } else { $false }
            IsFavorite = if ($detail) { [bool]$detail.Favorite } else { $false }
            IsVideo = $isVideo
            AlbumNames = @($albums)
            Warnings = $warnings
        })
    }

    $assets = New-Object System.Collections.Generic.List[object]
    foreach ($group in @($occurrences.ToArray() | Group-Object Hash)) {
        $groupOccurrences = @($group.Group)
        $primary = @($groupOccurrences | Where-Object { -not $_.IsTrash } | Sort-Object @{ Expression = { $_.File.FullName.Length } } | Select-Object -First 1)
        if (@($primary).Count -eq 0) { $primary = @($groupOccurrences | Sort-Object @{ Expression = { $_.File.FullName.Length } } | Select-Object -First 1) }
        $primary = @($primary)[0]
        $albumNames = @($groupOccurrences | ForEach-Object { $_.AlbumNames } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
        $warnings = New-Object System.Collections.Generic.List[string]
        foreach ($occurrence in $groupOccurrences) {
            foreach ($warning in @($occurrence.Warnings)) {
                if (-not @($warnings.ToArray()).Contains([string]$warning)) { $warnings.Add([string]$warning) }
            }
        }
        if (@($groupOccurrences | Where-Object { $_.IsTrash }).Count -gt 0) { $warnings.Add('Trash occurrence') }
        if (@($groupOccurrences | Where-Object { $_.IsHidden }).Count -gt 0) { $warnings.Add('Hidden asset') }
        if (@($groupOccurrences | Where-Object { $_.IsFavorite }).Count -gt 0) { $warnings.Add('Favorite asset') }

        $assets.Add([pscustomobject]@{
            Hash = [string]$group.Name
            Occurrences = @($groupOccurrences)
            OccurrenceCount = @($groupOccurrences).Count
            PrimaryOccurrence = $primary
            PrimaryPath = $primary.File.FullName
            ProviderChecksum = [string]$primary.ProviderChecksum
            ProviderDate = $primary.ProviderDate
            IsVideo = [bool]$primary.IsVideo
            AlbumNames = @($albumNames)
            Warnings = $warnings
            Status = 'Pending'
            DateSource = ''
            MetadataConfidence = 'Pending'
            ExifVerification = 'Pending'
            EmbeddedCaptureDateState = 'NotChecked'
            TargetPath = ''
        })
    }

    $metadataFiles = New-Object System.Collections.Generic.List[object]
    foreach ($asset in @($assets.ToArray())) {
        if (Test-ApplePhotosAssetNeedsExifVerification -Asset $asset) {
            $asset.ExifVerification = 'Queued'
            $metadataFiles.Add($asset.PrimaryOccurrence.File)
        }
        else {
            $asset.ExifVerification = 'SkippedProviderTrusted'
        }
    }
    $metadataSkipped = @($assets.ToArray() | Where-Object { $_.ExifVerification -eq 'SkippedProviderTrusted' }).Count
    Write-Log -Message ((Get-ImportProviderText -Key 'ExifPlan') -f 'Apple Photos / iCloud', $metadataFiles.Count, $metadataSkipped, $Diagnostic) -Phase 'ImportProvider'
    $metadataMap = Get-ExifMetadataBatch -Files @($metadataFiles.ToArray()) -ExifToolAvailable $ExifToolAvailable -ShowImportProviderProgress
    Load-ProcessedIndexLight
    if ($Apply) { Initialize-DestinationStructure }

    $importable = 0
    $conflicts = 0
    $trashOnly = 0
    $duplicateInIndex = 0
    foreach ($asset in @($assets.ToArray())) {
        $primary = $asset.PrimaryOccurrence
        $metadata = if ($metadataMap.ContainsKey($primary.File.FullName)) { $metadataMap[$primary.File.FullName] } else { New-EmptyMetadata }
        $embeddedMetadataRead = ([string]$metadata.ReadStatus -eq 'Read')
        $dateInfo = Get-ApplePhotosImportDateInfo -Asset $asset -Metadata $metadata -EmbeddedMetadataRead $embeddedMetadataRead
        $asset.DateSource = [string]$dateInfo.Source
        if ($dateInfo.PSObject.Properties.Name -contains 'MetadataConfidence') { $asset.MetadataConfidence = [string]$dateInfo.MetadataConfidence }
        if ($dateInfo.PSObject.Properties.Name -contains 'ExifVerification') { $asset.ExifVerification = [string]$dateInfo.ExifVerification }
        if ($dateInfo.PSObject.Properties.Name -contains 'Detail' -and $dateInfo.Detail) { $asset.Warnings.Add([string]$dateInfo.Detail) }

        if (@($asset.Occurrences | Where-Object { -not $_.IsTrash }).Count -eq 0) {
            $asset.Status = 'TrashOnlySkipped'
            $trashOnly++
            continue
        }
        if ($dateInfo.PSObject.Properties.Name -contains 'ProviderExifConflict' -and [bool]$dateInfo.ProviderExifConflict) {
            $asset.Status = 'ConflictNeedsReview'
            $conflicts++
        }
        elseif ($script:ProcessedByHash.ContainsKey($asset.Hash)) {
            $asset.Status = 'ExistingDuplicateSkipped'
            $duplicateInIndex++
            $Stats.ExactDuplicatesFound++
            continue
        }
        else {
            $asset.Status = 'Importable'
        }

        $item = [pscustomobject]@{
            File = $primary.File
            Extension = $primary.File.Extension.ToLowerInvariant()
            IsVideo = [bool]$primary.IsVideo
            IsRaw = $false
            Metadata = $metadata
            DateInfo = $dateInfo
            Sha256 = $asset.Hash
            GlobalDuplicate = $false
            PerceptualHash = $null
            Width = if ($metadata.ImageWidth) { [int]$metadata.ImageWidth } else { 0 }
            Height = if ($metadata.ImageHeight) { [int]$metadata.ImageHeight } else { 0 }
            DuplicateHandled = $false
            ProviderOccurrenceCount = $asset.OccurrenceCount
            ProviderAlbumNames = @($asset.AlbumNames)
            ProviderMetadataConfidence = $asset.MetadataConfidence
            ProviderExifVerification = $asset.ExifVerification
            EmbeddedMetadataReadStatus = [string]$metadata.ReadStatus
        }
        $asset.EmbeddedCaptureDateState = [string](Initialize-EmbeddedCaptureDateProbe -Item $item).State

        if ($asset.Status -eq 'ConflictNeedsReview') {
            $Stats.NeedsReview++
            $target = Copy-ProviderAssetToDestination -Item $item -DestinationDirectory $NeedsReviewRoot -ProviderName $providerName -ProviderRootPath $resolved.Root -Reason 'Apple Photos provider conflict'
            $asset.TargetPath = $target
            continue
        }

        $destinationDirectory = Get-DestinationPath -Item $item -RootPath $resolved.Root -OrganizedRoot $OrganizedRoot
        $targetPath = Copy-ProviderAssetToDestination -Item $item -DestinationDirectory $destinationDirectory -ProviderName $providerName -ProviderRootPath $resolved.Root -Reason 'Apple Photos import'
        $asset.TargetPath = $targetPath
        $importable++
    }

    if ($Apply) {
        Save-ProcessedDatabase
    }

    $summary = [pscustomobject]@{
        Mode = if ($Apply) { 'APPLY' } else { 'DRY RUN' }
        Provider = $providerName
        Root = $resolved.Root
        DetailsCsv = @($resolved.DetailsCsvs)
        MediaFiles = $mediaFiles.Count
        Videos = @($occurrences.ToArray() | Where-Object { $_.IsVideo }).Count
        AlbumCsv = @($albumInfo.AlbumFiles).Count
        AlbumReferences = $albumInfo.ReferenceCount
        LogicalAssets = @($assets.ToArray()).Count
        Occurrences = @($occurrences.ToArray()).Count
        InternalDuplicateOccurrences = [math]::Max(0, @($occurrences.ToArray()).Count - @($assets.ToArray()).Count)
        TrashOccurrences = @($occurrences.ToArray() | Where-Object { $_.IsTrash }).Count
        TrashOnlyAssets = $trashOnly
        LivePhotoCandidates = $livePhotoBaseNames.Count
        ImportableAssets = $importable
        ExistingDuplicates = $duplicateInIndex
        Conflicts = $conflicts
        HighConfidenceAssets = @($assets.ToArray() | Where-Object { $_.MetadataConfidence -eq 'HighConfidence' }).Count
        MediumConfidenceAssets = @($assets.ToArray() | Where-Object { $_.MetadataConfidence -eq 'MediumConfidence' }).Count
        LowConfidenceAssets = @($assets.ToArray() | Where-Object { $_.MetadataConfidence -eq 'LowConfidence' }).Count
        ExifVerificationRead = @($assets.ToArray() | Where-Object { $_.ExifVerification -eq 'Read' }).Count
        ExifVerificationSkippedProviderTrusted = @($assets.ToArray() | Where-Object { $_.ExifVerification -eq 'SkippedProviderTrusted' }).Count
        MediaWithoutDetails = $mediaWithoutDetails.Count
        DetailsCsvCount = @($resolved.DetailsCsvs).Count
        AppleExportParts = @($resolved.ExportPartRoots).Count
        AlbumCandidateCsv = $albumInfo.CandidateCsvCount
        MemoryCsv = $albumInfo.MemoryCsvCount
        FilesCopied = $Stats.FilesCopied
        DryRunActions = $Stats.DryRunActions
        Errors = $Stats.Errors
        SourceDeletionStatus = if ($DeleteImportProviderSourceAfterSuccess) { 'RequestedPending' } elseif ($Apply) { 'NotRequested' } else { 'NotRequested' }
        SourceDeletionPath = Resolve-FullPath $ProviderRootPath
    }
    $reports = Write-ApplePhotosImportReport -Summary $summary -Assets @($assets.ToArray()) -AlbumFiles @($albumInfo.AlbumFiles)
    Write-Log -Message ((Get-ImportProviderText -Key 'AppleSummary') -f $summary.MediaFiles, $summary.LogicalAssets, $summary.AlbumCsv, $summary.AlbumReferences, $summary.TrashOccurrences, $summary.Videos, $summary.LivePhotoCandidates, $summary.ImportableAssets, $summary.Conflicts, $summary.HighConfidenceAssets, $summary.MediumConfidenceAssets, $summary.LowConfidenceAssets, $summary.ExifVerificationRead, $summary.ExifVerificationSkippedProviderTrusted, $summary.MediaWithoutDetails, $summary.FilesCopied, $reports.HtmlPath) -Phase 'Complete' -Status 'Completed'

    Invoke-ImportProviderSourceDeletion -SelectedProviderPath $ProviderRootPath -Summary $summary
    $reports = Write-ApplePhotosImportReport -Summary $summary -Assets @($assets.ToArray()) -AlbumFiles @($albumInfo.AlbumFiles)
    Write-Log -Message ((Get-ImportProviderText -Key 'SourceDeletionStatusLog') -f $summary.SourceDeletionStatus, $summary.SourceDeletionPath, $reports.HtmlPath) -Phase 'Complete' -Status 'Completed'
    Cleanup-MetadataBackupsOnSuccess
}

function Invoke-ImportProviderXmpSidecarLibrary {
    param([string]$ProviderRootPath)

    $providerName = 'XmpSidecarLibrary'
    $displayName = Get-ImportProviderDisplayName -Provider $providerName
    $root = Resolve-XmpSidecarLibraryRoot -RootPath $ProviderRootPath
    Write-Log -Message ((Get-ImportProviderText -Key 'ImportStarted') -f $displayName, $(if ($Apply) { 'APPLY' } else { 'DRY RUN' }), $root) -Phase 'ImportProvider'

    $sidecarExts = @('.xmp', '.json', '.yaml', '.yml')
    $mediaExts = @($ImageExtensions + $VideoExtensions)
    $allFilesList = New-Object System.Collections.Generic.List[object]
    $directories = New-Object System.Collections.Generic.Stack[string]
    $directories.Push($root)
    while ($directories.Count -gt 0) {
        $directory = $directories.Pop()
        if (Test-IsExcludedPath -Path $directory) {
            Write-FolderProtectionSkipLog -Path $directory -Phase 'ImportProvider'
            continue
        }
        try {
            foreach ($childDirectory in @(Get-ChildItem -LiteralPath $directory -Directory -Force -ErrorAction Stop)) {
                if (Test-IsExcludedPath -Path $childDirectory.FullName) {
                    Write-FolderProtectionSkipLog -Path $childDirectory.FullName -Phase 'ImportProvider'
                    continue
                }
                $dirAvailability = Detect-StorageAvailability -Item $childDirectory -Directory
                if ($dirAvailability.State -eq 'CloudPlaceholder') {
                    Register-CloudPlaceholderSkipped -Path $childDirectory.FullName -Phase 'ImportProvider' -Availability $dirAvailability
                    continue
                }
                if ($dirAvailability.State -eq 'MissingReal') {
                    $Stats.MissingReal++
                    continue
                }
                $directories.Push($childDirectory.FullName)
            }
            foreach ($childFile in @(Get-ChildItem -LiteralPath $directory -File -Force -ErrorAction Stop)) {
                if (Test-IsExcludedPath -Path $childFile.FullName) { continue }
                $allFilesList.Add($childFile)
            }
        }
        catch {
            $Stats.Inaccessible++
            Write-Log -Message ((Get-ImportProviderText -Key 'ImportProviderInaccessibleFolder') -f $directory, $_.Exception.Message) -Phase 'ImportProvider'
        }
    }
    $allFiles = @($allFilesList.ToArray())
    $mediaFiles = @($allFiles | Where-Object { $mediaExts -contains $_.Extension.ToLowerInvariant() })
    $sidecarFiles = @($allFiles | Where-Object { $sidecarExts -contains $_.Extension.ToLowerInvariant() })
    $folderMetadataFiles = @($sidecarFiles | Where-Object { $_.BaseName -match '^(metadata|metadatos|album|folder)$' })
    Write-Log -Message ((Get-ImportProviderText -Key 'XmpScan') -f $mediaFiles.Count, $sidecarFiles.Count, $folderMetadataFiles.Count) -Phase 'ImportProvider'

    $sidecars = New-Object System.Collections.Generic.List[object]
    $sidecarsByKey = @{}
    foreach ($sidecarFile in $sidecarFiles) {
        $record = Read-XmpSidecarMetadata -Sidecar $sidecarFile
        $sidecars.Add($record)
        foreach ($key in @(Get-XmpSidecarAssociationKeys -Sidecar $sidecarFile)) {
            if (-not $sidecarsByKey.ContainsKey($key)) {
                $sidecarsByKey[$key] = New-Object System.Collections.Generic.List[object]
            }
            $sidecarsByKey[$key].Add($record)
        }
    }

    $folderMetadataByDirectory = @{}
    foreach ($folderMetadata in $folderMetadataFiles) {
        $record = @($sidecars.ToArray() | Where-Object { $_.Path -eq $folderMetadata.FullName } | Select-Object -First 1)
        if ($record.Count -gt 0) {
            $folderMetadataByDirectory[(Resolve-FullPath $folderMetadata.DirectoryName).ToLowerInvariant()] = $record[0]
        }
    }

    $Stats.FilesFound = $mediaFiles.Count
    $Stats.LocalFilesDetected = $mediaFiles.Count
    $hashMap = Get-Sha256Batch -Files $mediaFiles
    $assets = New-Object System.Collections.Generic.List[object]
    foreach ($file in $mediaFiles) {
        $Stats.FilesAnalyzed++
        $hashResult = if ($hashMap.ContainsKey($file.FullName)) { $hashMap[$file.FullName] } else { $null }
        if (-not $hashResult -or [string]::IsNullOrWhiteSpace($hashResult.Hash)) {
            $Stats.Errors++
            Write-Log -Message ((Get-ImportProviderText -Key 'ImportProviderHashFailed') -f $file.FullName, $hashResult.Error) -Phase 'ImportProvider' -Status 'Warning'
            continue
        }

        $dirKey = (Resolve-FullPath $file.DirectoryName).ToLowerInvariant()
        $fileKey = $dirKey + '|' + $file.Name.ToLowerInvariant()
        $stemKey = $dirKey + '|' + $file.BaseName.ToLowerInvariant()
        $matchedSidecars = New-Object System.Collections.Generic.List[object]
        foreach ($key in @($fileKey, $stemKey)) {
            if ($sidecarsByKey.ContainsKey($key)) {
                foreach ($candidate in @($sidecarsByKey[$key].ToArray())) {
                    if (-not (@($matchedSidecars.ToArray()) | Where-Object { $_.Path -eq $candidate.Path })) {
                        $matchedSidecars.Add($candidate)
                    }
                }
            }
        }

        $sidecarStatus = if ($matchedSidecars.Count -eq 1) { 'Matched' } elseif ($matchedSidecars.Count -gt 1) { 'Ambiguous' } else { 'Missing' }
        if ($matchedSidecars.Count -eq 1) { $matchedSidecars[0].Used = $true }
        $folderMetadata = if ($folderMetadataByDirectory.ContainsKey($dirKey)) { $folderMetadataByDirectory[$dirKey] } else { $null }
        $albumNames = @()
        if ($folderMetadata -and -not [string]::IsNullOrWhiteSpace([string]$folderMetadata.Title)) {
            $albumNames += [string]$folderMetadata.Title
        }
        else {
            $relativeFolder = ConvertTo-RelativePath -Path $file.DirectoryName -BasePath $root
            if (-not [string]::IsNullOrWhiteSpace($relativeFolder) -and $relativeFolder -ne '.') { $albumNames += $relativeFolder }
        }

        $warnings = New-Object System.Collections.Generic.List[string]
        if ($sidecarStatus -eq 'Ambiguous') { $warnings.Add('Ambiguous sidecar') }
        if ($sidecarStatus -eq 'Missing') { $warnings.Add('Media without sidecar; classic fallback') }
        $assets.Add([pscustomobject]@{
            Hash = $hashResult.Hash.ToUpperInvariant()
            File = $file
            IsVideo = ($VideoExtensions -contains $file.Extension.ToLowerInvariant())
            Sidecars = @($matchedSidecars.ToArray())
            Sidecar = if ($matchedSidecars.Count -eq 1) { $matchedSidecars[0] } else { $null }
            SidecarStatus = $sidecarStatus
            FolderMetadata = $folderMetadata
            AlbumNames = @($albumNames)
            Warnings = $warnings
            Status = 'Pending'
            DateSource = ''
            MetadataConfidence = 'Pending'
            ExifVerification = 'Pending'
            EmbeddedCaptureDateState = 'NotChecked'
            PrimaryPath = $file.FullName
            TargetPath = ''
        })
    }

    $metadataFiles = New-Object System.Collections.Generic.List[object]
    foreach ($asset in @($assets.ToArray())) {
        $needsExif = $Diagnostic -or $asset.SidecarStatus -ne 'Matched' -or -not $asset.Sidecar -or -not $asset.Sidecar.Date
        if ($needsExif) {
            $asset.ExifVerification = 'Queued'
            $metadataFiles.Add($asset.File)
        }
        else {
            $asset.ExifVerification = 'SkippedProviderTrusted'
        }
    }
    $metadataSkipped = @($assets.ToArray() | Where-Object { $_.ExifVerification -eq 'SkippedProviderTrusted' }).Count
    Write-Log -Message ((Get-ImportProviderText -Key 'ExifPlan') -f 'XMP / Sidecar Library', $metadataFiles.Count, $metadataSkipped, $Diagnostic) -Phase 'ImportProvider'
    $metadataMap = Get-ExifMetadataBatch -Files @($metadataFiles.ToArray()) -ExifToolAvailable $ExifToolAvailable -ShowImportProviderProgress
    Load-ProcessedIndexLight
    if ($Apply) { Initialize-DestinationStructure }

    $importable = 0
    $conflicts = 0
    $fallbackClassic = 0
    $duplicateInIndex = 0
    foreach ($asset in @($assets.ToArray())) {
        $metadata = if ($metadataMap.ContainsKey($asset.File.FullName)) { $metadataMap[$asset.File.FullName] } else { New-EmptyMetadata }
        $embeddedMetadataRead = ([string]$metadata.ReadStatus -eq 'Read')
        $dateInfo = Get-XmpSidecarLibraryDateInfo -Asset $asset -Metadata $metadata -EmbeddedMetadataRead $embeddedMetadataRead
        $asset.DateSource = [string]$dateInfo.Source
        if ($dateInfo.PSObject.Properties.Name -contains 'MetadataConfidence') { $asset.MetadataConfidence = [string]$dateInfo.MetadataConfidence }
        if ($dateInfo.PSObject.Properties.Name -contains 'ExifVerification') { $asset.ExifVerification = [string]$dateInfo.ExifVerification }
        if ($dateInfo.PSObject.Properties.Name -contains 'Detail' -and $dateInfo.Detail) { $asset.Warnings.Add([string]$dateInfo.Detail) }
        if ($asset.SidecarStatus -eq 'Missing' -or ($asset.SidecarStatus -eq 'Matched' -and -not $asset.Sidecar.Date)) { $fallbackClassic++ }

        if ($asset.SidecarStatus -eq 'Ambiguous' -or ($dateInfo.PSObject.Properties.Name -contains 'ProviderExifConflict' -and [bool]$dateInfo.ProviderExifConflict)) {
            $asset.Status = 'ConflictNeedsReview'
            $asset.MetadataConfidence = 'LowConfidence'
            $conflicts++
        }
        elseif ($script:ProcessedByHash.ContainsKey($asset.Hash)) {
            $asset.Status = 'ExistingDuplicateSkipped'
            $duplicateInIndex++
            $Stats.ExactDuplicatesFound++
            continue
        }
        else {
            $asset.Status = 'Importable'
        }

        $item = [pscustomobject]@{
            File = $asset.File
            Extension = $asset.File.Extension.ToLowerInvariant()
            IsVideo = [bool]$asset.IsVideo
            IsRaw = $false
            Metadata = $metadata
            DateInfo = $dateInfo
            Sha256 = $asset.Hash
            GlobalDuplicate = $false
            PerceptualHash = $null
            Width = if ($metadata.ImageWidth) { [int]$metadata.ImageWidth } else { 0 }
            Height = if ($metadata.ImageHeight) { [int]$metadata.ImageHeight } else { 0 }
            DuplicateHandled = $false
            ProviderOccurrenceCount = 1
            ProviderAlbumNames = @($asset.AlbumNames)
            ProviderMetadataConfidence = $asset.MetadataConfidence
            ProviderExifVerification = $asset.ExifVerification
            EmbeddedMetadataReadStatus = [string]$metadata.ReadStatus
        }
        $asset.EmbeddedCaptureDateState = [string](Initialize-EmbeddedCaptureDateProbe -Item $item).State

        if ($asset.Status -eq 'ConflictNeedsReview' -or $dateInfo.Confidence -lt $ExifRepairConfidence) {
            $Stats.NeedsReview++
            $target = Copy-ProviderAssetToDestination -Item $item -DestinationDirectory $NeedsReviewRoot -ProviderName $providerName -ProviderRootPath $root -Reason 'XMP sidecar provider review'
            $asset.TargetPath = $target
            continue
        }

        $destinationDirectory = Get-DestinationPath -Item $item -RootPath $root -OrganizedRoot $OrganizedRoot
        $targetPath = Copy-ProviderAssetToDestination -Item $item -DestinationDirectory $destinationDirectory -ProviderName $providerName -ProviderRootPath $root -Reason 'XMP sidecar import'
        $asset.TargetPath = $targetPath
        $importable++
    }

    if ($Apply) {
        Save-ProcessedDatabase
    }

    $folderMetadataPathSet = @{}
    foreach ($folderMetadataFile in $folderMetadataFiles) {
        $folderMetadataPathSet[(Resolve-FullPath $folderMetadataFile.FullName).ToLowerInvariant()] = $true
    }
    $orphanSidecars = @($sidecars.ToArray() | Where-Object { -not $_.Used -and -not $folderMetadataPathSet.ContainsKey((Resolve-FullPath $_.Path).ToLowerInvariant()) } | ForEach-Object { $_.File })
    $summary = [pscustomobject]@{
        Mode = if ($Apply) { 'APPLY' } else { 'DRY RUN' }
        Provider = $providerName
        Root = $root
        MediaFiles = $mediaFiles.Count
        Videos = @($assets.ToArray() | Where-Object { $_.IsVideo }).Count
        SidecarsFound = $sidecarFiles.Count
        SidecarsUsed = @($sidecars.ToArray() | Where-Object { $_.Used }).Count
        OrphanSidecars = $orphanSidecars.Count
        FolderMetadataSidecars = $folderMetadataFiles.Count
        MediaWithoutSidecar = @($assets.ToArray() | Where-Object { $_.SidecarStatus -eq 'Missing' }).Count
        AmbiguousSidecars = @($assets.ToArray() | Where-Object { $_.SidecarStatus -eq 'Ambiguous' }).Count
        LogicalAssets = $assets.Count
        ImportableAssets = $importable
        ExistingDuplicates = $duplicateInIndex
        Conflicts = $conflicts
        FallbackClassicAssets = $fallbackClassic
        HighConfidenceAssets = @($assets.ToArray() | Where-Object { $_.MetadataConfidence -eq 'HighConfidence' }).Count
        MediumConfidenceAssets = @($assets.ToArray() | Where-Object { $_.MetadataConfidence -eq 'MediumConfidence' }).Count
        LowConfidenceAssets = @($assets.ToArray() | Where-Object { $_.MetadataConfidence -eq 'LowConfidence' }).Count
        ExifVerificationRead = @($assets.ToArray() | Where-Object { $_.ExifVerification -eq 'Read' }).Count
        ExifVerificationSkippedProviderTrusted = @($assets.ToArray() | Where-Object { $_.ExifVerification -eq 'SkippedProviderTrusted' }).Count
        FilesCopied = $Stats.FilesCopied
        DryRunActions = $Stats.DryRunActions
        Errors = $Stats.Errors
        SourceDeletionStatus = if ($DeleteImportProviderSourceAfterSuccess) { 'RequestedPending' } elseif ($Apply) { 'NotRequested' } else { 'NotRequested' }
        SourceDeletionPath = Resolve-FullPath $ProviderRootPath
    }

    $reports = Write-XmpSidecarLibraryImportReport -Summary $summary -Assets @($assets.ToArray()) -OrphanSidecars $orphanSidecars -FolderMetadata $folderMetadataFiles
    Write-Log -Message ((Get-ImportProviderText -Key 'XmpSummary') -f $summary.MediaFiles, $summary.SidecarsFound, $summary.SidecarsUsed, $summary.OrphanSidecars, $summary.MediaWithoutSidecar, $summary.AmbiguousSidecars, $summary.FallbackClassicAssets, $summary.ImportableAssets, $summary.Conflicts, $summary.HighConfidenceAssets, $summary.MediumConfidenceAssets, $summary.LowConfidenceAssets, $summary.ExifVerificationRead, $summary.ExifVerificationSkippedProviderTrusted, $summary.FilesCopied, $reports.HtmlPath) -Phase 'Complete' -Status 'Completed'

    Invoke-ImportProviderSourceDeletion -SelectedProviderPath $ProviderRootPath -Summary $summary
    $reports = Write-XmpSidecarLibraryImportReport -Summary $summary -Assets @($assets.ToArray()) -OrphanSidecars $orphanSidecars -FolderMetadata $folderMetadataFiles
    Write-Log -Message ((Get-ImportProviderText -Key 'SourceDeletionStatusLog') -f $summary.SourceDeletionStatus, $summary.SourceDeletionPath, $reports.HtmlPath) -Phase 'Complete' -Status 'Completed'
    Cleanup-MetadataBackupsOnSuccess
}

function Invoke-ImportProviderGoogleTakeout {
    param([string]$ProviderRootPath)

    $providerName = 'GoogleTakeout'
    $displayName = Get-ImportProviderDisplayName -Provider $providerName
    $googleRoot = Resolve-GoogleTakeoutPhotosRoot -RootPath $ProviderRootPath
    Write-Log -Message ((Get-ImportProviderText -Key 'ImportStarted') -f $displayName, $(if ($Apply) { 'APPLY' } else { 'DRY RUN' }), $googleRoot) -Phase 'ImportProvider'

    $mediaExts = @($ImageExtensions + $VideoExtensions)
    $allFiles = @(Get-ChildItem -LiteralPath $googleRoot -File -Recurse -Force -ErrorAction SilentlyContinue)
    $mediaFiles = @($allFiles | Where-Object { $mediaExts -contains $_.Extension.ToLowerInvariant() })
    $albumMetadataFiles = @($allFiles | Where-Object { $_.Name -ieq 'metadatos.json' -or $_.Name -ieq 'metadata.json' })
    $candidateJsonFiles = @($allFiles | Where-Object { $_.Extension -ieq '.json' -and $_.Name -ine 'metadatos.json' -and $_.Name -ine 'metadata.json' })

    $sidecars = New-Object System.Collections.Generic.List[object]
    foreach ($jsonFile in $candidateJsonFiles) {
        $record = New-GoogleTakeoutSidecarRecord -JsonFile $jsonFile
        if ($record) { $sidecars.Add($record) }
    }
    $sidecarFiles = @($sidecars.ToArray() | ForEach-Object { $_.File })
    $sidecarPathSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($sidecarFile in $sidecarFiles) {
        [void]$sidecarPathSet.Add((Resolve-FullPath $sidecarFile.FullName))
    }
    $rootJsonFiles = @($candidateJsonFiles | Where-Object { -not $sidecarPathSet.Contains((Resolve-FullPath $_.FullName)) })

    $sidecarsByDirectoryAndTitle = @{}
    $copyIndexedSidecarsByDirectoryBaseAndIndex = @{}
    foreach ($sidecar in @($sidecars.ToArray())) {
        $key = (Resolve-FullPath $sidecar.Directory).ToLowerInvariant() + '|' + ([string]$sidecar.Title).ToLowerInvariant()
        if (-not $sidecarsByDirectoryAndTitle.ContainsKey($key)) {
            $sidecarsByDirectoryAndTitle[$key] = New-Object System.Collections.Generic.List[object]
        }
        $sidecarsByDirectoryAndTitle[$key].Add($sidecar)

        if ($sidecar.CopyInfo -and [bool]$sidecar.CopyInfo.HasCopyIndex) {
            $copyKey = (Resolve-FullPath $sidecar.Directory).ToLowerInvariant() + '|' + ([string]$sidecar.CopyInfo.BaseMediaName).ToLowerInvariant() + '|' + [string]$sidecar.CopyInfo.CopyIndex
            if (-not $copyIndexedSidecarsByDirectoryBaseAndIndex.ContainsKey($copyKey)) {
                $copyIndexedSidecarsByDirectoryBaseAndIndex[$copyKey] = New-Object System.Collections.Generic.List[object]
            }
            $copyIndexedSidecarsByDirectoryBaseAndIndex[$copyKey].Add($sidecar)
        }
    }

    $albumMetadataByDirectory = @{}
    foreach ($albumJson in $albumMetadataFiles) {
        $metadata = Get-GoogleTakeoutAlbumMetadata -DirectoryPath $albumJson.DirectoryName
        if ($metadata) {
            $albumMetadataByDirectory[(Resolve-FullPath $albumJson.DirectoryName).ToLowerInvariant()] = $metadata
        }
    }

    $Stats.FilesFound = $mediaFiles.Count
    $Stats.LocalFilesDetected = $mediaFiles.Count
    $Stats.FilesAnalyzed = 0
    Write-Log -Message ((Get-ImportProviderText -Key 'GoogleScan') -f $mediaFiles.Count, $sidecarFiles.Count, $albumMetadataByDirectory.Count, $rootJsonFiles.Count) -Phase 'ImportProvider'

    $hashMap = Get-Sha256Batch -Files $mediaFiles
    $occurrences = New-Object System.Collections.Generic.List[object]
    $mediaWithoutJson = New-Object System.Collections.Generic.List[object]
    foreach ($file in $mediaFiles) {
        $Stats.FilesAnalyzed++
        $hashResult = if ($hashMap.ContainsKey($file.FullName)) { $hashMap[$file.FullName] } else { $null }
        if (-not $hashResult -or [string]::IsNullOrWhiteSpace($hashResult.Hash)) {
            $Stats.Errors++
            Write-Log -Message ((Get-ImportProviderText -Key 'ImportProviderHashFailed') -f $file.FullName, $hashResult.Error) -Phase 'ImportProvider' -Status 'Warning'
            continue
        }

        $dirKey = (Resolve-FullPath $file.DirectoryName).ToLowerInvariant()
        $sidecarKeys = @(
            ($dirKey + '|' + $file.Name.ToLowerInvariant()),
            ($dirKey + '|' + $file.BaseName.ToLowerInvariant())
        ) | Select-Object -Unique
        $matchedSidecars = @()
        foreach ($sidecarKey in $sidecarKeys) {
            if ($sidecarsByDirectoryAndTitle.ContainsKey($sidecarKey)) {
                $matchedSidecars = @($sidecarsByDirectoryAndTitle[$sidecarKey].ToArray())
                break
            }
        }
        if (@($matchedSidecars).Count -eq 0) {
            $mediaCopyInfo = Get-GoogleTakeoutMediaCopyInfo -MediaFile $file
            if ($mediaCopyInfo.HasCopyIndex) {
                $copyKey = $dirKey + '|' + ([string]$mediaCopyInfo.BaseMediaName).ToLowerInvariant() + '|' + [string]$mediaCopyInfo.CopyIndex
                if ($copyIndexedSidecarsByDirectoryBaseAndIndex.ContainsKey($copyKey)) {
                    $copyIndexedMatches = @($copyIndexedSidecarsByDirectoryBaseAndIndex[$copyKey].ToArray())
                    $copyIndexedProviderDates = @($copyIndexedMatches | Where-Object { $_.PhotoTakenDate } | ForEach-Object { $_.PhotoTakenDate.Date } | Select-Object -Unique)
                    if ($copyIndexedMatches.Count -eq 1 -and $copyIndexedProviderDates.Count -eq 1) {
                        $matchedSidecars = @($copyIndexedMatches)
                    }
                }
            }
        }
        foreach ($sidecar in $matchedSidecars) { $sidecar.Used = $true }
        $matchedSidecarCount = @($matchedSidecars).Count
        if ($matchedSidecarCount -eq 0) { $mediaWithoutJson.Add($file) }

        $folderRole = Get-GoogleTakeoutFolderRole -DirectoryPath $file.DirectoryName -GoogleRoot $googleRoot
        $albumMetadata = if ($albumMetadataByDirectory.ContainsKey($dirKey)) { $albumMetadataByDirectory[$dirKey] } else { $null }
        $occurrences.Add([pscustomobject]@{
            File = $file
            Hash = $hashResult.Hash.ToUpperInvariant()
            Sidecars = @($matchedSidecars)
            FolderRole = $folderRole
            AlbumTitle = if ($albumMetadata) { [string]$albumMetadata.Title } else { '' }
            AlbumMetadata = $albumMetadata
            IsTrash = ($folderRole -eq 'Trash') -or (@($matchedSidecars | Where-Object { $_.IsTrashed }).Count -gt 0)
            IsVideo = $VideoExtensions -contains $file.Extension.ToLowerInvariant()
            SidecarStatus = if ($matchedSidecarCount -eq 1) { 'Matched' } elseif ($matchedSidecarCount -gt 1) { 'Ambiguous' } else { 'Missing' }
        })
    }

    $assets = New-Object System.Collections.Generic.List[object]
    foreach ($group in @($occurrences.ToArray() | Group-Object Hash)) {
        $groupOccurrences = @($group.Group)
        $primary = Select-GoogleTakeoutPrimaryOccurrence -Occurrences @($groupOccurrences | Where-Object { -not $_.IsTrash })
        if (-not $primary) { $primary = Select-GoogleTakeoutPrimaryOccurrence -Occurrences $groupOccurrences }
        $assetSidecars = @($groupOccurrences | ForEach-Object { $_.Sidecars } | Where-Object { $_ })
        $albumNames = @($groupOccurrences | Where-Object { -not [string]::IsNullOrWhiteSpace($_.AlbumTitle) } | ForEach-Object { $_.AlbumTitle } | Select-Object -Unique)
        $warnings = New-Object System.Collections.Generic.List[string]
        if (@($groupOccurrences | Where-Object { $_.FolderRole -eq 'Album' }).Count -gt 0) { $warnings.Add('Album occurrence') }
        if (@($groupOccurrences | Where-Object { $_.IsTrash }).Count -gt 0) { $warnings.Add('Trash occurrence') }
        if (@($groupOccurrences | Where-Object { $_.SidecarStatus -eq 'Ambiguous' }).Count -gt 0) { $warnings.Add('Ambiguous sidecar') }
        if (@($groupOccurrences | Where-Object { $_.SidecarStatus -eq 'Missing' }).Count -gt 0) { $warnings.Add('Media without exact sidecar') }

        $assets.Add([pscustomobject]@{
            Hash = [string]$group.Name
            Occurrences = @($groupOccurrences)
            OccurrenceCount = $groupOccurrences.Count
            PrimaryOccurrence = $primary
            PrimaryPath = $primary.File.FullName
            Sidecars = @($assetSidecars)
            AlbumNames = @($albumNames)
            Warnings = $warnings
            Status = 'Pending'
            DateSource = ''
            MetadataConfidence = 'Pending'
            ExifVerification = 'Pending'
            EmbeddedCaptureDateState = 'NotChecked'
            TargetPath = ''
        })
    }

    $metadataFiles = New-Object System.Collections.Generic.List[object]
    foreach ($asset in @($assets.ToArray())) {
        if (Test-GoogleTakeoutAssetNeedsExifVerification -Asset $asset) {
            $asset.ExifVerification = 'Queued'
            $metadataFiles.Add($asset.PrimaryOccurrence.File)
        }
        else {
            $asset.ExifVerification = if ($asset.PrimaryOccurrence.IsVideo) { 'SkippedProviderTrustedVideo' } else { 'SkippedProviderTrusted' }
        }
    }
    $metadataSkippedNonVideo = @($assets.ToArray() | Where-Object { $_.ExifVerification -eq 'SkippedProviderTrusted' }).Count
    $metadataSkippedVideo = @($assets.ToArray() | Where-Object { $_.ExifVerification -eq 'SkippedProviderTrustedVideo' }).Count
    $metadataSkipped = $metadataSkippedNonVideo + $metadataSkippedVideo
    Write-Log -Message ((Get-ImportProviderText -Key 'ExifPlan') -f 'Google Takeout', $metadataFiles.Count, $metadataSkipped, $Diagnostic) -Phase 'ImportProvider'
    Write-Log -Message ((Get-ImportProviderText -Key 'ExifPlanVideoTrusted') -f 'Google Takeout', $metadataSkippedVideo) -Phase 'ImportProvider'
    $metadataMap = Get-ExifMetadataBatch -Files @($metadataFiles.ToArray()) -ExifToolAvailable $ExifToolAvailable -ShowImportProviderProgress
    Load-ProcessedIndexLight
    if ($Apply) { Initialize-DestinationStructure }

    $importable = 0
    $conflicts = 0
    $trashOnly = 0
    $duplicateInIndex = 0
    foreach ($asset in @($assets.ToArray())) {
        $primary = $asset.PrimaryOccurrence
        $metadata = if ($metadataMap.ContainsKey($primary.File.FullName)) { $metadataMap[$primary.File.FullName] } else { New-EmptyMetadata }
        $embeddedMetadataRead = ([string]$metadata.ReadStatus -eq 'Read')
        $dateInfo = Get-GoogleTakeoutImportDateInfo -Asset $asset -PrimaryOccurrence $primary -Metadata $metadata -EmbeddedMetadataRead $embeddedMetadataRead
        $asset.DateSource = [string]$dateInfo.Source
        if ($dateInfo.PSObject.Properties.Name -contains 'MetadataConfidence') { $asset.MetadataConfidence = [string]$dateInfo.MetadataConfidence }
        if ($dateInfo.PSObject.Properties.Name -contains 'ExifVerification') { $asset.ExifVerification = [string]$dateInfo.ExifVerification }
        if ($dateInfo.PSObject.Properties.Name -contains 'Detail' -and $dateInfo.Detail) { $asset.Warnings.Add([string]$dateInfo.Detail) }

        if (@($asset.Occurrences | Where-Object { -not $_.IsTrash }).Count -eq 0) {
            $asset.Status = 'TrashOnlySkipped'
            $trashOnly++
            continue
        }

        if ($dateInfo.PSObject.Properties.Name -contains 'ProviderExifConflict' -and [bool]$dateInfo.ProviderExifConflict) {
            $asset.Status = 'ConflictNeedsReview'
            $conflicts++
        }
        elseif ($script:ProcessedByHash.ContainsKey($asset.Hash)) {
            $asset.Status = 'ExistingDuplicateSkipped'
            $duplicateInIndex++
            $Stats.ExactDuplicatesFound++
            continue
        }
        else {
            $asset.Status = 'Importable'
        }

        $item = [pscustomobject]@{
            File = $primary.File
            Extension = $primary.File.Extension.ToLowerInvariant()
            IsVideo = [bool]$primary.IsVideo
            IsRaw = $false
            Metadata = $metadata
            DateInfo = $dateInfo
            Sha256 = $asset.Hash
            GlobalDuplicate = $false
            PerceptualHash = $null
            Width = if ($metadata.ImageWidth) { [int]$metadata.ImageWidth } else { 0 }
            Height = if ($metadata.ImageHeight) { [int]$metadata.ImageHeight } else { 0 }
            DuplicateHandled = $false
            ProviderOccurrenceCount = $asset.OccurrenceCount
            ProviderAlbumNames = @($asset.AlbumNames)
            ProviderMetadataConfidence = $asset.MetadataConfidence
            ProviderExifVerification = $asset.ExifVerification
            EmbeddedMetadataReadStatus = [string]$metadata.ReadStatus
        }
        $asset.EmbeddedCaptureDateState = [string](Initialize-EmbeddedCaptureDateProbe -Item $item).State

        if ($asset.Status -eq 'ConflictNeedsReview') {
            $Stats.NeedsReview++
            $target = Copy-ProviderAssetToDestination -Item $item -DestinationDirectory $NeedsReviewRoot -ProviderName $providerName -ProviderRootPath $googleRoot -Reason 'Google Takeout provider conflict'
            $asset.TargetPath = $target
            continue
        }

        $destinationDirectory = Get-DestinationPath -Item $item -RootPath $googleRoot -OrganizedRoot $OrganizedRoot
        $targetPath = Copy-ProviderAssetToDestination -Item $item -DestinationDirectory $destinationDirectory -ProviderName $providerName -ProviderRootPath $googleRoot -Reason 'Google Takeout import'
        $asset.TargetPath = $targetPath
        $importable++
    }

    if ($Apply) {
        Save-ProcessedDatabase
    }

    $orphanJson = @($sidecars.ToArray() | Where-Object { -not $_.Used } | ForEach-Object { $_.File })
    $summary = [pscustomobject]@{
        Mode = if ($Apply) { 'APPLY' } else { 'DRY RUN' }
        Provider = $providerName
        Root = $googleRoot
        MediaFiles = $mediaFiles.Count
        Videos = @($occurrences.ToArray() | Where-Object { $_.IsVideo }).Count
        SupplementalJson = $sidecarFiles.Count
        RootJson = $rootJsonFiles.Count
        Albums = $albumMetadataByDirectory.Count
        LogicalAssets = $assets.Count
        Occurrences = $occurrences.Count
        InternalDuplicateOccurrences = [math]::Max(0, $occurrences.Count - $assets.Count)
        AlbumOccurrences = @($occurrences.ToArray() | Where-Object { $_.FolderRole -eq 'Album' }).Count
        TrashOccurrences = @($occurrences.ToArray() | Where-Object { $_.IsTrash }).Count
        TrashOnlyAssets = $trashOnly
        ImportableAssets = $importable
        ExistingDuplicates = $duplicateInIndex
        Conflicts = $conflicts
        HighConfidenceAssets = @($assets.ToArray() | Where-Object { $_.MetadataConfidence -eq 'HighConfidence' }).Count
        MediumConfidenceAssets = @($assets.ToArray() | Where-Object { $_.MetadataConfidence -eq 'MediumConfidence' }).Count
        LowConfidenceAssets = @($assets.ToArray() | Where-Object { $_.MetadataConfidence -eq 'LowConfidence' }).Count
        ExifVerificationRead = @($assets.ToArray() | Where-Object { $_.ExifVerification -eq 'Read' }).Count
        ExifVerificationSkippedProviderTrusted = @($assets.ToArray() | Where-Object { $_.ExifVerification -eq 'SkippedProviderTrusted' -or $_.ExifVerification -eq 'SkippedProviderTrustedVideo' }).Count
        ExifVerificationSkippedProviderTrustedNonVideo = @($assets.ToArray() | Where-Object { $_.ExifVerification -eq 'SkippedProviderTrusted' }).Count
        ExifVerificationSkippedProviderTrustedVideo = @($assets.ToArray() | Where-Object { $_.ExifVerification -eq 'SkippedProviderTrustedVideo' }).Count
        SidecarsUsed = @($sidecars.ToArray() | Where-Object { $_.Used }).Count
        SidecarsAmbiguous = @($occurrences.ToArray() | Where-Object { $_.SidecarStatus -eq 'Ambiguous' }).Count
        JsonOrphans = $orphanJson.Count
        MediaWithoutJson = $mediaWithoutJson.Count
        FilesCopied = $Stats.FilesCopied
        DryRunActions = $Stats.DryRunActions
        Errors = $Stats.Errors
        SourceDeletionStatus = if ($DeleteImportProviderSourceAfterSuccess) { 'RequestedPending' } elseif ($Apply) { 'NotRequested' } else { 'NotRequested' }
        SourceDeletionPath = Resolve-FullPath $ProviderRootPath
    }
    $reports = Write-GoogleTakeoutImportReport -Summary $summary -Assets @($assets.ToArray()) -OrphanJson $orphanJson -MediaWithoutJson @($mediaWithoutJson.ToArray())
    Write-Log -Message ((Get-ImportProviderText -Key 'GoogleSummary') -f $summary.MediaFiles, $summary.LogicalAssets, $summary.Occurrences, $summary.Albums, $summary.InternalDuplicateOccurrences, $summary.TrashOccurrences, $summary.Videos, $summary.ImportableAssets, $summary.Conflicts, $summary.HighConfidenceAssets, $summary.MediumConfidenceAssets, $summary.LowConfidenceAssets, $summary.ExifVerificationRead, $summary.ExifVerificationSkippedProviderTrusted, $summary.SidecarsUsed, $summary.SidecarsAmbiguous, $summary.JsonOrphans, $summary.MediaWithoutJson, $summary.FilesCopied, $reports.HtmlPath) -Phase 'Complete' -Status 'Completed'

    Invoke-ImportProviderSourceDeletion -SelectedProviderPath $ProviderRootPath -Summary $summary
    $reports = Write-GoogleTakeoutImportReport -Summary $summary -Assets @($assets.ToArray()) -OrphanJson $orphanJson -MediaWithoutJson @($mediaWithoutJson.ToArray())
    Write-Log -Message ((Get-ImportProviderText -Key 'SourceDeletionStatusLog') -f $summary.SourceDeletionStatus, $summary.SourceDeletionPath, $reports.HtmlPath) -Phase 'Complete' -Status 'Completed'
    Cleanup-MetadataBackupsOnSuccess
}

function Test-DedupeExcludedPath {
    param([string]$Path)

    return [bool]($script:DedupeExcludedRoots | Where-Object { Test-IsChildPath -Path $Path -ParentPath $_ })
}

function Get-DedupeCleanupFiles {
    param([string]$RootPath)

    Write-Log -Message "DedupeCleanup scanning files..." -Phase 'DedupeCleanup'
    $files = New-Object System.Collections.Generic.List[object]
    $directories = New-Object System.Collections.Generic.Stack[string]
    $directories.Push($RootPath)
    $lastScanLog = Get-Date

    while ($directories.Count -gt 0) {
        $directory = $directories.Pop()
        if (Test-DedupeExcludedPath -Path $directory) {
            Write-FolderProtectionSkipLog -Path $directory -Phase 'DedupeCleanup'
            continue
        }

        Write-Heartbeat -Phase 'DedupeCleanup' -Message "DedupeCleanup scanning folder: $directory" -EverySeconds 10

        try {
            $childDirectories = @(Get-ChildItem -LiteralPath $directory -Directory -Force -ErrorAction Stop)
        }
        catch {
            $Stats.Inaccessible++
            Write-Log -Message "DedupeCleanup inaccessible folder skipped: $directory - $($_.Exception.Message)" -Phase 'DedupeCleanup'
            continue
        }

        foreach ($childDirectory in $childDirectories) {
            if (Test-DedupeExcludedPath -Path $childDirectory.FullName) {
                Write-FolderProtectionSkipLog -Path $childDirectory.FullName -Phase 'DedupeCleanup'
                continue
            }
            $dirAvailability = Detect-StorageAvailability -Item $childDirectory -Directory
            if ($dirAvailability.State -eq 'CloudPlaceholder') {
                Register-CloudPlaceholderSkipped -Path $childDirectory.FullName -Phase 'DedupeCleanup' -Availability $dirAvailability
                continue
            }
            if ($dirAvailability.State -eq 'MissingReal') {
                $Stats.MissingReal++
                continue
            }
            $directories.Push($childDirectory.FullName)
        }

        try {
            $childFiles = @(Get-ChildItem -LiteralPath $directory -File -Force -ErrorAction Stop)
        }
        catch {
            $Stats.Inaccessible++
            Write-Log -Message "DedupeCleanup inaccessible files skipped: $directory - $($_.Exception.Message)" -Phase 'DedupeCleanup'
            continue
        }

        foreach ($file in $childFiles) {
            if (Test-DedupeExcludedPath -Path $file.FullName) { continue }
            if ($MediaExtensions -notcontains $file.Extension.ToLowerInvariant()) { continue }
            $availability = Detect-StorageAvailability -Item $file
            if ($availability.State -eq 'CloudPlaceholder') {
                Register-CloudPlaceholderSkipped -Path $file.FullName -Phase 'DedupeCleanup' -Availability $availability
                continue
            }
            if ($availability.State -eq 'MissingReal') {
                $Stats.MissingReal++
                continue
            }
            if (-not (Test-ReadableFile -Path $file.FullName)) {
                $Stats.Inaccessible++
                Write-Log -Message "DedupeCleanup inaccessible file skipped: $($file.FullName)" -Phase 'DedupeCleanup'
                continue
            }

            $files.Add($file)
            $Stats.LocalFilesDetected++
            if ($files.Count -eq 1 -or $files.Count % 1000 -eq 0 -or ((Get-Date) - $lastScanLog).TotalSeconds -ge 10) {
                $lastScanLog = Get-Date
                Write-Log -Message "DedupeCleanup scan found $($files.Count) media files so far. Current folder: $directory" -Phase 'DedupeCleanup'
            }
        }
    }

    Write-Log -Message "DedupeCleanup found $($files.Count) media files." -Phase 'DedupeCleanup'
    return $files.ToArray()
}

function Get-DedupeArea {
    param([string]$Path)

    if ($script:DedupeOrganizedRoots | Where-Object { Test-IsChildPath -Path $Path -ParentPath $_ }) { return 'Organized' }
    if ($script:DedupeDuplicatesRoots | Where-Object { Test-IsChildPath -Path $Path -ParentPath $_ }) { return 'DuplicatesReview' }
    if ($script:DedupeNeedsReviewRoots | Where-Object { Test-IsChildPath -Path $Path -ParentPath $_ }) { return 'NeedsReview' }
    return 'Source'
}

function Get-DedupeCanonicalScore {
    param([pscustomobject]$Entry)

    $score = 0
    switch ($Entry.Area) {
        'Organized' { $score += 10000 }
        'Source' { $score += 7000 }
        'NeedsReview' { $score += 2000 }
        'DuplicatesReview' { $score += 1000 }
    }

    if (-not $Entry.IsReviewArea) { $score += 5000 }
    if ($Entry.IsRaw) { $score += 250 }
    if (Test-IsQuarterlyOrganizedPath -Path $Entry.Path) { $score += 150 }
    $score += [math]::Min(100, [math]::Floor($Entry.Size / 1MB))
    return [int]$score
}

function Test-IsQuarterlyOrganizedPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return $false
    }

    $parts = @($Path -split '[\\/]' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    for ($i = 0; $i -lt ($parts.Count - 1); $i++) {
        if ($parts[$i] -notmatch '^(19|20)\d{2}$') {
            continue
        }

        if (Get-QuarterStartMonthFromFolderName -FolderName $parts[$i + 1]) {
            return $true
        }
    }

    return $false
}

function Get-DedupeCanonicalEntry {
    param([object[]]$Entries)

    return @($Entries | Sort-Object @{ Expression = { Get-DedupeCanonicalScore -Entry $_ }; Descending = $true }, @{ Expression = { $_.Path.Length }; Descending = $false }, Path | Select-Object -First 1)[0]
}

function Resolve-DedupeQuarantinePath {
    param([pscustomobject]$Entry)

    $relative = ConvertTo-RelativePath -Path $Entry.Path -BasePath $SourcePath
    if ([string]::IsNullOrWhiteSpace($relative) -or $relative -eq $Entry.Path) {
        $relative = Join-Path $Entry.Hash $Entry.File.Name
    }
    $target = Join-Path $ConfirmedDuplicatesQuarantineRoot $relative
    $targetDirectory = [System.IO.Path]::GetDirectoryName($target)
    return Resolve-UniquePath -Directory $targetDirectory -FileName ([System.IO.Path]::GetFileName($target)) -SourceHash '__dedupe_quarantine_never_skip__'
}

function Resolve-DedupePromotionPath {
    param([pscustomobject]$Entry)

    $date = Get-DateFromFileName -FileName $Entry.File.Name
    if (-not $date) {
        $date = $Entry.File.LastWriteTime
    }

    $targetDirectory = Get-QuarterlyDestinationDirectory -Date $date -OrganizedRoot $OrganizedRoot
    return Resolve-UniquePath -Directory $targetDirectory -FileName $Entry.File.Name -SourceHash '__dedupe_promotion_never_skip__'
}

function Move-DedupeEntry {
    param(
        [pscustomobject]$Entry,
        [string]$TargetPath,
        [string]$Action
    )

    if (-not $Apply) {
        return
    }

    $targetDirectory = [System.IO.Path]::GetDirectoryName($TargetPath)
    if (-not (Test-Path -LiteralPath $targetDirectory)) {
        New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
    }

    Move-Item -LiteralPath $Entry.Path -Destination $TargetPath -ErrorAction Stop
    Write-Log -Message "${Action}: $($Entry.Path) -> $TargetPath" -Phase 'DedupeCleanup'
}

function ConvertTo-HtmlText {
    param([object]$Text)
    if ($null -eq $Text) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Text)
}

function Write-DedupeCleanupReport {
    param(
        [hashtable]$Summary,
        [object[]]$Actions,
        [object[]]$ManualReview
    )

    $reportPath = Join-Path $LogRoot ("DedupeCleanupReport-{0}.html" -f $script:RunId)
    $actionRows = New-Object System.Text.StringBuilder
    foreach ($action in @($Actions | Select-Object -First 2000)) {
        [void]$actionRows.AppendLine(("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>" -f (ConvertTo-HtmlText $action.Action), (ConvertTo-HtmlText $action.Hash), (ConvertTo-HtmlText $action.Source), (ConvertTo-HtmlText $action.Target)))
    }

    $manualRows = New-Object System.Text.StringBuilder
    foreach ($item in @($ManualReview | Select-Object -First 1000)) {
        [void]$manualRows.AppendLine(("<tr><td>{0}</td><td>{1}</td><td>{2}</td></tr>" -f (ConvertTo-HtmlText $item.Hash), (ConvertTo-HtmlText $item.Path), (ConvertTo-HtmlText $item.Reason)))
    }

    $modeText = if ($Apply) { 'APPLY' } else { 'DRY RUN' }
    $html = @"
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>DedupeCleanupReport</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#1f2933}
table{border-collapse:collapse;width:100%;margin:16px 0}
th,td{border:1px solid #d7dde5;padding:6px 8px;text-align:left;font-size:13px}
th{background:#eef2f7}
.summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px}
.card{border:1px solid #d7dde5;padding:10px;background:#f8fafc}
</style>
</head>
<body>
<h1>DedupeCleanupReport</h1>
<p>Mode: $modeText</p>
<div class="summary">
<div class="card"><strong>Total files scanned</strong><br>$($Summary.TotalFiles)</div>
<div class="card"><strong>Total hashes</strong><br>$($Summary.TotalHashes)</div>
<div class="card"><strong>Total exact extra copies</strong><br>$($Summary.TotalExtraCopies)</div>
<div class="card"><strong>Duplicate hash groups</strong><br>$($Summary.DuplicateGroups)</div>
<div class="card"><strong>Confirmed duplicates</strong><br>$($Summary.ConfirmedDuplicates)</div>
<div class="card"><strong>Unique rescued from review</strong><br>$($Summary.PromotedUnique)</div>
<div class="card"><strong>Manual review</strong><br>$($Summary.ManualReview)</div>
<div class="card"><strong>Cloud placeholders skipped</strong><br>$($Stats.CloudPlaceholdersSkipped)</div>
<div class="card"><strong>Missing real</strong><br>$($Stats.MissingReal)</div>
</div>
<h2>Planned/applied actions</h2>
<table><thead><tr><th>Action</th><th>Hash</th><th>Source</th><th>Target</th></tr></thead><tbody>
$actionRows
</tbody></table>
<h2>Manual review</h2>
<table><thead><tr><th>Hash</th><th>Path</th><th>Reason</th></tr></thead><tbody>
$manualRows
</tbody></table>
</body>
</html>
"@
    Set-Content -LiteralPath $reportPath -Value $html -Encoding UTF8
    Write-Log -Message "DedupeCleanup report written: $reportPath" -Phase 'DedupeCleanup'
    return $reportPath
}

function Invoke-DedupeCleanup {
    Write-Log -Message ("DedupeCleanup started. Mode: {0}" -f $(if ($Apply) { 'APPLY' } else { 'DRY RUN' })) -Phase 'DedupeCleanup'
    Load-ProcessedDatabase
    Write-Log -Message "Global duplicate index loaded: $($script:ProcessedByHash.Count) hashes from ProcessedFiles.json" -Phase 'DedupeCleanup'

    if ($Apply -and -not (Test-Path -LiteralPath $ConfirmedDuplicatesQuarantineRoot)) {
        New-Item -ItemType Directory -Path $ConfirmedDuplicatesQuarantineRoot -Force | Out-Null
    }

    $files = @(Get-DedupeCleanupFiles -RootPath $SourcePath)
    $Stats.FilesFound = $files.Count
    if ($files.Count -eq 0) {
        Write-Log -Message "DedupeCleanup found no files to inspect." -Phase 'DedupeCleanup' -Status 'Completed'
        Close-Logging
        exit 0
    }

    Write-Log -Message "DedupeCleanup hashing $($files.Count) files with up to $script:WorkerCount workers." -Phase 'Hash queue'
    $hashResults = Get-Sha256Batch -Files $files
    $entriesByHash = @{}
    $analyzed = 0
    $lastLog = Get-Date
    foreach ($file in $files) {
        $analyzed++
        $Stats.FilesAnalyzed = $analyzed
        $hashResult = if ($hashResults.ContainsKey($file.FullName)) { $hashResults[$file.FullName] } else { $null }
        if ($null -eq $hashResult -or [string]::IsNullOrWhiteSpace($hashResult.Hash)) {
            $Stats.Inaccessible++
            $message = if ($hashResult -and $hashResult.Error) { $hashResult.Error } else { 'No hash result returned.' }
            Write-Log -Message "DedupeCleanup hash/read failed for $($file.FullName): $message" -Phase 'DedupeCleanup'
            continue
        }

        $hashKey = $hashResult.Hash.ToUpperInvariant()
        $area = Get-DedupeArea -Path $file.FullName
        $entry = [pscustomobject]@{
            File = $file
            Path = $file.FullName
            Hash = $hashKey
            Size = [int64]$file.Length
            Extension = $file.Extension.ToLowerInvariant()
            IsRaw = ($RawExtensions -contains $file.Extension.ToLowerInvariant())
            Area = $area
            IsReviewArea = ($area -eq 'DuplicatesReview' -or $area -eq 'NeedsReview')
        }
        if (-not $entriesByHash.ContainsKey($hashKey)) {
            $entriesByHash[$hashKey] = New-Object System.Collections.Generic.List[object]
        }
        $entriesByHash[$hashKey].Add($entry)

        if ($analyzed -eq 1 -or $analyzed % 1000 -eq 0 -or ((Get-Date) - $lastLog).TotalSeconds -ge 10) {
            $lastLog = Get-Date
            Write-Log -Message "DedupeCleanup hashed $analyzed/$($files.Count) files." -Phase 'DedupeCleanup'
        }
    }

    $actions = New-Object System.Collections.Generic.List[object]
    $manualReview = New-Object System.Collections.Generic.List[object]
    $duplicateGroups = 0
    $totalExtraCopies = 0
    $confirmedDuplicates = 0
    $promotedUnique = 0

    foreach ($hash in $entriesByHash.Keys) {
        $group = @($entriesByHash[$hash].ToArray())
        if ($group.Count -le 1) {
            $single = $group[0]
            if ($single.Area -eq 'DuplicatesReview') {
                $promotion = Resolve-DedupePromotionPath -Entry $single
                $actions.Add([pscustomobject]@{ Action = 'Promote unique from duplicates review'; Hash = $hash; Source = $single.Path; Target = $promotion.Path })
                $promotedUnique++
                if ($Apply) {
                    Move-DedupeEntry -Entry $single -TargetPath $promotion.Path -Action 'Promoted unique from duplicates review'
                    $Stats.FilesMoved++
                }
            }
            continue
        }

        $duplicateGroups++
        $totalExtraCopies += ($group.Count - 1)
        $canonical = Get-DedupeCanonicalEntry -Entries $group
        Write-Log -Message "Dedupe group hash=$hash canonical=$($canonical.Path) copies=$($group.Count)" -Phase 'DedupeCleanup'

        foreach ($entry in $group) {
            if ($entry.Path.Equals($canonical.Path, [StringComparison]::OrdinalIgnoreCase)) {
                continue
            }

            if ($entry.IsRaw) {
                $manualReview.Add([pscustomobject]@{ Hash = $hash; Path = $entry.Path; Reason = 'RAW/DNG exact duplicate not moved automatically' })
                continue
            }

            $quarantine = Resolve-DedupeQuarantinePath -Entry $entry
            $actions.Add([pscustomobject]@{ Action = 'Move confirmed exact duplicate to quarantine'; Hash = $hash; Source = $entry.Path; Target = $quarantine.Path })
            $confirmedDuplicates++
            if ($Apply) {
                Move-DedupeEntry -Entry $entry -TargetPath $quarantine.Path -Action 'Moved confirmed exact duplicate to quarantine'
                Add-ConfirmedDuplicateQuarantineEntry -QuarantinePath $quarantine.Path -CanonicalPath $canonical.Path -Hash $hash
                $Stats.FilesMoved++
            }
        }

        if ($canonical.Area -eq 'DuplicatesReview') {
            $hasOutsideReview = @($group | Where-Object { $_.Area -ne 'DuplicatesReview' }).Count -gt 0
            if (-not $hasOutsideReview) {
                $promotion = Resolve-DedupePromotionPath -Entry $canonical
                $actions.Add([pscustomobject]@{ Action = 'Promote canonical unique from duplicates review'; Hash = $hash; Source = $canonical.Path; Target = $promotion.Path })
                $promotedUnique++
                if ($Apply) {
                    Move-DedupeEntry -Entry $canonical -TargetPath $promotion.Path -Action 'Promoted canonical unique from duplicates review'
                    Update-ConfirmedDuplicateQuarantineCanonicalPath -Hash $hash -OldCanonicalPath $canonical.Path -NewCanonicalPath $promotion.Path
                    $Stats.FilesMoved++
                }
            }
        }
    }

    if ($Apply) {
        Save-ConfirmedDuplicateQuarantineRunManifest -Successful ($Stats.Errors -eq 0) | Out-Null
    }

    $summary = @{
        TotalFiles = $files.Count
        TotalHashes = $entriesByHash.Count
        TotalExtraCopies = $totalExtraCopies
        DuplicateGroups = $duplicateGroups
        ConfirmedDuplicates = $confirmedDuplicates
        PromotedUnique = $promotedUnique
        ManualReview = $manualReview.Count
    }
    $reportPath = Write-DedupeCleanupReport -Summary $summary -Actions @($actions.ToArray()) -ManualReview @($manualReview.ToArray())

    Write-Log -Message ("DedupeCleanup summary: files={0}; hashes={1}; extraCopies={2}; duplicateGroups={3}; confirmedDuplicates={4}; promotedUnique={5}; manualReview={6}; cloudPlaceholdersSkipped={7}; missingReal={8}; report={9}" -f $summary.TotalFiles, $summary.TotalHashes, $summary.TotalExtraCopies, $summary.DuplicateGroups, $summary.ConfirmedDuplicates, $summary.PromotedUnique, $summary.ManualReview, $Stats.CloudPlaceholdersSkipped, $Stats.MissingReal, $reportPath) -Phase 'Complete' -Status 'Completed'
    Close-Logging
    exit 0
}

function Update-ProcessedPathForKnownFileMove {
    param(
        [string]$OldPath,
        [string]$NewPath,
        [string]$Status = 'Known file path updated',
        [string]$Phase = 'ProcessedFiles'
    )

    if (-not $Apply) {
        return 0
    }

    if (-not $script:ProcessedRecords -or $script:ProcessedRecords.Count -eq 0) {
        return 0
    }

    $oldFull = (Resolve-FullPath $OldPath).TrimEnd('\')
    $updated = 0
    foreach ($record in @($script:ProcessedRecords.ToArray())) {
        if ($record.PSObject.Properties.Name -notcontains 'newRelativePath' -or [string]::IsNullOrWhiteSpace([string]$record.newRelativePath)) {
            continue
        }

        $recordFull = (Resolve-FullPath (Join-Path $DestinationBase (([string]$record.newRelativePath).Replace('/', '\').TrimStart('\')))).TrimEnd('\')
        if (-not $recordFull.Equals($oldFull, [StringComparison]::OrdinalIgnoreCase)) {
            continue
        }

        $record.newRelativePath = ConvertTo-RelativePath -Path $NewPath -BasePath $DestinationBase
        if ($record.PSObject.Properties.Name -contains 'status') {
            $record.status = $Status
        }
        if ($record.PSObject.Properties.Name -contains 'date') {
            $record.date = (Get-Date).ToString('o')
        }
        $updated++
    }

    if ($updated -gt 0) {
        $script:ProcessedDirtyCount += $updated
        Write-Log -Message "ProcessedFiles path updated for known file move: $OldPath -> $NewPath. Records updated: $updated" -Phase $Phase
    }

    return $updated
}

function Write-QuarterlyNormalizeReport {
    param(
        [object[]]$Actions,
        [object[]]$SkippedItems,
        [int]$ScannedFiles,
        [int]$SkippedUncertain,
        [int]$Errors
    )

    $reportPath = Join-Path $LogRoot ("NormalizeExistingFoldersReport-{0}.html" -f $script:RunId)
    $rows = New-Object System.Text.StringBuilder
    foreach ($action in @($Actions | Select-Object -First 10000)) {
        $rowHtml = '<tr><td>' + [string](ConvertTo-HtmlText $action.Action) +
            '</td><td>' + [string](ConvertTo-HtmlText $action.Source) +
            '</td><td>' + [string](ConvertTo-HtmlText $action.Target) +
            '</td><td>' + [string](ConvertTo-HtmlText $action.Reason) +
            '</td></tr>'
        [void]$rows.AppendLine($rowHtml)
    }

    $skippedRows = New-Object System.Text.StringBuilder
    foreach ($item in @($SkippedItems | Select-Object -First 5000)) {
        $rowHtml = '<tr><td>' + [string](ConvertTo-HtmlText $item.Path) +
            '</td><td>' + [string](ConvertTo-HtmlText $item.Reason) +
            '</td><td>' + [string](ConvertTo-HtmlText $item.Detail) +
            '</td></tr>'
        [void]$skippedRows.AppendLine($rowHtml)
    }

    $modeText = if ($Apply) { 'APPLY' } else { 'DRY RUN' }
    $html = @"
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>NormalizeExistingFoldersReport</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#1f2933}
table{border-collapse:collapse;width:100%;margin:16px 0}
th,td{border:1px solid #d7dde5;padding:6px 8px;text-align:left;font-size:13px}
th{background:#eef2f7}
.summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px}
.card{border:1px solid #d7dde5;padding:10px;background:#f8fafc}
</style>
</head>
<body>
<h1>NormalizeExistingFoldersReport</h1>
<p>Mode: $modeText</p>
<p>Organization profile: QuarterlyFolders</p>
<div class="summary">
<div class="card"><strong>Files scanned</strong><br>$ScannedFiles</div>
<div class="card"><strong>Files to move/rename</strong><br>$($Actions.Count)</div>
<div class="card"><strong>Cloud placeholders skipped</strong><br>$($Stats.CloudPlaceholdersSkipped)</div>
<div class="card"><strong>Missing real</strong><br>$($Stats.MissingReal)</div>
<div class="card"><strong>Skipped uncertain names</strong><br>$SkippedUncertain</div>
<div class="card"><strong>JSON paths updated</strong><br>$($Stats.JsonPathsUpdated)</div>
<div class="card"><strong>Zombie branches removed</strong><br>$($Stats.ZombieNormalizeFoldersRemoved)</div>
<div class="card"><strong>Errors</strong><br>$Errors</div>
</div>
<table><thead><tr><th>Action</th><th>Source</th><th>Target</th><th>Reason</th></tr></thead><tbody>
$rows
</tbody></table>
<h2>Skipped / uncertain</h2>
<table><thead><tr><th>Path</th><th>Reason</th><th>Detail</th></tr></thead><tbody>
$skippedRows
</tbody></table>
</body>
</html>
"@
    Set-Content -LiteralPath $reportPath -Value $html -Encoding UTF8
    Write-Log -Message "NormalizeExistingFolders report written: $reportPath" -Phase 'NormalizeExistingFolders'
    return $reportPath
}

function Get-QuarterStartMonthFromFolderName {
    param([string]$FolderName)

    if ([string]::IsNullOrWhiteSpace($FolderName)) { return $null }
    $trimmed = $FolderName.Trim()
    $monthMatch = [regex]::Match($trimmed, '^(?<m>0?[1-9]|1[0-2])\s*-')
    if ($monthMatch.Success) {
        return [int]$monthMatch.Groups['m'].Value
    }

    foreach ($languageKey in $QuarterFolderNamesByLanguage.Keys) {
        foreach ($quarter in $QuarterFolderNamesByLanguage[$languageKey].Keys) {
            if ($trimmed.Equals([string]$QuarterFolderNamesByLanguage[$languageKey][$quarter], [StringComparison]::OrdinalIgnoreCase)) {
                return (([int]$quarter - 1) * 3) + 1
            }
        }
    }

    return $null
}

function Get-NormalizeQuarterlyDateFromPath {
    param(
        [System.IO.FileInfo]$File,
        [string]$OrganizedRoot
    )

    $relative = ConvertTo-RelativePath -Path $File.FullName -BasePath $OrganizedRoot
    $parts = @($relative -split '[\\/]' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    for ($i = 0; $i -lt ($parts.Count - 1); $i++) {
        if ($parts[$i] -notmatch '^(19|20)\d{2}$') { continue }
        if ($i + 1 -ge $parts.Count) { continue }

        $month = Get-QuarterStartMonthFromFolderName -FolderName $parts[$i + 1]
        if ($null -eq $month) { continue }

        $year = [int]$parts[$i]
        $fileNameDate = Get-DateFromFileName -FileName $File.Name
        if ($fileNameDate) {
            $pathQuarter = [int][math]::Ceiling($month / 3.0)
            $fileQuarter = [int][math]::Ceiling($fileNameDate.Month / 3.0)
            if ($fileNameDate.Year -eq $year -and $fileQuarter -eq $pathQuarter) {
                return [pscustomobject]@{ Date = $fileNameDate; Confidence = 99; Source = 'Current path + file name' }
            }

            return [pscustomobject]@{
                Date = $fileNameDate
                Confidence = 40
                Source = 'Current path conflicts with file name'
                Detail = ('Path={0}/{1:00}; FileName={2:yyyy-MM-dd}' -f $year, $month, $fileNameDate)
            }
        }

        return [pscustomobject]@{ Date = (Get-Date -Year $year -Month $month -Day 1 -Hour 12); Confidence = 98; Source = 'Current organized path' }
    }

    return $null
}

function Get-NormalizeQuarterlyDateFromProcessedFiles {
    param([System.IO.FileInfo]$File)

    if (-not $script:ProcessedRecords -or $script:ProcessedRecords.Count -eq 0) {
        return $null
    }

    $currentFull = (Resolve-FullPath $File.FullName).TrimEnd('\')
    foreach ($record in @($script:ProcessedRecords.ToArray())) {
        foreach ($propertyName in @('newRelativePath', 'originalRelativePath')) {
            if ($record.PSObject.Properties.Name -notcontains $propertyName -or [string]::IsNullOrWhiteSpace([string]$record.$propertyName)) {
                continue
            }

            $basePath = if ($propertyName -eq 'newRelativePath') { $DestinationBase } else { $SourcePath }
            $recordFull = (Resolve-FullPath (Join-Path $basePath (([string]$record.$propertyName).Replace('/', '\').TrimStart('\')))).TrimEnd('\')
            if (-not $recordFull.Equals($currentFull, [StringComparison]::OrdinalIgnoreCase)) {
                continue
            }

            $dateFromRecordPath = Get-DateFromFolderName -Path $recordFull
            if ($dateFromRecordPath) {
                return [pscustomobject]@{ Date = $dateFromRecordPath; Confidence = 95; Source = 'ProcessedFiles path' }
            }
        }
    }

    return $null
}

function Get-NormalizeQuarterlyDateInfo {
    param(
        [System.IO.FileInfo]$File,
        [string]$OrganizedRoot
    )

    $fromPath = Get-NormalizeQuarterlyDateFromPath -File $File -OrganizedRoot $OrganizedRoot
    if ($fromPath) { return $fromPath }

    $fromProcessed = Get-NormalizeQuarterlyDateFromProcessedFiles -File $File
    if ($fromProcessed) { return $fromProcessed }

    $fromName = Get-DateFromFileName -FileName $File.Name
    if ($fromName) {
        return [pscustomobject]@{ Date = $fromName; Confidence = 96; Source = 'File name' }
    }

    if ($ExifToolAvailable) {
        Write-DiagnosticLog "Normalize falling back to EXIF for date only: $($File.FullName)"
        $metadata = Get-ExifMetadata -Path $File.FullName -ExifToolAvailable $true -TimeoutSeconds 10
        $item = [pscustomobject]@{
            File = $File
            Metadata = $metadata
            DateInfo = $null
        }
        $dateInfo = Get-PrimaryDate -Item $item -IsVideo ($VideoExtensions -contains $File.Extension.ToLowerInvariant())
        Write-DateInfoDiagnostic -File $File -DateInfo $dateInfo -Context 'NormalizeExistingFolders EXIF fallback'
        if ($dateInfo -and $dateInfo.Source -ne 'WindowsLastWriteTime') {
            return [pscustomobject]@{ Date = $dateInfo.Date; Confidence = $dateInfo.Confidence; Source = "EXIF fallback: $($dateInfo.Source)" }
        }
    }

    return $null
}

function Get-NormalizeQuarterlyFileName {
    param([string]$FileName)

    $clean = Remove-RedundantCaptureDatePrefixFromFileName -FileName $FileName
    if ([string]::IsNullOrWhiteSpace($clean)) {
        return $FileName
    }

    return $clean
}

function Invoke-NormalizeExistingFoldersQuarterly {
    param([string[]]$OrganizedRoots)

    Write-Log -Message "NormalizeExistingFolders using organization profile: QuarterlyFolders" -Phase 'NormalizeExistingFolders'
    $actions = New-Object System.Collections.Generic.List[object]
    $skippedItems = New-Object System.Collections.Generic.List[object]
    $scannedFiles = 0
    $skippedUncertain = 0
    $errors = 0
    $rootFileSets = New-Object System.Collections.Generic.List[object]
    $totalFilesToInspect = 0
    $actualNormalizeMoves = 0

    foreach ($root in @($OrganizedRoots)) {
        Write-Log -Message "NormalizeExistingFolders scanning organized root: $root" -Phase 'NormalizeExistingFolders'
        $files = @(Get-MediaFilesWithoutProtectedTraversal -RootPath $root -Phase 'NormalizeExistingFolders' |
            Where-Object { -not (Test-IsProtectedInternalFolderSegment -Path $_.FullName -RootPath $root) })
        $rootFileSets.Add([pscustomobject]@{ Root = $root; Files = $files })
        $totalFilesToInspect += $files.Count
    }

    $normalizeProgressMessage = if ($Apply) {
        'Scanning structure first. ETA is based on completed real moves after warmup, not on fast inspection.'
    }
    else {
        'DryRun inspection only. ETA is based on inspected items.'
    }
    Start-OperationalProgress -Name 'NormalizeExistingFolders' -Total $totalFilesToInspect -Phase 'NormalizeExistingFolders' -Message $normalizeProgressMessage -WarmupItems 100 -WarmupMinutes 3
    if ($Apply) {
        Write-OperationalStage -Name 'NormalizeExistingFolders' -Stage 'Scanning structure and deciding which files need real movement. ETA will stabilize after real moves begin.' -Phase 'NormalizeExistingFolders'
    }

    foreach ($fileSet in @($rootFileSets.ToArray())) {
        $root = [string]$fileSet.Root
        $files = @($fileSet.Files)
        foreach ($file in $files) {
            $scannedFiles++
            if ($Apply) {
                Update-OperationalProgress -Current $actualNormalizeMoves -Total $totalFilesToInspect -Phase 'NormalizeExistingFolders' -Stage 'Scanning structure; real-move ETA pending' -EveryItems 500 -EveryMinutes 5
            }
            else {
                Update-OperationalProgress -Current $scannedFiles -Total $totalFilesToInspect -Phase 'NormalizeExistingFolders' -Stage 'Scanning structure / DryRun planning' -EveryItems 500 -EveryMinutes 5
            }
            try {
                $availability = Detect-StorageAvailability -Item $file
                if ($availability.State -eq 'CloudPlaceholder') {
                    Register-CloudPlaceholderSkipped -Path $file.FullName -Phase 'NormalizeExistingFolders' -Availability $availability
                    $skippedItems.Add([pscustomobject]@{
                        Path = $file.FullName
                        Reason = 'CloudPlaceholder'
                        Detail = ('Provider={0}; Reason={1}' -f $availability.ProviderHint, $availability.Reason)
                    })
                    continue
                }
                if ($availability.State -eq 'MissingReal') {
                    $Stats.MissingReal++
                    $skippedItems.Add([pscustomobject]@{
                        Path = $file.FullName
                        Reason = 'MissingReal'
                        Detail = [string]$availability.Reason
                    })
                    continue
                }

                $Stats.LocalFilesDetected++
                $dateInfo = Get-NormalizeQuarterlyDateInfo -File $file -OrganizedRoot $root
                Write-DateInfoDiagnostic -File $file -DateInfo $dateInfo -Context 'NormalizeExistingFolders'
                if ($null -eq $dateInfo -or $dateInfo.Confidence -lt 90) {
                    $skippedUncertain++
                    $Stats.SkippedUncertainNames++
                    $reason = if ($dateInfo -and $dateInfo.Source) { [string]$dateInfo.Source } else { 'No reliable date source' }
                    $detail = if ($dateInfo -and $dateInfo.PSObject.Properties.Name -contains 'Detail') { [string]$dateInfo.Detail } else { '' }
                    $skippedItems.Add([pscustomobject]@{
                        Path = $file.FullName
                        Reason = $reason
                        Detail = $detail
                    })
                    if ($reason -eq 'Current path conflicts with file name' -or $Diagnostic) {
                        Write-Log -Message "Normalize skipped uncertain/conflicting date: $($file.FullName). Reason=$reason. $detail" -Phase 'NormalizeExistingFolders'
                    }
                    continue
                }

                $targetDirectory = Get-QuarterlyDestinationDirectory -Date $dateInfo.Date -OrganizedRoot $root
                $targetName = Get-NormalizeQuarterlyFileName -FileName $file.Name
                $targetPath = Join-Path $targetDirectory $targetName
                if ((Resolve-FullPath $file.FullName).Equals((Resolve-FullPath $targetPath), [StringComparison]::OrdinalIgnoreCase)) {
                    continue
                }

                $actions.Add([pscustomobject]@{
                    Action = 'Move to quarterly folder'
                    Source = $file.FullName
                    Target = $targetPath
                    Reason = "QuarterlyFolders - $($dateInfo.Source)"
                })
                if ($Diagnostic -or $actions.Count -le 200) {
                    Write-Log -Message "Normalize quarterly plan: $($file.FullName) -> $targetPath" -Phase 'NormalizeExistingFolders'
                }

                if ($Apply) {
                    if (-not (Test-Path -LiteralPath $targetDirectory)) {
                        New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
                    }

                    $sourceHash = Get-Sha256 -Path $file.FullName
                    $resolved = Resolve-UniquePath -Directory $targetDirectory -FileName $targetName -SourceHash $sourceHash
                    $finalTarget = $resolved.Path
                    $sourceParent = [System.IO.Path]::GetDirectoryName($file.FullName)
                    if ($resolved.SkipExistingIdentical) {
                        Remove-Item -LiteralPath $file.FullName -Force
                        $actualNormalizeMoves++
                    }
                    else {
                        Move-Item -LiteralPath $file.FullName -Destination $finalTarget -ErrorAction Stop
                        $Stats.FilesMoved++
                        $actualNormalizeMoves++
                    }
                    $estimatedMoveTotal = if ($scannedFiles -gt 0) {
                        [int][math]::Ceiling(($actions.Count / [double]$scannedFiles) * $totalFilesToInspect)
                    }
                    else {
                        $totalFilesToInspect
                    }
                    $estimatedMoveTotal = [math]::Max($actions.Count, [math]::Min($estimatedMoveTotal, $totalFilesToInspect))
                    Update-OperationalProgress -Current $actualNormalizeMoves -Total $estimatedMoveTotal -Phase 'NormalizeExistingFolders' -Stage 'Moving/renaming files. Index synchronization is deferred to post-operation Reconcile.' -EveryItems 250 -EveryMinutes 5
                    Remove-EmptyFolderChain -StartPath $sourceParent -StopRoot $root -Phase 'NormalizeExistingFolders'
                }
            }
            catch {
                $errors++
                $Stats.Errors++
                Write-Log -Message "NormalizeExistingFolders quarterly failed for $($file.FullName): $($_.Exception.Message)" -Phase 'NormalizeExistingFolders'
            }
        }
    }
    Complete-OperationalProgress -Phase 'NormalizeExistingFolders' -Message 'Normalize file movement phase completed. Entering cleanup and post-operation validation.'

    Remove-EmptySourceFolders
    $Stats.FoldersReduced = $Stats.EmptyFoldersRemoved
    $reportPath = Write-QuarterlyNormalizeReport -Actions @($actions.ToArray()) -SkippedItems @($skippedItems.ToArray()) -ScannedFiles $scannedFiles -SkippedUncertain $skippedUncertain -Errors $errors
    if ($Apply) {
        Write-Log -Message "Movements completed. Entering Post-operation index validation. Normalize intentionally defers ProcessedFiles path synchronization to this Reconcile phase." -Phase 'JSON reconciliation'
        Invoke-PostMutationProcessedDatabaseValidation -Roots $OrganizedRoots -Reason 'NormalizeExistingFolders QuarterlyFolders post-Apply'
        Save-ProcessedDatabase
    }
    Write-Log -Message ("NormalizeExistingFolders summary: profile=QuarterlyFolders; scannedFiles={0}; plannedMoves={1}; cloudPlaceholdersSkipped={2}; missingReal={3}; foldersReduced={4}; zombieBranchesRemoved={5}; junkOnlyFoldersRemoved={6}; junkOnlySmallMarkerFoldersRemoved={7}; jsonPathsUpdated={8}; skippedUncertainNames={9}; errors={10}; report={11}" -f $scannedFiles, $actions.Count, $Stats.CloudPlaceholdersSkipped, $Stats.MissingReal, $Stats.FoldersReduced, $Stats.ZombieNormalizeFoldersRemoved, $Stats.JunkOnlyFoldersRemoved, $Stats.JunkOnlySmallMarkerFoldersRemoved, $Stats.JsonPathsUpdated, $Stats.SkippedUncertainNames, $errors, $reportPath) -Phase 'Complete' -Status 'Completed'
    Close-Logging
    exit 0
}

function Invoke-NormalizeExistingFolders {
    Write-Log -Message ("NormalizeExistingFolders started. Mode: {0}" -f $(if ($Apply) { 'APPLY' } else { 'DRY RUN' })) -Phase 'NormalizeExistingFolders'
    Load-ProcessedDatabase

    $organizedRoots = @($script:DedupeOrganizedRoots | Where-Object { Test-Path -LiteralPath $_ -PathType Container } | Select-Object -Unique)
    if ($organizedRoots.Count -eq 0 -and (Test-Path -LiteralPath $OrganizedRoot -PathType Container)) {
        $organizedRoots = @($OrganizedRoot)
    }
    if ($organizedRoots.Count -eq 0) {
        Write-Log -Message "No organized folder found for NormalizeExistingFolders." -Phase 'NormalizeExistingFolders' -Status 'Completed'
        Close-Logging
        exit 0
    }

    Invoke-NormalizeExistingFoldersQuarterly -OrganizedRoots $organizedRoots
}

function Get-RecoveryLogCandidates {
    if (-not [string]::IsNullOrWhiteSpace($RecoveryLogPath)) {
        $candidate = Resolve-FullPath $RecoveryLogPath
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return @($candidate)
        }
        Write-Log -Message "Recovery log path was provided but not found: $candidate" -Phase 'Recovery'
        return @()
    }

    if (-not (Test-Path -LiteralPath $LogRoot -PathType Container)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $LogRoot -File -Filter '*.log' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 20 |
        ForEach-Object { $_.FullName })
}

function Get-RecoveryMapFromLogs {
    $map = @{}
    $logs = @(Get-RecoveryLogCandidates)
    foreach ($log in $logs) {
        Write-Log -Message "Recovery parsing log: $log" -Phase 'Recovery'
        try {
            $lines = Get-Content -LiteralPath $log -Encoding UTF8 -ErrorAction Stop
            foreach ($line in $lines) {
                $match = [regex]::Match($line, 'Removed source because identical target already exists:\s*(?<source>.+?)\s*->\s*(?<target>.+?)\.\s*Reason:')
                if (-not $match.Success) { continue }
                $source = Resolve-FullPath ([string]$match.Groups['source'].Value)
                $target = Resolve-FullPath ([string]$match.Groups['target'].Value)
                $map[$target.ToLowerInvariant()] = $source
            }
        }
        catch {
            Write-Log -Message "Recovery could not read log ${log}: $($_.Exception.Message)" -Phase 'Recovery'
        }
    }

    Write-Log -Message "Recovery log map loaded: $($map.Count) entries." -Phase 'Recovery'
    return $map
}

function Get-RecoveryFiles {
    if (-not (Test-Path -LiteralPath $DuplicatesRoot -PathType Container)) {
        return @()
    }

    $files = New-Object System.Collections.Generic.List[object]
    foreach ($file in @(Get-ChildItem -LiteralPath $DuplicatesRoot -File -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { $MediaExtensions -contains $_.Extension.ToLowerInvariant() })) {
        $availability = Detect-StorageAvailability -Item $file
        if ($availability.State -eq 'CloudPlaceholder') {
            Register-CloudPlaceholderSkipped -Path $file.FullName -Phase 'Recovery' -Availability $availability
            continue
        }
        if ($availability.State -eq 'MissingReal') {
            $Stats.MissingReal++
            continue
        }
        $Stats.LocalFilesDetected++
        $files.Add($file)
    }

    return @($files.ToArray())
}

function Test-RecoveryTargetUsable {
    param(
        [string]$TargetPath,
        [switch]$RequireOrganizedRoot
    )

    if ([string]::IsNullOrWhiteSpace($TargetPath)) {
        return $false
    }

    $resolvedTarget = Resolve-FullPath $TargetPath
    if (Test-IsChildPath -Path $resolvedTarget -ParentPath $DuplicatesRoot) { return $false }
    if (Test-IsChildPath -Path $resolvedTarget -ParentPath $ConfirmedDuplicatesQuarantineRoot) { return $false }
    if (Test-IsChildPath -Path $resolvedTarget -ParentPath $NeedsReviewRoot) { return $false }

    if ($RequireOrganizedRoot -and -not (Test-IsChildPath -Path $resolvedTarget -ParentPath $OrganizedRoot)) {
        return $false
    }

    return $true
}

function Resolve-RecoveredConflictPath {
    param([string]$TargetPath)

    $directory = [System.IO.Path]::GetDirectoryName($TargetPath)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($TargetPath)
    $ext = [System.IO.Path]::GetExtension($TargetPath)
    $candidate = Join-Path $directory ("{0}_RECUPERADO{1}" -f $base, $ext)
    $index = 1
    while (Test-Path -LiteralPath $candidate) {
        $candidate = Join-Path $directory ("{0}_RECUPERADO ({1}){2}" -f $base, $index, $ext)
        $index++
    }
    return $candidate
}

function Resolve-RecoveryDestinationFromExif {
    param(
        [System.IO.FileInfo]$File,
        [string]$Hash,
        [hashtable]$MetadataMap
    )

    $metadata = if ($MetadataMap.ContainsKey($File.FullName)) { $MetadataMap[$File.FullName] } else { New-EmptyMetadata }
    $ext = $File.Extension.ToLowerInvariant()
    $isVideo = $VideoExtensions -contains $ext
    $item = [pscustomobject]@{
        File = $File
        Extension = $ext
        IsVideo = $isVideo
        IsRaw = $RawExtensions -contains $ext
        Metadata = $metadata
        DateInfo = $null
        Sha256 = $Hash
        GlobalDuplicate = $false
        PerceptualHash = $null
        Width = if ($metadata.ImageWidth) { [int]$metadata.ImageWidth } else { 0 }
        Height = if ($metadata.ImageHeight) { [int]$metadata.ImageHeight } else { 0 }
    }
    $item.DateInfo = Get-PrimaryDate -Item $item -IsVideo $isVideo
    Write-DateInfoDiagnostic -File $File -DateInfo $item.DateInfo -Context 'Recovery EXIF rebuild'
    if ($item.DateInfo.Confidence -lt $ExifRepairConfidence) {
        return $null
    }

    return Join-Path (Get-DestinationPath -Item $item -RootPath $SourcePath -OrganizedRoot $OrganizedRoot) $File.Name
}

function Move-RecoveryFile {
    param(
        [System.IO.FileInfo]$File,
        [string]$TargetPath,
        [string]$Hash,
        [string]$Reason
    )

    $finalTarget = $TargetPath
    $action = 'Recover'
    if (Test-Path -LiteralPath $finalTarget -PathType Leaf) {
        try {
            $existingHash = Get-Sha256 -Path $finalTarget
            if ($existingHash.Equals($Hash, [StringComparison]::OrdinalIgnoreCase)) {
                $finalTarget = Join-Path $ConfirmedDuplicatesQuarantineRoot (ConvertTo-RelativePath -Path $File.FullName -BasePath $DuplicatesRoot)
                $action = 'Quarantine identical existing'
            }
            else {
                $finalTarget = Resolve-RecoveredConflictPath -TargetPath $finalTarget
                $action = 'Recover with _RECUPERADO suffix'
            }
        }
        catch {
            $finalTarget = Resolve-RecoveredConflictPath -TargetPath $finalTarget
            $action = 'Recover with _RECUPERADO suffix'
        }
    }

    if (-not $Apply) {
        return [pscustomobject]@{ Action = $action; Target = $finalTarget }
    }

    $targetDirectory = [System.IO.Path]::GetDirectoryName($finalTarget)
    if (-not (Test-Path -LiteralPath $targetDirectory)) {
        New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
    }

    Move-Item -LiteralPath $File.FullName -Destination $finalTarget -ErrorAction Stop
    if ($action -like 'Quarantine*') {
        Add-ConfirmedDuplicateQuarantineEntry -QuarantinePath $finalTarget -CanonicalPath $TargetPath -Hash $Hash
        Write-Log -Message "Recovery moved identical existing duplicate to quarantine: $($File.FullName) -> $finalTarget. Reason=$Reason" -Phase 'Recovery'
    }
    else {
        Update-ProcessedPathForKnownFileMove -OldPath $File.FullName -NewPath $finalTarget -Status "Recovered from wrong duplicate move - $Reason" -Phase 'Recovery' | Out-Null
        Write-Log -Message "Recovery moved file: $($File.FullName) -> $finalTarget. Reason=$Reason" -Phase 'Recovery'
    }

    return [pscustomobject]@{ Action = $action; Target = $finalTarget }
}

function Write-RecoveryReport {
    param(
        [hashtable]$Summary,
        [object[]]$Rows
    )

    $reportPath = Join-Path $LogRoot ("RecoverWrongDuplicateMoveReport-{0}.html" -f $script:RunId)
    $htmlRows = New-Object System.Text.StringBuilder
    foreach ($row in @($Rows | Select-Object -First 20000)) {
        [void]$htmlRows.AppendLine(("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>" -f (ConvertTo-HtmlText $row.Method), (ConvertTo-HtmlText $row.Source), (ConvertTo-HtmlText $row.Target), (ConvertTo-HtmlText $row.Action)))
    }
    $modeText = if ($Apply) { 'APPLY' } else { 'DRY RUN' }
    $html = @"
<!doctype html>
<html><head><meta charset="utf-8"><title>RecoverWrongDuplicateMoveReport</title>
<style>body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#1f2933}table{border-collapse:collapse;width:100%;margin:16px 0}th,td{border:1px solid #d7dde5;padding:6px 8px;text-align:left;font-size:13px}th{background:#eef2f7}.summary{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px}.card{border:1px solid #d7dde5;padding:10px;background:#f8fafc}</style>
</head><body>
<h1>RecoverWrongDuplicateMoveReport</h1>
<p>Mode: $modeText</p>
<div class="summary">
<div class="card"><strong>Total in duplicates</strong><br>$($Summary.Total)</div>
<div class="card"><strong>Recovered by log</strong><br>$($Summary.ByLog)</div>
<div class="card"><strong>Recovered by ProcessedFiles</strong><br>$($Summary.ByProcessed)</div>
<div class="card"><strong>Rebuilt by EXIF</strong><br>$($Summary.ByExif)</div>
<div class="card"><strong>RecoveryPathUnknown</strong><br>$($Summary.Unknown)</div>
<div class="card"><strong>Conflicts</strong><br>$($Summary.Conflicts)</div>
<div class="card"><strong>Identical existing</strong><br>$($Summary.IdenticalExisting)</div>
<div class="card"><strong>Errors</strong><br>$($Summary.Errors)</div>
</div>
<table><thead><tr><th>Method</th><th>Source</th><th>Target</th><th>Action</th></tr></thead><tbody>
$htmlRows
</tbody></table>
</body></html>
"@
    Set-Content -LiteralPath $reportPath -Value $html -Encoding UTF8
    Write-Log -Message "Recovery report written: $reportPath" -Phase 'Recovery'
    return $reportPath
}

function Invoke-RecoverFromWrongDuplicateMove {
    Write-Log -Message ("RecoverFromWrongDuplicateMove started. Mode: {0}" -f $(if ($Apply) { 'APPLY' } else { 'DRY RUN' })) -Phase 'Recovery'
    Load-ProcessedDatabase
    $logMap = Get-RecoveryMapFromLogs
    $files = @(Get-RecoveryFiles)
    Write-Log -Message "Recovery detected files in duplicates folder: $($files.Count)" -Phase 'Recovery'

    $hashResults = Get-Sha256Batch -Files $files
    $fallbackFiles = New-Object System.Collections.Generic.List[object]
    $rows = New-Object System.Collections.Generic.List[object]
    $summary = @{ Total = $files.Count; ByLog = 0; ByProcessed = 0; ByExif = 0; Unknown = 0; Conflicts = 0; IdenticalExisting = 0; Errors = 0 }

    foreach ($file in $files) {
        $hashResult = if ($hashResults.ContainsKey($file.FullName)) { $hashResults[$file.FullName] } else { $null }
        if ($null -eq $hashResult -or [string]::IsNullOrWhiteSpace($hashResult.Hash)) {
            $summary.Errors++
            $rows.Add([pscustomobject]@{ Method = 'Error'; Source = $file.FullName; Target = ''; Action = 'Hash failed' })
            continue
        }

        $hash = $hashResult.Hash.ToUpperInvariant()
        $target = $null
        $method = $null
        $logKey = (Resolve-FullPath $file.FullName).ToLowerInvariant()
        if ($logMap.ContainsKey($logKey)) {
            $target = $logMap[$logKey]
            if (Test-RecoveryTargetUsable -TargetPath $target) {
                $method = 'Log'
                $summary.ByLog++
            }
            else {
                Write-Log -Message "Recovery log target ignored because it points to a protected/review folder: $target" -Phase 'Recovery'
                $target = $null
            }
        }
        $processedHintTarget = $null
        if ([string]::IsNullOrWhiteSpace($target) -and $script:ProcessedByHash.ContainsKey($hash) -and -not [string]::IsNullOrWhiteSpace([string]$script:ProcessedByHash[$hash].newRelativePath)) {
            $candidateProcessedTarget = Join-Path $DestinationBase (([string]$script:ProcessedByHash[$hash].newRelativePath).Replace('/', '\').TrimStart('\'))
            if (Test-RecoveryTargetUsable -TargetPath $candidateProcessedTarget -RequireOrganizedRoot) {
                $processedHintTarget = $candidateProcessedTarget
            }
            else {
                Write-Log -Message "Recovery ProcessedFiles hint ignored because it is not an organized-library path: $candidateProcessedTarget" -Phase 'Recovery'
            }
        }

        if ([string]::IsNullOrWhiteSpace($target)) {
            $fallbackFiles.Add([pscustomobject]@{ File = $file; Hash = $hash; ProcessedHintTarget = $processedHintTarget })
            continue
        }

        try {
            $move = Move-RecoveryFile -File $file -TargetPath $target -Hash $hash -Reason $method
            if ($move.Action -like '*RECUPERADO*') { $summary.Conflicts++ }
            if ($move.Action -like 'Quarantine*') { $summary.IdenticalExisting++ }
            $rows.Add([pscustomobject]@{ Method = $method; Source = $file.FullName; Target = $move.Target; Action = $move.Action })
        }
        catch {
            $summary.Errors++
            $rows.Add([pscustomobject]@{ Method = $method; Source = $file.FullName; Target = $target; Action = "Error: $($_.Exception.Message)" })
        }
    }

    if ($fallbackFiles.Count -gt 0) {
        Write-Log -Message "Recovery rebuilding $($fallbackFiles.Count) paths from EXIF/date rules." -Phase 'Recovery'
        $metadataMap = Get-ExifMetadataBatch -Files @($fallbackFiles | ForEach-Object { $_.File }) -ExifToolAvailable $ExifToolAvailable
        foreach ($entry in $fallbackFiles) {
            $target = Resolve-RecoveryDestinationFromExif -File $entry.File -Hash $entry.Hash -MetadataMap $metadataMap
            $method = 'EXIF current rules'
            if (-not [string]::IsNullOrWhiteSpace($target) -and (Test-RecoveryTargetUsable -TargetPath $target -RequireOrganizedRoot)) {
                $summary.ByExif++
            }
            elseif (-not [string]::IsNullOrWhiteSpace([string]$entry.ProcessedHintTarget) -and (Test-RecoveryTargetUsable -TargetPath ([string]$entry.ProcessedHintTarget) -RequireOrganizedRoot)) {
                $target = [string]$entry.ProcessedHintTarget
                $method = 'ProcessedFiles fallback'
                $summary.ByProcessed++
                Write-Log -Message "Recovery used ProcessedFiles as fallback after current EXIF/path rebuild was unavailable: $($entry.File.FullName) -> $target" -Phase 'Recovery'
            }
            else {
                $target = Join-Path (Join-Path $NeedsReviewRoot 'RecoveryPathUnknown') $entry.File.Name
                $method = 'RecoveryPathUnknown'
                $summary.Unknown++
            }

            try {
                $move = Move-RecoveryFile -File $entry.File -TargetPath $target -Hash $entry.Hash -Reason $method
                if ($move.Action -like '*RECUPERADO*') { $summary.Conflicts++ }
                if ($move.Action -like 'Quarantine*') { $summary.IdenticalExisting++ }
                $rows.Add([pscustomobject]@{ Method = $method; Source = $entry.File.FullName; Target = $move.Target; Action = $move.Action })
            }
            catch {
                $summary.Errors++
                $rows.Add([pscustomobject]@{ Method = $method; Source = $entry.File.FullName; Target = $target; Action = "Error: $($_.Exception.Message)" })
            }
        }
    }

    if ($Apply) {
        Save-ProcessedDatabase
        Save-ConfirmedDuplicateQuarantineRunManifest -Successful ($summary.Errors -eq 0) | Out-Null
    }
    $reportPath = Write-RecoveryReport -Summary $summary -Rows @($rows.ToArray())
    Write-Log -Message ("Recovery summary: total={0}; byLog={1}; byProcessed={2}; byExif={3}; unknown={4}; conflicts={5}; identicalExisting={6}; errors={7}; report={8}" -f $summary.Total, $summary.ByLog, $summary.ByProcessed, $summary.ByExif, $summary.Unknown, $summary.Conflicts, $summary.IdenticalExisting, $summary.Errors, $reportPath) -Phase 'Complete' -Status 'Completed'
    Close-Logging
    exit 0
}

function Get-RepairOnlyFiles {
    param([string]$RootPath)
    return @(Get-MediaFilesWithoutProtectedTraversal -RootPath $RootPath -Phase 'RepairOnly' -CountLocalFiles)
}

function Invoke-MetadataAuditOrRepair {
    param([switch]$CloseAndExit)

    $phase = if ($MetadataRepair) { 'MetadataRepair' } else { 'MetadataAudit' }
    if ($MetadataAudit -and $Apply) {
        Write-Log -Message 'MetadataAudit is always DryRun. Ignoring -Apply for this mode.' -Phase $phase -Status 'Warning'
        $Apply = $false
    }
    Write-Log -Message ("{0} started. Mode={1}. Scope=OrganizedRoot. Root={2}" -f $phase, $(if ($Apply -and $MetadataRepair) { 'APPLY' } else { 'DRY RUN' }), $OrganizedRoot) -Phase $phase
    if ($MetadataRepair -and -not $Apply) {
        Write-Log -Message "MetadataRepair without -Apply runs as DryRun. No files will be changed." -Phase $phase
    }

    if ($Apply -and $MetadataRepair -and (-not $script:ProcessedRecords -or $script:ProcessedRecords.Count -eq 0)) {
        Load-ProcessedIndexLight
    }

    $repairFiles = @(Get-RepairOnlyFiles -RootPath $OrganizedRoot)
    Write-Log -Message "Metadata materialization files found: $($repairFiles.Count). Cloud placeholders skipped: $($Stats.CloudPlaceholdersSkipped). Missing real: $($Stats.MissingReal)" -Phase $phase
    $reportRows = New-Object System.Collections.Generic.List[object]
    $BatchSize = [math]::Max(1, $BatchSize)
    Start-OperationalProgress -Name $phase -Total $repairFiles.Count -Phase $phase -Message 'Auditing visible capture date materialization.'

    for ($batchStart = 0; $batchStart -lt $repairFiles.Count; $batchStart += $BatchSize) {
        $batchEnd = [math]::Min($batchStart + $BatchSize - 1, $repairFiles.Count - 1)
        $batch = @($repairFiles[$batchStart..$batchEnd])
        Update-OperationalProgress -Current $batchStart -Total $repairFiles.Count -Phase $phase -Stage 'Reading current visible metadata' -EveryItems ([math]::Max(1, $BatchSize * 2)) -EveryMinutes 5
        $metadataMap = Get-ExifMetadataBatch -Files $batch -ExifToolAvailable $ExifToolAvailable
        foreach ($file in $batch) {
            $Stats.FilesAnalyzed++
            Update-OperationalProgress -Current $Stats.FilesAnalyzed -Total $repairFiles.Count -Phase $phase -Stage 'Auditing/materializing capture dates' -EveryItems 500 -EveryMinutes 5
            $metadata = if ($metadataMap.ContainsKey($file.FullName)) { $metadataMap[$file.FullName] } else { New-EmptyMetadata }
            $ext = $file.Extension.ToLowerInvariant()
            $item = [pscustomobject]@{
                File = $file
                Extension = $ext
                IsVideo = ($VideoExtensions -contains $ext)
                IsRaw = ($RawExtensions -contains $ext)
                Metadata = $metadata
                DateInfo = $null
                Sha256 = $null
                Width = if ($metadata.ImageWidth) { [int]$metadata.ImageWidth } else { 0 }
                Height = if ($metadata.ImageHeight) { [int]$metadata.ImageHeight } else { 0 }
                EmbeddedMetadataReadStatus = [string]$metadata.ReadStatus
            }
            $item.DateInfo = Get-PrimaryDate -Item $item -IsVideo $item.IsVideo
            [void](Initialize-EmbeddedCaptureDateProbe -Item $item)
            $processedRecordBeforeRepair = Find-ProcessedRecordByPath -Path $file.FullName
            $oldHashBeforeRepair = if ($processedRecordBeforeRepair -and $processedRecordBeforeRepair.hash) { [string]$processedRecordBeforeRepair.hash } else { '' }
            if (-not [string]::IsNullOrWhiteSpace($oldHashBeforeRepair)) { $item.Sha256 = $oldHashBeforeRepair }
            $materializationResult = Invoke-CaptureDateMaterialization -Item $item -RootPath $OrganizedRoot -MetadataBackupRoot $MetadataBackupRoot -ExifToolAvailable $ExifToolAvailable -UpdateProcessedIndex -Reason $phase
            if ($materializationResult.Candidate -or $materializationResult.Status -ne 'NotNeeded') {
                $reportRows.Add([pscustomobject]@{
                    Path = $file.FullName
                    Extension = $ext
                    CaptureDate = if ($item.DateInfo -and $item.DateInfo.Date) { $item.DateInfo.Date.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
                    Source = if ($item.DateInfo) { [string]$item.DateInfo.Source } else { '' }
                    Confidence = if ($item.DateInfo) { [int]$item.DateInfo.Confidence } else { 0 }
                    EmbeddedCaptureDateState = $materializationResult.EmbeddedCaptureDateState
                    Status = $materializationResult.Status
                    EmbeddedMetadataWritten = $materializationResult.EmbeddedMetadataWritten
                    FileSystemDatesSynced = $materializationResult.FileSystemDatesSynced
                    DateKnownButMetadataNotWritten = $materializationResult.DateKnownButMetadataNotWritten
                    HashChanged = $materializationResult.HashChanged
                    OldHash = $materializationResult.OldHash
                    NewHash = $materializationResult.NewHash
                })
            }
        }
    }

    Complete-OperationalProgress -Phase $phase -Message "$phase completed."
    $reportPath = Join-Path $LogRoot ("CaptureDateMaterialization-{0}.csv" -f $script:RunId)
    $reportRows | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8
    if ($Apply -and $MetadataRepair) { Save-ProcessedDatabase }
    Write-Log -Message ("{0} completed. Candidates={1}; embeddedWritten={2}; filesystemDatesSynced={3}; dateKnownButMetadataNotWritten={4}; report={5}" -f $phase, $Stats.CaptureDateMaterializationCandidates, $Stats.CaptureMetadataWritten, $Stats.FileSystemDatesSynced, $Stats.DateKnownButMetadataNotWritten, $reportPath) -Phase $phase -Status 'Completed'
    if ($CloseAndExit) {
        if ($Apply -and $MetadataRepair) { Cleanup-MetadataBackupsOnSuccess }
        Close-Logging
        exit 0
    }
}
function Invoke-RepairOnlyExistingOrganizedLibrary {
    param([switch]$CloseAndExit)

    Write-Log -Message "RepairOnlyExistingOrganizedLibrary started. No move, no dedupe, no organization." -Phase 'RepairOnly'
    if (-not $RepairExif) {
        Write-Log -Message "RepairOnlyExistingOrganizedLibrary requires -RepairExif. Nothing to do." -Phase 'RepairOnly'
        if ($CloseAndExit) { Close-Logging; exit 0 }
        return
    }
    if (-not $Apply) {
        Write-Log -Message "RepairOnlyExistingOrganizedLibrary dry run. Candidates will be detected and logged; no files changed." -Phase 'RepairOnly'
    }

    if (-not $script:ProcessedRecords -or $script:ProcessedRecords.Count -eq 0) {
        Load-ProcessedDatabase
    }

    $repairFiles = @(Get-RepairOnlyFiles -RootPath $OrganizedRoot)
    Write-Log -Message "RepairOnly files found: $($repairFiles.Count). Cloud placeholders skipped: $($Stats.CloudPlaceholdersSkipped). Missing real: $($Stats.MissingReal)" -Phase 'RepairOnly'
    $Stats.FilesFound = $repairFiles.Count
    Start-OperationalProgress -Name 'RepairOnlyExistingOrganizedLibrary' -Total $repairFiles.Count -Phase 'RepairOnly' -Message 'Repairing EXIF in-place without moving, dedupe or organization.'
    $repairCandidates = 0
    $BatchSize = [math]::Max(1, $BatchSize)
    for ($batchStart = 0; $batchStart -lt $repairFiles.Count; $batchStart += $BatchSize) {
        $batchEnd = [math]::Min($batchStart + $BatchSize - 1, $repairFiles.Count - 1)
        $batch = @($repairFiles[$batchStart..$batchEnd])
        Update-OperationalProgress -Current $batchStart -Total $repairFiles.Count -Phase 'RepairOnly' -Stage 'Reading EXIF metadata batches' -EveryItems ([math]::Max(1, $BatchSize * 2)) -EveryMinutes 5
        $metadataMap = Get-ExifMetadataBatch -Files $batch -ExifToolAvailable $ExifToolAvailable
        foreach ($file in $batch) {
            $Stats.FilesAnalyzed++
            Update-OperationalProgress -Current $Stats.FilesAnalyzed -Total $repairFiles.Count -Phase 'RepairOnly' -Stage 'Repairing EXIF in-place' -EveryItems 500 -EveryMinutes 5
            $metadata = if ($metadataMap.ContainsKey($file.FullName)) { $metadataMap[$file.FullName] } else { New-EmptyMetadata }
            $ext = $file.Extension.ToLowerInvariant()
            $item = [pscustomobject]@{
                File = $file
                Extension = $ext
                IsVideo = ($VideoExtensions -contains $ext)
                IsRaw = ($RawExtensions -contains $ext)
                Metadata = $metadata
                DateInfo = $null
                Sha256 = $null
                Width = if ($metadata.ImageWidth) { [int]$metadata.ImageWidth } else { 0 }
                Height = if ($metadata.ImageHeight) { [int]$metadata.ImageHeight } else { 0 }
                EmbeddedMetadataReadStatus = [string]$metadata.ReadStatus
            }
            $item.DateInfo = Get-PrimaryDate -Item $item -IsVideo $item.IsVideo
            [void](Initialize-EmbeddedCaptureDateProbe -Item $item)
            Write-DateInfoDiagnostic -File $file -DateInfo $item.DateInfo -Context 'RepairOnlyExistingOrganizedLibrary'
            $processedRecordBeforeRepair = Find-ProcessedRecordByPath -Path $file.FullName
            $oldHashBeforeRepair = if ($processedRecordBeforeRepair -and $processedRecordBeforeRepair.hash) { [string]$processedRecordBeforeRepair.hash } else { '' }
            $materializationResult = Invoke-CaptureDateMaterialization -Item $item -RootPath $OrganizedRoot -MetadataBackupRoot $MetadataBackupRoot -ExifToolAvailable $ExifToolAvailable -UpdateProcessedIndex -Reason 'RepairOnlyExistingOrganizedLibrary'
            if ($materializationResult.Candidate) {
                $repairCandidates++
            }
        }
        Write-Log -Message "RepairOnly progress: $([math]::Min($batchStart + $BatchSize, $repairFiles.Count))/$($repairFiles.Count)" -Phase 'RepairOnly'
    }
    Complete-OperationalProgress -Phase 'RepairOnly' -Message 'RepairOnly processing phase completed.'
    if ($Apply) {
        Save-ProcessedDatabase
    }
    Write-Log -Message "RepairOnlyExistingOrganizedLibrary completed. EXIF candidates: $repairCandidates. EXIF repaired: $($Stats.ExifRepaired). Cloud placeholders skipped: $($Stats.CloudPlaceholdersSkipped). Missing real: $($Stats.MissingReal)" -Phase 'RepairOnly'
    if ($CloseAndExit) {
        if ($Apply) {
            Cleanup-MetadataBackupsOnSuccess
        }
        Close-Logging
        exit 0
    }
}

function Get-ImagePhash {
    param([string]$Path)

    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        $bitmap = [System.Drawing.Bitmap]::FromFile($Path)
        try {
            $small = New-Object System.Drawing.Bitmap 8, 8
            $graphics = [System.Drawing.Graphics]::FromImage($small)
            try {
                $graphics.DrawImage($bitmap, 0, 0, 8, 8)
            }
            finally {
                $graphics.Dispose()
            }

            $values = New-Object System.Collections.Generic.List[int]
            for ($y = 0; $y -lt 8; $y++) {
                for ($x = 0; $x -lt 8; $x++) {
                    $pixel = $small.GetPixel($x, $y)
                    $gray = [int](($pixel.R * 0.299) + ($pixel.G * 0.587) + ($pixel.B * 0.114))
                    $values.Add($gray)
                }
            }

            $avg = ($values | Measure-Object -Average).Average
            $bits = foreach ($value in $values) {
                if ($value -ge $avg) { '1' } else { '0' }
            }
            return -join $bits
        }
        finally {
            if ($small) { $small.Dispose() }
            $bitmap.Dispose()
        }
    }
    catch {
        return $null
    }
}

function Get-HammingDistance {
    param(
        [string]$A,
        [string]$B
    )
    if ([string]::IsNullOrEmpty($A) -or [string]::IsNullOrEmpty($B) -or $A.Length -ne $B.Length) {
        return 999
    }
    $distance = 0
    for ($i = 0; $i -lt $A.Length; $i++) {
        if ($A[$i] -ne $B[$i]) { $distance++ }
    }
    return $distance
}

function Select-BestVersion {
    param([object[]]$Items)

    return $Items |
        Sort-Object `
            @{ Expression = { [int]($_.Width) * [int]($_.Height) }; Descending = $true },
            @{ Expression = { $_.File.Length }; Descending = $true },
            @{ Expression = { $_.Metadata.ExifScore }; Descending = $true },
            @{ Expression = { $_.DateInfo.Confidence }; Descending = $true } |
        Select-Object -First 1
}

function Backup-OriginalForMetadataChange {
    param(
        [pscustomobject]$Item,
        [string]$RootPath,
        [string]$MetadataBackupRoot
    )

    $relative = $Item.File.FullName.Substring($RootPath.TrimEnd('\').Length).TrimStart('\')
    $backupPath = Join-Path $MetadataBackupRoot $relative
    $backupDir = [System.IO.Path]::GetDirectoryName($backupPath)

    if (-not (Test-Path -LiteralPath $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    }

    if (-not (Test-Path -LiteralPath $backupPath)) {
        Copy-Item -LiteralPath $Item.File.FullName -Destination $backupPath
        Update-MetadataBackupSize | Out-Null
    }
}

function Cleanup-MetadataBackupsOnSuccess {
    if (-not $Apply -or (-not $RepairExif -and [int]$Stats.CaptureMetadataWritten -le 0)) {
        return
    }

    Update-MetadataBackupSize | Out-Null
    Write-Log -Message ("Metadata backup size: {0} GB" -f $Stats.MetadataBackupSizeGB) -Phase 'Complete'

    if ($Stats.Errors -gt 0) {
        Write-Log -Message "Run completed with errors. Verifying metadata backups before cleanup." -Phase 'Complete'
        if (-not (Test-MetadataBackupsHaveVerifiedTargets)) {
            Write-Log -Message "Metadata backups preserved because verification failed: $MetadataBackupRoot" -Phase 'Complete'
            return
        }
        Write-Log -Message "Metadata backups are safe to clean despite non-critical run errors: $MetadataBackupRoot" -Phase 'Complete'
    }

    if (Test-Path -LiteralPath $MetadataBackupRoot) {
        try {
            Remove-Item -LiteralPath $MetadataBackupRoot -Recurse -Force
            Write-Log -Message "Metadata backups cleaned after successful run: $MetadataBackupRoot" -Phase 'Complete'
        }
        catch {
            $Stats.Errors++
            Write-Log -Message "Could not clean metadata backups: $($_.Exception.Message)" -Phase 'Complete'
        }
    }
}

function Get-ExifRepairWriteMode {
    param(
        [string]$Extension,
        [bool]$IsVideo
    )

    $ext = ([string]$Extension).ToLowerInvariant()
    if ($ext -in $RawExtensions) { return '' }
    if ($ext -in @('.jpg', '.jpeg')) { return 'JpegExif' }
    if ($ext -in @('.mp4', '.mov', '.m4v', '.3gp', '.heic', '.heif')) { return 'QuickTime' }
    if ($ext -eq '.png') { return 'PngXmp' }
    if ($ext -in @('.tif', '.tiff', '.webp', '.gif')) { return 'Xmp' }
    return ''
}

function Get-ExistingEmbeddedCaptureDateForRepair {
    param(
        [pscustomobject]$Item
    )

    if ($null -eq $Item -or $null -eq $Item.Metadata) { return $null }

    $fields = if ($Item.IsVideo) {
        @('MediaCreateDate', 'CreateDate', 'TrackCreateDate', 'QuickTimeCreateDate', 'DateTimeOriginal', 'XMPDateCreated', 'XMPCreateDate')
    }
    else {
        @('DateTimeOriginal', 'CreateDate', 'XMPDateCreated', 'XMPCreateDate', 'PNGCreationTime', 'MediaCreateDate')
    }

    foreach ($field in $fields) {
        if ($Item.Metadata.PSObject.Properties.Name -contains $field) {
            $date = ConvertTo-MediaDate $Item.Metadata.$field
            if ($date) { return $date }
        }
    }

    return $null
}

function Get-EmbeddedCaptureDateProbe {
    param([pscustomobject]$Item)

    $result = [ordered]@{
        State = 'NotChecked'
        ExistingDate = $null
        Detail = 'Embedded capture metadata was not checked.'
    }
    if ($null -eq $Item) {
        return [pscustomobject]$result
    }

    $readStatus = if ($Item.PSObject.Properties.Name -contains 'EmbeddedMetadataReadStatus') {
        [string]$Item.EmbeddedMetadataReadStatus
    }
    else {
        'NotChecked'
    }
    if ($readStatus -eq 'NotChecked') {
        return [pscustomobject]$result
    }
    if ($readStatus -ne 'Read') {
        $result.State = 'Unreadable'
        $result.Detail = 'Embedded capture metadata could not be read reliably.'
        return [pscustomobject]$result
    }

    $providerConflict = ($Item.DateInfo -and
        $Item.DateInfo.PSObject.Properties.Name -contains 'ProviderExifConflict' -and
        [bool]$Item.DateInfo.ProviderExifConflict)
    $filenameConflict = ($Item.DateInfo -and
        $Item.DateInfo.PSObject.Properties.Name -contains 'FilenameDateConflictsWithExif' -and
        [bool]$Item.DateInfo.FilenameDateConflictsWithExif)
    if ($providerConflict -or $filenameConflict) {
        $result.State = 'Conflict'
        $result.Detail = if ($providerConflict) { 'Provider and embedded capture dates conflict.' } else { 'Filename and embedded capture dates conflict.' }
        return [pscustomobject]$result
    }

    $existingDate = Get-ExistingEmbeddedCaptureDateForRepair -Item $Item
    if ($existingDate) {
        $result.State = 'PresentValid'
        $result.ExistingDate = $existingDate
        $result.Detail = 'A valid embedded capture date already exists.'
        return [pscustomobject]$result
    }

    $writeMode = Get-ExifRepairWriteMode -Extension ([string]$Item.Extension) -IsVideo ([bool]$Item.IsVideo)
    if ([string]::IsNullOrWhiteSpace($writeMode)) {
        $result.State = 'Unsupported'
        $result.Detail = 'The media format has no supported capture-date write mode.'
        return [pscustomobject]$result
    }

    $result.State = 'Absent'
    $result.Detail = 'Embedded metadata was read successfully and no valid capture date was found.'
    return [pscustomobject]$result
}

function Initialize-EmbeddedCaptureDateProbe {
    param([pscustomobject]$Item)

    $probe = Get-EmbeddedCaptureDateProbe -Item $Item
    if ($null -ne $Item) {
        $Item | Add-Member -NotePropertyName EmbeddedCaptureDateState -NotePropertyValue ([string]$probe.State) -Force
        $Item | Add-Member -NotePropertyName EmbeddedCaptureDateProbeDetail -NotePropertyValue ([string]$probe.Detail) -Force
    }
    return $probe
}

function Get-ExifRepairWriteArguments {
    param(
        [string]$WriteMode,
        [string]$DateText,
        [string]$Path
    )

    switch ($WriteMode) {
        'JpegExif' {
            return @(
                '-overwrite_original',
                "-DateTimeOriginal=$DateText",
                "-CreateDate=$DateText",
                "-ModifyDate=$DateText",
                $Path
            )
        }
        'QuickTime' {
            return @(
                '-overwrite_original',
                "-QuickTime:CreateDate=$DateText",
                "-QuickTime:ModifyDate=$DateText",
                "-TrackCreateDate=$DateText",
                "-TrackModifyDate=$DateText",
                "-MediaCreateDate=$DateText",
                "-MediaModifyDate=$DateText",
                $Path
            )
        }
        'PngXmp' {
            return @(
                '-overwrite_original',
                "-PNG:CreationTime=$DateText",
                "-XMP:DateCreated=$DateText",
                "-XMP:CreateDate=$DateText",
                "-XMP:ModifyDate=$DateText",
                $Path
            )
        }
        'Xmp' {
            return @(
                '-overwrite_original',
                "-XMP:DateCreated=$DateText",
                "-XMP:CreateDate=$DateText",
                "-XMP:ModifyDate=$DateText",
                $Path
            )
        }
    }

    return @()
}

function Get-ReliableFilenameExifRepairDecision {
    param([pscustomobject]$Item)

    $decision = [ordered]@{
        CanRepair = $false
        Kind = $null
        SyntheticTime = $false
        Conflict = $false
        ConflictSource = $null
        WriteMode = ''
    }

    if ($null -eq $Item -or $null -eq $Item.DateInfo) {
        return [pscustomobject]$decision
    }

    $extension = ([string]$Item.Extension).ToLowerInvariant()
    $writeMode = Get-ExifRepairWriteMode -Extension $extension -IsVideo ([bool]$Item.IsVideo)
    if ([string]::IsNullOrWhiteSpace($writeMode)) {
        return [pscustomobject]$decision
    }

    if ([string]$Item.DateInfo.Source -ne 'FileName') {
        return [pscustomobject]$decision
    }

    if ($Item.DateInfo.PSObject.Properties.Name -notcontains 'FilenameDateKind') {
        return [pscustomobject]$decision
    }

    $kind = [string]$Item.DateInfo.FilenameDateKind
    if ($kind -notin @('ReliableDateTime', 'ReliableDateOnly')) {
        return [pscustomobject]$decision
    }

    foreach ($metadataDateName in @('CreateDate', 'MediaCreateDate')) {
        $metadataDate = ConvertTo-MediaDate $Item.Metadata.$metadataDateName
        if ($metadataDate -and $metadataDate.Date -ne $Item.DateInfo.Date.Date) {
            $decision.Conflict = $true
            $decision.ConflictSource = $metadataDateName
            return [pscustomobject]$decision
        }
    }

    $decision.CanRepair = $true
    $decision.Kind = $kind
    $decision.SyntheticTime = ($kind -eq 'ReliableDateOnly')
    $decision.WriteMode = $writeMode
    return [pscustomobject]$decision
}

function Test-ProviderDateInfoCanMaterialize {
    param([pscustomobject]$DateInfo)

    if ($null -eq $DateInfo -or $null -eq $DateInfo.Date) { return $false }
    if ($DateInfo.Confidence -lt 97) { return $false }
    if ($DateInfo.PSObject.Properties.Name -contains 'ProviderExifConflict' -and [bool]$DateInfo.ProviderExifConflict) { return $false }
    return (@('GoogleTakeout', 'ApplePhotos', 'XmpSidecarLibrary') -contains [string]$DateInfo.Source)
}

function Test-DateInfoCanDriveVisibleDateMaterialization {
    param([pscustomobject]$DateInfo)

    if ($null -eq $DateInfo -or $null -eq $DateInfo.Date) { return $false }
    if ($DateInfo.Confidence -lt $ExifRepairConfidence) { return $false }
    if ([string]$DateInfo.Source -eq 'WindowsLastWriteTime') { return $false }
    if ($DateInfo.PSObject.Properties.Name -contains 'FilenameDateConflictsWithExif' -and [bool]$DateInfo.FilenameDateConflictsWithExif) { return $false }
    if ($DateInfo.PSObject.Properties.Name -contains 'ProviderExifConflict' -and [bool]$DateInfo.ProviderExifConflict) { return $false }
    return $true
}

function Test-FileSystemDatesNeedSync {
    param(
        [System.IO.FileInfo]$File,
        [datetime]$CaptureDate
    )

    if ($null -eq $File -or $null -eq $CaptureDate) { return $false }
    if ($File.CreationTime.Date -ne $CaptureDate.Date) { return $true }
    if ($File.LastWriteTime.Date -ne $CaptureDate.Date) { return $true }
    return $false
}

function Sync-FileSystemDatesFromValidatedCaptureDate {
    param(
        [pscustomobject]$Item,
        [string]$Reason = 'CaptureDateMaterialization'
    )

    $result = [pscustomobject]@{
        Candidate = $false
        Synced = $false
        Reason = ''
    }

    if ($null -eq $Item -or $null -eq $Item.File -or -not (Test-DateInfoCanDriveVisibleDateMaterialization -DateInfo $Item.DateInfo)) {
        $result.Reason = 'NoReliableCaptureDate'
        return $result
    }

    $captureDate = [datetime]$Item.DateInfo.Date
    if (-not (Test-FileSystemDatesNeedSync -File $Item.File -CaptureDate $captureDate)) {
        $result.Reason = 'FileSystemDatesAlreadyAligned'
        return $result
    }

    $result.Candidate = $true
    if (-not $Apply) {
        Write-Log -Message ("Filesystem date sync candidate: Path={0}; CaptureDate={1:yyyy-MM-dd HH:mm:ss}; CreationTime={2:yyyy-MM-dd HH:mm:ss}; LastWriteTime={3:yyyy-MM-dd HH:mm:ss}; Reason={4}" -f $Item.File.FullName, $captureDate, $Item.File.CreationTime, $Item.File.LastWriteTime, $Reason) -Phase 'CaptureDateMaterialization'
        return $result
    }

    try {
        $Item.File.CreationTime = $captureDate
        $Item.File.LastWriteTime = $captureDate
        $Stats.FileSystemDatesSynced++
        $result.Synced = $true
        Write-Log -Message ("Filesystem dates synced from validated capture date: Path={0}; CaptureDate={1:yyyy-MM-dd HH:mm:ss}; Source={2}; Confidence={3}; Reason={4}" -f $Item.File.FullName, $captureDate, $Item.DateInfo.Source, $Item.DateInfo.Confidence, $Reason) -Phase 'CaptureDateMaterialization'
    }
    catch {
        $Stats.Errors++
        $result.Reason = $_.Exception.Message
        Write-Log -Message "Filesystem date sync failed: Path=$($Item.File.FullName); Error=$($_.Exception.Message)" -Phase 'CaptureDateMaterialization'
    }

    return $result
}

function Set-CaptureDateMaterializationProperties {
    param(
        [pscustomobject]$Item,
        [string]$Status,
        [bool]$EmbeddedWritten,
        [bool]$FileSystemSynced,
        [bool]$DateKnownButMetadataNotWritten
    )

    $Item | Add-Member -NotePropertyName CaptureDateMaterializationStatus -NotePropertyValue $Status -Force
    $Item | Add-Member -NotePropertyName EmbeddedCaptureMetadataWritten -NotePropertyValue $EmbeddedWritten -Force
    $Item | Add-Member -NotePropertyName FileSystemDatesSynced -NotePropertyValue $FileSystemSynced -Force
    $Item | Add-Member -NotePropertyName DateKnownButMetadataNotWritten -NotePropertyValue $DateKnownButMetadataNotWritten -Force
}

function Invoke-CaptureDateMaterialization {
    param(
        [pscustomobject]$Item,
        [string]$RootPath,
        [string]$MetadataBackupRoot,
        [bool]$ExifToolAvailable,
        [switch]$UpdateProcessedIndex,
        [switch]$AllowMetadataWriteWithoutRepairExif,
        [string]$Reason = 'CaptureDateMaterialization'
    )

    $result = [pscustomobject]@{
        Candidate = $false
        EmbeddedMetadataWritten = $false
        FileSystemDatesSynced = $false
        DateKnownButMetadataNotWritten = $false
        HashChanged = $false
        OldHash = if ($Item -and $Item.PSObject.Properties.Name -contains 'Sha256') { [string]$Item.Sha256 } else { '' }
        NewHash = if ($Item -and $Item.PSObject.Properties.Name -contains 'Sha256') { [string]$Item.Sha256 } else { '' }
        EmbeddedCaptureDateState = 'NotChecked'
        Status = 'NotNeeded'
    }

    $embeddedDateProbe = Initialize-EmbeddedCaptureDateProbe -Item $Item
    $result.EmbeddedCaptureDateState = [string]$embeddedDateProbe.State

    if ($null -eq $Item -or -not (Test-DateInfoCanDriveVisibleDateMaterialization -DateInfo $Item.DateInfo)) {
        if ($null -ne $Item) {
            Set-CaptureDateMaterializationProperties -Item $Item -Status 'NoReliableCaptureDate' -EmbeddedWritten $false -FileSystemSynced $false -DateKnownButMetadataNotWritten $false
        }
        $result.Status = 'NoReliableCaptureDate'
        return $result
    }

    $repairResult = Repair-ExifDate -Item $Item -RootPath $RootPath -MetadataBackupRoot $MetadataBackupRoot -ExifToolAvailable $ExifToolAvailable -AllowMetadataWriteWithoutRepairExif:$AllowMetadataWriteWithoutRepairExif
    if ($repairResult.Candidate) {
        $result.Candidate = $true
        $Stats.CaptureDateMaterializationCandidates++
    }

    if ($repairResult.Repaired) {
        $result.EmbeddedMetadataWritten = $true
        $Stats.CaptureMetadataWritten++
        try {
            $newHash = Get-Sha256 -Path $Item.File.FullName
            if (-not [string]::IsNullOrWhiteSpace($newHash)) {
                $newHash = $newHash.ToUpperInvariant()
                $oldHash = $result.OldHash
                if (-not [string]::IsNullOrWhiteSpace($oldHash) -and -not $newHash.Equals($oldHash, [StringComparison]::OrdinalIgnoreCase)) {
                    $Item.Sha256 = $newHash
                    $result.HashChanged = $true
                    $result.NewHash = $newHash
                    if ($UpdateProcessedIndex) {
                        [void](Update-ProcessedIndexHashAfterExifRepair -Path $Item.File.FullName -OldHash $oldHash -NewHash $newHash)
                    }
                    else {
                        Write-Log -Message "File hash recalculated after capture metadata materialization before indexing: oldHash=$oldHash newHash=$newHash path=$($Item.File.FullName)" -Phase 'JSON reconciliation'
                    }
                }
            }
        }
        catch {
            $Stats.Errors++
            Write-Log -Message "Post-capture metadata hash recalculation failed for $($Item.File.FullName): $($_.Exception.Message)" -Phase 'JSON reconciliation'
        }
    }
    elseif ($repairResult.Candidate -and -not $Apply) {
        $result.DateKnownButMetadataNotWritten = $true
        $Stats.DateKnownButMetadataNotWritten++
        Write-Log -Message "DateKnownButMetadataNotWritten: Path=$($Item.File.FullName); Date=$($Item.DateInfo.Date.ToString('yyyy-MM-dd HH:mm:ss')); Source=$($Item.DateInfo.Source); Reason=DryRun" -Phase 'CaptureDateMaterialization'
    }
    elseif ((-not $repairResult.Repaired) -and [string]$repairResult.Reason -notin @('', 'ExistingEmbeddedDate')) {
        $result.DateKnownButMetadataNotWritten = $true
        $Stats.DateKnownButMetadataNotWritten++
        Write-Log -Message "DateKnownButMetadataNotWritten: Path=$($Item.File.FullName); Date=$($Item.DateInfo.Date.ToString('yyyy-MM-dd HH:mm:ss')); Source=$($Item.DateInfo.Source); Reason=$($repairResult.Reason)" -Phase 'CaptureDateMaterialization'
    }

    $materializationStateAllowsWrite = ([string]$embeddedDateProbe.State -eq 'Absent')
    $shouldSyncFs = ($materializationStateAllowsWrite -and ($SyncFileSystemDates -or $MetadataRepair -or $RepairExif -or $MetadataAudit -or $AllowMetadataWriteWithoutRepairExif))
    if ($shouldSyncFs) {
        $fsResult = Sync-FileSystemDatesFromValidatedCaptureDate -Item $Item -Reason $Reason
        if ($fsResult.Candidate) { $result.Candidate = $true }
        if ($fsResult.Synced) { $result.FileSystemDatesSynced = $true }
    }

    $result.Status = if ($result.EmbeddedMetadataWritten -or $result.FileSystemDatesSynced) {
        'Materialized'
    }
    elseif ($repairResult.Reason -eq 'ExistingEmbeddedDate') {
        'AlreadyHasEmbeddedCaptureDate'
    }
    elseif ($result.DateKnownButMetadataNotWritten) {
        'DateKnownButMetadataNotWritten'
    }
    elseif ($result.Candidate) {
        'Candidate'
    }
    else {
        'NotNeeded'
    }

    Set-CaptureDateMaterializationProperties -Item $Item -Status $result.Status -EmbeddedWritten ([bool]$result.EmbeddedMetadataWritten) -FileSystemSynced ([bool]$result.FileSystemDatesSynced) -DateKnownButMetadataNotWritten ([bool]$result.DateKnownButMetadataNotWritten)
    return $result
}

function Repair-ExifDate {
    param(
        [pscustomobject]$Item,
        [string]$RootPath,
        [string]$MetadataBackupRoot,
        [bool]$ExifToolAvailable,
        [switch]$AllowMetadataWriteWithoutRepairExif
    )

    $result = [pscustomobject]@{
        Repaired = $false
        Candidate = $false
        Reason = ''
        WriteMode = ''
    }

    if (-not $RepairExif -and -not $MetadataRepair -and -not $MetadataAudit -and -not $AllowMetadataWriteWithoutRepairExif) {
        return $result
    }

    $embeddedDateProbe = Initialize-EmbeddedCaptureDateProbe -Item $Item
    if ([string]$embeddedDateProbe.State -eq 'PresentValid') {
        $result.Reason = 'ExistingEmbeddedDate'
        return $result
    }
    if ([string]$embeddedDateProbe.State -ne 'Absent') {
        $result.Reason = 'EmbeddedCaptureDate' + [string]$embeddedDateProbe.State
        return $result
    }
    if (-not $ExifToolAvailable) {
        $result.Reason = 'EmbeddedCaptureDateUnreadable'
        return $result
    }

    $availability = Detect-StorageAvailability -Item $Item.File
    if ($availability.State -eq 'CloudPlaceholder') {
        Register-CloudPlaceholderSkipped -Path $Item.File.FullName -Phase 'EXIF' -Availability $availability
        return $result
    }
    if ($availability.State -eq 'MissingReal') {
        $Stats.MissingReal++
        Write-Log -Message "EXIF repair skipped missing file: $($Item.File.FullName). Reason=$($availability.Reason)" -Phase 'EXIF'
        return $result
    }

    if ($Item.DateInfo.Confidence -lt $ExifRepairConfidence) {
        return $result
    }

    $filenameRepairDecision = Get-ReliableFilenameExifRepairDecision -Item $Item
    if ($filenameRepairDecision.Conflict) {
        Write-Log -Message "EXIF repair from reliable filename skipped due to existing metadata conflict: $($Item.File.FullName). ConflictSource=$($filenameRepairDecision.ConflictSource)." -Phase 'EXIF'
        $result.Reason = 'MetadataConflict'
        return $result
    }

    $hasMultipleSources = @($Item.DateInfo.Sources).Count -ge 2
    $providerCanRepair = Test-ProviderDateInfoCanMaterialize -DateInfo $Item.DateInfo
    if ($Item.DateInfo.Confidence -lt $AutoActionConfidence -and -not $hasMultipleSources -and -not $filenameRepairDecision.CanRepair -and -not $providerCanRepair) {
        $result.Reason = 'InsufficientConfidence'
        return $result
    }

    $writeMode = if ($filenameRepairDecision.CanRepair) { [string]$filenameRepairDecision.WriteMode } else { Get-ExifRepairWriteMode -Extension $Item.Extension -IsVideo ([bool]$Item.IsVideo) }
    if ([string]::IsNullOrWhiteSpace($writeMode)) {
        $result.Reason = 'UnsupportedFormat'
        return $result
    }

    $dateText = $Item.DateInfo.Date.ToString('yyyy:MM:dd HH:mm:ss')
    $result.Candidate = $true
    $result.WriteMode = $writeMode

    if (-not $Apply) {
        $syntheticNote = if ($filenameRepairDecision.SyntheticTime) { ' Time component is synthetic/default.' } else { '' }
        Write-Log -Message ("EXIF repair candidate from {0}: {1} -> {2}. WriteMode={3}.{4}" -f $(if ($filenameRepairDecision.CanRepair) { $filenameRepairDecision.Kind } else { $Item.DateInfo.Source }), $Item.File.FullName, $dateText, $writeMode, $syntheticNote) -Phase 'EXIF'
        return $result
    }

    try {
        Backup-OriginalForMetadataChange -Item $Item -RootPath $RootPath -MetadataBackupRoot $MetadataBackupRoot
        $writeArguments = Get-ExifRepairWriteArguments -WriteMode $writeMode -DateText $dateText -Path $Item.File.FullName
        if ($writeArguments.Count -eq 0) {
            $result.Reason = 'UnsupportedFormat'
            return $result
        }
        $exif = Invoke-ExifTool -Path $ExifToolPath -Arguments $writeArguments
        if ($exif.TimedOut) {
            Write-Log -Message "ExifTool timeout while repairing metadata: $($Item.File.FullName)" -Phase 'EXIF'
            $Stats.Errors++
            return $result
        }
        if (-not $exif.Success) {
            Write-Log -Message "EXIF repair failed for $($Item.File.FullName): $($exif.Error)" -Phase 'EXIF'
            $Stats.Errors++
            return $result
        }
        $Stats.ExifRepaired++
        $result.Repaired = $true
        if ($filenameRepairDecision.CanRepair) {
            $syntheticNote = if ($filenameRepairDecision.SyntheticTime) { ' Time component is synthetic/default.' } else { '' }
            Write-Log -Message ("EXIF repaired from {0}: {1} -> {2}. WriteMode={3}.{4}" -f $filenameRepairDecision.Kind, $Item.File.Name, $dateText, $writeMode, $syntheticNote) -Phase 'EXIF'
        }
        else {
            Write-Log -Message ("EXIF repaired from {0}: {1} -> {2}. WriteMode={3}." -f $Item.DateInfo.Source, $Item.File.Name, $dateText, $writeMode) -Phase 'EXIF'
        }
    }
    catch {
        $Stats.Errors++
        Write-Warning "EXIF repair failed for $($Item.File.FullName): $($_.Exception.Message)"
    }

    return $result
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $defaultLogRoot = Join-Path (Get-UserDataRootPath) (Get-InternalFolderName -Key 'LogsFolder')
    $LogPath = Join-Path $defaultLogRoot ("PhotoOrganizer-{0}.log" -f $script:RunId)
}
if ([string]::IsNullOrWhiteSpace($ProgressPath)) {
    $defaultProgressRoot = [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($LogPath))
    $ProgressPath = Join-Path $defaultProgressRoot ("PhotoOrganizer-{0}.progress.json" -f $script:RunId)
}

Initialize-Logging
Write-Log -Message "Start PhotoOrganizer. PID: $PID" -Phase 'Start'
Write-Log -Message ("Mode requested: {0}" -f $(if (-not [string]::IsNullOrWhiteSpace($ImportProvider)) { if ($Apply) { "IMPORT PROVIDER $ImportProvider APPLY" } else { "IMPORT PROVIDER $ImportProvider DRY RUN" } } elseif ($RecoverFromWrongDuplicateMove) { if ($Apply) { 'RECOVER WRONG DUPLICATE MOVE APPLY' } else { 'RECOVER WRONG DUPLICATE MOVE DRY RUN' } } elseif ($MetadataRepair) { if ($Apply) { 'METADATA REPAIR APPLY' } else { 'METADATA REPAIR DRY RUN' } } elseif ($MetadataAudit) { 'METADATA AUDIT DRY RUN' } elseif ($RepairOnlyExistingOrganizedLibrary) { 'REPAIR ONLY EXISTING ORGANIZED LIBRARY' } elseif ($NormalizeExistingFolders) { if ($Apply) { 'NORMALIZE EXISTING FOLDERS APPLY' } else { 'NORMALIZE EXISTING FOLDERS DRY RUN' } } elseif ($RetentionCleanup) { if ($Apply) { 'RETENTION CLEANUP APPLY' } else { 'RETENTION CLEANUP DRY RUN' } } elseif ($DedupeCleanup) { if ($Apply) { 'DEDUPE CLEANUP APPLY' } else { 'DEDUPE CLEANUP DRY RUN' } } elseif ($ReconcileProcessedDatabase) { if ($Apply) { 'RECONCILE PROCESSED DATABASE APPLY' } else { 'RECONCILE PROCESSED DATABASE DRY RUN' } } elseif ($PurgeMissingFromProcessedDatabase) { if ($Apply) { 'PURGE MISSING PROCESSED DATABASE APPLY' } else { 'PURGE MISSING PROCESSED DATABASE DRY RUN' } } elseif ($RenameExistingFoldersToCurrentLanguage -or $RenameInternalFoldersToCurrentLanguage) { if ($Apply) { 'FOLDER LANGUAGE RENAME APPLY' } else { 'FOLDER LANGUAGE RENAME DRY RUN' } } elseif ($TestScan) { 'TEST SCAN' } elseif ($Apply) { 'APPLY' } else { 'DRY RUN' })) -Phase 'Start'

$SourcePath = Resolve-FullPath $SourcePath
Write-Log -Message "Validating paths..." -Phase 'Validation'
if (-not (Test-Path -LiteralPath $SourcePath -PathType Container)) {
    Stop-WithError "SourcePath does not exist: $SourcePath"
}

$script:WorkerCount = Resolve-WorkerCount
$script:DriveKind = Resolve-DriveKind -Path $SourcePath
Write-Log -Message ("Performance mode: {0}. CPU threads: {1}. Workers: {2}. Runtime: {3}. Storage: {4}." -f $PerformanceMode, $script:ProcessorCount, $script:WorkerCount, $script:PowerShellRuntime, $script:DriveKind) -Phase 'Validation'

$DestinationPath = Resolve-LocalizedDestinationPath -RequestedDestinationPath $DestinationPath -ResolvedSourcePath $SourcePath
$DestinationPath = Resolve-FullPath $DestinationPath
$DestinationBase = [System.IO.Path]::GetDirectoryName($DestinationPath.TrimEnd('\'))
if ([string]::IsNullOrWhiteSpace($DestinationBase)) {
    Stop-WithError "DestinationPath is invalid: $DestinationPath"
}

if ([string]::IsNullOrWhiteSpace($ProcessedDbPath)) {
    $ProcessedDbPath = Join-Path (Get-UserDataRootPath) 'ProcessedFiles.json'
}
$ProcessedDbPath = Resolve-FullPath $ProcessedDbPath

$OrganizedRoot = $DestinationPath
$IndexMaintenanceMode = [bool]($ReconcileProcessedDatabase -or $PurgeMissingFromProcessedDatabase)
$NeedsReviewRoot = Join-Path $DestinationBase (Get-InternalFolderName -Key 'NeedsReviewFolder')
$MediaMetadataIssuesRoot = Join-Path $NeedsReviewRoot (Get-InternalFolderName -Key 'MediaMetadataIssuesFolder')
$DuplicatesRoot = Join-Path $DestinationBase (Get-InternalFolderName -Key 'DuplicatesFolder')
$ConfirmedDuplicatesQuarantineRoot = Join-Path $DestinationBase (Get-InternalFolderName -Key 'ConfirmedDuplicatesQuarantineFolder')
$PreviousMetadataBackupRoot = Join-Path $DestinationBase (Get-InternalFolderName -Key 'MetadataBackupFolder')
$metadataBackupBase = if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) {
    Join-Path (Join-Path $env:LOCALAPPDATA 'PhotoOrganizer') (Get-InternalFolderName -Key 'MetadataBackupFolder')
}
else {
    Join-Path (Join-Path ([System.IO.Path]::GetTempPath()) 'PhotoOrganizer') (Get-InternalFolderName -Key 'MetadataBackupFolder')
}
$MetadataBackupRoot = Join-Path $metadataBackupBase $script:RunId
$ScriptRoot = Get-ScriptRootPath
$LogRoot = if (-not [string]::IsNullOrWhiteSpace($LogPath)) { [System.IO.Path]::GetDirectoryName([System.IO.Path]::GetFullPath($LogPath)) } else { Join-Path (Get-UserDataRootPath) (Get-InternalFolderName -Key 'LogsFolder') }
Initialize-FolderProtectionRules -ResolvedSourcePath $SourcePath -ResolvedDestinationPath $DestinationPath -ResolvedScriptRoot $ScriptRoot
$excludedRootList = @(
    $PreviousMetadataBackupRoot,
    $MetadataBackupRoot,
    $LogRoot
)
$excludedRootList += @($script:FolderProtectionRules | ForEach-Object { $_.Path })
if (-not $IndexMaintenanceMode) {
    $excludedRootList += @(
        $NeedsReviewRoot,
        $DuplicatesRoot,
        $ConfirmedDuplicatesQuarantineRoot
    )
}

$internalBaseRoots = @($DestinationBase, $SourcePath) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { Resolve-FullPath $_ } | Select-Object -Unique
foreach ($internalBaseRoot in $internalBaseRoots) {
    $excludedFolderKeys = if ($IndexMaintenanceMode) {
        @('MetadataBackupFolder', 'LogsFolder')
    }
    else {
        @('NeedsReviewFolder', 'DuplicatesFolder', 'ConfirmedDuplicatesQuarantineFolder', 'MetadataBackupFolder', 'LogsFolder')
    }
    foreach ($folderKey in $excludedFolderKeys) {
        foreach ($folderName in (Get-AllInternalFolderNames -Key $folderKey)) {
            $excludedRootList += (Join-Path $internalBaseRoot $folderName)
        }
    }

    if (-not $IndexMaintenanceMode) {
        foreach ($organizedName in (Get-AllInternalFolderNames -Key 'OrganizedFolder')) {
            $candidateOrganizedRoot = Join-Path $internalBaseRoot $organizedName
            if (-not ((Resolve-FullPath $candidateOrganizedRoot).TrimEnd('\').Equals((Resolve-FullPath $SourcePath).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase))) {
                $excludedRootList += $candidateOrganizedRoot
            }
        }
    }
}

if ((-not $IndexMaintenanceMode) -and -not ((Resolve-FullPath $OrganizedRoot).TrimEnd('\').Equals((Resolve-FullPath $SourcePath).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase))) {
    $excludedRootList += $OrganizedRoot
}
if ((Test-IsChildPath -Path $ScriptRoot -ParentPath $SourcePath) -and -not ((Resolve-FullPath $ScriptRoot).TrimEnd('\').Equals((Resolve-FullPath $SourcePath).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase))) {
    $excludedRootList += $ScriptRoot
}
$resolvedSourceForExclusions = (Resolve-FullPath $SourcePath).TrimEnd('\')
$ExcludedRoots = @(
    $excludedRootList |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { Resolve-FullPath $_ } |
        Where-Object {
            $candidate = $_.TrimEnd('\')
            (Test-IsChildPath -Path $candidate -ParentPath $resolvedSourceForExclusions) -and
                -not $candidate.Equals($resolvedSourceForExclusions, [StringComparison]::OrdinalIgnoreCase)
        } |
        Select-Object -Unique
)
$ExcludedRoots = @(
    @($ExcludedRoots) +
    @($script:FolderProtectionRules | ForEach-Object { Resolve-FullPath $_.Path })
) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

Write-Log -Message "Exclusions applied: $($ExcludedRoots -join '; ')" -Phase 'Validation'
foreach ($rule in @($script:FolderProtectionRules)) {
    Write-Log -Message ("External protected folder configured: Role={0}; Path={1}; Label={2}; Reason={3}; Exists={4}" -f $rule.Role, $rule.Path, $rule.Label, $rule.Reason, (Test-Path -LiteralPath $rule.Path -PathType Container)) -Phase 'Validation'
}

if (Test-FolderLanguageRenameOnlyMode) {
    Write-Notice ("Mode: {0}" -f $(if ($Apply) { 'APPLY' } else { 'DRY RUN' }))
    Write-Notice "Source: $SourcePath"
    Write-Notice "Destination: $OrganizedRoot"
    Write-Notice "Language selected: $Language"
    Write-Notice "Internal folders language: $Language"
    Rename-InternalFoldersToCurrentLanguage
    Rename-ExistingFoldersToCurrentLanguage
    Write-Log -Message "Folder language rename mode completed. No content scan, EXIF, hash, index maintenance or cloud traversal was run." -Phase 'Complete' -Status 'Completed'
    Close-Logging
    exit 0
}

Write-Log -Message "Checking ExifTool..." -Phase 'ExifTool'
$ExifToolPath = Resolve-ExifToolPath -PreferredPath $ExifToolPath
$ExifToolAvailable = Test-ExifTool -Path $ExifToolPath
if (-not $ExifToolAvailable) {
    Write-Log -Message "ExifTool was not found or did not respond. Metadata reading will be limited and EXIF repair will be skipped." -Phase 'ExifTool'
}
else {
    Write-Log -Message "ExifTool available: $ExifToolPath" -Phase 'ExifTool'
}

if ($RepairExif -and -not $Apply) {
    Write-Log -Message "-RepairExif was specified without -Apply. No EXIF changes will be made." -Phase 'Validation'
}
if ($MetadataAudit -and $Apply) {
    Write-Log -Message 'MetadataAudit is always DryRun. Ignoring -Apply for this mode.' -Phase 'Validation' -Status 'Warning'
    $Apply = $false
}

Write-Notice ("Mode: {0}" -f $(if ($Apply) { 'APPLY' } else { 'DRY RUN' }))
Write-Notice "Source: $SourcePath"
Write-Notice "Destination: $OrganizedRoot"
Write-Notice "Processed database: $ProcessedDbPath"
Write-Notice "Metadata backup: $MetadataBackupRoot"
Write-Notice "Confirmed duplicates quarantine: $ConfirmedDuplicatesQuarantineRoot"
Write-Notice "Organization profile: $OrganizationProfile"
Write-Notice ("Performance: {0}, workers: {1}, runtime: {2}" -f $PerformanceMode, $script:WorkerCount, $script:PowerShellRuntime)
Write-Notice "Language selected: $Language"
Write-Notice "Internal folders language: $Language"
Write-Notice "Excluded: protected internal folders and configured external/user excluded folders"

Rename-InternalFoldersToCurrentLanguage
Rename-ExistingFoldersToCurrentLanguage

$dedupeOperationalExcludedRoots = @(
    $PreviousMetadataBackupRoot,
    $MetadataBackupRoot,
    $LogRoot,
    $ConfirmedDuplicatesQuarantineRoot
)
if ((Test-IsChildPath -Path $ScriptRoot -ParentPath $SourcePath) -and -not ((Resolve-FullPath $ScriptRoot).TrimEnd('\').Equals((Resolve-FullPath $SourcePath).TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase))) {
    $dedupeOperationalExcludedRoots += $ScriptRoot
}
foreach ($internalBaseRoot in $internalBaseRoots) {
    foreach ($folderKey in @('MetadataBackupFolder', 'LogsFolder', 'ConfirmedDuplicatesQuarantineFolder')) {
        foreach ($folderName in (Get-AllInternalFolderNames -Key $folderKey)) {
            $dedupeOperationalExcludedRoots += (Join-Path $internalBaseRoot $folderName)
        }
    }
}
$script:DedupeExcludedRoots = @(
    @(
        $dedupeOperationalExcludedRoots |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { Resolve-FullPath $_ } |
            Where-Object {
                $candidate = $_.TrimEnd('\')
                (Test-IsChildPath -Path $candidate -ParentPath $resolvedSourceForExclusions) -and
                    -not $candidate.Equals($resolvedSourceForExclusions, [StringComparison]::OrdinalIgnoreCase)
            }
    ) +
    @($script:FolderProtectionRules | ForEach-Object { $_.Path }) |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { Resolve-FullPath $_ } |
        Select-Object -Unique
)

$script:DedupeOrganizedRoots = @()
$script:DedupeDuplicatesRoots = @()
$script:DedupeNeedsReviewRoots = @()
foreach ($internalBaseRoot in $internalBaseRoots) {
    foreach ($folderName in (Get-AllInternalFolderNames -Key 'OrganizedFolder')) {
        $script:DedupeOrganizedRoots += (Join-Path $internalBaseRoot $folderName)
    }
    foreach ($folderName in (Get-AllInternalFolderNames -Key 'DuplicatesFolder')) {
        $script:DedupeDuplicatesRoots += (Join-Path $internalBaseRoot $folderName)
    }
    foreach ($folderName in (Get-AllInternalFolderNames -Key 'NeedsReviewFolder')) {
        $script:DedupeNeedsReviewRoots += (Join-Path $internalBaseRoot $folderName)
    }
}
$script:DedupeOrganizedRoots = @($script:DedupeOrganizedRoots | ForEach-Object { Resolve-FullPath $_ } | Select-Object -Unique)
$script:DedupeDuplicatesRoots = @($script:DedupeDuplicatesRoots | ForEach-Object { Resolve-FullPath $_ } | Select-Object -Unique)
$script:DedupeNeedsReviewRoots = @($script:DedupeNeedsReviewRoots | ForEach-Object { Resolve-FullPath $_ } | Select-Object -Unique)

if ($RetentionCleanup) {
    Invoke-RetentionCleanup
    Close-Logging
    exit 0
}

if (-not [string]::IsNullOrWhiteSpace($ImportProvider)) {
    if ([string]::IsNullOrWhiteSpace($ImportProviderPath)) {
        Stop-WithError "ImportProviderPath is required when -ImportProvider is used."
    }
    $importProviderSpec = Assert-ImportProviderAvailable -Provider $ImportProvider
    switch ([string]$importProviderSpec.Id) {
        'GoogleTakeout' { Invoke-ImportProviderGoogleTakeout -ProviderRootPath $ImportProviderPath }
        'ApplePhotos' { Invoke-ImportProviderApplePhotos -ProviderRootPath $ImportProviderPath }
        'XmpSidecarLibrary' { Invoke-ImportProviderXmpSidecarLibrary -ProviderRootPath $ImportProviderPath }
        default { Stop-WithError "ImportProvider $($importProviderSpec.DisplayName) is registered but has no executable implementation yet." }
    }
    Close-Logging
    exit 0
}

if ($DedupeCleanup) {
    Invoke-DedupeCleanup
}

if ($RecoverFromWrongDuplicateMove) {
    Invoke-RecoverFromWrongDuplicateMove
}

if ($NormalizeExistingFolders) {
    Invoke-NormalizeExistingFolders
}

if ($MetadataAudit -or $MetadataRepair) {
    Invoke-MetadataAuditOrRepair -CloseAndExit
}

if ($RepairOnlyExistingOrganizedLibrary) {
    Invoke-RepairOnlyExistingOrganizedLibrary -CloseAndExit
}

if ($ReconcileProcessedDatabase) {
    $scanRoots = New-Object System.Collections.Generic.List[string]
    $scanRoots.Add((Resolve-FullPath $SourcePath))
    $resolvedOrganizedRoot = Resolve-FullPath $OrganizedRoot
    if (-not (Test-IsChildPath -Path $resolvedOrganizedRoot -ParentPath $SourcePath) -and
        -not $scanRoots.Contains($resolvedOrganizedRoot)) {
        $scanRoots.Add($resolvedOrganizedRoot)
    }

    $fileByPath = @{}
    foreach ($scanRoot in @($scanRoots.ToArray())) {
        if (-not (Test-Path -LiteralPath $scanRoot -PathType Container)) { continue }
        Write-Log -Message "Index reconciliation scan root: $scanRoot" -Phase 'JSON reconciliation'
        foreach ($file in @(Get-CompatibleFiles -RootPath $scanRoot)) {
            $fileByPath[(Resolve-FullPath $file.FullName).ToLowerInvariant()] = $file
        }
    }
    $files = @($fileByPath.Values)
    Load-ProcessedDatabase
    Invoke-ProcessedDatabaseSmartReconciliation -Files $files
    Close-Logging
    exit 0
}

if ($PurgeMissingFromProcessedDatabase) {
    $scanRoots = New-Object System.Collections.Generic.List[string]
    $scanRoots.Add((Resolve-FullPath $SourcePath))
    $resolvedOrganizedRoot = Resolve-FullPath $OrganizedRoot
    if (-not (Test-IsChildPath -Path $resolvedOrganizedRoot -ParentPath $SourcePath) -and
        -not $scanRoots.Contains($resolvedOrganizedRoot)) {
        $scanRoots.Add($resolvedOrganizedRoot)
    }

    $fileByPath = @{}
    foreach ($scanRoot in @($scanRoots.ToArray())) {
        if (-not (Test-Path -LiteralPath $scanRoot -PathType Container)) { continue }
        Write-Log -Message "Purge preflight reconciliation scan root: $scanRoot" -Phase 'JSON reconciliation'
        foreach ($file in @(Get-CompatibleFiles -RootPath $scanRoot)) {
            $fileByPath[(Resolve-FullPath $file.FullName).ToLowerInvariant()] = $file
        }
    }
    $files = @($fileByPath.Values)

    Load-ProcessedDatabase
    if ($script:StaleProcessedRecords.Count -gt 0) {
        Write-Log -Message "Purge preflight: reconciling stale entries by hash before any purge decision. Stale entries: $($script:StaleProcessedRecords.Count)." -Phase 'JSON reconciliation'
        Invoke-ProcessedDatabaseSmartReconciliation -Files $files
    }
    else {
        Write-Log -Message "Purge preflight: no stale entries found before purge." -Phase 'JSON reconciliation'
    }
    Invoke-PurgeMissingFromProcessedDatabase
    Close-Logging
    exit 0
}

$files = @(Get-CompatibleFiles -RootPath $SourcePath)

if ($TestScan) {
    Write-Log -Message "Test scan complete. Found $($files.Count) local files. Cloud placeholders skipped: $($Stats.CloudPlaceholdersSkipped). Missing real: $($Stats.MissingReal). Inaccessible: $($Stats.Inaccessible)." -Phase 'Complete' -Status 'Completed'
    Write-Host ''
    Write-Host (Get-SummaryText -Key 'FinalSummary')
    Write-Host '-------------'
    Write-SummaryHostLine -Key 'FilesFound' -Value $Stats.FilesFound
    Write-SummaryHostLine -Key 'LocalFilesDetected' -Value $Stats.LocalFilesDetected
    Write-SummaryHostLine -Key 'CloudPlaceholdersSkipped' -Value $Stats.CloudPlaceholdersSkipped
    Write-SummaryHostLine -Key 'MissingReal' -Value $Stats.MissingReal
    Write-SummaryHostLine -Key 'SkippedOneDrive' -Value $Stats.SkippedOneDrive
    Write-SummaryHostLine -Key 'Inaccessible' -Value $Stats.Inaccessible
    Close-Logging
    exit 0
}

if ($files.Count -eq 0) {
    Write-Log -Message "Organize found no processable files in this run. Skipping destination initialization, ProcessedFiles load, duplicate index build and JSON reconciliation. Use -ReconcileProcessedDatabase explicitly to synchronize the index." -Phase 'Complete' -Status 'Completed'
    Write-Host ''
    Write-Log -Message (Get-SummaryText -Key 'FinalSummary') -Phase 'Complete'
    Write-Log -Message '-------------' -Phase 'Complete'
    Write-SummaryLine -Key 'FilesFound' -Value $Stats.FilesFound
    Write-SummaryLine -Key 'LocalFilesDetected' -Value $Stats.LocalFilesDetected
    Write-SummaryLine -Key 'CloudPlaceholdersSkipped' -Value $Stats.CloudPlaceholdersSkipped
    Write-SummaryLine -Key 'MissingReal' -Value $Stats.MissingReal
    Write-SummaryLine -Key 'SkippedOneDrive' -Value $Stats.SkippedOneDrive
    Write-SummaryLine -Key 'Inaccessible' -Value $Stats.Inaccessible
    Write-Log -Message ('{0}: {1}' -f (Get-SummaryText -Key 'Errors'), $Stats.Errors) -Phase 'Complete' -Status 'Completed'
    if (-not $Apply) {
        Write-Host ''
        Write-Log -Message (Get-SummaryText -Key 'DryRunNoChanges') -Phase 'Complete' -Status 'Completed'
    }
    Close-Logging
    exit 0
}

Write-Log -Message "Starting processing..." -Phase 'Processing'
Initialize-DestinationStructure
Load-ProcessedIndexLight

$seenHashes = @{}
foreach ($knownHash in $script:ProcessedByHash.Keys) {
    $seenHashes[$knownHash] = $true
}
Write-Log -Message "Global duplicate index loaded: $($seenHashes.Count) hashes" -Phase 'Incremental'

$BatchSize = [math]::Max(1, $BatchSize)
$script:TotalBatches = [int][math]::Ceiling($files.Count / [double]$BatchSize)
$lastProcessLog = Get-Date
Start-OperationalProgress -Name 'Organize' -Total $files.Count -Phase 'Processing' -Message 'Hashing, reading metadata, moving files and saving incremental progress.'

for ($batchStart = 0; $batchStart -lt $files.Count; $batchStart += $BatchSize) {
    $script:CurrentBatch++
    $batchEnd = [math]::Min($batchStart + $BatchSize - 1, $files.Count - 1)
    $batch = @($files[$batchStart..$batchEnd])
    $script:QueueSize = [math]::Max(0, $files.Count - $batchStart)
    Write-Log -Message "Processing batch $script:CurrentBatch/$script:TotalBatches ($($batch.Count) files)." -Phase 'Processing'

    Write-Log -Message "Hashing batch $script:CurrentBatch/$script:TotalBatches with up to $script:WorkerCount workers." -Phase 'Hash queue'
    Update-OperationalProgress -Current $Stats.FilesAnalyzed -Total $files.Count -Phase 'Hash queue' -Stage 'Hashing batch' -EveryItems 1000 -EveryMinutes 5
    $hashResults = Get-Sha256Batch -Files $batch
    $batchCandidates = New-Object System.Collections.Generic.List[object]
    foreach ($file in $batch) {
        try {
            $Stats.FilesAnalyzed++
            $script:QueueSize = [math]::Max(0, $Stats.FilesFound - $Stats.FilesAnalyzed)
            Update-Throughput
            Update-OperationalProgress -Current $Stats.FilesAnalyzed -Total $files.Count -Phase 'Processing' -Stage 'Processing/hash classification' -EveryItems 1000 -EveryMinutes 5
            if ($Stats.FilesAnalyzed -eq 1) {
                Write-Log -Message "First file processed: $($file.FullName)" -Phase 'Processing'
            }
            elseif ($Stats.FilesAnalyzed % 100 -eq 0) {
                Write-Log -Message ("Progress: processed {0} of {1} files. {2} files/sec. ETA: {3}" -f $Stats.FilesAnalyzed, $Stats.FilesFound, $script:FilesPerSecond, $script:EtaText) -Phase 'Processing'
            }
            elseif (((Get-Date) - $lastProcessLog).TotalSeconds -ge 10) {
                $lastProcessLog = Get-Date
                Write-Log -Message ("Progress: processed {0} of {1} files. Current file: {2}. {3} files/sec. ETA: {4}" -f $Stats.FilesAnalyzed, $Stats.FilesFound, $file.Name, $script:FilesPerSecond, $script:EtaText) -Phase 'Processing'
            }

            $hashResult = if ($hashResults.ContainsKey($file.FullName)) { $hashResults[$file.FullName] } else { $null }
            if ($null -eq $hashResult -or [string]::IsNullOrWhiteSpace($hashResult.Hash)) {
                $message = if ($hashResult -and $hashResult.Error) { [string]$hashResult.Error } else { 'No hash result returned.' }
                if ($message -eq 'CloudPlaceholder') {
                    Write-Log -Message "Hash skipped cloud placeholder: $($file.FullName)" -Phase 'Processing'
                    continue
                }
                if ($message -eq 'MissingReal') {
                    Write-Log -Message "Hash skipped missing file: $($file.FullName)" -Phase 'Processing'
                    continue
                }
                $Stats.Errors++
                $Stats.Inaccessible++
                Write-Log -Message "Hash/read failed for $($file.FullName): $message" -Phase 'Processing'
                continue
            }

            $sha256 = $hashResult.Hash
            $hashKey = $sha256.ToUpperInvariant()
            $processedRecord = if ($script:ProcessedByHash.ContainsKey($hashKey)) { $script:ProcessedByHash[$hashKey] } else { $null }
            if ($null -ne $processedRecord) {
                $processedRecord = Resolve-ProcessedRecordForOrganizeHash -HashKey $hashKey -Record $processedRecord -File $file
                if ($null -eq $processedRecord) {
                    $seenHashes.Remove($hashKey) | Out-Null
                }
            }
            if (Test-ProcessedRecordMatchesCurrentPath -Record $processedRecord -File $file) {
                $Stats.IncrementalSkipped++
                if ($Stats.IncrementalSkipped -eq 1 -or $Stats.IncrementalSkipped % 100 -eq 0 -or $Diagnostic) {
                    Write-Log -Message "Incremental skip: same file already processed for $($file.FullName)" -Phase 'Incremental'
                }
                continue
            }

            $batchCandidates.Add([pscustomobject]@{
                File = $file
                Sha256 = $hashKey
                GlobalDuplicate = ($null -ne $processedRecord -or $seenHashes.ContainsKey($hashKey))
            })
        }
        catch {
            $Stats.Errors++
            Write-Log -Message "Hash/read failed for $($file.FullName): $($_.Exception.Message)" -Phase 'Processing'
        }
    }

    if ($batchCandidates.Count -eq 0) {
        Write-Log -Message "Saving progress after batch $script:CurrentBatch/$script:TotalBatches." -Phase 'Saving progress'
        Save-ProcessedDatabase
        continue
    }

    $metadataFiles = @($batchCandidates | Where-Object { -not $_.GlobalDuplicate } | ForEach-Object { $_.File })
    $script:ActiveWorkers = [math]::Min($script:WorkerCount, $metadataFiles.Count)
    Write-Log -Message "Reading EXIF for batch $script:CurrentBatch/$script:TotalBatches." -Phase 'Metadata queue'
    Update-OperationalProgress -Current $Stats.FilesAnalyzed -Total $files.Count -Phase 'Metadata queue' -Stage 'Reading EXIF metadata. Slow EXIF candidates are logged separately when detected.' -EveryItems 1000 -EveryMinutes 5
    $metadataMap = Get-ExifMetadataBatch -Files $metadataFiles -ExifToolAvailable $ExifToolAvailable
    $script:ActiveWorkers = 0

    $items = New-Object System.Collections.Generic.List[object]
    foreach ($candidate in $batchCandidates) {
        try {
            $file = $candidate.File
            $ext = $file.Extension.ToLowerInvariant()
            $isVideo = $VideoExtensions -contains $ext
            $metadata = if ($metadataMap.ContainsKey($file.FullName)) { $metadataMap[$file.FullName] } else { New-EmptyMetadata }
            $item = [pscustomobject]@{
                File = $file
                Extension = $ext
                IsVideo = $isVideo
                IsRaw = $RawExtensions -contains $ext
                Metadata = $metadata
                DateInfo = $null
                Sha256 = $candidate.Sha256
                GlobalDuplicate = [bool]$candidate.GlobalDuplicate
                PerceptualHash = $null
                Width = if ($metadata.ImageWidth) { [int]$metadata.ImageWidth } else { 0 }
                Height = if ($metadata.ImageHeight) { [int]$metadata.ImageHeight } else { 0 }
                DuplicateHandled = $false
                EmbeddedMetadataReadStatus = [string]$metadata.ReadStatus
            }
            $item.DateInfo = Get-PrimaryDate -Item $item -IsVideo $isVideo
            [void](Initialize-EmbeddedCaptureDateProbe -Item $item)
            Write-DateInfoDiagnostic -File $file -DateInfo $item.DateInfo -Context 'Organize'
            $items.Add($item)
        }
        catch {
            $Stats.Errors++
            Write-Log -Message "Analyze failed for $($candidate.File.FullName): $($_.Exception.Message)" -Phase 'Processing'
        }
    }

    Write-Log -Message "Moving batch $script:CurrentBatch/$script:TotalBatches." -Phase 'Moving'
    Update-OperationalProgress -Current $Stats.FilesAnalyzed -Total $files.Count -Phase 'Moving' -Stage 'Moving/copying files. OneDrive and Explorer may be the main bottleneck during mass moves.' -EveryItems 1000 -EveryMinutes 5
    $script:ActiveWorkers = 1
    foreach ($item in $items) {
        try {
            $resolvedItemPath = Resolve-FullPath $item.File.FullName
            if (-not $item.GlobalDuplicate -and $script:ExifProblemFiles.ContainsKey($resolvedItemPath)) {
                $Stats.NeedsReview++
                $problemReason = $script:ExifProblemFiles[$resolvedItemPath]
                $problemDirectory = if ($problemReason -eq 'Media corruption / WhatsApp metadata issue') { $MediaMetadataIssuesRoot } else { $NeedsReviewRoot }
                Invoke-SafeTransfer -Item $item -DestinationDirectory $problemDirectory -Reason $problemReason
                continue
            }

            if ($item.GlobalDuplicate -or $seenHashes.ContainsKey($item.Sha256)) {
                $Stats.ExactDuplicatesFound++
                $item.DuplicateHandled = $true
                Invoke-SafeTransfer -Item $item -DestinationDirectory $DuplicatesRoot -Reason 'Exact duplicate - Global index'
                continue
            }

            $seenHashes[$item.Sha256] = $true
            if ($item.IsRaw) {
                Write-DiagnosticLog "RAW master kept and organized independently: $($item.File.FullName)"
            }

            $materializationResult = Invoke-CaptureDateMaterialization -Item $item -RootPath $SourcePath -MetadataBackupRoot $MetadataBackupRoot -ExifToolAvailable $ExifToolAvailable -Reason 'Organize'
            if ($materializationResult.HashChanged -and -not [string]::IsNullOrWhiteSpace($materializationResult.NewHash)) {
                $seenHashes[$materializationResult.NewHash] = $true
            }

            if ($item.DateInfo.Confidence -lt $ExifRepairConfidence) {
                $Stats.NeedsReview++
                Invoke-SafeTransfer -Item $item -DestinationDirectory $NeedsReviewRoot -Reason 'Low confidence date'
                continue
            }

            $destinationDirectory = Get-DestinationPath -Item $item -RootPath $SourcePath -OrganizedRoot $OrganizedRoot
            Invoke-SafeTransfer -Item $item -DestinationDirectory $destinationDirectory -Reason 'Organize'
        }
        catch {
            $Stats.Errors++
            Write-Log -Message "Action failed for $($item.File.FullName): $($_.Exception.Message)" -Phase 'Moving'
        }
    }
    $script:ActiveWorkers = 0

    Write-Log -Message "Saving progress after batch $script:CurrentBatch/$script:TotalBatches." -Phase 'Saving progress'
    Save-ProcessedDatabase
}
Complete-OperationalProgress -Phase 'Processing' -Message 'Organize processing phase completed. Entering optional repair/cleanup/finalization phases.'

if ($RepairExif -and $Apply) {
    Write-Log -Message "RepairExif was applied only to files processed during this Organize run. Use -RepairOnlyExistingOrganizedLibrary explicitly for a full organized-library maintenance pass." -Phase 'RepairExif'
}

Remove-EmptySourceFolders
Update-MetadataBackupSize | Out-Null
Cleanup-MetadataBackupsOnSuccess

Write-Log -Message "Finishing run..." -Phase 'Complete'
Write-Host ''
Write-Log -Message (Get-SummaryText -Key 'FinalSummary') -Phase 'Complete'
Write-Log -Message '-------------' -Phase 'Complete'
Write-SummaryLine -Key 'FilesFound' -Value $Stats.FilesFound
Write-SummaryLine -Key 'LocalFilesDetected' -Value $Stats.LocalFilesDetected
Write-SummaryLine -Key 'CloudPlaceholdersSkipped' -Value $Stats.CloudPlaceholdersSkipped
Write-SummaryLine -Key 'CloudPlaceholdersInIndex' -Value $Stats.CloudPlaceholdersInIndex
Write-SummaryLine -Key 'MissingReal' -Value $Stats.MissingReal
Write-SummaryLine -Key 'FilesAnalyzed' -Value $Stats.FilesAnalyzed
Write-SummaryLine -Key 'IncrementalSkipped' -Value $Stats.IncrementalSkipped
Write-SummaryLine -Key 'ExactDuplicatesFound' -Value $Stats.ExactDuplicatesFound
Write-SummaryLine -Key 'NearDuplicatesFound' -Value $Stats.NearDuplicatesFound
Write-SummaryLine -Key 'ExifRepaired' -Value $Stats.ExifRepaired
Write-Log -Message ("Capture date materialization: candidates={0}; embeddedWritten={1}; filesystemDatesSynced={2}; dateKnownButMetadataNotWritten={3}" -f $Stats.CaptureDateMaterializationCandidates, $Stats.CaptureMetadataWritten, $Stats.FileSystemDatesSynced, $Stats.DateKnownButMetadataNotWritten) -Phase 'Complete'
Write-SummaryLine -Key 'FilesMoved' -Value $Stats.FilesMoved
Write-SummaryLine -Key 'FilesCopied' -Value $Stats.FilesCopied
Write-SummaryLine -Key 'ExistingIdenticalSkipped' -Value $Stats.ExistingIdenticalSkipped
Write-SummaryLine -Key 'NeedsReview' -Value $Stats.NeedsReview
Write-SummaryLine -Key 'EmptyFoldersRemoved' -Value $Stats.EmptyFoldersRemoved
Write-SummaryLine -Key 'JunkOnlyFoldersRemoved' -Value $Stats.JunkOnlyFoldersRemoved
Write-SummaryLine -Key 'JunkOnlySmallMarkerFoldersRemoved' -Value $Stats.JunkOnlySmallMarkerFoldersRemoved
Write-SummaryLine -Key 'ZombieNormalizeFoldersRemoved' -Value $Stats.ZombieNormalizeFoldersRemoved
Write-SummaryLine -Key 'OrganizationProfile' -Value $OrganizationProfile
Write-SummaryLine -Key 'FoldersReduced' -Value $Stats.FoldersReduced
Write-SummaryLine -Key 'JsonPathsUpdated' -Value $Stats.JsonPathsUpdated
Write-SummaryLine -Key 'SkippedUncertainNames' -Value $Stats.SkippedUncertainNames
Write-SummaryLine -Key 'SkippedOneDrive' -Value $Stats.SkippedOneDrive
Write-SummaryLine -Key 'Inaccessible' -Value $Stats.Inaccessible
Write-SummaryLine -Key 'LockedFiles' -Value $Stats.LockedFiles
Write-SummaryLine -Key 'RetryCount' -Value $Stats.RetryCount
Write-SummaryLine -Key 'MetadataCorruptedMedia' -Value $Stats.MetadataCorruptedMedia
Write-SummaryLine -Key 'SlowExifCandidates' -Value $Stats.SlowExifCandidates
Write-SummaryLine -Key 'SlowExifDetections' -Value $Stats.SlowExifDetections
Write-SummaryLine -Key 'ExifBatchTimeouts' -Value $Stats.ExifBatchTimeouts
Write-SummaryLine -Key 'ExifBatchTimeoutAffectedFiles' -Value $Stats.ExifBatchTimeoutAffectedFiles
Write-SummaryLine -Key 'ExifBatchFallbacks' -Value $Stats.ExifBatchFallbacks
Write-SummaryLine -Key 'ExifBatchFallbackAffectedFiles' -Value $Stats.ExifBatchFallbackAffectedFiles
Write-SummaryLine -Key 'MetadataBackupSizeGB' -Value $Stats.MetadataBackupSizeGB
Write-SummaryLine -Key 'RetentionDeletedItems' -Value $Stats.RetentionDeletedItems
Write-SummaryLine -Key 'RetentionRecoveredGB' -Value $Stats.RetentionRecoveredGB
Write-SummaryLine -Key 'JsonReconcileValid' -Value $Stats.JsonReconcileValid
Write-SummaryLine -Key 'JsonReconcileStale' -Value $Stats.JsonReconcileStale
Write-SummaryLine -Key 'JsonReconcilePathsUpdated' -Value $Stats.JsonReconcilePathsUpdated
Write-SummaryLine -Key 'JsonReconcileMissing' -Value $Stats.JsonReconcileMissing
Write-SummaryLine -Key 'JsonReconcilePurged' -Value $Stats.JsonReconcilePurged
Write-SummaryLine -Key 'JsonReconcileEntriesRemoved' -Value $Stats.JsonReconcileEntriesRemoved
Write-SummaryLine -Key 'JsonReconcileConflicts' -Value $Stats.JsonReconcileConflicts
Write-SummaryLine -Key 'DryRunActions' -Value $Stats.DryRunActions
Write-Log -Message ('{0}: {1}' -f (Get-SummaryText -Key 'Errors'), $Stats.Errors) -Phase 'Complete' -Status 'Completed'

if (-not $Apply) {
    Write-Host ''
    Write-Log -Message (Get-SummaryText -Key 'DryRunNoChanges') -Phase 'Complete' -Status 'Completed'
}

Save-ProcessedDatabase
Close-Logging














