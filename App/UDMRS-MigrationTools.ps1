function Add-FileToZipArchive {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$FilePath,
        [string]$EntryName
    )

    if ([string]::IsNullOrWhiteSpace($FilePath) -or -not (Test-Path -LiteralPath $FilePath -PathType Leaf)) {
        return
    }

    $normalizedEntry = $EntryName.Replace('\', '/').TrimStart('/')
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($Archive, $FilePath, $normalizedEntry, [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null
}

function Get-PortableRelativePath {
    param(
        [string]$BasePath,
        [string]$Path
    )

    $baseFull = [System.IO.Path]::GetFullPath($BasePath).TrimEnd('\')
    $pathFull = [System.IO.Path]::GetFullPath($Path)
    if ($pathFull.Length -le $baseFull.Length) {
        return [System.IO.Path]::GetFileName($pathFull)
    }
    return $pathFull.Substring($baseFull.Length).TrimStart('\')
}

function Get-UDMRSDefaultMigrationOutputRoot {
    $userProfile = [Environment]::GetFolderPath('UserProfile')
    if ([string]::IsNullOrWhiteSpace($userProfile)) {
        $userProfile = $env:USERPROFILE
    }
    if ([string]::IsNullOrWhiteSpace($userProfile)) {
        return $null
    }

    return (Join-Path (Join-Path $userProfile 'Downloads') 'UDMRS-MigrationPackages')
}

function Get-UDMRSMigrationText {
    param([string]$LanguageCode)

    switch (($LanguageCode + '').ToLowerInvariant()) {
        'ro' {
            return [pscustomobject]@{
                InstallZipSegment = 'InstalarePartajata'
                UserZipSegment = 'StareUtilizator'
                Title = 'Ghid de migrare UDMRS'
                Generated = 'Generat'
                SharedZip = 'ZIP instalare partajata'
                UserZip = 'ZIP stare utilizator curent'
                SourcePaths = 'Rute sursa folosite de acest pachet'
                InstallRoot = 'Radacina instalare'
                UserRoot = 'Radacina date utilizator'
                LocalRoot = 'Radacina date locale'
                Included = 'Ce este inclus'
                InstallIncluded = 'ZIP instalare partajata: App, Docs, Tools, Branding, Templates, Releases, Config, README.md si lansatoarele existente.'
                UserIncluded = 'ZIP stare utilizator: ProcessedFiles.json, Config, IndexBackups, setari dashboard si fisiere JSON de stare necesare pentru continuarea istoricului acelui utilizator.'
                NotIncluded = 'Ce nu este inclus'
                ProgressFiles = 'fisiere progress.json'
                QueueFiles = 'queue.jsonl'
                StatusFiles = 'fisiere pid/status/lock'
                TempState = 'stare temporara a consolei tehnice'
                ExifBackups = 'backup-uri EXIF locale'
                Steps = 'Pasi de migrare'
                Step1 = 'Copiaza ZIP-ul de instalare pe noul PC.'
                Step2 = 'Extrage-l in folderul unde vrei sa folosesti UDMRS.'
                Step3 = 'Copiaza ZIP-ul de stare utilizator pe noul PC.'
                Step4 = 'Extrage-l in folderul AppData corespunzator utilizatorului Windows tinta.'
                TypicalTarget = 'Tinta tipica'
                Step5 = 'Porneste dashboard-ul din Start-PhotoOrganizer.cmd in instalarea extrasa.'
                Step6 = 'Ruleaza mai intai Test Scan.'
                Step7 = 'Ruleaza ReconcileProcessedDatabase in DryRun pentru a valida rutele si sanatatea indexului inainte de orice Apply.'
                Notes = 'Note'
                NoteLogs = 'Logs si Runtime nu se migreaza deoarece UDMRS le regenereaza automat.'
                NoteSingleRun = 'Pastreaza o singura operatie activa a motorului cand folosesti o biblioteca sincronizata.'
                NoteCloud = 'Daca folosesti OneDrive/Dropbox/iCloud/Google Drive, asigura-te ca fisierele pe care vrei sa le procesezi sunt disponibile local inainte de Apply.'
            }
        }
        'en' {
            return [pscustomobject]@{
                InstallZipSegment = 'SharedInstall'
                UserZipSegment = 'UserState'
                Title = 'UDMRS migration guide'
                Generated = 'Generated'
                SharedZip = 'Shared installation ZIP'
                UserZip = 'Current user state ZIP'
                SourcePaths = 'Source paths used by this package'
                InstallRoot = 'Installation root'
                UserRoot = 'User data root'
                LocalRoot = 'Local data root'
                Included = 'What is included'
                InstallIncluded = 'Shared installation ZIP: App, Docs, Tools, Branding, Templates, Releases, Config, README.md and launcher files that exist in the current installation.'
                UserIncluded = 'User state ZIP: ProcessedFiles.json, Config, IndexBackups, dashboard settings and JSON state files needed to continue the same user history.'
                NotIncluded = 'What is not included'
                ProgressFiles = 'progress.json files'
                QueueFiles = 'queue.jsonl'
                StatusFiles = 'pid/status/lock files'
                TempState = 'temporary technical console state'
                ExifBackups = 'local EXIF metadata backups'
                Steps = 'Migration steps'
                Step1 = 'Copy the shared installation ZIP to the new PC.'
                Step2 = 'Extract it into the folder you want to use for UDMRS.'
                Step3 = 'Copy the user state ZIP to the new PC.'
                Step4 = 'Extract it into the corresponding AppData folder for the target Windows user.'
                TypicalTarget = 'Typical target'
                Step5 = 'Start the dashboard from Start-PhotoOrganizer.cmd in the extracted installation.'
                Step6 = 'Run Test Scan first.'
                Step7 = 'Run ReconcileProcessedDatabase in DryRun to validate paths and index health before any Apply operation.'
                Notes = 'Notes'
                NoteLogs = 'Logs and Runtime are not migrated because UDMRS regenerates them automatically.'
                NoteSingleRun = 'Keep one engine operation active at a time when using a synced library.'
                NoteCloud = 'If OneDrive/Dropbox/iCloud/Google Drive is involved, make sure the content you want to process is locally available before running Apply.'
            }
        }
        default {
            return [pscustomobject]@{
                InstallZipSegment = 'InstalacionCompartida'
                UserZipSegment = 'EstadoUsuario'
                Title = 'Guia de migracion UDMRS'
                Generated = 'Generado'
                SharedZip = 'ZIP de instalacion compartida'
                UserZip = 'ZIP de estado del usuario actual'
                SourcePaths = 'Rutas de origen usadas por este paquete'
                InstallRoot = 'Raiz de instalacion'
                UserRoot = 'Raiz de datos de usuario'
                LocalRoot = 'Raiz de datos locales'
                Included = 'Que se incluye'
                InstallIncluded = 'ZIP de instalacion compartida: App, Docs, Tools, Branding, Templates, Releases, Config, README.md y lanzadores existentes.'
                UserIncluded = 'ZIP de estado del usuario: ProcessedFiles.json, Config, IndexBackups, ajustes del dashboard y JSON de estado necesarios para continuar el mismo historial de usuario.'
                NotIncluded = 'Que no se incluye'
                ProgressFiles = 'archivos progress.json'
                QueueFiles = 'queue.jsonl'
                StatusFiles = 'archivos pid/status/lock'
                TempState = 'estado temporal de consola tecnica'
                ExifBackups = 'backups EXIF locales'
                Steps = 'Pasos de migracion'
                Step1 = 'Copia el ZIP de instalacion compartida al nuevo PC.'
                Step2 = 'Extraelo en la carpeta donde quieras usar UDMRS.'
                Step3 = 'Copia el ZIP de estado del usuario al nuevo PC.'
                Step4 = 'Extraelo en la carpeta AppData correspondiente del usuario Windows de destino.'
                TypicalTarget = 'Destino habitual'
                Step5 = 'Inicia el dashboard desde Start-PhotoOrganizer.cmd en la instalacion extraida.'
                Step6 = 'Ejecuta primero Test Scan.'
                Step7 = 'Ejecuta ReconcileProcessedDatabase en DryRun para validar rutas y salud del indice antes de cualquier Apply.'
                Notes = 'Notas'
                NoteLogs = 'Logs y Runtime no se migran porque UDMRS los regenera automaticamente.'
                NoteSingleRun = 'Mantén una sola operacion activa del motor cuando uses una biblioteca sincronizada.'
                NoteCloud = 'Si usas OneDrive/Dropbox/iCloud/Google Drive, asegúrate de que el contenido que quieres procesar esta disponible localmente antes de ejecutar Apply.'
            }
        }
    }
}

function New-ZipFromSelectedPaths {
    param(
        [string]$ZipPath,
        [string]$BasePath,
        [string[]]$IncludeNames,
        [scriptblock]$ExcludePredicate
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zipDirectory = [System.IO.Path]::GetDirectoryName($ZipPath)
    if (-not (Test-Path -LiteralPath $zipDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $zipDirectory -Force | Out-Null
    }
    if (Test-Path -LiteralPath $ZipPath -PathType Leaf) {
        Remove-Item -LiteralPath $ZipPath -Force
    }

    $stream = [System.IO.File]::Open($ZipPath, [System.IO.FileMode]::CreateNew)
    try {
        $archive = New-Object System.IO.Compression.ZipArchive($stream, [System.IO.Compression.ZipArchiveMode]::Create)
        try {
            foreach ($name in @($IncludeNames)) {
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                $candidate = Join-Path $BasePath $name
                if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                    if ($ExcludePredicate -and (& $ExcludePredicate $candidate $false)) { continue }
                    Add-FileToZipArchive -Archive $archive -FilePath $candidate -EntryName $name
                    continue
                }
                if (-not (Test-Path -LiteralPath $candidate -PathType Container)) { continue }

                foreach ($file in @(Get-ChildItem -LiteralPath $candidate -File -Recurse -Force -ErrorAction SilentlyContinue)) {
                    if ($ExcludePredicate -and (& $ExcludePredicate $file.FullName $false)) { continue }
                    $relative = Get-PortableRelativePath -BasePath $BasePath -Path $file.FullName
                    Add-FileToZipArchive -Archive $archive -FilePath $file.FullName -EntryName $relative
                }
            }
        }
        finally {
            $archive.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function New-UDMRSMigrationPackage {
    param(
        [Parameter(Mandatory = $true)][string]$InstallRoot,
        [Parameter(Mandatory = $true)][string]$UserDataRoot,
        [string]$LocalDataRoot,
        [string]$OutputRoot,
        [string]$LanguageCode = 'es',
        [string]$PackagePrefix = 'UDMRS-Migration'
    )

    $resolvedInstallRoot = [System.IO.Path]::GetFullPath($InstallRoot)
    $resolvedUserDataRoot = [System.IO.Path]::GetFullPath($UserDataRoot)
    if (-not (Test-Path -LiteralPath $resolvedInstallRoot -PathType Container)) {
        throw "Install root not found: $resolvedInstallRoot"
    }
    if (-not (Test-Path -LiteralPath $resolvedUserDataRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $resolvedUserDataRoot -Force | Out-Null
    }
    if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
        $OutputRoot = Get-UDMRSDefaultMigrationOutputRoot
        if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
            $OutputRoot = Join-Path $resolvedUserDataRoot 'MigrationPackages'
        }
    }
    $resolvedOutputRoot = [System.IO.Path]::GetFullPath($OutputRoot)
    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $packageRoot = Join-Path $resolvedOutputRoot $stamp
    New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null
    $text = Get-UDMRSMigrationText -LanguageCode $LanguageCode

    $installLeaf = Split-Path -Leaf $resolvedInstallRoot.TrimEnd('\')
    if ([string]::IsNullOrWhiteSpace($installLeaf)) { $installLeaf = 'UDMRS' }
    $userLeaf = if ($env:USERNAME) { $env:USERNAME } else { 'CurrentUser' }

    $installZip = Join-Path $packageRoot ("{0}-{1}-{2}-{3}.zip" -f $PackagePrefix, $text.InstallZipSegment, $installLeaf, $stamp)
    $userZip = Join-Path $packageRoot ("{0}-{1}-{2}-{3}.zip" -f $PackagePrefix, $text.UserZipSegment, $userLeaf, $stamp)
    $guidePath = Join-Path $packageRoot 'MigrationGuide.txt'

    $installIncludes = @(
        'App',
        'Docs',
        'Tools',
        'Branding',
        'Templates',
        'Releases',
        'Config',
        'README.md',
        'Start-PhotoOrganizer.cmd',
        'Start-PhotoOrganizer.lnk'
    )
    $installExcludedNames = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
    foreach ($name in @('Logs', 'Runtime', 'tmp', 'temp', 'MigrationPackages', '.git', '__pycache__')) {
        [void]$installExcludedNames.Add($name)
    }
    $installExclude = {
        param([string]$Path, [bool]$IsDirectory)
        $parts = [System.IO.Path]::GetFullPath($Path).Substring($resolvedInstallRoot.Length).TrimStart('\').Split('\')
        foreach ($part in $parts) {
            if ($installExcludedNames.Contains($part)) { return $true }
            if ($part -like 'tmp-*') { return $true }
        }
        return $false
    }.GetNewClosure()

    New-ZipFromSelectedPaths -ZipPath $installZip -BasePath $resolvedInstallRoot -IncludeNames $installIncludes -ExcludePredicate $installExclude

    $userIncludes = @()
    foreach ($name in @('ProcessedFiles.json', 'settings.json', 'dashboard-settings.json', 'Config', 'IndexBackups')) {
        if (Test-Path -LiteralPath (Join-Path $resolvedUserDataRoot $name)) {
            $userIncludes += $name
        }
    }
    foreach ($file in @(Get-ChildItem -LiteralPath $resolvedUserDataRoot -File -Force -ErrorAction SilentlyContinue)) {
        if ($file.Extension -ieq '.json' -and
            $file.Name -notlike '*.progress.json' -and
            $file.Name -notmatch '(?i)(queue|status|pid|lock)') {
            if ($userIncludes -notcontains $file.Name) {
                $userIncludes += $file.Name
            }
        }
    }

    $userExclude = {
        param([string]$Path, [bool]$IsDirectory)
        $relative = Get-PortableRelativePath -BasePath $resolvedUserDataRoot -Path ([System.IO.Path]::GetFullPath($Path))
        $parts = $relative.Split('\')
        foreach ($part in $parts) {
            if ($part -in @('Logs', 'Runtime', 'MigrationPackages')) { return $true }
            if ($part -match '(?i)(queue|technical-console|\.pid$|status\.json$|\.progress\.json$)') { return $true }
        }
        return $false
    }.GetNewClosure()

    New-ZipFromSelectedPaths -ZipPath $userZip -BasePath $resolvedUserDataRoot -IncludeNames $userIncludes -ExcludePredicate $userExclude

    $guide = @"
$($text.Title)
$($text.Generated): $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

$($text.SharedZip):
$installZip

$($text.UserZip):
$userZip

$($text.SourcePaths):
$($text.InstallRoot): $resolvedInstallRoot
$($text.UserRoot): $resolvedUserDataRoot
$($text.LocalRoot): $LocalDataRoot

$($text.Included)
1. $($text.InstallIncluded)

2. $($text.UserIncluded)

$($text.NotIncluded)
- Logs
- Runtime / TechnicalConsole
- $($text.ProgressFiles)
- $($text.QueueFiles)
- $($text.StatusFiles)
- $($text.TempState)
- $($text.ExifBackups)

$($text.Steps)
1. $($text.Step1)
2. $($text.Step2)
3. $($text.Step3)
4. $($text.Step4)
   $($text.TypicalTarget): %APPDATA%\PhotoOrganizer
5. $($text.Step5)
6. $($text.Step6)
7. $($text.Step7)

$($text.Notes)
- $($text.NoteLogs)
- $($text.NoteSingleRun)
- $($text.NoteCloud)
"@
    $guide | Set-Content -LiteralPath $guidePath -Encoding UTF8

    return [pscustomobject]@{
        PackageRoot = $packageRoot
        InstallZip = $installZip
        UserStateZip = $userZip
        GuidePath = $guidePath
        InstallZipBytes = (Get-Item -LiteralPath $installZip).Length
        UserStateZipBytes = (Get-Item -LiteralPath $userZip).Length
    }
}
