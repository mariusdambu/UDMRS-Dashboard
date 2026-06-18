Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:AppRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$script:InstallRoot = if ([System.IO.Path]::GetFileName($script:AppRoot) -ieq 'App') { [System.IO.Path]::GetDirectoryName($script:AppRoot) } else { $script:AppRoot }
$script:Root = $script:AppRoot
$script:EnginePath = Join-Path $script:AppRoot 'PhotoOrganizer.ps1'
$script:LanguageRoot = Join-Path $script:AppRoot 'Languages'
$script:LanguageResourcesPath = Join-Path $script:AppRoot 'LanguageResources.json'
$script:SettingsRoot = Join-Path $env:APPDATA 'PhotoOrganizer'
$script:ConfigRoot = Join-Path $script:SettingsRoot 'Config'
$script:UserExcludedFoldersConfigPath = Join-Path $script:ConfigRoot 'UserExcludedFolders.json'
$script:LogRoot = Join-Path $script:SettingsRoot 'Logs'
$script:IndexBackupRoot = Join-Path $script:SettingsRoot 'IndexBackups'
$script:RuntimeRoot = Join-Path $script:SettingsRoot 'Runtime'
$script:DocsRoot = Join-Path $script:InstallRoot 'Docs'
$script:MigrationToolPath = Join-Path $script:AppRoot 'UDMRS-MigrationTools.ps1'
$script:SettingsPath = Join-Path $script:SettingsRoot 'settings.json'
$script:PreviousSettingsPath = Join-Path $script:SettingsRoot 'dashboard-settings.json'
$script:DashboardSettingsPath = Join-Path $script:SettingsRoot 'dashboard-settings.json'
$script:CurrentProcess = $null
$script:CurrentLogPath = $null
$script:CurrentProgressPath = $null
$script:MonitorTimer = $null
$script:TechnicalConsoleProcess = $null
$script:TechnicalConsoleRoot = Join-Path $script:RuntimeRoot 'TechnicalConsole'
$script:TechnicalConsoleQueuePath = Join-Path $script:TechnicalConsoleRoot 'queue.jsonl'
$script:TechnicalConsoleRunnerPath = Join-Path $script:TechnicalConsoleRoot 'UDMRS-TechnicalConsole.ps1'
$script:TechnicalConsolePidPath = Join-Path $script:TechnicalConsoleRoot 'technical-console.pid'
$script:TechnicalConsoleStatusPath = Join-Path $script:TechnicalConsoleRoot 'technical-console.status.json'
$script:AdvancedRunActive = $false
$script:RunWasCancelled = $false
$script:LastLogLength = 0
$script:FallbackLanguage = $null
$script:EnglishLanguage = $null
$script:CurrentLanguage = $null
$script:CurrentLanguageCode = 'es'
$script:LanguageResources = $null
$script:MaxVisibleLogLines = 5000
$script:AdvancedModeDescriptionLabels = @{}
$script:ToolTip = New-Object System.Windows.Forms.ToolTip
$script:ToolTip.AutoPopDelay = 12000
$script:ToolTip.InitialDelay = 400
$script:ToolTip.ReshowDelay = 150
$script:LogRetentionDays = 7
$script:LogRetentionSafetyHours = 1
$script:IndexBackupRetentionDays = 7
$script:IndexBackupMaxFiles = 10
$script:LastLogRetentionCleanupSummary = ''

if (Test-Path -LiteralPath $script:MigrationToolPath -PathType Leaf) {
    . $script:MigrationToolPath
}

$script:DefaultLanguageResources = @{
    es = @{
        OrganizedFolder = 'Fotos_Organizadas'
        NeedsReviewFolder = '_NecesitaRevision'
        DuplicatesFolder = '_Duplicados_Para_Revisar'
        MetadataBackupFolder = '_CopiaSeguridadMetadatos'
        LogsFolder = 'Logs'
        MediaMetadataIssuesFolder = 'MediaMetadataIssues'
    }
    ro = @{
        OrganizedFolder = 'Poze_Organizate'
        NeedsReviewFolder = '_De_Revizuit'
        DuplicatesFolder = '_Duplicate_De_Revizuit'
        MetadataBackupFolder = '_Backup_Metadate'
        LogsFolder = 'Logs'
        MediaMetadataIssuesFolder = 'MediaMetadataIssues'
    }
    en = @{
        OrganizedFolder = 'Organized_Photos'
        NeedsReviewFolder = '_NeedsReview'
        DuplicatesFolder = '_Duplicates_To_Review'
        MetadataBackupFolder = '_MetadataBackup'
        LogsFolder = 'Logs'
        MediaMetadataIssuesFolder = 'MediaMetadataIssues'
    }
}

function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    return Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Save-Settings {
    if (-not (Test-Path -LiteralPath $script:SettingsRoot)) {
        New-Item -ItemType Directory -Path $script:SettingsRoot -Force | Out-Null
    }

    [pscustomobject]@{
        Language = $script:CurrentLanguageCode
        PerformanceMode = if ($script:cmbPerformance) { [string]$script:cmbPerformance.SelectedItem } else { 'Balanced' }
        MaxParallelJobs = if ($script:txtMaxParallel) { [string]$script:txtMaxParallel.Text } else { '' }
        RenameInternalFolders = if ($script:chkRenameInternal) { [bool]$script:chkRenameInternal.Checked } else { $false }
    } | ConvertTo-Json | Set-Content -LiteralPath $script:SettingsPath -Encoding UTF8
}

function Load-LanguageResources {
    $resources = @{}
    foreach ($code in $script:DefaultLanguageResources.Keys) {
        $resources[$code] = @{}
        foreach ($key in $script:DefaultLanguageResources[$code].Keys) {
            $resources[$code][$key] = $script:DefaultLanguageResources[$code][$key]
        }
    }

    if (Test-Path -LiteralPath $script:LanguageResourcesPath -PathType Leaf) {
        try {
            $loaded = Get-Content -LiteralPath $script:LanguageResourcesPath -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($languageProperty in $loaded.PSObject.Properties) {
                $code = [string]$languageProperty.Name
                if (-not $resources.ContainsKey($code)) { $resources[$code] = @{} }
                foreach ($folderProperty in $languageProperty.Value.PSObject.Properties) {
                    $resources[$code][[string]$folderProperty.Name] = [string]$folderProperty.Value
                }
            }
        }
        catch {
        }
    }

    $script:LanguageResources = $resources
}

function Load-Language {
    param([string]$Code)

    $script:FallbackLanguage = Read-JsonFile -Path (Join-Path $script:LanguageRoot 'es.json')
    $script:EnglishLanguage = Read-JsonFile -Path (Join-Path $script:LanguageRoot 'en.json')
    $languagePath = Join-Path $script:LanguageRoot ($Code + '.json')
    $loaded = Read-JsonFile -Path $languagePath

    if ($null -eq $loaded) {
        $Code = 'es'
        $loaded = $script:FallbackLanguage
    }

    $script:CurrentLanguageCode = $Code
    $script:CurrentLanguage = $loaded
}

function T {
    param([string]$Key)

    if ($script:CurrentLanguage -and $script:CurrentLanguage.PSObject.Properties[$Key]) {
        return [string]$script:CurrentLanguage.PSObject.Properties[$Key].Value
    }

    if ($script:EnglishLanguage -and $script:EnglishLanguage.PSObject.Properties[$Key]) {
        return [string]$script:EnglishLanguage.PSObject.Properties[$Key].Value
    }

    if ($script:FallbackLanguage -and $script:FallbackLanguage.PSObject.Properties[$Key]) {
        return [string]$script:FallbackLanguage.PSObject.Properties[$Key].Value
    }

    return $Key
}

function UT {
    param([string]$Key)
    return T $Key
}

function Get-AvailableLanguages {
    $languages = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $script:LanguageRoot -PathType Container)) {
        $languages.Add([pscustomobject]@{ Code = 'es'; DisplayName = 'Español' })
        return @($languages.ToArray())
    }

    foreach ($file in @(Get-ChildItem -LiteralPath $script:LanguageRoot -Filter '*.json' -File -ErrorAction SilentlyContinue | Sort-Object BaseName)) {
        $code = [string]$file.BaseName
        $data = Read-JsonFile -Path $file.FullName
        $displayName = $code
        if ($data -and $data.PSObject.Properties['language_name'] -and -not [string]::IsNullOrWhiteSpace([string]$data.language_name)) {
            $displayName = [string]$data.language_name
        }
        $languages.Add([pscustomobject]@{ Code = $code; DisplayName = $displayName })
    }

    if ($languages.Count -eq 0) {
        $languages.Add([pscustomobject]@{ Code = 'es'; DisplayName = 'Español' })
    }

    return @($languages.ToArray())
}

function Set-LanguageComboItems {
    param([System.Windows.Forms.ComboBox]$Combo)

    $languages = @(Get-AvailableLanguages)
    $Combo.Items.Clear()
    $Combo.DisplayMember = 'DisplayName'
    $Combo.ValueMember = 'Code'

    $selectedIndex = 0
    for ($i = 0; $i -lt $languages.Count; $i++) {
        [void]$Combo.Items.Add($languages[$i])
        if ([string]$languages[$i].Code -eq $script:CurrentLanguageCode) {
            $selectedIndex = $i
        }
    }

    if ($Combo.Items.Count -gt 0) {
        $Combo.SelectedIndex = $selectedIndex
    }
}

function Get-SelectedLanguageCode {
    param([System.Windows.Forms.ComboBox]$Combo)

    if ($Combo.SelectedItem -and $Combo.SelectedItem.PSObject.Properties['Code']) {
        return [string]$Combo.SelectedItem.Code
    }

    return $script:CurrentLanguageCode
}

function Get-InternalFolderName {
    param([string]$Key)

    if (-not $script:LanguageResources) {
        Load-LanguageResources
    }

    if ($script:LanguageResources.ContainsKey($script:CurrentLanguageCode) -and $script:LanguageResources[$script:CurrentLanguageCode].ContainsKey($Key)) {
        return [string]$script:LanguageResources[$script:CurrentLanguageCode][$Key]
    }

    return [string]$script:DefaultLanguageResources.es[$Key]
}

function Get-AllInternalFolderNames {
    param([string]$Key)

    if (-not $script:LanguageResources) {
        Load-LanguageResources
    }

    $names = New-Object System.Collections.Generic.List[string]
    foreach ($code in $script:LanguageResources.Keys) {
        if ($script:LanguageResources[$code].ContainsKey($Key)) {
            $value = [string]$script:LanguageResources[$code][$Key]
            if (-not [string]::IsNullOrWhiteSpace($value) -and -not $names.Contains($value)) {
                $names.Add($value)
            }
        }
    }

    return @($names)
}

function Resolve-LocalizedDestinationPath {
    param(
        [string]$DestinationPath,
        [string]$SourcePath
    )

    $organizedFolder = Get-InternalFolderName -Key 'OrganizedFolder'
    if ([string]::IsNullOrWhiteSpace($DestinationPath)) {
        if ([string]::IsNullOrWhiteSpace($SourcePath)) { return '' }
        return Join-Path $SourcePath $organizedFolder
    }

    $trimmed = $DestinationPath.TrimEnd('\')
    $leaf = [System.IO.Path]::GetFileName($trimmed)
    if ((Get-AllInternalFolderNames -Key 'OrganizedFolder') -contains $leaf) {
        $parent = [System.IO.Path]::GetDirectoryName($trimmed)
        if (-not [string]::IsNullOrWhiteSpace($parent)) {
            return Join-Path $parent $organizedFolder
        }
    }

    return $DestinationPath
}

function Get-DefaultSourcePath {
    $preferred = Join-Path $env:USERPROFILE 'OneDrive\Imagini'
    $spanish = Join-Path $env:USERPROFILE ('OneDrive\Im' + [char]0x00E1 + 'genes')
    $pictures = [Environment]::GetFolderPath('MyPictures')

    if (Test-Path -LiteralPath $preferred -PathType Container) { return $preferred }
    if (Test-Path -LiteralPath $spanish -PathType Container) { return $spanish }
    if (-not [string]::IsNullOrWhiteSpace($pictures)) { return $pictures }
    return $preferred
}

function Get-DefaultDestinationPath {
    param([string]$SourcePath)
    if ([string]::IsNullOrWhiteSpace($SourcePath)) { return '' }
    return Join-Path $SourcePath (Get-InternalFolderName -Key 'OrganizedFolder')
}

function Get-DefaultUserExcludedFoldersConfig {
    [pscustomobject]@{
        version = 1
        userExcludedFolders = @()
        vendorPresets = @(
            [pscustomobject]@{
                path = '<SourcePath>\Samsung Gallery'
                label = 'Samsung Gallery'
                reason = 'External app managed folder'
                role = 'VendorManaged'
            },
            [pscustomobject]@{
                path = '<SourcePath>\Camera Roll'
                label = 'Camera Roll'
                reason = 'External camera upload folder'
                role = 'VendorManaged'
            }
        )
    }
}

function Read-UserExcludedFoldersConfig {
    if (Test-Path -LiteralPath $script:UserExcludedFoldersConfigPath -PathType Leaf) {
        try {
            return Get-Content -LiteralPath $script:UserExcludedFoldersConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
        }
        catch {
        }
    }

    $config = Get-DefaultUserExcludedFoldersConfig
    Save-UserExcludedFoldersConfig -Config $config
    return $config
}

function Save-UserExcludedFoldersConfig {
    param([object]$Config)

    if (-not (Test-Path -LiteralPath $script:ConfigRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $script:ConfigRoot -Force | Out-Null
    }

    $Config | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $script:UserExcludedFoldersConfigPath -Encoding UTF8
}

function Expand-DashboardConfiguredPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $expanded = [Environment]::ExpandEnvironmentVariables($Path)
    $expanded = $expanded.Replace('<SourcePath>', $script:txtSource.Text.Trim())
    $expanded = $expanded.Replace('<DestinationPath>', $script:txtDestination.Text.Trim())
    $expanded = $expanded.Replace('<AppRoot>', $script:InstallRoot)
    $expanded = $expanded.Replace('<InstallRoot>', $script:InstallRoot)
    $expanded = $expanded.Replace('<ScriptRoot>', $script:AppRoot)
    try {
        return [System.IO.Path]::GetFullPath($expanded)
    }
    catch {
        return $expanded
    }
}

function ConvertTo-PortableConfiguredPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    try {
        $fullPath = [System.IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($Path)).TrimEnd('\')
        $sourcePath = [System.IO.Path]::GetFullPath($script:txtSource.Text.Trim()).TrimEnd('\')
        $destinationPath = [System.IO.Path]::GetFullPath($script:txtDestination.Text.Trim()).TrimEnd('\')
        if ($fullPath.Equals($sourcePath, [StringComparison]::OrdinalIgnoreCase)) { return '<SourcePath>' }
        if ($fullPath.StartsWith($sourcePath + '\', [StringComparison]::OrdinalIgnoreCase)) {
            return '<SourcePath>\' + $fullPath.Substring($sourcePath.Length + 1)
        }
        if ($fullPath.Equals($destinationPath, [StringComparison]::OrdinalIgnoreCase)) { return '<DestinationPath>' }
        if ($fullPath.StartsWith($destinationPath + '\', [StringComparison]::OrdinalIgnoreCase)) {
            return '<DestinationPath>\' + $fullPath.Substring($destinationPath.Length + 1)
        }
    }
    catch {
    }

    return $Path
}

function Get-ExcludedFolderRoleDisplayText {
    param([string]$Role)
    switch ($Role) {
        'VendorManaged' { return (UT 'role_vendor_managed') }
        'UserExcluded' { return (UT 'role_user_excluded') }
        default { return $Role }
    }
}

function Get-ExcludedFolderReasonDisplayText {
    param([string]$Reason)
    switch ($Reason) {
        'External app managed folder' { return (UT 'reason_external_app_managed') }
        'External camera upload folder' { return (UT 'reason_external_camera_upload') }
        'User protected folder' { return (UT 'reason_user_protected') }
        default { return $Reason }
    }
}

function Get-ExcludedFolderConfigSummary {
    $config = Read-UserExcludedFoldersConfig
    $entries = @($config.userExcludedFolders)
    $enabled = 0
    $found = 0
    $missing = 0
    foreach ($entry in $entries) {
        if ($null -eq $entry) { continue }
        $isEnabled = $true
        if ($entry.PSObject.Properties['enabled']) { $isEnabled = [bool]$entry.enabled }
        if (-not $isEnabled) { continue }
        $enabled++
        $resolved = Expand-DashboardConfiguredPath -Path ([string]$entry.path)
        if (-not [string]::IsNullOrWhiteSpace($resolved) -and (Test-Path -LiteralPath $resolved -PathType Container)) {
            $found++
        }
        else {
            $missing++
        }
    }

    if ($enabled -eq 0) {
        return (UT 'folder_protection_summary_none')
    }

    return (UT 'folder_protection_summary') -f $enabled, $found, $missing
}

function Update-FolderProtectionSummary {
    if ($script:lblFolderProtectionSummary) {
        $script:lblFolderProtectionSummary.Text = Get-ExcludedFolderConfigSummary
    }
}

function Quote-Arg {
    param([string]$Value)
    if ($null -eq $Value) { return '""' }
    return '"' + ($Value -replace '"', '\"') + '"'
}

function Get-EngineHost {
    try {
        $pwsh = Get-Command pwsh.exe -ErrorAction Stop
        if ($pwsh -and -not [string]::IsNullOrWhiteSpace($pwsh.Source)) {
            return $pwsh.Source
        }
    }
    catch {
    }

    return 'powershell.exe'
}

function Test-EngineHostIsPowerShell7 {
    param([string]$EngineHost)
    if ([string]::IsNullOrWhiteSpace($EngineHost)) { return $false }
    return ([System.IO.Path]::GetFileName($EngineHost) -ieq 'pwsh.exe')
}

function Get-EngineRuntimeMessage {
    $engineHost = Get-EngineHost
    if (Test-EngineHostIsPowerShell7 -EngineHost $engineHost) {
        $version = '7.x'
        try {
            $detected = & $engineHost -NoProfile -Command '$PSVersionTable.PSVersion.ToString()' 2>$null
            if (-not [string]::IsNullOrWhiteSpace($detected)) { $version = [string]$detected }
        }
        catch {
        }
        return ("{0} PowerShell {1} ({2})" -f (UT 'engine_prefix'), $version, (UT 'engine_recommended'))
    }

    return ("{0} {1}" -f (UT 'engine_prefix'), (UT 'engine_ps51'))
}

function Open-PowerShell7Download {
    [System.Diagnostics.Process]::Start('https://github.com/PowerShell/PowerShell/releases') | Out-Null
}

function Get-TechnicalConsoleRunnerContent {
    return @'
param(
    [Parameter(Mandatory = $true)][string]$QueuePath,
    [Parameter(Mandatory = $true)][string]$RootPath,
    [Parameter(Mandatory = $true)][string]$StatusPath
)

$ErrorActionPreference = 'Continue'
try { $host.UI.RawUI.WindowTitle = 'UDMRS Technical Console' } catch {}

if (-not (Test-Path -LiteralPath $RootPath -PathType Container)) {
    New-Item -ItemType Directory -Path $RootPath -Force | Out-Null
}
if (-not (Test-Path -LiteralPath $QueuePath -PathType Leaf)) {
    New-Item -ItemType File -Path $QueuePath -Force | Out-Null
}

function Write-TechnicalConsoleStatus {
    param(
        [string]$State,
        [object]$Job = $null,
        [int]$ExitCode = 0
    )
    try {
        [pscustomobject]@{
            state = $State
            runId = if ($Job) { [string]$Job.RunId } else { '' }
            label = if ($Job) { [string]$Job.Label } else { '' }
            started = if ($State -eq 'Running') { (Get-Date).ToString('o') } else { '' }
            updated = (Get-Date).ToString('o')
            exitCode = $ExitCode
        } | ConvertTo-Json -Compress | Set-Content -LiteralPath $StatusPath -Encoding UTF8
    }
    catch {
    }
}

Write-TechnicalConsoleStatus -State 'Idle'

$position = 0L
Write-Host ''
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host ' UDMRS Technical Console' -ForegroundColor Cyan
Write-Host ' Persistent queue mode. Close this window only when idle.' -ForegroundColor DarkGray
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host ''

while ($true) {
    try {
        if (-not (Test-Path -LiteralPath $QueuePath -PathType Leaf)) {
            Start-Sleep -Milliseconds 800
            continue
        }

        $fs = [System.IO.File]::Open($QueuePath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            if ($fs.Length -gt $position) {
                $fs.Seek($position, [System.IO.SeekOrigin]::Begin) | Out-Null
                $reader = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
                $text = $reader.ReadToEnd()
                $position = $fs.Position
                foreach ($line in ($text -split "`r?`n")) {
                    if ([string]::IsNullOrWhiteSpace($line)) { continue }
                    try {
                        $job = $line | ConvertFrom-Json
                        $arguments = @()
                        foreach ($arg in @($job.Arguments)) { $arguments += [string]$arg }
                        Write-TechnicalConsoleStatus -State 'Running' -Job $job
                        Write-Host ''
                        Write-Host '------------------------------------------------------------' -ForegroundColor Yellow
                        Write-Host ("Action: {0}" -f $job.Label) -ForegroundColor Yellow
                        Write-Host ("RunId : {0}" -f $job.RunId) -ForegroundColor DarkGray
                        Write-Host ("Log   : {0}" -f $job.LogPath) -ForegroundColor DarkGray
                        Write-Host ("Mode  : {0}" -f $(if ($job.Apply) { 'APPLY' } else { 'DRY RUN' })) -ForegroundColor $(if ($job.Apply) { 'Red' } else { 'Green' })
                        Write-Host ("Command: {0} {1}" -f $job.EngineHost, ($arguments -join ' ')) -ForegroundColor Gray
                        Write-Host '------------------------------------------------------------' -ForegroundColor Yellow
                        & ([string]$job.EngineHost) @arguments
                        $exitCode = if ($null -ne $global:LASTEXITCODE) { [int]$global:LASTEXITCODE } else { 0 }
                        Write-Host ''
                        Write-Host ("Completed: {0}; exit code: {1}" -f $job.Label, $exitCode) -ForegroundColor Cyan
                        Write-Host 'Window stays open. Launch another advanced action from the dashboard to reuse this console.' -ForegroundColor DarkGray
                        Write-TechnicalConsoleStatus -State 'Idle' -Job $job -ExitCode $exitCode
                    }
                    catch {
                        Write-Host ("Technical console job failed: {0}" -f $_.Exception.Message) -ForegroundColor Red
                        try { Write-TechnicalConsoleStatus -State 'Idle' -Job $job -ExitCode 1 } catch {}
                    }
                }
            }
        }
        finally {
            $fs.Close()
        }
    }
    catch {
        Write-Host ("Technical console loop warning: {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow
    }

    Start-Sleep -Milliseconds 800
}
'@
}

function Get-TechnicalConsoleProcess {
    if ($script:TechnicalConsoleProcess -and -not $script:TechnicalConsoleProcess.HasExited) {
        return $script:TechnicalConsoleProcess
    }

    if (Test-Path -LiteralPath $script:TechnicalConsolePidPath -PathType Leaf) {
        try {
            $pidText = Get-Content -LiteralPath $script:TechnicalConsolePidPath -Raw -ErrorAction Stop
            $pid = 0
            if ([int]::TryParse($pidText.Trim(), [ref]$pid) -and $pid -gt 0) {
                $existing = Get-Process -Id $pid -ErrorAction SilentlyContinue
                if ($existing) {
                    $commandLine = ''
                    try {
                        $cimProcess = Get-CimInstance Win32_Process -Filter "ProcessId=$pid" -ErrorAction SilentlyContinue
                        if ($cimProcess) { $commandLine = [string]$cimProcess.CommandLine }
                    }
                    catch {
                    }
                    if ([string]::IsNullOrWhiteSpace($commandLine) -or $commandLine.IndexOf($script:TechnicalConsoleRunnerPath, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                        $script:TechnicalConsoleProcess = $existing
                        return $existing
                    }
                }
            }
        }
        catch {
        }
    }

    return $null
}

function Stop-ControlledProcessTree {
    param([int]$ProcessId)

    if ($ProcessId -le 0) { return }
    try {
        $children = @(Get-CimInstance Win32_Process -Filter "ParentProcessId=$ProcessId" -ErrorAction SilentlyContinue)
        foreach ($child in $children) {
            Stop-ControlledProcessTree -ProcessId ([int]$child.ProcessId)
        }
    }
    catch {
    }

    try {
        Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
    }
    catch {
    }
}

function Ensure-TechnicalConsole {
    if (-not (Test-Path -LiteralPath $script:TechnicalConsoleRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $script:TechnicalConsoleRoot -Force | Out-Null
    }

    $runnerContent = Get-TechnicalConsoleRunnerContent
    Set-Content -LiteralPath $script:TechnicalConsoleRunnerPath -Value $runnerContent -Encoding UTF8

    $existing = Get-TechnicalConsoleProcess
    if ($existing) { return $existing }

    Set-Content -LiteralPath $script:TechnicalConsoleQueuePath -Value '' -Encoding UTF8

    $engineHost = Get-EngineHost
    $arguments = @(
        '-NoExit',
        '-ExecutionPolicy', 'Bypass',
        '-NoProfile',
        '-File', (Quote-Arg $script:TechnicalConsoleRunnerPath),
        '-QueuePath', (Quote-Arg $script:TechnicalConsoleQueuePath),
        '-RootPath', (Quote-Arg $script:TechnicalConsoleRoot),
        '-StatusPath', (Quote-Arg $script:TechnicalConsoleStatusPath)
    )

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $engineHost
    $psi.Arguments = ($arguments -join ' ')
    $psi.UseShellExecute = $true
    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Normal

    $process = [System.Diagnostics.Process]::Start($psi)
    $script:TechnicalConsoleProcess = $process
    Set-Content -LiteralPath $script:TechnicalConsolePidPath -Value ([string]$process.Id) -Encoding ASCII
    Start-Sleep -Milliseconds 300
    return $process
}

function Test-AdvancedRunActive {
    $process = Get-TechnicalConsoleProcess
    if (-not $process) {
        $script:AdvancedRunActive = $false
        return $false
    }

    if (Test-Path -LiteralPath $script:TechnicalConsoleStatusPath -PathType Leaf) {
        try {
            $status = Get-Content -LiteralPath $script:TechnicalConsoleStatusPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json
            if ($status -and [string]$status.state -eq 'Running') {
                $script:AdvancedRunActive = $true
                return $true
            }
        }
        catch {
        }
    }

    $script:AdvancedRunActive = $false
    return $false
}

function Clear-TechnicalConsoleState {
    foreach ($path in @($script:TechnicalConsoleStatusPath, $script:TechnicalConsolePidPath, $script:TechnicalConsoleQueuePath)) {
        try {
            if (Test-Path -LiteralPath $path -PathType Leaf) {
                Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
        }
    }
    $script:AdvancedRunActive = $false
}

function Stop-TechnicalConsoleSession {
    $process = Get-TechnicalConsoleProcess
    if ($process) {
        try {
            Stop-ControlledProcessTree -ProcessId ([int]$process.Id)
        }
        catch {
        }
    }

    $script:TechnicalConsoleProcess = $null
    Clear-TechnicalConsoleState
}


function Get-ExternalEngineRunForCurrentGallery {
    $source = $script:txtSource.Text.Trim()
    $destination = $script:txtDestination.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($source) -and [string]::IsNullOrWhiteSpace($destination)) { return $null }

    $ownPid = [System.Diagnostics.Process]::GetCurrentProcess().Id
    $needles = @()
    foreach ($path in @($source, $destination)) {
        if (-not [string]::IsNullOrWhiteSpace($path)) {
            try { $needles += [System.IO.Path]::GetFullPath($path).TrimEnd('\').ToLowerInvariant() }
            catch { $needles += $path.TrimEnd('\').ToLowerInvariant() }
        }
    }

    try {
        $processes = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {
            $_.ProcessId -ne $ownPid -and
            $_.CommandLine -and
            $_.CommandLine -match 'PhotoOrganizer\.ps1'
        }
        foreach ($proc in $processes) {
            $cmd = $proc.CommandLine.ToLowerInvariant()
            foreach ($needle in $needles) {
                if ($needle -and $cmd.Contains($needle)) { return $proc }
            }
        }
    }
    catch {
    }
    return $null
}
function Assert-CanStartNormalEngineRun {
    if (Get-ExternalEngineRunForCurrentGallery) {
        [System.Windows.Forms.MessageBox]::Show((UT 'external_run_active_msg'), (T 'warning_title'), 'OK', 'Warning') | Out-Null
        return $false
    }
    if (Test-AdvancedRunActive) {
        [System.Windows.Forms.MessageBox]::Show((UT 'advanced_run_active_msg'), (T 'warning_title'), 'OK', 'Warning') | Out-Null
        return $false
    }
    return $true
}

function Assert-CanStartAdvancedEngineRun {
    if (Get-ExternalEngineRunForCurrentGallery) {
        [System.Windows.Forms.MessageBox]::Show((UT 'external_run_active_msg'), (T 'warning_title'), 'OK', 'Warning') | Out-Null
        return $false
    }
    if ($script:CurrentProcess -and -not $script:CurrentProcess.HasExited) {
        [System.Windows.Forms.MessageBox]::Show((UT 'normal_run_active_msg'), (T 'warning_title'), 'OK', 'Warning') | Out-Null
        return $false
    }
    if (Test-AdvancedRunActive) {
        [System.Windows.Forms.MessageBox]::Show((UT 'advanced_run_active_msg'), (T 'warning_title'), 'OK', 'Warning') | Out-Null
        return $false
    }
    return $true
}

function Get-AutoWorkerCount {
    param([string]$Mode)

    $threads = [math]::Max(1, [Environment]::ProcessorCount)
    $pct = switch ($Mode) {
        'Safe' { 0.25 }
        'HighPerformance' { 0.75 }
        default { 0.50 }
    }

    $workers = [int][math]::Floor($threads * $pct)
    return [math]::Min($threads, [math]::Max(1, $workers))
}

function Update-MaxWorkerDisplay {
    if (-not $script:cmbPerformance -or -not $script:txtMaxParallel) { return }
    $script:txtMaxParallel.Text = [string](Get-AutoWorkerCount -Mode ([string]$script:cmbPerformance.SelectedItem))
    Update-ModeSummary
}

function Set-ControlTooltip {
    param(
        [System.Windows.Forms.Control]$Control,
        [string]$Text
    )
    if ($script:ToolTip -and $Control) {
        $script:ToolTip.SetToolTip($Control, $Text)
    }
}

function Set-LogTextBoxStyle {
    param([System.Windows.Forms.TextBox]$TextBox)

    if (-not $TextBox) { return }
    $TextBox.Multiline = $true
    $TextBox.ReadOnly = $true
    $TextBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $TextBox.WordWrap = $false
    $TextBox.AcceptsReturn = $true
    $TextBox.AcceptsTab = $false
    $TextBox.HideSelection = $false
    $TextBox.Font = New-Object System.Drawing.Font('Consolas', 9)
}

function Update-ExecuteHint {
    if (-not $script:lblExecuteHint) { return }

    $selectedAction = if ([string]::IsNullOrWhiteSpace([string]$script:SelectedActionKey)) { 'execute' } else { [string]$script:SelectedActionKey }
    if ($selectedAction -eq 'test_scan') {
        $script:lblExecuteHint.Text = (UT 'action_selected') + [Environment]::NewLine + (UT 'action_hint_test_scan')
        return
    }
    if ($selectedAction -eq 'reconcile') {
        $script:lblExecuteHint.Text = (UT 'action_selected') + [Environment]::NewLine + (UT 'action_hint_reconcile')
        return
    }
    if ($selectedAction -eq 'purge_missing') {
        $script:lblExecuteHint.Text = (UT 'action_selected') + [Environment]::NewLine + (UT 'action_hint_purge_missing')
        return
    }
    if (($selectedAction -like 'advanced_*') -or ($selectedAction -like 'import_*')) {
        $hintKey = 'action_hint_' + $selectedAction
        $script:lblExecuteHint.Text = (UT 'action_selected') + [Environment]::NewLine + (UT $hintKey)
        return
    }

    $isApply = ($script:chkApply -and $script:chkApply.Checked -and -not ($script:chkDryRun -and $script:chkDryRun.Checked))
    $repairExif = ($script:chkRepairExif -and $script:chkRepairExif.Checked)
    $copyMode = ($script:chkCopyInstead -and $script:chkCopyInstead.Checked)

    if (-not $isApply -and -not $repairExif) {
        $text = UT 'hint_dryrun'
    }
    elseif (-not $isApply -and $repairExif) {
        $text = UT 'hint_dryrun_exif'
    }
    elseif ($isApply -and -not $repairExif) {
        $text = UT 'hint_apply'
    }
    else {
        $text = UT 'hint_apply_exif'
    }

    if ($copyMode) {
        $text += [Environment]::NewLine + (UT 'hint_copy')
    }

    $script:lblExecuteHint.Text = (UT 'action_selected') + [Environment]::NewLine + $text
}

function Set-SelectedActionHint {
    param([string]$ActionKey)

    $script:SelectedActionKey = $ActionKey
    Update-ExecuteHint
}

function Update-ReconcileHint {
    Update-ExecuteHint
}

function Update-ModeSummary {
    if (-not $script:txtModeSummary) { return }

    $isApply = ($script:chkApply -and $script:chkApply.Checked -and -not ($script:chkDryRun -and $script:chkDryRun.Checked))
    $modeText = if ($isApply) { UT 'mode_apply' } else { UT 'mode_dryrun' }
    $repairText = if ($script:chkRepairExif -and $script:chkRepairExif.Checked) { UT 'yes' } else { UT 'no' }
    $actionText = if ($script:chkCopyInstead -and $script:chkCopyInstead.Checked) { UT 'action_copy' } else { UT 'action_move' }
    $workersText = if ($script:txtMaxParallel -and -not [string]::IsNullOrWhiteSpace($script:txtMaxParallel.Text)) {
        $script:txtMaxParallel.Text
    }
    elseif ($script:cmbPerformance) {
        [string](Get-AutoWorkerCount -Mode ([string]$script:cmbPerformance.SelectedItem))
    }
    else {
        ''
    }
    $engineText = (Get-EngineRuntimeMessage) -replace ('^{0}\s*' -f [regex]::Escape((UT 'engine_prefix'))), ''

    $script:txtModeSummary.Text = @(
        (UT 'mode_summary_title'),
        ("- {0}: {1}" -f (T 'status'), $modeText),
        ("- {0}: {1}" -f (T 'btn_run'), (UT 'organize_yes')),
        ("- {0}: {1}" -f (UT 'repair_exif'), $repairText),
        ("- {0}: {1}" -f (T 'phase'), $actionText),
        ("- PowerShell: {0}" -f $engineText),
        ("- {0}: {1}" -f (UT 'workers'), $workersText)
    ) -join [Environment]::NewLine

    Update-ExecuteHint
    Update-ReconcileHint
}

function Toggle-AdvancedTools {
    if (-not $script:advancedToolsGroup -or -not $script:btnToggleAdvancedTools) { return }
    $script:advancedToolsGroup.Visible = -not $script:advancedToolsGroup.Visible
    $script:btnToggleAdvancedTools.Text = if ($script:advancedToolsGroup.Visible) { UT 'toggle_advanced_hide' } else { UT 'toggle_advanced_show' }
}

function Toggle-AdvancedModes {
    if (-not $script:advancedModesGroup -or -not $script:btnAdvancedModes) { return }
    $script:advancedModesGroup.Visible = -not $script:advancedModesGroup.Visible
    $script:btnAdvancedModes.Text = if ($script:advancedModesGroup.Visible) { UT 'advanced_modes_hide' } else { UT 'advanced_modes_show' }
}

function Append-LogLine {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($script:CurrentLogPath)) { return }

    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            Add-Content -LiteralPath $script:CurrentLogPath -Value $Line -Encoding UTF8 -ErrorAction Stop
            return
        }
        catch {
            if ($attempt -lt 3) {
                Start-Sleep -Milliseconds 120
            }
        }
    }
}

function Open-Folder {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show((T 'folder_not_found'), (T 'error_title'), 'OK', 'Warning') | Out-Null
        return
    }

    [System.Diagnostics.Process]::Start('explorer.exe', (Quote-Arg $Path)) | Out-Null
}

function Open-HelpManual {
    $manualName = 'Manual_{0}.md' -f $script:CurrentLanguageCode.ToUpperInvariant()
    $manualPath = Join-Path (Join-Path $script:DocsRoot 'Manuals') $manualName
    if (-not (Test-Path -LiteralPath $manualPath -PathType Leaf)) {
        $manualPath = Join-Path (Join-Path $script:DocsRoot 'Manuals') 'Manual_ES.md'
    }
    if (-not (Test-Path -LiteralPath $manualPath -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show((UT 'help_missing'), (T 'error_title'), 'OK', 'Warning') | Out-Null
        return
    }

    try {
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $manualPath
        $startInfo.UseShellExecute = $true
        [System.Diagnostics.Process]::Start($startInfo) | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, (T 'error_title'), 'OK', 'Warning') | Out-Null
    }
}

function Get-QuotedCommandArgumentValue {
    param(
        [string]$CommandLine,
        [string]$Name
    )

    if ([string]::IsNullOrWhiteSpace($CommandLine) -or [string]::IsNullOrWhiteSpace($Name)) {
        return ''
    }

    $pattern = '(?i)(?:^|\s)-' + [regex]::Escape($Name) + '\s+(?:"([^"]+)"|([^\s]+))'
    $match = [regex]::Match($CommandLine, $pattern)
    if (-not $match.Success) {
        return ''
    }

    if ($match.Groups[1].Success) {
        return [string]$match.Groups[1].Value
    }
    return [string]$match.Groups[2].Value
}

function Add-ProtectedLogPath {
    param(
        [System.Collections.Generic.HashSet[string]]$ProtectedPaths,
        [string]$Path
    )

    if (-not $ProtectedPaths -or [string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    try {
        $resolved = [System.IO.Path]::GetFullPath($Path)
        [void]$ProtectedPaths.Add($resolved)
    }
    catch {
    }
}

function Get-ProtectedLogRetentionPaths {
    $protected = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)

    Add-ProtectedLogPath -ProtectedPaths $protected -Path $script:CurrentLogPath
    Add-ProtectedLogPath -ProtectedPaths $protected -Path $script:CurrentProgressPath

    if (Test-Path -LiteralPath $script:TechnicalConsoleStatusPath -PathType Leaf) {
        try {
            $status = Get-Content -LiteralPath $script:TechnicalConsoleStatusPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json
            if ($status -and [string]$status.state -eq 'Running' -and -not [string]::IsNullOrWhiteSpace([string]$status.runId)) {
                $runId = [string]$status.runId
                Add-ProtectedLogPath -ProtectedPaths $protected -Path (Join-Path $script:LogRoot ("PhotoOrganizer-$runId.log"))
                Add-ProtectedLogPath -ProtectedPaths $protected -Path (Join-Path $script:LogRoot ("PhotoOrganizer-$runId.progress.json"))
            }
        }
        catch {
        }
    }

    try {
        $processes = @(Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {
                $_.CommandLine -and
                $_.CommandLine.IndexOf('PhotoOrganizer.ps1', [StringComparison]::OrdinalIgnoreCase) -ge 0 -and
                $_.CommandLine.IndexOf('-LogPath', [StringComparison]::OrdinalIgnoreCase) -ge 0
            })
        foreach ($process in $processes) {
            $commandLine = [string]$process.CommandLine
            Add-ProtectedLogPath -ProtectedPaths $protected -Path (Get-QuotedCommandArgumentValue -CommandLine $commandLine -Name 'LogPath')
            Add-ProtectedLogPath -ProtectedPaths $protected -Path (Get-QuotedCommandArgumentValue -CommandLine $commandLine -Name 'ProgressPath')
        }
    }
    catch {
    }

    return $protected
}

function Test-LogRetentionCandidate {
    param(
        [System.IO.FileInfo]$File,
        [datetime]$Cutoff,
        [datetime]$SafetyCutoff,
        [System.Collections.Generic.HashSet[string]]$ProtectedPaths
    )

    if (-not $File) { return $false }
    if ($File.LastWriteTime -ge $Cutoff) { return $false }
    if ($File.LastWriteTime -ge $SafetyCutoff) { return $false }

    try {
        if ($ProtectedPaths -and $ProtectedPaths.Contains([System.IO.Path]::GetFullPath($File.FullName))) {
            return $false
        }
    }
    catch {
        return $false
    }

    $name = [string]$File.Name
    $extension = [string]$File.Extension
    if ($extension -ieq '.log') { return $true }
    if ($extension -ieq '.html') { return $true }
    if ($name -like '*.progress.json') { return $true }

    return $false
}

function Remove-OldFileQuietly {
    param([System.IO.FileInfo]$File)

    if (-not $File) {
        return [int64]0
    }

    $bytes = [int64]$File.Length
    try {
        Remove-Item -LiteralPath $File.FullName -Force -ErrorAction Stop
        return $bytes
    }
    catch {
        return [int64]0
    }
}

function Get-UniqueDashboardBackupTargetPath {
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

function Get-DashboardRelativePathCompat {
    param(
        [string]$Path,
        [string]$BasePath
    )

    try {
        $fullPath = ([System.IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($Path))).TrimEnd('\')
        $fullBase = ([System.IO.Path]::GetFullPath([Environment]::ExpandEnvironmentVariables($BasePath))).TrimEnd('\')
        if ($fullPath.Equals($fullBase, [StringComparison]::OrdinalIgnoreCase)) { return '' }
        if ($fullPath.StartsWith($fullBase + '\', [StringComparison]::OrdinalIgnoreCase)) {
            return $fullPath.Substring($fullBase.Length + 1)
        }
        return $fullPath
    }
    catch {
        return $Path
    }
}
function Move-LegacyJsonBackupsToIndexBackups {
    $legacyRoot = Join-Path $script:LogRoot 'JsonBackups'
    if (-not (Test-Path -LiteralPath $legacyRoot -PathType Container)) {
        return 0
    }

    $moved = 0
    try {
        if (-not (Test-Path -LiteralPath $script:IndexBackupRoot -PathType Container)) {
            New-Item -ItemType Directory -Path $script:IndexBackupRoot -Force | Out-Null
        }

        foreach ($file in @(Get-ChildItem -LiteralPath $legacyRoot -File -Recurse -Force -ErrorAction SilentlyContinue)) {
            try {
                $relative = Get-DashboardRelativePathCompat -Path $file.FullName -BasePath $legacyRoot
                $target = Join-Path $script:IndexBackupRoot $relative
                $targetDirectory = [System.IO.Path]::GetDirectoryName($target)
                if (-not (Test-Path -LiteralPath $targetDirectory -PathType Container)) {
                    New-Item -ItemType Directory -Path $targetDirectory -Force | Out-Null
                }
                Move-Item -LiteralPath $file.FullName -Destination (Get-UniqueDashboardBackupTargetPath -Path $target) -Force
                $moved++
            }
            catch {
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
    }
    catch {
    }

    return $moved
}

function Invoke-IndexBackupRetentionCleanup {
    if ([string]::IsNullOrWhiteSpace($script:IndexBackupRoot) -or -not (Test-Path -LiteralPath $script:IndexBackupRoot -PathType Container)) {
        return [pscustomobject]@{ Deleted = 0; RecoveredBytes = [int64]0 }
    }

    $deleted = 0
    $recoveredBytes = [int64]0
    try {
        $files = @(Get-ChildItem -LiteralPath $script:IndexBackupRoot -File -Recurse -Force -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
        if ($files.Count -gt 1) {
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

            foreach ($file in $files) {
                if ($protected.Contains($file.FullName)) { continue }
                $bytes = Remove-OldFileQuietly -File $file
                if ($bytes -gt 0) {
                    $deleted++
                    $recoveredBytes += $bytes
                }
            }
        }

        foreach ($directory in @(Get-ChildItem -LiteralPath $script:IndexBackupRoot -Directory -Recurse -Force -ErrorAction SilentlyContinue | Sort-Object FullName -Descending)) {
            try {
                if (@(Get-ChildItem -LiteralPath $directory.FullName -Force -ErrorAction Stop).Count -eq 0) {
                    Remove-Item -LiteralPath $directory.FullName -Force -ErrorAction Stop
                }
            }
            catch {
            }
        }
    }
    catch {
    }

    return [pscustomobject]@{ Deleted = $deleted; RecoveredBytes = $recoveredBytes }
}

function Invoke-DashboardRetentionCleanup {
    $deleted = 0
    $recoveredBytes = [int64]0
    $now = Get-Date
    $cutoff = $now.AddDays(-1 * [math]::Max(1, $script:LogRetentionDays))
    $safetyCutoff = $now.AddHours(-1 * [math]::Max(1, $script:LogRetentionSafetyHours))
    $protected = Get-ProtectedLogRetentionPaths
    $migratedIndexBackups = Move-LegacyJsonBackupsToIndexBackups

    if (-not [string]::IsNullOrWhiteSpace($script:LogRoot) -and (Test-Path -LiteralPath $script:LogRoot -PathType Container)) {
        try {
            $files = @(Get-ChildItem -LiteralPath $script:LogRoot -File -Recurse -Force -ErrorAction SilentlyContinue)
            foreach ($file in $files) {
                if (Test-LogRetentionCandidate -File $file -Cutoff $cutoff -SafetyCutoff $safetyCutoff -ProtectedPaths $protected) {
                    $bytes = Remove-OldFileQuietly -File $file
                    if ($bytes -gt 0) {
                        $deleted++
                        $recoveredBytes += $bytes
                    }
                }
            }
        }
        catch {
        }
    }

    $indexResult = Invoke-IndexBackupRetentionCleanup
    $deleted += [int]$indexResult.Deleted
    $recoveredBytes += [int64]$indexResult.RecoveredBytes

    $recoveredMb = [math]::Round($recoveredBytes / 1MB, 2)
    $script:LastLogRetentionCleanupSummary = (UT 'dashboard_retention_cleanup_completed') -f $deleted, $recoveredMb, $migratedIndexBackups
    return [pscustomobject]@{
        Deleted = $deleted
        RecoveredMB = $recoveredMb
        MigratedIndexBackups = $migratedIndexBackups
    }
}

function Get-DestinationBase {
    param([string]$DestinationPath)
    if ([string]::IsNullOrWhiteSpace($DestinationPath)) { return '' }
    return [System.IO.Path]::GetDirectoryName($DestinationPath.TrimEnd('\'))
}

function Update-RunState {
    param([bool]$Running)

    $advancedActive = Test-AdvancedRunActive
    $blockNormalLaunch = $Running -or $advancedActive
    $blockAdvancedLaunch = $Running -or $advancedActive

    $script:btnDryRun.Enabled = -not $blockNormalLaunch
    $script:btnRun.Enabled = -not $blockNormalLaunch
    $script:btnRunExif.Enabled = -not $blockNormalLaunch
    if ($script:btnExecute) { $script:btnExecute.Enabled = -not $blockNormalLaunch }
    $script:btnTestScan.Enabled = -not $blockNormalLaunch
    if ($script:btnReconcile) { $script:btnReconcile.Enabled = -not $blockNormalLaunch }
    if ($script:btnPurgeMissing) { $script:btnPurgeMissing.Enabled = -not $blockNormalLaunch }
    foreach ($button in @($script:btnAdvancedModes, $script:btnRetentionCleanup, $script:btnRecoverWrongDuplicateMove, $script:btnRenameExistingFolders, $script:btnRenameInternalFolders, $script:btnNormalizeExistingFolders, $script:btnDedupeCleanup, $script:btnRepairOnlyLibrary, $script:btnMetadataAudit, $script:btnMetadataRepair, $script:btnMigrateUdmrs, $script:btnImportGoogleTakeout, $script:btnImportApplePhotos, $script:btnImportSamsungGallery, $script:btnImportImmich, $script:btnImportXmpSidecar)) {
        if ($button) { $button.Enabled = -not $blockAdvancedLaunch }
    }
    $script:btnCancel.Enabled = $Running
    $script:btnLogs.Enabled = $true
    $script:btnSettings.Enabled = -not $Running
    $script:btnExit.Enabled = -not $Running
    $script:btnBrowseSource.Enabled = -not $Running
    $script:btnBrowseDestination.Enabled = -not $Running
    if ($script:cmbPerformance) { $script:cmbPerformance.Enabled = -not $Running }
    if ($script:txtMaxParallel) { $script:txtMaxParallel.Enabled = -not $Running }
    if ($script:chkRenameInternal) { $script:chkRenameInternal.Enabled = -not $Running }
    $script:progress.Style = if ($Running) { 'Marquee' } else { 'Blocks' }
    $script:progress.Value = if ($Running) { 0 } else { 100 }

    if ($advancedActive -and -not $Running -and $script:txtLastMessage) {
        $script:txtLastMessage.Text = UT 'advanced_run_active_msg'
    }
}

function Update-DashboardFromProgress {
    if ([string]::IsNullOrWhiteSpace($script:CurrentProgressPath) -or -not (Test-Path -LiteralPath $script:CurrentProgressPath)) {
        return
    }

    $progress = $null
    for ($attempt = 1; $attempt -le 3; $attempt++) {
        try {
            $progress = Get-Content -LiteralPath $script:CurrentProgressPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            break
        }
        catch {
            if ($attempt -lt 3) {
                Start-Sleep -Milliseconds 100
            }
        }
    }

    if ($null -eq $progress) {
        return
    }

    if ($progress.pid) { $script:txtPid.Text = [string]$progress.pid }
    if ($progress.status) { $script:lblStatusValue.Text = [string]$progress.status }
    if ($progress.phase) { $script:txtPhase.Text = [string]$progress.phase }
    if ($progress.message) { $script:txtLastMessage.Text = [string]$progress.message }
    if ($progress.filesAnalyzed -ne $null) { $script:txtAnalyzed.Text = [string]$progress.filesAnalyzed }
    if ($progress.duplicatesFound -ne $null) { $script:txtDuplicates.Text = [string]$progress.duplicatesFound }
    if ($progress.exifRepaired -ne $null) { $script:txtExif.Text = [string]$progress.exifRepaired }
    if ($progress.filesMoved -ne $null) { $script:txtMoved.Text = [string]$progress.filesMoved }
    if ($progress.filesCopied -ne $null) { $script:txtCopied.Text = [string]$progress.filesCopied }
    if ($progress.needsReview -ne $null) { $script:txtNeedsReview.Text = [string]$progress.needsReview }
    if ($progress.emptyFoldersRemoved -ne $null) { $script:txtEmptyRemoved.Text = [string]$progress.emptyFoldersRemoved }
    if ($progress.filesPerSecond -ne $null) { $script:txtFilesPerSecond.Text = ([double]$progress.filesPerSecond).ToString('0.00') }
    if ($progress.eta) { $script:txtEta.Text = [string]$progress.eta }
    if ($progress.elapsedSeconds -ne $null) { $script:txtElapsed.Text = ([TimeSpan]::FromSeconds([int]$progress.elapsedSeconds)).ToString() }
    if ($progress.currentBatch -ne $null -and $progress.totalBatches -ne $null) { $script:txtBatch.Text = ("{0}/{1}" -f $progress.currentBatch, $progress.totalBatches) }
    if ($progress.cpuPercent -ne $null) { $script:txtCpu.Text = ([int]$progress.cpuPercent).ToString() + '%' }
    if ($progress.ramPercent -ne $null) { $script:txtRam.Text = ([int]$progress.ramPercent).ToString() + '%' }
    if ($progress.activeWorkers -ne $null -and $progress.workerCount -ne $null) { $script:txtWorkers.Text = ("{0}/{1}" -f $progress.activeWorkers, $progress.workerCount) }
    if ($progress.queueSize -ne $null) { $script:txtQueue.Text = [string]$progress.queueSize }
    $skippedTotal = 0
    if ($progress.skippedOneDrive -ne $null) { $skippedTotal += [int]$progress.skippedOneDrive }
    if ($progress.incrementalSkipped -ne $null) { $skippedTotal += [int]$progress.incrementalSkipped }
    $script:txtSkipped.Text = [string]$skippedTotal
    if ($progress.inaccessible -ne $null) { $script:txtInaccessible.Text = [string]$progress.inaccessible }
    if ($progress.lockedFiles -ne $null) { $script:txtLocked.Text = [string]$progress.lockedFiles }
    if ($progress.retryCount -ne $null) { $script:txtRetries.Text = [string]$progress.retryCount }
    if ($progress.metadataCorruptedMedia -ne $null) { $script:txtMetadataCorrupt.Text = [string]$progress.metadataCorruptedMedia }
    if ($progress.metadataBackupSizeGb -ne $null) { $script:txtMetadataBackupSize.Text = ([double]$progress.metadataBackupSizeGb).ToString('0.000') }
}

function Update-LogTail {
    if ([string]::IsNullOrWhiteSpace($script:CurrentLogPath) -or -not (Test-Path -LiteralPath $script:CurrentLogPath)) {
        return
    }

    try {
        $file = Get-Item -LiteralPath $script:CurrentLogPath
        if ($file.Length -lt $script:LastLogLength) {
            $script:LastLogLength = 0
        }
        if ($file.Length -eq $script:LastLogLength) {
            return
        }

        $stream = [System.IO.File]::Open($script:CurrentLogPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            $stream.Seek($script:LastLogLength, [System.IO.SeekOrigin]::Begin) | Out-Null
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
            $newText = $reader.ReadToEnd()
            $script:LastLogLength = $stream.Position
        }
        finally {
            $stream.Dispose()
        }

        if (-not [string]::IsNullOrWhiteSpace($newText)) {
            $script:txtSummary.AppendText($newText)
            if ($script:txtSummary.Lines.Count -gt $script:MaxVisibleLogLines) {
                $script:txtSummary.Lines = @($script:txtSummary.Lines | Select-Object -Last $script:MaxVisibleLogLines)
                $script:txtSummary.SelectionStart = $script:txtSummary.TextLength
                $script:txtSummary.ScrollToCaret()
            }
        }
    }
    catch {
    }
}

function Complete-Run {
    param(
        [string]$Status,
        [string]$Message
    )

    if ($script:MonitorTimer) {
        $script:MonitorTimer.Stop()
    }
    Update-DashboardFromProgress
    Update-LogTail
    Update-RunState -Running $false
    $script:lblStatusValue.Text = $Status
    $script:txtLastMessage.Text = $Message

    if (-not [string]::IsNullOrWhiteSpace($script:CurrentLogPath)) {
        Append-LogLine -Line ("Status: " + $Status)
    }

    if ($script:chkOpenOutput.Checked) { Open-Folder -Path $script:txtDestination.Text.Trim() }
    $base = Get-DestinationBase -DestinationPath $script:txtDestination.Text.Trim()
    if ($script:chkOpenNeeds.Checked) { Open-Folder -Path (Join-Path $base (Get-InternalFolderName -Key 'NeedsReviewFolder')) }
    if ($script:chkOpenDuplicates.Checked) { Open-Folder -Path (Join-Path $base (Get-InternalFolderName -Key 'DuplicatesFolder')) }
    if ($script:chkOpenLog.Checked) { Open-Folder -Path $script:LogRoot }

    Update-RunState -Running $false
}

function Monitor-Run {
    Update-RunState -Running ($script:CurrentProcess -ne $null -and -not $script:CurrentProcess.HasExited)
    Update-DashboardFromProgress
    Update-LogTail

    if ($script:CurrentProcess -eq $null) {
        return
    }

    try {
        $hasExited = $script:CurrentProcess.HasExited
    }
    catch {
        Complete-Run -Status (T 'failed') -Message (T 'process_disappeared')
        return
    }

    if ($hasExited) {
        $exitCode = $script:CurrentProcess.ExitCode
        if ($script:RunWasCancelled) {
            Complete-Run -Status (T 'cancelled') -Message (T 'cancelled_message')
        }
        elseif ($exitCode -eq 0) {
            Complete-Run -Status (T 'completed') -Message (T 'completed_message')
        }
        else {
            Complete-Run -Status (T 'failed') -Message ((T 'failed_message') + " $exitCode")
        }
    }
}

function Stop-OrganizerRun {
    if ($script:CurrentProcess -eq $null -or $script:CurrentProcess.HasExited) {
        return
    }

    $script:RunWasCancelled = $true
    Append-LogLine -Line 'Cancel requested by user.'
    try {
        $childProcesses = Get-CimInstance Win32_Process -Filter "ParentProcessId=$($script:CurrentProcess.Id)" -ErrorAction SilentlyContinue
        foreach ($child in $childProcesses) {
            try {
                Stop-Process -Id $child.ProcessId -Force -ErrorAction SilentlyContinue
            }
            catch {
            }
        }
    }
    catch {
    }

    try {
        Stop-Process -Id $script:CurrentProcess.Id -Force -ErrorAction SilentlyContinue
        Append-LogLine -Line 'Process stopped.'
        Append-LogLine -Line 'Status: Cancelled.'
    }
    catch {
        Append-LogLine -Line ("Cancel error: " + $_.Exception.Message)
    }

    Complete-Run -Status (T 'cancelled') -Message (T 'cancelled_message')
}

function Set-LastMessage {
    param([string]$Message)
    if ($script:form.IsHandleCreated) {
        $script:form.BeginInvoke([Action]{
            $script:txtLastMessage.Text = $Message
        }) | Out-Null
    }
}

function Process-OutputLine {
    param([string]$Line)

    if ([string]::IsNullOrWhiteSpace($Line)) { return }
    Append-LogLine -Line $Line

    $script:form.BeginInvoke([Action]{
        $script:txtLastMessage.Text = $Line

        if ($Line -match '^Fisiere analizate:\s+(\d+)') {
            $script:txtAnalyzed.Text = $Matches[1]
        }
        elseif ($Line -match '^Duplicate exacte gasite:\s+(\d+)') {
            $script:txtDuplicates.Text = $Matches[1]
        }
        elseif ($Line -match '^Duplicate apropiate gasite:\s+(\d+)') {
            $current = 0
            [int]::TryParse($script:txtDuplicates.Text, [ref]$current) | Out-Null
            $script:txtDuplicates.Text = [string]($current + [int]$Matches[1])
        }
        elseif ($Line -match '^EXIF reparate:\s+(\d+)') {
            $script:txtExif.Text = $Matches[1]
        }
        elseif ($Line -match '^Fisiere mutate:\s+(\d+)') {
            $script:txtMoved.Text = $Matches[1]
        }
        elseif ($Line -match '^Fisiere copiate:\s+(\d+)') {
            $script:txtCopied.Text = $Matches[1]
        }
        elseif ($Line -match '^Cazuri NeedsReview:\s+(\d+)') {
            $script:txtNeedsReview.Text = $Matches[1]
        }

        $script:txtSummary.AppendText($Line + [Environment]::NewLine)
    }) | Out-Null
}

function Validate-Inputs {
    $source = $script:txtSource.Text.Trim()
    $destination = Resolve-LocalizedDestinationPath -DestinationPath $script:txtDestination.Text.Trim() -SourcePath $source
    $script:txtDestination.Text = $destination

    if (-not (Test-Path -LiteralPath $script:EnginePath -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show((T 'missing_engine'), (T 'error_title'), 'OK', 'Error') | Out-Null
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($source) -or -not (Test-Path -LiteralPath $source -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show((T 'invalid_source'), (T 'error_title'), 'OK', 'Warning') | Out-Null
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($destination) -or -not [System.IO.Path]::IsPathRooted($destination)) {
        [System.Windows.Forms.MessageBox]::Show((T 'invalid_destination'), (T 'error_title'), 'OK', 'Warning') | Out-Null
        return $false
    }

    $destinationBase = Get-DestinationBase -DestinationPath $destination
    if ([string]::IsNullOrWhiteSpace($destinationBase) -or -not (Test-Path -LiteralPath $destinationBase -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show((T 'invalid_destination_parent'), (T 'error_title'), 'OK', 'Warning') | Out-Null
        return $false
    }

    if ($script:chkRepairExif.Checked -and -not $script:chkApply.Checked -and -not $script:chkDryRun.Checked) {
        [System.Windows.Forms.MessageBox]::Show((UT 'repair_requires_apply_msg'), (T 'error_title'), 'OK', 'Warning') | Out-Null
        return $false
    }

    return $true
}

function Confirm-Run {
    if ($script:chkApply.Checked) {
        $answer = [System.Windows.Forms.MessageBox]::Show((UT 'confirm_apply_msg'), (T 'warning_title'), 'YesNo', 'Warning')
        if ($answer -ne 'Yes') { return $false }
    }

    if ($script:chkRepairExif.Checked -and $script:chkApply.Checked) {
        $answer = [System.Windows.Forms.MessageBox]::Show((UT 'confirm_repair_exif_msg'), (T 'warning_title'), 'YesNo', 'Warning')
        if ($answer -ne 'Yes') { return $false }
    }

    return $true
}

function Confirm-AdvancedAction {
    param(
        [string]$ActionLabel,
        [string[]]$Switches,
        [bool]$UseRepairExif = $false,
        [bool]$ForceDryRun = $false
    )

    $effectiveApply = $script:chkApply.Checked -and -not $script:chkDryRun.Checked -and -not $ForceDryRun
    if (-not $effectiveApply) { return $true }

    $switchText = ($Switches -join ' ')
    if ($UseRepairExif -and ($switchText -notmatch '(^|\s)-RepairExif(\s|$)')) {
        $switchText = ($switchText + ' -RepairExif').Trim()
    }

    $message = (UT 'confirm_advanced_apply_msg') -f $ActionLabel, $switchText
    $answer = [System.Windows.Forms.MessageBox]::Show($message, (T 'warning_title'), 'YesNo', 'Warning')
    return ($answer -eq 'Yes')
}

function Start-OrganizerRun {
    param(
        [bool]$UseApply,
        [bool]$UseRepairExif,
        [bool]$UseTestScan = $false,
        [bool]$UseReconcile = $false,
        [bool]$UsePurgeMissing = $false,
        [string[]]$ExtraEngineSwitches = @(),
        [bool]$SuppressOrganizeOptions = $false,
        [bool]$AllowKeepEmptyFolders = $false,
        [bool]$SkipConfirmation = $false
    )

    if (-not (Assert-CanStartNormalEngineRun)) { return }

    $maintenanceMode = $UseReconcile -or $UsePurgeMissing -or $SuppressOrganizeOptions
    $effectiveApply = $UseApply -and -not $script:chkDryRun.Checked
    $effectiveRepairExif = $UseRepairExif -and -not $maintenanceMode

    $script:chkApply.Checked = $effectiveApply
    $script:chkDryRun.Checked = -not $effectiveApply
    if ($maintenanceMode) {
        $script:chkRepairExif.Checked = $false
    }
    else {
        $script:chkRepairExif.Checked = $effectiveRepairExif
    }

    if (-not (Validate-Inputs)) { return }
    if (-not $UseTestScan -and -not $SkipConfirmation -and -not (Confirm-Run)) { return }

    if (-not (Test-Path -LiteralPath $script:LogRoot)) {
        New-Item -ItemType Directory -Path $script:LogRoot -Force | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $script:CurrentLogPath = Join-Path $script:LogRoot ("PhotoOrganizer-$timestamp.log")
    $script:CurrentProgressPath = Join-Path $script:LogRoot ("PhotoOrganizer-$timestamp.progress.json")
    $script:RunWasCancelled = $false
    $script:LastLogLength = 0

    Update-RunState -Running $false

    $script:txtAnalyzed.Text = '0'
    $script:txtDuplicates.Text = '0'
    $script:txtExif.Text = '0'
    $script:txtMoved.Text = '0'
    $script:txtCopied.Text = '0'
    $script:txtNeedsReview.Text = '0'
    $script:txtSkipped.Text = '0'
    $script:txtInaccessible.Text = '0'
    $script:txtLocked.Text = '0'
    $script:txtRetries.Text = '0'
    $script:txtMetadataCorrupt.Text = '0'
    $script:txtMetadataBackupSize.Text = '0.000'
    $script:txtEmptyRemoved.Text = '0'
    $script:txtPhase.Text = ''
    $script:txtFilesPerSecond.Text = ''
    $script:txtEta.Text = ''
    $script:txtElapsed.Text = ''
    $script:txtBatch.Text = ''
    $script:txtCpu.Text = ''
    $script:txtRam.Text = ''
    $script:txtWorkers.Text = ''
    $script:txtQueue.Text = ''
    $script:txtPid.Text = ''
    $script:txtLogPath.Text = $script:CurrentLogPath
    $script:txtSummary.Clear()
    $script:txtLastMessage.Text = T 'starting'
    $script:lblStatusValue.Text = T 'running'

    $args = @(
        '-ExecutionPolicy', 'Bypass',
        '-NoProfile',
        '-File', (Quote-Arg $script:EnginePath),
        '-SourcePath', (Quote-Arg $script:txtSource.Text.Trim()),
        '-DestinationPath', (Quote-Arg $script:txtDestination.Text.Trim()),
        '-LogPath', (Quote-Arg $script:CurrentLogPath),
        '-ProgressPath', (Quote-Arg $script:CurrentProgressPath),
        '-PerformanceMode', (Quote-Arg ([string]$script:cmbPerformance.SelectedItem)),
        '-Language', (Quote-Arg $script:CurrentLanguageCode)
    )

    $maxParallelText = $script:txtMaxParallel.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($maxParallelText)) {
        $maxParallel = 0
        if ([int]::TryParse($maxParallelText, [ref]$maxParallel) -and $maxParallel -gt 0) {
            $args += @('-MaxParallelJobs', [string]$maxParallel)
        }
    }

    if ($UseTestScan) { $args += '-TestScan' }
    if ($UseReconcile) { $args += '-ReconcileProcessedDatabase' }
    if ($UsePurgeMissing) { $args += '-PurgeMissingFromProcessedDatabase' }
    foreach ($extraSwitch in @($ExtraEngineSwitches)) {
        if (-not [string]::IsNullOrWhiteSpace($extraSwitch)) {
            $args += $extraSwitch
        }
    }
    if ($effectiveApply) { $args += '-Apply' }
    if ($UseRepairExif -and ($ExtraEngineSwitches -contains '-RepairOnlyExistingOrganizedLibrary')) { $args += '-RepairExif' }
    elseif (-not $maintenanceMode -and $effectiveRepairExif) { $args += '-RepairExif' }
    if (-not $maintenanceMode -and $script:chkCopyInstead.Checked) { $args += '-CopyInsteadOfMove' }
    if ($script:chkDiagnostic.Checked) { $args += '-Diagnostic' }
    if ((-not $maintenanceMode -or $AllowKeepEmptyFolders) -and $script:chkKeepEmpty.Checked) { $args += '-KeepEmptyFolders' }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $engineHost = Get-EngineHost
    $psi.FileName = $engineHost
    $psi.Arguments = ($args -join ' ')
    $psi.UseShellExecute = $false
    $psi.RedirectStandardOutput = $false
    $psi.RedirectStandardError = $false
    $psi.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $psi
    $process.EnableRaisingEvents = $false
    $script:CurrentProcess = $process

    Update-RunState -Running $true
    Append-LogLine -Line ("Command: " + $engineHost + " " + $psi.Arguments)
    $process.Start() | Out-Null
    $script:txtPid.Text = [string]$process.Id
    Append-LogLine -Line ("PID: " + $process.Id)

    if ($script:MonitorTimer) {
        $script:MonitorTimer.Stop()
    }
    $script:MonitorTimer = New-Object System.Windows.Forms.Timer
    $script:MonitorTimer.Interval = 1000
    $script:MonitorTimer.Add_Tick({ Monitor-Run })
    $script:MonitorTimer.Start()
}

function Start-AdvancedDashboardRun {
    param(
        [string]$ActionKey,
        [string]$LabelKey,
        [string[]]$Switches,
        [bool]$UseRepairExif = $false,
        [bool]$AllowKeepEmptyFolders = $false,
        [bool]$ForceDryRun = $false,
        [bool]$RunDashboardRetentionCleanup = $false
    )

    Set-SelectedActionHint -ActionKey $ActionKey
    if (-not (Assert-CanStartAdvancedEngineRun)) { return }
    $label = UT $LabelKey
    if (-not (Confirm-AdvancedAction -ActionLabel $label -Switches $Switches -UseRepairExif:$UseRepairExif -ForceDryRun:$ForceDryRun)) { return }

    if (-not (Validate-Inputs)) { return }
    if (-not (Test-Path -LiteralPath $script:LogRoot)) {
        New-Item -ItemType Directory -Path $script:LogRoot -Force | Out-Null
    }

    $effectiveApply = $script:chkApply.Checked -and -not $script:chkDryRun.Checked -and -not $ForceDryRun
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $advancedRunId = "Advanced-$timestamp"
    $logPath = Join-Path $script:LogRoot ("PhotoOrganizer-$advancedRunId.log")
    $progressPath = Join-Path $script:LogRoot ("PhotoOrganizer-$advancedRunId.progress.json")
    $engineHost = Get-EngineHost

    $engineArguments = @(
        '-ExecutionPolicy', 'Bypass',
        '-NoProfile',
        '-File', $script:EnginePath,
        '-SourcePath', $script:txtSource.Text.Trim(),
        '-DestinationPath', $script:txtDestination.Text.Trim(),
        '-LogPath', $logPath,
        '-ProgressPath', $progressPath,
        '-PerformanceMode', ([string]$script:cmbPerformance.SelectedItem),
        '-Language', $script:CurrentLanguageCode
    )

    $maxParallelText = $script:txtMaxParallel.Text.Trim()
    if (-not [string]::IsNullOrWhiteSpace($maxParallelText)) {
        $maxParallel = 0
        if ([int]::TryParse($maxParallelText, [ref]$maxParallel) -and $maxParallel -gt 0) {
            $engineArguments += @('-MaxParallelJobs', [string]$maxParallel)
        }
    }

    foreach ($extraSwitch in @($Switches)) {
        foreach ($switchPart in @($extraSwitch)) {
            $switchText = [string]$switchPart
            if (-not [string]::IsNullOrWhiteSpace($switchText)) {
                $engineArguments += $switchText
            }
        }
    }
    if ($effectiveApply) { $engineArguments += '-Apply' }
    if ($UseRepairExif) { $engineArguments += '-RepairExif' }
    if ($script:chkDiagnostic.Checked) { $engineArguments += '-Diagnostic' }
    if ($AllowKeepEmptyFolders -and $script:chkKeepEmpty.Checked) { $engineArguments += '-KeepEmptyFolders' }

    try {
        $script:CurrentLogPath = $logPath
        $script:CurrentProgressPath = $progressPath
        if ($script:txtLogPath) { $script:txtLogPath.Text = $logPath }

        if ($RunDashboardRetentionCleanup) {
            if ($effectiveApply) {
                try {
                    Invoke-DashboardRetentionCleanup | Out-Null
                    if ($script:txtSummary -and -not [string]::IsNullOrWhiteSpace($script:LastLogRetentionCleanupSummary)) {
                        $script:txtSummary.AppendText($script:LastLogRetentionCleanupSummary + [Environment]::NewLine)
                    }
                }
                catch {
                    $script:LastLogRetentionCleanupSummary = ''
                }
            }
            elseif ($script:txtSummary) {
                $script:txtSummary.AppendText((UT 'dashboard_retention_cleanup_dryrun_skipped') + [Environment]::NewLine)
            }
        }

        Ensure-TechnicalConsole | Out-Null
        $job = [pscustomobject]@{
            RunId = $advancedRunId
            Label = $label
            EngineHost = $engineHost
            Arguments = @($engineArguments)
            LogPath = $logPath
            ProgressPath = $progressPath
            Apply = [bool]$effectiveApply
        }
        $line = $job | ConvertTo-Json -Compress -Depth 5
        Add-Content -LiteralPath $script:TechnicalConsoleQueuePath -Value $line -Encoding UTF8

        if ($script:lblStatusValue) { $script:lblStatusValue.Text = T 'ready' }
        if ($script:txtLastMessage) { $script:txtLastMessage.Text = (UT 'technical_console_queued') -f $label }
        if ($script:txtSummary) {
            $displayArgs = (@($engineArguments) | ForEach-Object { Quote-Arg $_ }) -join ' '
            $script:txtSummary.AppendText(((UT 'technical_console_queued') -f $label) + [Environment]::NewLine)
            $script:txtSummary.AppendText(("Command: {0} {1}" -f $engineHost, $displayArgs) + [Environment]::NewLine)
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, (T 'error_title'), 'OK', 'Error') | Out-Null
    }
}

function Start-UdmrsMigrationPackage {
    Set-SelectedActionHint -ActionKey 'advanced_migrate_udmrs'

    $normalRunActive = $script:CurrentProcess -ne $null -and -not $script:CurrentProcess.HasExited
    if ($normalRunActive) {
        [System.Windows.Forms.MessageBox]::Show((UT 'normal_run_active_msg'), (T 'warning_title'), 'OK', 'Warning') | Out-Null
        return
    }
    if (Test-AdvancedRunActive) {
        [System.Windows.Forms.MessageBox]::Show((UT 'advanced_run_active_msg'), (T 'warning_title'), 'OK', 'Warning') | Out-Null
        return
    }
    if (-not (Get-Command -Name New-UDMRSMigrationPackage -ErrorAction SilentlyContinue)) {
        [System.Windows.Forms.MessageBox]::Show((UT 'migration_tool_missing_msg'), (T 'error_title'), 'OK', 'Error') | Out-Null
        return
    }

    try {
        if (-not (Test-Path -LiteralPath $script:SettingsRoot -PathType Container)) {
            New-Item -ItemType Directory -Path $script:SettingsRoot -Force | Out-Null
        }
        $outputRoot = if (Get-Command -Name Get-UDMRSDefaultMigrationOutputRoot -ErrorAction SilentlyContinue) {
            Get-UDMRSDefaultMigrationOutputRoot
        }
        else {
            Join-Path (Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads') 'UDMRS-MigrationPackages'
        }
        if ([string]::IsNullOrWhiteSpace($outputRoot)) {
            $outputRoot = Join-Path $script:SettingsRoot 'MigrationPackages'
        }
        $localDataRoot = Join-Path $env:LOCALAPPDATA 'PhotoOrganizer'
        $result = New-UDMRSMigrationPackage `
            -InstallRoot $script:InstallRoot `
            -UserDataRoot $script:SettingsRoot `
            -LocalDataRoot $localDataRoot `
            -OutputRoot $outputRoot `
            -LanguageCode $script:CurrentLanguageCode

        if ($script:txtLastMessage) {
            $script:txtLastMessage.Text = (UT 'migration_complete_short') -f $result.PackageRoot
        }
        if ($script:txtSummary) {
            $script:txtSummary.AppendText(((UT 'migration_complete_short') -f $result.PackageRoot) + [Environment]::NewLine)
            $script:txtSummary.AppendText(("{0}: {1}" -f (UT 'migration_label_install_zip'), $result.InstallZip) + [Environment]::NewLine)
            $script:txtSummary.AppendText(("{0}: {1}" -f (UT 'migration_label_user_state_zip'), $result.UserStateZip) + [Environment]::NewLine)
            $script:txtSummary.AppendText(("{0}: {1}" -f (UT 'migration_label_guide'), $result.GuidePath) + [Environment]::NewLine)
        }

        $message = (UT 'migration_complete_msg') -f $result.PackageRoot, $result.InstallZip, $result.UserStateZip, $result.GuidePath
        [System.Windows.Forms.MessageBox]::Show($message, (UT 'migration_complete_title'), 'OK', 'Information') | Out-Null
        Open-Folder -Path $result.PackageRoot
    }
    catch {
        $message = (UT 'migration_failed_msg') -f $_.Exception.Message
        [System.Windows.Forms.MessageBox]::Show($message, (T 'error_title'), 'OK', 'Error') | Out-Null
    }
}

function Start-GoogleTakeoutImportWizard {
    Set-SelectedActionHint -ActionKey 'import_provider_google'
    $chosen = Choose-Folder -InitialPath $script:txtSource.Text -Description (UT 'select_google_takeout')
    if ([string]::IsNullOrWhiteSpace($chosen)) { return }

    $switches = @('-ImportProvider', 'GoogleTakeout', '-ImportProviderPath', $chosen)
    $switches = @(Confirm-ImportProviderApplyOptions -Switches $switches -ProviderLabelKey 'import_provider_google' -ProviderPath $chosen)
    if (-not $switches -or $switches.Count -eq 0) { return }

    Start-AdvancedDashboardRun `
        -ActionKey 'import_provider_google' `
        -LabelKey 'import_provider_google' `
        -Switches $switches
}

function Start-ApplePhotosImportWizard {
    Set-SelectedActionHint -ActionKey 'import_provider_apple'
    $chosen = Choose-Folder -InitialPath $script:txtSource.Text -Description (UT 'select_apple_photos')
    if ([string]::IsNullOrWhiteSpace($chosen)) { return }

    $switches = @('-ImportProvider', 'ApplePhotos', '-ImportProviderPath', $chosen)
    $switches = @(Confirm-ImportProviderApplyOptions -Switches $switches -ProviderLabelKey 'import_provider_apple' -ProviderPath $chosen)
    if (-not $switches -or $switches.Count -eq 0) { return }

    Start-AdvancedDashboardRun `
        -ActionKey 'import_provider_apple' `
        -LabelKey 'import_provider_apple' `
        -Switches $switches
}

function Start-XmpSidecarImportWizard {
    Set-SelectedActionHint -ActionKey 'import_provider_xmp'
    $chosen = Choose-Folder -InitialPath $script:txtSource.Text -Description (UT 'select_xmp_sidecar_library')
    if ([string]::IsNullOrWhiteSpace($chosen)) { return }

    $switches = @('-ImportProvider', 'XmpSidecarLibrary', '-ImportProviderPath', $chosen)
    $switches = @(Confirm-ImportProviderApplyOptions -Switches $switches -ProviderLabelKey 'import_provider_xmp' -ProviderPath $chosen)
    if (-not $switches -or $switches.Count -eq 0) { return }

    Start-AdvancedDashboardRun `
        -ActionKey 'import_provider_xmp' `
        -LabelKey 'import_provider_xmp' `
        -Switches $switches
}

function Confirm-ImportProviderApplyOptions {
    param(
        [string[]]$Switches,
        [string]$ProviderLabelKey,
        [string]$ProviderPath
    )

    if ($script:chkApply.Checked -and -not $script:chkDryRun.Checked) {
        $provider = UT $ProviderLabelKey
        $answer = [System.Windows.Forms.MessageBox]::Show(((UT 'confirm_import_provider_apply_msg') -f $provider), (T 'warning_title'), 'YesNo', 'Warning')
        if ($answer -ne 'Yes') { return $null }

        $deleteAnswer = [System.Windows.Forms.MessageBox]::Show(((UT 'confirm_import_provider_delete_source_msg') -f $provider, $ProviderPath), (T 'warning_title'), 'YesNo', 'Warning')
        if ($deleteAnswer -eq 'Yes') {
            $Switches += '-DeleteImportProviderSourceAfterSuccess'
        }
    }

    return @($Switches)
}

function Show-PlannedImportProviderNotice {
    param(
        [string]$ProviderLabelKey,
        [string]$StatusKey
    )

    Set-SelectedActionHint -ActionKey 'import_gallery_coming'
    $provider = UT $ProviderLabelKey
    $status = UT $StatusKey
    $message = (UT 'import_provider_coming_msg') -f $provider, $status
    [System.Windows.Forms.MessageBox]::Show($message, (UT 'import_gallery_title'), 'OK', 'Information') | Out-Null
}

function Choose-Folder {
    param(
        [string]$InitialPath,
        [string]$Description
    )

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    if (-not [string]::IsNullOrWhiteSpace($InitialPath) -and (Test-Path -LiteralPath $InitialPath -PathType Container)) {
        $dialog.SelectedPath = $InitialPath
    }

    if ($dialog.ShowDialog() -eq 'OK') {
        return $dialog.SelectedPath
    }
    return $null
}

function Show-LanguageSelector {
    $selector = New-Object System.Windows.Forms.Form
    $selector.Text = T 'language'
    $selector.StartPosition = 'CenterScreen'
    $selector.FormBorderStyle = 'FixedDialog'
    $selector.MaximizeBox = $false
    $selector.MinimizeBox = $false
    $selector.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $selector.MinimumSize = New-Object System.Drawing.Size(360, 190)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = T 'select_language'
    $label.AutoSize = $true

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.DropDownStyle = 'DropDownList'
    Set-LanguageComboItems -Combo $combo
    $combo.Dock = 'Fill'

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'OK'
    $ok.MinimumSize = New-Object System.Drawing.Size(0, 40)
    $ok.AutoSize = $true
    $ok.Add_Click({
        $script:CurrentLanguageCode = Get-SelectedLanguageCode -Combo $combo
        Save-Settings
        $selector.DialogResult = 'OK'
        $selector.Close()
    })

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.Padding = New-Object System.Windows.Forms.Padding(16)
    $layout.ColumnCount = 1
    $layout.RowCount = 3
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    foreach ($control in @($label, $combo, $ok)) {
        $control.Margin = New-Object System.Windows.Forms.Padding(6)
        $control.Dock = 'Top'
    }
    $layout.Controls.Add($label, 0, 0)
    $layout.Controls.Add($combo, 0, 1)
    $layout.Controls.Add($ok, 0, 2)
    $selector.Controls.Add($layout)
    $selector.ShowDialog() | Out-Null
}

function Show-Settings {
    $settings = New-Object System.Windows.Forms.Form
    $settings.Text = T 'settings'
    $settings.StartPosition = 'CenterParent'
    $settings.FormBorderStyle = 'Sizable'
    $settings.MaximizeBox = $false
    $settings.MinimizeBox = $false
    $settings.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $settings.MinimumSize = New-Object System.Drawing.Size(760, 560)
    $settings.ClientSize = New-Object System.Drawing.Size(820, 600)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = T 'language'
    $label.AutoSize = $true

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.DropDownStyle = 'DropDownList'
    Set-LanguageComboItems -Combo $combo
    $combo.Dock = 'Fill'

    $excludedConfig = Read-UserExcludedFoldersConfig
    $excludedEntries = New-Object System.Collections.Generic.List[object]
    foreach ($entry in @($excludedConfig.userExcludedFolders)) {
        if ($null -ne $entry) {
            $excludedEntries.Add([pscustomobject]@{
                    path = [string]$entry.path
                    enabled = if ($entry.PSObject.Properties['enabled']) { [bool]$entry.enabled } else { $true }
                    label = if ($entry.PSObject.Properties['label']) { [string]$entry.label } else { '' }
                    reason = if ($entry.PSObject.Properties['reason']) { [string]$entry.reason } else { 'User protected folder' }
                    role = if ($entry.PSObject.Properties['role']) { [string]$entry.role } else { 'UserExcluded' }
                })
        }
    }

    $excludedLabel = New-Object System.Windows.Forms.Label
    $excludedLabel.Text = UT 'excluded_folders_title'
    $excludedLabel.AutoSize = $true

    $excludedInfo = New-Object System.Windows.Forms.Label
    $excludedInfo.Text = UT 'excluded_folders_info'
    $excludedInfo.AutoSize = $false
    $excludedInfo.Height = 48
    $excludedInfo.Dock = 'Top'

    $excludedList = New-Object System.Windows.Forms.ListView
    $excludedList.View = [System.Windows.Forms.View]::Details
    $excludedList.CheckBoxes = $true
    $excludedList.FullRowSelect = $true
    $excludedList.GridLines = $true
    $excludedList.Dock = 'Fill'
    [void]$excludedList.Columns.Add((UT 'enabled'), 70)
    [void]$excludedList.Columns.Add((UT 'label'), 150)
    [void]$excludedList.Columns.Add((UT 'role'), 110)
    [void]$excludedList.Columns.Add((UT 'path'), 260)
    [void]$excludedList.Columns.Add((UT 'reason'), 210)
    [void]$excludedList.Columns.Add((UT 'exists'), 80)

    $refreshExcludedList = {
        $excludedList.Items.Clear()
        foreach ($entry in @($excludedEntries.ToArray())) {
            $resolved = Expand-DashboardConfiguredPath -Path ([string]$entry.path)
            $existsText = if (-not [string]::IsNullOrWhiteSpace($resolved) -and (Test-Path -LiteralPath $resolved -PathType Container)) { UT 'yes' } else { UT 'no' }
            $item = New-Object System.Windows.Forms.ListViewItem('')
            $item.Checked = [bool]$entry.enabled
            [void]$item.SubItems.Add([string]$entry.label)
            [void]$item.SubItems.Add((Get-ExcludedFolderRoleDisplayText -Role ([string]$entry.role)))
            [void]$item.SubItems.Add([string]$entry.path)
            [void]$item.SubItems.Add((Get-ExcludedFolderReasonDisplayText -Reason ([string]$entry.reason)))
            [void]$item.SubItems.Add($existsText)
            $item.Tag = $entry
            [void]$excludedList.Items.Add($item)
        }
    }

    $excludedList.Add_ItemChecked({
        if ($_.Item -and $_.Item.Tag) {
            $_.Item.Tag.enabled = [bool]$_.Item.Checked
        }
    })

    $addExcluded = New-Object System.Windows.Forms.Button
    $addExcluded.Text = UT 'add_folder'
    $addExcluded.MinimumSize = New-Object System.Drawing.Size(0, 36)
    $addExcluded.AutoSize = $true
    $addExcluded.Add_Click({
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = UT 'add_excluded_folder_prompt'
        if (Test-Path -LiteralPath $script:txtSource.Text.Trim() -PathType Container) {
            $dialog.SelectedPath = $script:txtSource.Text.Trim()
        }
        if ($dialog.ShowDialog($settings) -eq [System.Windows.Forms.DialogResult]::OK) {
            $portablePath = ConvertTo-PortableConfiguredPath -Path $dialog.SelectedPath
            $labelText = [System.IO.Path]::GetFileName($dialog.SelectedPath.TrimEnd('\'))
            $exists = $false
            foreach ($entry in @($excludedEntries.ToArray())) {
                if ([string]$entry.path -eq $portablePath) { $exists = $true; break }
            }
            if (-not $exists) {
                $excludedEntries.Add([pscustomobject]@{
                        path = $portablePath
                        enabled = $true
                        label = $labelText
                        reason = 'User protected folder'
                        role = 'UserExcluded'
                    })
                & $refreshExcludedList
            }
        }
    })

    $removeExcluded = New-Object System.Windows.Forms.Button
    $removeExcluded.Text = UT 'remove_folder'
    $removeExcluded.MinimumSize = New-Object System.Drawing.Size(0, 36)
    $removeExcluded.AutoSize = $true
    $removeExcluded.Add_Click({
        foreach ($item in @($excludedList.SelectedItems)) {
            if ($item.Tag) {
                [void]$excludedEntries.Remove($item.Tag)
            }
        }
        & $refreshExcludedList
    })

    $restoreExcluded = New-Object System.Windows.Forms.Button
    $restoreExcluded.Text = UT 'restore_defaults'
    $restoreExcluded.MinimumSize = New-Object System.Drawing.Size(0, 36)
    $restoreExcluded.AutoSize = $true
    $restoreExcluded.Add_Click({
        $excludedEntries.Clear()
        $defaults = Get-DefaultUserExcludedFoldersConfig
        foreach ($entry in @($defaults.userExcludedFolders)) {
            $excludedEntries.Add([pscustomobject]@{
                    path = [string]$entry.path
                    enabled = [bool]$entry.enabled
                    label = [string]$entry.label
                    reason = [string]$entry.reason
                    role = [string]$entry.role
                })
        }
        & $refreshExcludedList
    })

    $excludedButtons = New-Object System.Windows.Forms.FlowLayoutPanel
    $excludedButtons.Dock = 'Top'
    $excludedButtons.AutoSize = $true
    $excludedButtons.WrapContents = $true
    foreach ($control in @($addExcluded, $removeExcluded, $restoreExcluded)) {
        $control.Margin = New-Object System.Windows.Forms.Padding(6)
        $excludedButtons.Controls.Add($control)
    }

    & $refreshExcludedList

    $save = New-Object System.Windows.Forms.Button
    $save.Text = T 'save'
    $save.MinimumSize = New-Object System.Drawing.Size(0, 40)
    $save.AutoSize = $true
    $save.Add_Click({
        $newCode = Get-SelectedLanguageCode -Combo $combo
        $currentDestination = $script:txtDestination.Text.Trim()
        Load-Language -Code $newCode
        $script:txtDestination.Text = Resolve-LocalizedDestinationPath -DestinationPath $currentDestination -SourcePath $script:txtSource.Text.Trim()
        Save-Settings
        $entriesToSave = @()
        foreach ($entry in @($excludedEntries.ToArray())) {
            $entriesToSave += [pscustomobject]@{
                path = [string]$entry.path
                enabled = [bool]$entry.enabled
                label = [string]$entry.label
                reason = [string]$entry.reason
                role = [string]$entry.role
            }
        }
        $configToSave = [pscustomobject]@{
            version = 1
            userExcludedFolders = $entriesToSave
            vendorPresets = @((Get-DefaultUserExcludedFoldersConfig).vendorPresets)
        }
        Save-UserExcludedFoldersConfig -Config $configToSave
        Update-Texts
        Update-FolderProtectionSummary
        $settings.Close()
    })

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = T 'cancel'
    $cancel.MinimumSize = New-Object System.Drawing.Size(0, 40)
    $cancel.AutoSize = $true
    $cancel.Add_Click({ $settings.Close() })

    $layout = New-Object System.Windows.Forms.TableLayoutPanel
    $layout.Dock = 'Fill'
    $layout.Padding = New-Object System.Windows.Forms.Padding(16)
    $layout.ColumnCount = 1
    $layout.RowCount = 7
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $buttons = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttons.Dock = 'Top'
    $buttons.AutoSize = $true
    $buttons.FlowDirection = 'RightToLeft'
    foreach ($control in @($save, $cancel)) {
        $control.Margin = New-Object System.Windows.Forms.Padding(6)
        $buttons.Controls.Add($control)
    }
    foreach ($control in @($label, $combo, $excludedLabel, $excludedInfo, $excludedList, $excludedButtons, $buttons)) {
        $control.Margin = New-Object System.Windows.Forms.Padding(6)
        $control.Dock = 'Top'
    }
    $excludedList.Dock = 'Fill'
    $layout.Controls.Add($label, 0, 0)
    $layout.Controls.Add($combo, 0, 1)
    $layout.Controls.Add($excludedLabel, 0, 2)
    $layout.Controls.Add($excludedInfo, 0, 3)
    $layout.Controls.Add($excludedList, 0, 4)
    $layout.Controls.Add($excludedButtons, 0, 5)
    $layout.Controls.Add($buttons, 0, 6)
    $settings.Controls.Add($layout)
    $settings.ShowDialog($script:form) | Out-Null
}

function New-ResponsivePanel {
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Fill'
    $panel.AutoScroll = $true
    $panel.Padding = New-Object System.Windows.Forms.Padding(12)
    return $panel
}

function Set-TouchControl {
    param([System.Windows.Forms.Control]$Control)
    $Control.Margin = New-Object System.Windows.Forms.Padding(6)
    if ($Control -is [System.Windows.Forms.Button]) {
        $Control.MinimumSize = New-Object System.Drawing.Size(0, 40)
        $Control.AutoSize = $true
    }
    elseif ($Control -is [System.Windows.Forms.CheckBox]) {
        $Control.AutoSize = $true
        $Control.MinimumSize = New-Object System.Drawing.Size(0, 32)
    }
}

function Add-ResponsiveRow {
    param(
        [System.Windows.Forms.TableLayoutPanel]$Table,
        [System.Windows.Forms.Control]$Label,
        [System.Windows.Forms.Control]$InputControl,
        [System.Windows.Forms.Control]$Button = $null
    )

    $row = $Table.RowCount
    $Table.RowCount++
    $Table.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null

    $Label.Dock = 'Fill'
    $Label.TextAlign = 'MiddleLeft'
    $Label.Margin = New-Object System.Windows.Forms.Padding(6)
    $InputControl.Dock = 'Fill'
    $InputControl.Margin = New-Object System.Windows.Forms.Padding(6)
    $Table.Controls.Add($Label, 0, $row)
    $Table.Controls.Add($InputControl, 1, $row)

    if ($Button) {
        Set-TouchControl -Control $Button
        $Button.Dock = 'Fill'
        $Table.Controls.Add($Button, 2, $row)
    }
}

function New-MetricCard {
    param(
        [System.Windows.Forms.Label]$Label,
        [System.Windows.Forms.TextBox]$TextBox
    )

    $panel = New-Object System.Windows.Forms.TableLayoutPanel
    $panel.ColumnCount = 1
    $panel.RowCount = 2
    $panel.AutoSize = $true
    $panel.MinimumSize = New-Object System.Drawing.Size(190, 64)
    $panel.Margin = New-Object System.Windows.Forms.Padding(6)
    $panel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $panel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null

    $Label.AutoSize = $true
    $Label.Margin = New-Object System.Windows.Forms.Padding(3)
    $TextBox.ReadOnly = $true
    $TextBox.Dock = 'Top'
    $TextBox.Margin = New-Object System.Windows.Forms.Padding(3)
    $panel.Controls.Add($Label, 0, 0)
    $panel.Controls.Add($TextBox, 0, 1)
    return $panel
}

function Save-DashboardWindowSettings {
    if (-not $script:form) { return }
    if (-not (Test-Path -LiteralPath $script:SettingsRoot)) {
        New-Item -ItemType Directory -Path $script:SettingsRoot -Force | Out-Null
    }

    $bounds = if ($script:form.WindowState -eq 'Normal') { $script:form.Bounds } else { $script:form.RestoreBounds }
    [pscustomobject]@{
        WindowX = $bounds.X
        WindowY = $bounds.Y
        WindowWidth = $bounds.Width
        WindowHeight = $bounds.Height
        WindowState = [string]$script:form.WindowState
    } | ConvertTo-Json | Set-Content -LiteralPath $script:DashboardSettingsPath -Encoding UTF8
}

function Get-DashboardWorkingArea {
    try {
        if ($script:form) {
            return [System.Windows.Forms.Screen]::FromControl($script:form).WorkingArea
        }
        return [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    }
    catch {
        return New-Object System.Drawing.Rectangle(0, 0, 1024, 768)
    }
}

function Get-DashboardSafeSize {
    param(
        [int]$DesiredWidth,
        [int]$DesiredHeight,
        [int]$MinimumWidth = 800,
        [int]$MinimumHeight = 600
    )

    $workingArea = Get-DashboardWorkingArea
    $margin = 40
    $maxWidth = [math]::Max(640, [int]$workingArea.Width - $margin)
    $maxHeight = [math]::Max(480, [int]$workingArea.Height - $margin)
    $safeMinimumWidth = [math]::Min($MinimumWidth, $maxWidth)
    $safeMinimumHeight = [math]::Min($MinimumHeight, $maxHeight)
    $safeWidth = [math]::Min([math]::Max($safeMinimumWidth, $DesiredWidth), $maxWidth)
    $safeHeight = [math]::Min([math]::Max($safeMinimumHeight, $DesiredHeight), $maxHeight)

    return [pscustomobject]@{
        MinimumSize = New-Object System.Drawing.Size($safeMinimumWidth, $safeMinimumHeight)
        ClientSize = New-Object System.Drawing.Size($safeWidth, $safeHeight)
        WorkingArea = $workingArea
    }
}

function Get-DashboardSafeBounds {
    param(
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height
    )

    $safe = Get-DashboardSafeSize -DesiredWidth $Width -DesiredHeight $Height
    $workingArea = $safe.WorkingArea
    $safeWidth = $safe.ClientSize.Width
    $safeHeight = $safe.ClientSize.Height
    $safeX = [math]::Max($workingArea.Left, [math]::Min($X, $workingArea.Right - $safeWidth))
    $safeY = [math]::Max($workingArea.Top, [math]::Min($Y, $workingArea.Bottom - $safeHeight))
    return New-Object System.Drawing.Rectangle($safeX, $safeY, $safeWidth, $safeHeight)
}

function Apply-DashboardWindowSettings {
    param([object]$WindowSettings)

    if (-not $WindowSettings) { return }
    if ($WindowSettings.WindowWidth -and $WindowSettings.WindowHeight) {
        $width = [math]::Max(800, [int]$WindowSettings.WindowWidth)
        $height = [math]::Max(600, [int]$WindowSettings.WindowHeight)
        $script:form.StartPosition = 'Manual'
        $x = if ($WindowSettings.WindowX -ne $null) { [int]$WindowSettings.WindowX } else { 40 }
        $y = if ($WindowSettings.WindowY -ne $null) { [int]$WindowSettings.WindowY } else { 40 }
        $script:form.Bounds = Get-DashboardSafeBounds -X $x -Y $y -Width $width -Height $height
    }

    if ($WindowSettings.WindowState -eq 'Maximized') {
        $script:form.WindowState = 'Maximized'
    }
}
function Build-ResponsiveLayout {
    $script:form.SuspendLayout()
    $script:form.Controls.Clear()
    $script:form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
    $script:form.AutoScroll = $true
    $safeDashboardSize = Get-DashboardSafeSize -DesiredWidth 1000 -DesiredHeight 720 -MinimumWidth 800 -MinimumHeight 600
    $script:form.MinimumSize = $safeDashboardSize.MinimumSize
    $script:form.ClientSize = $safeDashboardSize.ClientSize
    $script:form.Padding = New-Object System.Windows.Forms.Padding(8)

    $rootLayout = New-Object System.Windows.Forms.TableLayoutPanel
    $rootLayout.Dock = 'Fill'
    $rootLayout.ColumnCount = 1
    $rootLayout.RowCount = 2
    $rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $rootLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null

    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Dock = 'Fill'

    $script:tabConfiguration = New-Object System.Windows.Forms.TabPage
    $script:tabImportGallery = New-Object System.Windows.Forms.TabPage
    $script:tabExecution = New-Object System.Windows.Forms.TabPage
    $script:tabProgress = New-Object System.Windows.Forms.TabPage
    $script:tabLogs = New-Object System.Windows.Forms.TabPage

    $tabs.TabPages.AddRange(@($script:tabConfiguration, $script:tabImportGallery, $script:tabProgress, $script:tabLogs, $script:tabExecution))
    $rootLayout.Controls.Add($tabs, 0, 0)
    if ($script:statusStripInfo) {
        $script:statusStripInfo.Dock = 'Fill'
        $rootLayout.Controls.Add($script:statusStripInfo, 0, 1)
    }
    $script:form.Controls.Add($rootLayout)

    $configPanel = New-ResponsivePanel
    $script:tabConfiguration.Controls.Add($configPanel)

    $configTable = New-Object System.Windows.Forms.TableLayoutPanel
    $configTable.Dock = 'Top'
    $configTable.AutoSize = $true
    $configTable.ColumnCount = 3
    $configTable.RowCount = 0
    $configTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $configTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $configTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $configPanel.Controls.Add($configTable)

    $script:txtSource.Text = Get-DefaultSourcePath
    Add-ResponsiveRow -Table $configTable -Label $script:lblSource -InputControl $script:txtSource -Button $script:btnBrowseSource
    $script:txtDestination.Text = Get-DefaultDestinationPath -SourcePath $script:txtSource.Text
    Add-ResponsiveRow -Table $configTable -Label $script:lblDestination -InputControl $script:txtDestination -Button $script:btnBrowseDestination

    $script:grpOptions.Dock = 'Top'
    $script:grpOptions.AutoSize = $true
    $script:grpOptions.Padding = New-Object System.Windows.Forms.Padding(10)
    $script:grpOptions.Text = 'Modo de ejecucion'
    $configTable.RowCount++
    $configTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $configTable.Controls.Add($script:grpOptions, 0, 2)
    $configTable.SetColumnSpan($script:grpOptions, 3)

    $modeFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $modeFlow.Dock = 'Top'
    $modeFlow.AutoSize = $true
    $modeFlow.WrapContents = $true
    $modeFlow.FlowDirection = 'LeftToRight'
    foreach ($control in @($script:chkDryRun, $script:chkApply, $script:chkRepairExif)) {
        Set-TouchControl -Control $control
        $modeFlow.Controls.Add($control)
    }

    $performanceTable = New-Object System.Windows.Forms.TableLayoutPanel
    $performanceTable.Dock = 'Top'
    $performanceTable.AutoSize = $true
    $performanceTable.ColumnCount = 4
    $performanceTable.RowCount = 1
    $performanceTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $performanceTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $performanceTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $performanceTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    foreach ($control in @($script:lblPerformance, $script:cmbPerformance, $script:lblMaxParallel, $script:txtMaxParallel)) {
        $control.Margin = New-Object System.Windows.Forms.Padding(6)
        $control.Dock = 'Fill'
    }
    $performanceTable.Controls.Add($script:lblPerformance, 0, 0)
    $performanceTable.Controls.Add($script:cmbPerformance, 1, 0)
    $performanceTable.Controls.Add($script:lblMaxParallel, 2, 0)
    $performanceTable.Controls.Add($script:txtMaxParallel, 3, 0)

    $modeTable = New-Object System.Windows.Forms.TableLayoutPanel
    $modeTable.Dock = 'Top'
    $modeTable.AutoSize = $true
    $modeTable.ColumnCount = 1
    $modeTable.RowCount = 2
    $modeTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $modeTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $modeTable.Controls.Add($modeFlow, 0, 0)
    $modeTable.Controls.Add($performanceTable, 0, 1)

    $script:grpOptions.Controls.Clear()
    $script:grpOptions.Controls.Add($modeTable)

    $script:grpAdvancedOptions = New-Object System.Windows.Forms.GroupBox
    $script:grpAdvancedOptions.Text = UT 'advanced_options'
    $script:grpAdvancedOptions.Dock = 'Top'
    $script:grpAdvancedOptions.AutoSize = $true
    $script:grpAdvancedOptions.Padding = New-Object System.Windows.Forms.Padding(10)
    $advancedTable = New-Object System.Windows.Forms.TableLayoutPanel
    $advancedTable.Dock = 'Top'
    $advancedTable.AutoSize = $true
    $advancedTable.ColumnCount = 1
    $advancedTable.RowCount = 2
    $advancedFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $advancedFlow.Dock = 'Top'
    $advancedFlow.AutoSize = $true
    $advancedFlow.WrapContents = $true
    foreach ($control in @($script:chkCopyInstead, $script:chkKeepEmpty, $script:chkDiagnostic, $script:chkOpenOutput, $script:chkOpenNeeds, $script:chkOpenDuplicates, $script:chkOpenLog, $script:chkRenameInternal)) {
        Set-TouchControl -Control $control
        $advancedFlow.Controls.Add($control)
    }
    Set-TouchControl -Control $script:btnPurgeMissing
    $advancedFlow.Controls.Add($script:btnPurgeMissing)
    $script:lblAdvancedNote = New-Object System.Windows.Forms.Label
    $script:lblAdvancedNote.AutoSize = $true
    $script:lblAdvancedNote.Margin = New-Object System.Windows.Forms.Padding(8)
    $script:lblAdvancedNote.Text = UT 'advanced_note'
    $advancedTable.Controls.Add($advancedFlow, 0, 0)
    $advancedTable.Controls.Add($script:lblAdvancedNote, 0, 1)
    $script:grpAdvancedOptions.Controls.Add($advancedTable)
    $configTable.RowCount++
    $configTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $configTable.Controls.Add($script:grpAdvancedOptions, 0, 3)
    $configTable.SetColumnSpan($script:grpAdvancedOptions, 3)

    $script:grpFolderProtection = New-Object System.Windows.Forms.GroupBox
    $script:grpFolderProtection.Text = UT 'folder_protection_group'
    $script:grpFolderProtection.Dock = 'Top'
    $script:grpFolderProtection.AutoSize = $true
    $script:grpFolderProtection.Padding = New-Object System.Windows.Forms.Padding(10)

    $folderProtectionTable = New-Object System.Windows.Forms.TableLayoutPanel
    $folderProtectionTable.Dock = 'Top'
    $folderProtectionTable.AutoSize = $true
    $folderProtectionTable.ColumnCount = 2
    $folderProtectionTable.RowCount = 2
    $folderProtectionTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $folderProtectionTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $folderProtectionTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $folderProtectionTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null

    $script:lblFolderProtectionSummary = New-Object System.Windows.Forms.Label
    $script:lblFolderProtectionSummary.AutoSize = $true
    $script:lblFolderProtectionSummary.MaximumSize = New-Object System.Drawing.Size(650, 0)
    $script:lblFolderProtectionSummary.Margin = New-Object System.Windows.Forms.Padding(6)

    $script:lblFolderProtectionHint = New-Object System.Windows.Forms.Label
    $script:lblFolderProtectionHint.AutoSize = $true
    $script:lblFolderProtectionHint.MaximumSize = New-Object System.Drawing.Size(900, 0)
    $script:lblFolderProtectionHint.ForeColor = [System.Drawing.Color]::DimGray
    $script:lblFolderProtectionHint.Margin = New-Object System.Windows.Forms.Padding(6)

    $script:btnManageExcludedFolders = New-Object System.Windows.Forms.Button
    $script:btnManageExcludedFolders.AutoSize = $true
    $script:btnManageExcludedFolders.MinimumSize = New-Object System.Drawing.Size(210, 40)
    $script:btnManageExcludedFolders.Margin = New-Object System.Windows.Forms.Padding(6)
    $script:btnManageExcludedFolders.Add_Click({ Show-Settings })

    $folderProtectionTable.Controls.Add($script:lblFolderProtectionSummary, 0, 0)
    $folderProtectionTable.Controls.Add($script:btnManageExcludedFolders, 1, 0)
    $folderProtectionTable.Controls.Add($script:lblFolderProtectionHint, 0, 1)
    $folderProtectionTable.SetColumnSpan($script:lblFolderProtectionHint, 2)
    $script:grpFolderProtection.Controls.Add($folderProtectionTable)

    $configTable.RowCount++
    $configTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $configTable.Controls.Add($script:grpFolderProtection, 0, 4)
    $configTable.SetColumnSpan($script:grpFolderProtection, 3)

    $script:grpModeSummary = New-Object System.Windows.Forms.GroupBox
    $script:grpModeSummary.Text = UT 'mode_summary_group'
    $script:grpModeSummary.Dock = 'Top'
    $script:grpModeSummary.AutoSize = $true
    $script:grpModeSummary.Padding = New-Object System.Windows.Forms.Padding(10)
    $script:txtModeSummary.Dock = 'Top'
    $script:txtModeSummary.Height = 150
    $script:grpModeSummary.Controls.Add($script:txtModeSummary)
    $configTable.RowCount++
    $configTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $configTable.Controls.Add($script:grpModeSummary, 0, 5)
    $configTable.SetColumnSpan($script:grpModeSummary, 3)

    $executeFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $executeFlow.Dock = 'Top'
    $executeFlow.AutoSize = $true
    $executeFlow.FlowDirection = 'LeftToRight'
    $script:btnExecute.Text = UT 'execute'
    Set-TouchControl -Control $script:btnExecute
    $executeFlow.Controls.Add($script:btnExecute)
    Set-TouchControl -Control $script:btnCancel
    $executeFlow.Controls.Add($script:btnCancel)
    Set-TouchControl -Control $script:btnTestScan
    $executeFlow.Controls.Add($script:btnTestScan)
    Set-TouchControl -Control $script:btnReconcile
    $executeFlow.Controls.Add($script:btnReconcile)

    $script:lblExecuteHint = New-Object System.Windows.Forms.Label
    $script:lblExecuteHint.AutoSize = $true
    $script:lblExecuteHint.MaximumSize = New-Object System.Drawing.Size(900, 0)
    $script:lblExecuteHint.Margin = New-Object System.Windows.Forms.Padding(12, 14, 6, 8)
    $executeFlow.Controls.Add($script:lblExecuteHint)
    $configTable.RowCount++
    $configTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $configTable.Controls.Add($executeFlow, 0, 6)
    $configTable.SetColumnSpan($executeFlow, 3)

    $statusTable = New-Object System.Windows.Forms.TableLayoutPanel
    $statusTable.Dock = 'Top'
    $statusTable.AutoSize = $true
    $statusTable.ColumnCount = 4
    $statusTable.RowCount = 2
    $statusTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $statusTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    $statusTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $statusTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50))) | Out-Null
    foreach ($control in @($script:lblStatus, $script:lblStatusValue, $script:lblPid, $script:txtPid, $script:lblLastMessage, $script:txtLastMessage)) {
        $control.Margin = New-Object System.Windows.Forms.Padding(6)
        $control.Dock = 'Fill'
    }
    $statusTable.Controls.Add($script:lblStatus, 0, 0)
    $statusTable.Controls.Add($script:lblStatusValue, 1, 0)
    $statusTable.Controls.Add($script:lblPid, 2, 0)
    $statusTable.Controls.Add($script:txtPid, 3, 0)
    $statusTable.Controls.Add($script:lblLastMessage, 0, 1)
    $statusTable.Controls.Add($script:txtLastMessage, 1, 1)
    $statusTable.SetColumnSpan($script:txtLastMessage, 3)
    $script:grpCurrentStatus = New-Object System.Windows.Forms.GroupBox
    $script:grpCurrentStatus.Text = UT 'status_group'
    $script:grpCurrentStatus.Dock = 'Top'
    $script:grpCurrentStatus.AutoSize = $true
    $script:grpCurrentStatus.Padding = New-Object System.Windows.Forms.Padding(10)
    $script:grpCurrentStatus.Controls.Add($statusTable)
    $configTable.RowCount++
    $configTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $configTable.Controls.Add($script:grpCurrentStatus, 0, 7)
    $configTable.SetColumnSpan($script:grpCurrentStatus, 3)

    $script:advancedToolsGroup = $null

    $importPanel = New-ResponsivePanel
    $script:tabImportGallery.Controls.Add($importPanel)
    $importTable = New-Object System.Windows.Forms.TableLayoutPanel
    $importTable.Dock = 'Top'
    $importTable.AutoSize = $true
    $importTable.ColumnCount = 2
    $importTable.RowCount = 0
    $importTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $importTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $importPanel.Controls.Add($importTable)

    $script:lblImportGalleryTitle = New-Object System.Windows.Forms.Label
    $script:lblImportGalleryTitle.AutoSize = $true
    $script:lblImportGalleryTitle.Font = New-Object System.Drawing.Font('Segoe UI', 14, [System.Drawing.FontStyle]::Bold)
    $script:lblImportGalleryTitle.Margin = New-Object System.Windows.Forms.Padding(8, 8, 8, 4)
    $script:lblImportGalleryTitle.Text = UT 'import_gallery_title'
    $importTable.RowCount++
    $importTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $importTable.Controls.Add($script:lblImportGalleryTitle, 0, 0)
    $importTable.SetColumnSpan($script:lblImportGalleryTitle, 2)

    $script:lblImportGalleryIntro = New-Object System.Windows.Forms.Label
    $script:lblImportGalleryIntro.AutoSize = $true
    $script:lblImportGalleryIntro.MaximumSize = New-Object System.Drawing.Size(900, 0)
    $script:lblImportGalleryIntro.Margin = New-Object System.Windows.Forms.Padding(8, 4, 8, 16)
    $script:lblImportGalleryIntro.Text = UT 'import_gallery_intro'
    $importTable.RowCount++
    $importTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $importTable.Controls.Add($script:lblImportGalleryIntro, 0, 1)
    $importTable.SetColumnSpan($script:lblImportGalleryIntro, 2)

    $script:ImportProviderDescriptionLabels = @{}
    $importProviderRows = @(
        @($script:btnImportGoogleTakeout, 'import_provider_google_desc', 'import_provider_status_available'),
        @($script:btnImportApplePhotos, 'import_provider_apple_desc', 'import_provider_status_available'),
        @($script:btnImportSamsungGallery, 'import_provider_samsung_desc', 'import_provider_status_coming_sample'),
        @($script:btnImportImmich, 'import_provider_immich_desc', 'import_provider_status_coming_sample'),
        @($script:btnImportXmpSidecar, 'import_provider_xmp_desc', 'import_provider_status_available')
    )
    foreach ($row in $importProviderRows) {
        $button = $row[0]
        $descriptionKey = [string]$row[1]
        $statusKey = [string]$row[2]
        Set-TouchControl -Control $button
        $label = New-Object System.Windows.Forms.Label
        $label.AutoSize = $true
        $label.MaximumSize = New-Object System.Drawing.Size(720, 0)
        $label.Margin = New-Object System.Windows.Forms.Padding(8, 12, 8, 10)
        $label.Text = ('{0} - {1}' -f (UT $statusKey), (UT $descriptionKey))
        $script:ImportProviderDescriptionLabels[$descriptionKey] = [pscustomobject]@{ Label = $label; StatusKey = $statusKey }
        $rowIndex = $importTable.RowCount
        $importTable.RowCount++
        $importTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
        $importTable.Controls.Add($button, 0, $rowIndex)
        $importTable.Controls.Add($label, 1, $rowIndex)
    }

    $progressPanel = New-ResponsivePanel
    $script:tabProgress.Controls.Add($progressPanel)
    $progressTable = New-Object System.Windows.Forms.TableLayoutPanel
    $progressTable.Dock = 'Top'
    $progressTable.AutoSize = $true
    $progressTable.ColumnCount = 1
    $progressTable.RowCount = 3
    $progressPanel.Controls.Add($progressTable)

    $script:progress.Dock = 'Top'
    $script:progress.MinimumSize = New-Object System.Drawing.Size(0, 28)
    $script:progress.Margin = New-Object System.Windows.Forms.Padding(6, 6, 6, 12)
    $progressTable.Controls.Add($script:progress, 0, 0)

    $script:grpMainProgress = New-Object System.Windows.Forms.GroupBox
    $script:grpMainProgress.Text = UT 'main_progress'
    $script:grpMainProgress.Dock = 'Top'
    $script:grpMainProgress.AutoSize = $true
    $script:grpMainProgress.Padding = New-Object System.Windows.Forms.Padding(10)
    $mainMetricFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $mainMetricFlow.Dock = 'Top'
    $mainMetricFlow.AutoSize = $true
    $mainMetricFlow.WrapContents = $true
    foreach ($metric in @(
        @($script:lblAnalyzed, $script:txtAnalyzed), @($script:lblMoved, $script:txtMoved),
        @($script:lblDuplicates, $script:txtDuplicates), @($script:lblExif, $script:txtExif),
        @($script:lblPhase, $script:txtPhase), @($script:lblFps, $script:txtFilesPerSecond),
        @($script:lblEta, $script:txtEta), @($script:lblMetadataBackupSize, $script:txtMetadataBackupSize)
    )) {
        $mainMetricFlow.Controls.Add((New-MetricCard -Label $metric[0] -TextBox $metric[1]))
    }
    $script:grpMainProgress.Controls.Add($mainMetricFlow)
    $progressTable.Controls.Add($script:grpMainProgress, 0, 1)

    $script:grpAdvancedProgress = New-Object System.Windows.Forms.GroupBox
    $script:grpAdvancedProgress.Text = UT 'advanced_details'
    $script:grpAdvancedProgress.Dock = 'Top'
    $script:grpAdvancedProgress.AutoSize = $true
    $script:grpAdvancedProgress.Padding = New-Object System.Windows.Forms.Padding(10)
    $advancedMetricFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $advancedMetricFlow.Dock = 'Top'
    $advancedMetricFlow.AutoSize = $true
    $advancedMetricFlow.WrapContents = $true
    foreach ($metric in @(
        @($script:lblCopied, $script:txtCopied), @($script:lblNeedsReview, $script:txtNeedsReview),
        @($script:lblElapsed, $script:txtElapsed), @($script:lblBatch, $script:txtBatch),
        @($script:lblSkipped, $script:txtSkipped), @($script:lblInaccessible, $script:txtInaccessible),
        @($script:lblEmptyRemoved, $script:txtEmptyRemoved), @($script:lblLocked, $script:txtLocked),
        @($script:lblRetries, $script:txtRetries), @($script:lblMetadataCorrupt, $script:txtMetadataCorrupt),
        @($script:lblCpu, $script:txtCpu), @($script:lblRam, $script:txtRam),
        @($script:lblWorkers, $script:txtWorkers), @($script:lblQueue, $script:txtQueue)
    )) {
        $advancedMetricFlow.Controls.Add((New-MetricCard -Label $metric[0] -TextBox $metric[1]))
    }
    $script:grpAdvancedProgress.Controls.Add($advancedMetricFlow)
    $progressTable.Controls.Add($script:grpAdvancedProgress, 0, 2)

    $logsPanel = New-ResponsivePanel
    $script:tabLogs.Controls.Add($logsPanel)
    $logsTable = New-Object System.Windows.Forms.TableLayoutPanel
    $logsTable.Dock = 'Fill'
    $logsTable.ColumnCount = 1
    $logsTable.RowCount = 2
    $logsTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $logsTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $logsPanel.Controls.Add($logsTable)

    $logsButtonFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $logsButtonFlow.Dock = 'Top'
    $logsButtonFlow.AutoSize = $true
    Set-TouchControl -Control $script:btnLogs
    $script:btnLogs.Text = UT 'open_logs_folder'
    $logsButtonFlow.Controls.Add($script:btnLogs)
    $logsTable.Controls.Add($logsButtonFlow, 0, 0)

    $script:grpSummary.Dock = 'Fill'
    $script:grpSummary.Padding = New-Object System.Windows.Forms.Padding(10)
    $script:txtSummary.Dock = 'Fill'
    Set-LogTextBoxStyle -TextBox $script:txtSummary
    $script:grpSummary.Controls.Clear()
    $script:grpSummary.Controls.Add($script:txtSummary)
    $logsTable.Controls.Add($script:grpSummary, 0, 1)

    $settingsPanel = New-ResponsivePanel
    $script:tabExecution.Controls.Add($settingsPanel)
    $settingsTable = New-Object System.Windows.Forms.TableLayoutPanel
    $settingsTable.Dock = 'Top'
    $settingsTable.AutoSize = $true
    $settingsTable.ColumnCount = 1
    $settingsTable.RowCount = 0
    $settingsPanel.Controls.Add($settingsTable)

    $settingsFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $settingsFlow.Dock = 'Top'
    $settingsFlow.AutoSize = $true
    $settingsFlow.WrapContents = $true
    foreach ($button in @($script:btnHelp, $script:btnSettings, $script:btnAdvancedModes, $script:btnExit)) {
        Set-TouchControl -Control $button
        $settingsFlow.Controls.Add($button)
    }
    $settingsTable.RowCount++
    $settingsTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $settingsTable.Controls.Add($settingsFlow, 0, 0)

    $script:advancedModesGroup = New-Object System.Windows.Forms.GroupBox
    $script:advancedModesGroup.Text = UT 'advanced_modes_group'
    $script:advancedModesGroup.Dock = 'Top'
    $script:advancedModesGroup.AutoSize = $true
    $script:advancedModesGroup.Padding = New-Object System.Windows.Forms.Padding(10)
    $script:advancedModesGroup.Visible = $false
    $script:AdvancedModeDescriptionLabels = @{}

    $advancedModesTable = New-Object System.Windows.Forms.TableLayoutPanel
    $advancedModesTable.Dock = 'Top'
    $advancedModesTable.AutoSize = $true
    $advancedModesTable.ColumnCount = 2
    $advancedModesTable.RowCount = 0
    $advancedModesTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $advancedModesTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

    $script:lblAdvancedModesNote = New-Object System.Windows.Forms.Label
    $script:lblAdvancedModesNote.AutoSize = $true
    $script:lblAdvancedModesNote.MaximumSize = New-Object System.Drawing.Size(900, 0)
    $script:lblAdvancedModesNote.Margin = New-Object System.Windows.Forms.Padding(8)
    $script:lblAdvancedModesNote.Text = UT 'advanced_modes_note'
    $advancedModesTable.RowCount++
    $advancedModesTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $advancedModesTable.Controls.Add($script:lblAdvancedModesNote, 0, 0)
    $advancedModesTable.SetColumnSpan($script:lblAdvancedModesNote, 2)

    $advancedModeRows = @(
        @($script:btnRetentionCleanup, 'advanced_retention_cleanup_desc'),
        @($script:btnRecoverWrongDuplicateMove, 'advanced_recover_wrong_duplicate_move_desc'),
        @($script:btnRenameExistingFolders, 'advanced_rename_existing_folders_desc'),
        @($script:btnRenameInternalFolders, 'advanced_rename_internal_folders_desc'),
        @($script:btnNormalizeExistingFolders, 'advanced_normalize_existing_folders_desc'),
        @($script:btnDedupeCleanup, 'advanced_dedupe_cleanup_desc'),
        @($script:btnRepairOnlyLibrary, 'advanced_repair_only_library_desc'),
        @($script:btnMetadataAudit, 'advanced_metadata_audit_desc'),
        @($script:btnMetadataRepair, 'advanced_metadata_repair_desc'),
        @($script:btnMigrateUdmrs, 'advanced_migrate_udmrs_desc')
    )
    foreach ($row in $advancedModeRows) {
        $button = $row[0]
        $labelKey = [string]$row[1]
        Set-TouchControl -Control $button
        $label = New-Object System.Windows.Forms.Label
        $label.AutoSize = $true
        $label.MaximumSize = New-Object System.Drawing.Size(650, 0)
        $label.Margin = New-Object System.Windows.Forms.Padding(8, 12, 8, 8)
        $label.Text = UT $labelKey
        $script:AdvancedModeDescriptionLabels[$labelKey] = $label
        $rowIndex = $advancedModesTable.RowCount
        $advancedModesTable.RowCount++
        $advancedModesTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
        $advancedModesTable.Controls.Add($button, 0, $rowIndex)
        $advancedModesTable.Controls.Add($label, 1, $rowIndex)
    }

    $script:advancedModesGroup.Controls.Add($advancedModesTable)
    $settingsTable.RowCount++
    $settingsTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
    $settingsTable.Controls.Add($script:advancedModesGroup, 0, 1)

    $script:form.ResumeLayout($true)
}

function Update-Texts {
    $script:form.Text = UT 'app_title'
    if ($script:tabConfiguration) { $script:tabConfiguration.Text = UT 'tab_home' }
    if ($script:tabImportGallery) { $script:tabImportGallery.Text = UT 'tab_import_gallery' }
    if ($script:tabExecution) { $script:tabExecution.Text = UT 'tab_settings' }
    if ($script:tabProgress) { $script:tabProgress.Text = T 'tab_progress' }
    if ($script:tabLogs) { $script:tabLogs.Text = UT 'tab_logs_summary' }
    $script:lblSource.Text = T 'source_folder'
    $script:lblDestination.Text = T 'destination_folder'
    $script:btnBrowseSource.Text = T 'browse'
    $script:btnBrowseDestination.Text = T 'browse'
    $script:grpOptions.Text = UT 'mode_group'
    if ($script:grpAdvancedOptions) { $script:grpAdvancedOptions.Text = UT 'advanced_options' }
    if ($script:lblAdvancedNote) { $script:lblAdvancedNote.Text = UT 'advanced_note' }
    if ($script:grpFolderProtection) { $script:grpFolderProtection.Text = UT 'folder_protection_group' }
    if ($script:lblFolderProtectionHint) { $script:lblFolderProtectionHint.Text = UT 'folder_protection_hint' }
    if ($script:btnManageExcludedFolders) { $script:btnManageExcludedFolders.Text = UT 'manage_excluded_folders' }
    if ($script:grpModeSummary) { $script:grpModeSummary.Text = UT 'mode_summary_group' }
    if ($script:grpCurrentStatus) { $script:grpCurrentStatus.Text = UT 'status_group' }
    if ($script:grpMainProgress) { $script:grpMainProgress.Text = UT 'main_progress' }
    if ($script:grpAdvancedProgress) { $script:grpAdvancedProgress.Text = UT 'advanced_details' }
    $script:chkDryRun.Text = UT 'dryrun'
    $script:chkApply.Text = UT 'apply_changes'
    $script:chkRepairExif.Text = UT 'repair_exif'
    $script:chkCopyInstead.Text = UT 'copy_instead'
    $script:chkOpenOutput.Text = UT 'open_output'
    $script:chkOpenNeeds.Text = UT 'open_needs'
    $script:chkOpenDuplicates.Text = UT 'open_duplicates'
    $script:chkOpenLog.Text = UT 'open_log'
    $script:chkDiagnostic.Text = UT 'diagnostic_logging'
    $script:chkKeepEmpty.Text = UT 'keep_empty_folders'
    $script:chkRenameInternal.Text = UT 'translate_internal'
    $script:lblPerformance.Text = T 'performance_mode'
    $script:lblMaxParallel.Text = UT 'max_workers'
    $script:btnTestScan.Text = UT 'test_scan'
    $script:btnDryRun.Text = UT 'dryrun'
    $script:btnRun.Text = T 'btn_run'
    $script:btnRunExif.Text = T 'btn_run_exif'
    if ($script:btnExecute) { $script:btnExecute.Text = UT 'execute' }
    if ($script:btnReconcile) { $script:btnReconcile.Text = UT 'reconcile_index' }
    if ($script:btnPurgeMissing) { $script:btnPurgeMissing.Text = UT 'purge_missing_index' }
    if ($script:btnAdvancedModes) { $script:btnAdvancedModes.Text = if ($script:advancedModesGroup -and $script:advancedModesGroup.Visible) { UT 'advanced_modes_hide' } else { UT 'advanced_modes_show' } }
    if ($script:advancedModesGroup) { $script:advancedModesGroup.Text = UT 'advanced_modes_group' }
    if ($script:lblAdvancedModesNote) { $script:lblAdvancedModesNote.Text = UT 'advanced_modes_note' }
    if ($script:btnRetentionCleanup) { $script:btnRetentionCleanup.Text = UT 'advanced_retention_cleanup' }
    if ($script:btnRecoverWrongDuplicateMove) { $script:btnRecoverWrongDuplicateMove.Text = UT 'advanced_recover_wrong_duplicate_move' }
    if ($script:btnRenameExistingFolders) { $script:btnRenameExistingFolders.Text = UT 'advanced_rename_existing_folders' }
    if ($script:btnRenameInternalFolders) { $script:btnRenameInternalFolders.Text = UT 'advanced_rename_internal_folders' }
    if ($script:btnNormalizeExistingFolders) { $script:btnNormalizeExistingFolders.Text = UT 'advanced_normalize_existing_folders' }
    if ($script:btnDedupeCleanup) { $script:btnDedupeCleanup.Text = UT 'advanced_dedupe_cleanup' }
    if ($script:btnRepairOnlyLibrary) { $script:btnRepairOnlyLibrary.Text = UT 'advanced_repair_only_library' }
    if ($script:btnMetadataAudit) { $script:btnMetadataAudit.Text = UT 'advanced_metadata_audit' }
    if ($script:btnMetadataRepair) { $script:btnMetadataRepair.Text = UT 'advanced_metadata_repair' }
    if ($script:btnImportGoogleTakeout) { $script:btnImportGoogleTakeout.Text = UT 'import_provider_google' }
    if ($script:btnImportApplePhotos) { $script:btnImportApplePhotos.Text = UT 'import_provider_apple' }
    if ($script:btnImportSamsungGallery) { $script:btnImportSamsungGallery.Text = UT 'import_provider_samsung' }
    if ($script:btnImportImmich) { $script:btnImportImmich.Text = UT 'import_provider_immich' }
    if ($script:btnImportXmpSidecar) { $script:btnImportXmpSidecar.Text = UT 'import_provider_xmp' }
    if ($script:lblImportGalleryTitle) { $script:lblImportGalleryTitle.Text = UT 'import_gallery_title' }
    if ($script:lblImportGalleryIntro) { $script:lblImportGalleryIntro.Text = UT 'import_gallery_intro' }
    if ($script:ImportProviderDescriptionLabels) {
        foreach ($entry in $script:ImportProviderDescriptionLabels.GetEnumerator()) {
            if ($entry.Value -and $entry.Value.Label) {
                $entry.Value.Label.Text = ('{0} - {1}' -f (UT ([string]$entry.Value.StatusKey)), (UT ([string]$entry.Key)))
            }
        }
    }
    if ($script:btnMigrateUdmrs) { $script:btnMigrateUdmrs.Text = UT 'advanced_migrate_udmrs' }
    if ($script:AdvancedModeDescriptionLabels) {
        foreach ($entry in $script:AdvancedModeDescriptionLabels.GetEnumerator()) {
            if ($entry.Value) { $entry.Value.Text = UT ([string]$entry.Key) }
        }
    }
    $script:btnCancel.Text = UT 'cancel'
    $script:btnLogs.Text = UT 'open_logs_folder'
    if ($script:btnHelp) { $script:btnHelp.Text = UT 'help' }
    if ($script:btnToggleAdvancedTools) { $script:btnToggleAdvancedTools.Text = if ($script:advancedToolsGroup -and $script:advancedToolsGroup.Visible) { UT 'toggle_advanced_hide' } else { UT 'toggle_advanced_show' } }
    $script:btnSettings.Text = UT 'settings'
    $script:btnExit.Text = UT 'exit'
    $script:lblStatus.Text = T 'status'
    $script:lblLastMessage.Text = T 'last_message'
    $script:lblAnalyzed.Text = T 'files_analyzed'
    $script:lblDuplicates.Text = T 'duplicates_found'
    $script:lblExif.Text = T 'exif_repaired'
    $script:lblMoved.Text = T 'files_moved'
    $script:lblCopied.Text = T 'files_copied'
    $script:lblNeedsReview.Text = UT 'needs_review'
    $script:lblSkipped.Text = T 'skipped_files'
    $script:lblInaccessible.Text = T 'inaccessible_files'
    $script:lblLocked.Text = T 'locked_files'
    $script:lblRetries.Text = T 'retry_count'
    $script:lblMetadataCorrupt.Text = UT 'metadata_corrupted_media'
    $script:lblMetadataBackupSize.Text = UT 'metadata_backup_size'
    $script:lblPid.Text = T 'pid'
    $script:lblPhase.Text = T 'phase'
    $script:lblFps.Text = T 'files_per_second'
    $script:lblEta.Text = T 'eta'
    $script:lblElapsed.Text = T 'elapsed'
    $script:lblBatch.Text = T 'batch'
    $script:lblCpu.Text = T 'cpu_percent'
    $script:lblRam.Text = T 'ram_percent'
    $script:lblWorkers.Text = UT 'workers'
    $script:lblQueue.Text = T 'queue_size'
    $script:lblEmptyRemoved.Text = T 'empty_removed'
    $script:lblLogPath.Text = UT 'log_location'
    $script:grpSummary.Text = T 'final_summary'
    if ($script:lblEngineRuntime) {
        $script:lblEngineRuntime.Text = Get-EngineRuntimeMessage
        $script:lblEngineRuntime.ToolTipText = $script:lblEngineRuntime.Text
    }
    if ($script:lblFirstRunRecommendation) {
        $script:lblFirstRunRecommendation.Text = UT 'first_run'
        $script:lblFirstRunRecommendation.ToolTipText = UT 'first_run_tip'
    }
    if ($script:lblDownloadPowerShell) {
        $script:lblDownloadPowerShell.Text = UT 'download_ps7'
        $script:lblDownloadPowerShell.ToolTipText = UT 'download_ps7_tip'
    }
    if ($script:lblStatusValue.Text -eq '') { $script:lblStatusValue.Text = T 'ready' }
    Apply-OptionTooltips
    Update-FolderProtectionSummary
    Update-ModeSummary
    Update-ReconcileHint
}

function Apply-OptionTooltips {
    Set-ControlTooltip -Control $script:chkCopyInstead -Text (UT 'tip_copy')
    Set-ControlTooltip -Control $script:chkKeepEmpty -Text (UT 'tip_keep_empty')
    Set-ControlTooltip -Control $script:chkDiagnostic -Text (UT 'tip_diagnostic')
    Set-ControlTooltip -Control $script:chkOpenOutput -Text (UT 'tip_open_output')
    Set-ControlTooltip -Control $script:chkOpenNeeds -Text (UT 'tip_open_needs')
    Set-ControlTooltip -Control $script:chkOpenDuplicates -Text (UT 'tip_open_duplicates')
    Set-ControlTooltip -Control $script:chkOpenLog -Text (UT 'tip_open_log')
    Set-ControlTooltip -Control $script:chkRenameInternal -Text (UT 'tip_rename_internal')
    Set-ControlTooltip -Control $script:btnTestScan -Text (UT 'tip_test_scan')
    Set-ControlTooltip -Control $script:btnReconcile -Text (UT 'tip_reconcile_index')
    Set-ControlTooltip -Control $script:btnPurgeMissing -Text (UT 'tip_purge_missing')
    Set-ControlTooltip -Control $script:btnAdvancedModes -Text (UT 'tip_advanced_modes')
    Set-ControlTooltip -Control $script:btnRetentionCleanup -Text (UT 'advanced_retention_cleanup_desc')
    Set-ControlTooltip -Control $script:btnRecoverWrongDuplicateMove -Text (UT 'advanced_recover_wrong_duplicate_move_desc')
    Set-ControlTooltip -Control $script:btnRenameExistingFolders -Text (UT 'advanced_rename_existing_folders_desc')
    Set-ControlTooltip -Control $script:btnRenameInternalFolders -Text (UT 'advanced_rename_internal_folders_desc')
    Set-ControlTooltip -Control $script:btnNormalizeExistingFolders -Text (UT 'advanced_normalize_existing_folders_desc')
    Set-ControlTooltip -Control $script:btnDedupeCleanup -Text (UT 'advanced_dedupe_cleanup_desc')
    Set-ControlTooltip -Control $script:btnRepairOnlyLibrary -Text (UT 'advanced_repair_only_library_desc')
    Set-ControlTooltip -Control $script:btnMetadataAudit -Text (UT 'advanced_metadata_audit_desc')
    Set-ControlTooltip -Control $script:btnMetadataRepair -Text (UT 'advanced_metadata_repair_desc')
    Set-ControlTooltip -Control $script:btnImportGoogleTakeout -Text (UT 'import_provider_google_desc')
    Set-ControlTooltip -Control $script:btnImportApplePhotos -Text (UT 'import_provider_apple_desc')
    Set-ControlTooltip -Control $script:btnImportSamsungGallery -Text (UT 'import_provider_samsung_desc')
    Set-ControlTooltip -Control $script:btnImportImmich -Text (UT 'import_provider_immich_desc')
    Set-ControlTooltip -Control $script:btnImportXmpSidecar -Text (UT 'import_provider_xmp_desc')
    Set-ControlTooltip -Control $script:btnMigrateUdmrs -Text (UT 'advanced_migrate_udmrs_desc')
    Set-ControlTooltip -Control $script:btnLogs -Text (UT 'tip_logs')
    Set-ControlTooltip -Control $script:btnHelp -Text (UT 'help_tooltip')
}

$settings = Read-JsonFile -Path $script:SettingsPath
if ($null -eq $settings) {
    $settings = Read-JsonFile -Path $script:PreviousSettingsPath
}
$dashboardSettings = Read-JsonFile -Path $script:DashboardSettingsPath
if ($settings -and $settings.Language) {
    Load-Language -Code ([string]$settings.Language)
}
else {
    Load-Language -Code 'es'
    Show-LanguageSelector
    Load-Language -Code $script:CurrentLanguageCode
}

$script:form = New-Object System.Windows.Forms.Form
$script:form.StartPosition = 'CenterScreen'
$initialDashboardSize = Get-DashboardSafeSize -DesiredWidth 900 -DesiredHeight 800 -MinimumWidth 800 -MinimumHeight 600
$script:form.MinimumSize = $initialDashboardSize.MinimumSize
$script:form.ClientSize = $initialDashboardSize.ClientSize

$script:lblSource = New-Object System.Windows.Forms.Label
$script:lblSource.Location = New-Object System.Drawing.Point(15, 18)
$script:lblSource.AutoSize = $true

$script:txtSource = New-Object System.Windows.Forms.TextBox
$script:txtSource.Location = New-Object System.Drawing.Point(150, 15)
$script:txtSource.Width = 600
$script:txtSource.Text = Get-DefaultSourcePath

$script:btnBrowseSource = New-Object System.Windows.Forms.Button
$script:btnBrowseSource.Location = New-Object System.Drawing.Point(765, 13)
$script:btnBrowseSource.Width = 105
$script:btnBrowseSource.Add_Click({
    $chosen = Choose-Folder -InitialPath $script:txtSource.Text -Description (T 'select_source')
    if ($chosen) {
        $script:txtSource.Text = $chosen
        $script:txtDestination.Text = Get-DefaultDestinationPath -SourcePath $chosen
    }
})

$script:lblDestination = New-Object System.Windows.Forms.Label
$script:lblDestination.Location = New-Object System.Drawing.Point(15, 55)
$script:lblDestination.AutoSize = $true

$script:txtDestination = New-Object System.Windows.Forms.TextBox
$script:txtDestination.Location = New-Object System.Drawing.Point(150, 52)
$script:txtDestination.Width = 600
$script:txtDestination.Text = Get-DefaultDestinationPath -SourcePath $script:txtSource.Text

$script:btnBrowseDestination = New-Object System.Windows.Forms.Button
$script:btnBrowseDestination.Location = New-Object System.Drawing.Point(765, 50)
$script:btnBrowseDestination.Width = 105
$script:btnBrowseDestination.Add_Click({
    $chosen = Choose-Folder -InitialPath (Get-DestinationBase -DestinationPath $script:txtDestination.Text) -Description (T 'select_destination')
    if ($chosen) { $script:txtDestination.Text = $chosen }
})

$script:grpOptions = New-Object System.Windows.Forms.GroupBox
$script:grpOptions.Location = New-Object System.Drawing.Point(15, 92)
$script:grpOptions.Size = New-Object System.Drawing.Size(855, 175)

$script:chkDryRun = New-Object System.Windows.Forms.CheckBox
$script:chkDryRun.Location = New-Object System.Drawing.Point(15, 28)
$script:chkDryRun.Width = 190
$script:chkDryRun.Checked = $true
$script:chkDryRun.Add_CheckedChanged({
    if ($script:chkDryRun.Checked) {
        $script:chkApply.Checked = $false
    }
    Update-ModeSummary
    Update-ReconcileHint
})

$script:chkApply = New-Object System.Windows.Forms.CheckBox
$script:chkApply.Location = New-Object System.Drawing.Point(220, 28)
$script:chkApply.Width = 190
$script:chkApply.Add_CheckedChanged({
    if ($script:chkApply.Checked) { $script:chkDryRun.Checked = $false }
    if (-not $script:chkApply.Checked -and -not $script:chkDryRun.Checked) { $script:chkDryRun.Checked = $true }
    Update-ModeSummary
    Update-ReconcileHint
})

$script:chkRepairExif = New-Object System.Windows.Forms.CheckBox
$script:chkRepairExif.Location = New-Object System.Drawing.Point(425, 28)
$script:chkRepairExif.Width = 160
$script:chkRepairExif.Enabled = $true
$script:chkRepairExif.Add_CheckedChanged({ Update-ModeSummary })

$script:chkCopyInstead = New-Object System.Windows.Forms.CheckBox
$script:chkCopyInstead.Location = New-Object System.Drawing.Point(610, 28)
$script:chkCopyInstead.Width = 210
$script:chkCopyInstead.Add_CheckedChanged({ Update-ModeSummary })

$script:chkOpenOutput = New-Object System.Windows.Forms.CheckBox
$script:chkOpenOutput.Location = New-Object System.Drawing.Point(15, 65)
$script:chkOpenOutput.Width = 185

$script:chkOpenNeeds = New-Object System.Windows.Forms.CheckBox
$script:chkOpenNeeds.Location = New-Object System.Drawing.Point(220, 65)
$script:chkOpenNeeds.Width = 190

$script:chkOpenDuplicates = New-Object System.Windows.Forms.CheckBox
$script:chkOpenDuplicates.Location = New-Object System.Drawing.Point(425, 65)
$script:chkOpenDuplicates.Width = 230

$script:chkOpenLog = New-Object System.Windows.Forms.CheckBox
$script:chkOpenLog.Location = New-Object System.Drawing.Point(15, 100)
$script:chkOpenLog.Width = 170

$script:chkDiagnostic = New-Object System.Windows.Forms.CheckBox
$script:chkDiagnostic.Location = New-Object System.Drawing.Point(220, 100)
$script:chkDiagnostic.Width = 180

$script:chkKeepEmpty = New-Object System.Windows.Forms.CheckBox
$script:chkKeepEmpty.Location = New-Object System.Drawing.Point(425, 100)
$script:chkKeepEmpty.Width = 190

$script:lblPerformance = New-Object System.Windows.Forms.Label
$script:lblPerformance.Location = New-Object System.Drawing.Point(15, 136)
$script:lblPerformance.Width = 130

$script:cmbPerformance = New-Object System.Windows.Forms.ComboBox
$script:cmbPerformance.DropDownStyle = 'DropDownList'
$script:cmbPerformance.Items.AddRange(@('Safe', 'Balanced', 'HighPerformance'))
$script:cmbPerformance.SelectedItem = 'Balanced'
$script:cmbPerformance.Location = New-Object System.Drawing.Point(150, 132)
$script:cmbPerformance.Width = 155

$script:lblMaxParallel = New-Object System.Windows.Forms.Label
$script:lblMaxParallel.Location = New-Object System.Drawing.Point(330, 136)
$script:lblMaxParallel.Width = 160

$script:txtMaxParallel = New-Object System.Windows.Forms.TextBox
$script:txtMaxParallel.Location = New-Object System.Drawing.Point(500, 132)
$script:txtMaxParallel.Width = 65
$script:txtMaxParallel.Add_TextChanged({ Update-ModeSummary })

$script:chkRenameInternal = New-Object System.Windows.Forms.CheckBox
$script:chkRenameInternal.Location = New-Object System.Drawing.Point(585, 132)
$script:chkRenameInternal.Width = 255
$script:chkRenameInternal.Visible = $false
$script:chkRenameInternal.Add_CheckedChanged({ Save-Settings })

$script:grpOptions.Controls.AddRange(@(
    $script:chkDryRun, $script:chkApply, $script:chkRepairExif, $script:chkCopyInstead,
    $script:chkOpenOutput, $script:chkOpenNeeds,
    $script:chkOpenDuplicates, $script:chkOpenLog, $script:chkDiagnostic, $script:chkKeepEmpty,
    $script:lblPerformance, $script:cmbPerformance, $script:lblMaxParallel, $script:txtMaxParallel
))

$buttonY = 282
$script:btnTestScan = New-Object System.Windows.Forms.Button
$script:btnTestScan.Location = New-Object System.Drawing.Point(15, $buttonY)
$script:btnTestScan.Size = New-Object System.Drawing.Size(105, 34)
$script:btnTestScan.Add_Click({
    Set-SelectedActionHint -ActionKey 'test_scan'
    Start-OrganizerRun -UseApply $false -UseRepairExif $false -UseTestScan $true
})

$script:btnReconcile = New-Object System.Windows.Forms.Button
$script:btnReconcile.Location = New-Object System.Drawing.Point(15, ($buttonY + 42))
$script:btnReconcile.Size = New-Object System.Drawing.Size(230, 34)
$script:btnReconcile.Add_Click({
    Set-SelectedActionHint -ActionKey 'reconcile'
    Start-OrganizerRun -UseApply $script:chkApply.Checked -UseRepairExif $false -UseReconcile $true
})

$script:btnPurgeMissing = New-Object System.Windows.Forms.Button
$script:btnPurgeMissing.Location = New-Object System.Drawing.Point(255, ($buttonY + 42))
$script:btnPurgeMissing.Size = New-Object System.Drawing.Size(210, 34)
$script:btnPurgeMissing.Add_Click({
    Set-SelectedActionHint -ActionKey 'purge_missing'
    Start-OrganizerRun -UseApply $script:chkApply.Checked -UseRepairExif $false -UsePurgeMissing $true
})

$script:btnDryRun = New-Object System.Windows.Forms.Button
$script:btnDryRun.Location = New-Object System.Drawing.Point(130, $buttonY)
$script:btnDryRun.Size = New-Object System.Drawing.Size(120, 34)
$script:btnDryRun.Add_Click({ Start-OrganizerRun -UseApply $false -UseRepairExif $false })

$script:btnRun = New-Object System.Windows.Forms.Button
$script:btnRun.Location = New-Object System.Drawing.Point(260, $buttonY)
$script:btnRun.Size = New-Object System.Drawing.Size(120, 34)
$script:btnRun.Add_Click({ Start-OrganizerRun -UseApply $script:chkApply.Checked -UseRepairExif $false })

$script:btnRunExif = New-Object System.Windows.Forms.Button
$script:btnRunExif.Location = New-Object System.Drawing.Point(390, $buttonY)
$script:btnRunExif.Size = New-Object System.Drawing.Size(165, 34)
$script:btnRunExif.Add_Click({ Start-OrganizerRun -UseApply $script:chkApply.Checked -UseRepairExif $true })

$script:btnCancel = New-Object System.Windows.Forms.Button
$script:btnCancel.Location = New-Object System.Drawing.Point(565, $buttonY)
$script:btnCancel.Size = New-Object System.Drawing.Size(90, 34)
$script:btnCancel.Enabled = $false
$script:btnCancel.Add_Click({ Stop-OrganizerRun })

$script:btnLogs = New-Object System.Windows.Forms.Button
$script:btnLogs.Location = New-Object System.Drawing.Point(665, $buttonY)
$script:btnLogs.Size = New-Object System.Drawing.Size(75, 34)
$script:btnLogs.Add_Click({ Open-Folder -Path $script:LogRoot })

$script:btnSettings = New-Object System.Windows.Forms.Button
$script:btnSettings.Location = New-Object System.Drawing.Point(750, $buttonY)
$script:btnSettings.Size = New-Object System.Drawing.Size(60, 34)
$script:btnSettings.Add_Click({ Show-Settings })

$script:btnHelp = New-Object System.Windows.Forms.Button
$script:btnHelp.Location = New-Object System.Drawing.Point(750, ($buttonY + 42))
$script:btnHelp.Size = New-Object System.Drawing.Size(80, 34)
$script:btnHelp.Add_Click({ Open-HelpManual })

$script:btnExit = New-Object System.Windows.Forms.Button
$script:btnExit.Location = New-Object System.Drawing.Point(820, $buttonY)
$script:btnExit.Size = New-Object System.Drawing.Size(50, 34)
$script:btnExit.Add_Click({ $script:form.Close() })

$script:btnExecute = New-Object System.Windows.Forms.Button
$script:btnExecute.Text = UT 'execute'
$script:btnExecute.MinimumSize = New-Object System.Drawing.Size(180, 48)
$script:btnExecute.AutoSize = $true
$script:btnExecute.Add_Click({
    Set-SelectedActionHint -ActionKey 'execute'
    Start-OrganizerRun -UseApply $script:chkApply.Checked -UseRepairExif $script:chkRepairExif.Checked
})

$script:btnReconcile.Text = UT 'reconcile_index'
$script:btnReconcile.AutoSize = $true
$script:btnPurgeMissing.Text = UT 'purge_missing_index'
$script:btnPurgeMissing.AutoSize = $true

$script:btnExecute.Add_MouseEnter({ Set-SelectedActionHint -ActionKey 'execute' })
$script:btnExecute.Add_GotFocus({ Set-SelectedActionHint -ActionKey 'execute' })
$script:btnTestScan.Add_MouseEnter({ Set-SelectedActionHint -ActionKey 'test_scan' })
$script:btnTestScan.Add_GotFocus({ Set-SelectedActionHint -ActionKey 'test_scan' })
$script:btnReconcile.Add_MouseEnter({ Set-SelectedActionHint -ActionKey 'reconcile' })
$script:btnReconcile.Add_GotFocus({ Set-SelectedActionHint -ActionKey 'reconcile' })
$script:btnPurgeMissing.Add_MouseEnter({ Set-SelectedActionHint -ActionKey 'purge_missing' })
$script:btnPurgeMissing.Add_GotFocus({ Set-SelectedActionHint -ActionKey 'purge_missing' })

$script:btnToggleAdvancedTools = New-Object System.Windows.Forms.Button
$script:btnToggleAdvancedTools.Text = UT 'toggle_advanced_show'
$script:btnToggleAdvancedTools.MinimumSize = New-Object System.Drawing.Size(210, 40)
$script:btnToggleAdvancedTools.AutoSize = $true
$script:btnToggleAdvancedTools.Add_Click({ Toggle-AdvancedTools })

$script:btnAdvancedModes = New-Object System.Windows.Forms.Button
$script:btnAdvancedModes.Text = UT 'advanced_modes_show'
$script:btnAdvancedModes.MinimumSize = New-Object System.Drawing.Size(190, 40)
$script:btnAdvancedModes.AutoSize = $true
$script:btnAdvancedModes.Add_Click({ Toggle-AdvancedModes })

$script:btnRetentionCleanup = New-Object System.Windows.Forms.Button
$script:btnRetentionCleanup.AutoSize = $true
$script:btnRetentionCleanup.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnRetentionCleanup.Add_Click({
    Start-AdvancedDashboardRun -ActionKey 'advanced_retention_cleanup' -LabelKey 'advanced_retention_cleanup' -Switches @('-RetentionCleanup') -RunDashboardRetentionCleanup $true
})

$script:btnRecoverWrongDuplicateMove = New-Object System.Windows.Forms.Button
$script:btnRecoverWrongDuplicateMove.AutoSize = $true
$script:btnRecoverWrongDuplicateMove.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnRecoverWrongDuplicateMove.Add_Click({
    Start-AdvancedDashboardRun -ActionKey 'advanced_recover_wrong_duplicate_move' -LabelKey 'advanced_recover_wrong_duplicate_move' -Switches @('-RecoverFromWrongDuplicateMove')
})

$script:btnRenameExistingFolders = New-Object System.Windows.Forms.Button
$script:btnRenameExistingFolders.AutoSize = $true
$script:btnRenameExistingFolders.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnRenameExistingFolders.Add_Click({
    Start-AdvancedDashboardRun -ActionKey 'advanced_rename_existing_folders' -LabelKey 'advanced_rename_existing_folders' -Switches @('-RenameExistingFoldersToCurrentLanguage')
})

$script:btnRenameInternalFolders = New-Object System.Windows.Forms.Button
$script:btnRenameInternalFolders.AutoSize = $true
$script:btnRenameInternalFolders.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnRenameInternalFolders.Add_Click({
    Start-AdvancedDashboardRun -ActionKey 'advanced_rename_internal_folders' -LabelKey 'advanced_rename_internal_folders' -Switches @('-RenameInternalFoldersToCurrentLanguage')
})

$script:btnNormalizeExistingFolders = New-Object System.Windows.Forms.Button
$script:btnNormalizeExistingFolders.AutoSize = $true
$script:btnNormalizeExistingFolders.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnNormalizeExistingFolders.Add_Click({
    Start-AdvancedDashboardRun -ActionKey 'advanced_normalize_existing_folders' -LabelKey 'advanced_normalize_existing_folders' -Switches @('-NormalizeExistingFolders') -AllowKeepEmptyFolders $true
})

$script:btnDedupeCleanup = New-Object System.Windows.Forms.Button
$script:btnDedupeCleanup.AutoSize = $true
$script:btnDedupeCleanup.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnDedupeCleanup.Add_Click({
    Start-AdvancedDashboardRun -ActionKey 'advanced_dedupe_cleanup' -LabelKey 'advanced_dedupe_cleanup' -Switches @('-DedupeCleanup')
})

$script:btnRepairOnlyLibrary = New-Object System.Windows.Forms.Button
$script:btnRepairOnlyLibrary.AutoSize = $true
$script:btnRepairOnlyLibrary.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnRepairOnlyLibrary.Add_Click({
    Start-AdvancedDashboardRun -ActionKey 'advanced_repair_only_library' -LabelKey 'advanced_repair_only_library' -Switches @('-RepairOnlyExistingOrganizedLibrary') -UseRepairExif $true
})

$script:btnMetadataAudit = New-Object System.Windows.Forms.Button
$script:btnMetadataAudit.AutoSize = $true
$script:btnMetadataAudit.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnMetadataAudit.Add_Click({
    Start-AdvancedDashboardRun -ActionKey 'advanced_metadata_audit' -LabelKey 'advanced_metadata_audit' -Switches @('-MetadataAudit') -ForceDryRun $true
})

$script:btnMetadataRepair = New-Object System.Windows.Forms.Button
$script:btnMetadataRepair.AutoSize = $true
$script:btnMetadataRepair.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnMetadataRepair.Add_Click({
    Start-AdvancedDashboardRun -ActionKey 'advanced_metadata_repair' -LabelKey 'advanced_metadata_repair' -Switches @('-MetadataRepair')
})

$script:btnImportGoogleTakeout = New-Object System.Windows.Forms.Button
$script:btnImportGoogleTakeout.AutoSize = $true
$script:btnImportGoogleTakeout.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnImportGoogleTakeout.Add_Click({
    Start-GoogleTakeoutImportWizard
})

$script:btnImportApplePhotos = New-Object System.Windows.Forms.Button
$script:btnImportApplePhotos.AutoSize = $true
$script:btnImportApplePhotos.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnImportApplePhotos.Add_Click({
    Start-ApplePhotosImportWizard
})

$script:btnImportSamsungGallery = New-Object System.Windows.Forms.Button
$script:btnImportSamsungGallery.AutoSize = $true
$script:btnImportSamsungGallery.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnImportSamsungGallery.Add_Click({
    Show-PlannedImportProviderNotice -ProviderLabelKey 'import_provider_samsung' -StatusKey 'import_provider_status_coming_sample'
})

$script:btnImportImmich = New-Object System.Windows.Forms.Button
$script:btnImportImmich.AutoSize = $true
$script:btnImportImmich.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnImportImmich.Add_Click({
    Show-PlannedImportProviderNotice -ProviderLabelKey 'import_provider_immich' -StatusKey 'import_provider_status_coming_sample'
})

$script:btnImportXmpSidecar = New-Object System.Windows.Forms.Button
$script:btnImportXmpSidecar.AutoSize = $true
$script:btnImportXmpSidecar.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnImportXmpSidecar.Add_Click({
    Start-XmpSidecarImportWizard
})

$script:btnMigrateUdmrs = New-Object System.Windows.Forms.Button
$script:btnMigrateUdmrs.AutoSize = $true
$script:btnMigrateUdmrs.MinimumSize = New-Object System.Drawing.Size(220, 40)
$script:btnMigrateUdmrs.Add_Click({
    Start-UdmrsMigrationPackage
})

foreach ($binding in @(
    @($script:btnRetentionCleanup, 'advanced_retention_cleanup'),
    @($script:btnRecoverWrongDuplicateMove, 'advanced_recover_wrong_duplicate_move'),
    @($script:btnRenameExistingFolders, 'advanced_rename_existing_folders'),
    @($script:btnRenameInternalFolders, 'advanced_rename_internal_folders'),
    @($script:btnNormalizeExistingFolders, 'advanced_normalize_existing_folders'),
    @($script:btnDedupeCleanup, 'advanced_dedupe_cleanup'),
    @($script:btnRepairOnlyLibrary, 'advanced_repair_only_library'),
    @($script:btnMetadataAudit, 'advanced_metadata_audit'),
    @($script:btnMetadataRepair, 'advanced_metadata_repair'),
    @($script:btnImportGoogleTakeout, 'import_provider_google'),
    @($script:btnImportApplePhotos, 'import_provider_apple'),
    @($script:btnImportSamsungGallery, 'import_gallery_coming'),
    @($script:btnImportImmich, 'import_gallery_coming'),
    @($script:btnImportXmpSidecar, 'import_provider_xmp'),
    @($script:btnMigrateUdmrs, 'advanced_migrate_udmrs')
)) {
    $button = $binding[0]
    $actionKey = [string]$binding[1]
    if ($button) {
        $button.Add_MouseEnter({ Set-SelectedActionHint -ActionKey $actionKey }.GetNewClosure())
        $button.Add_GotFocus({ Set-SelectedActionHint -ActionKey $actionKey }.GetNewClosure())
    }
}

$script:txtModeSummary = New-Object System.Windows.Forms.TextBox
$script:txtModeSummary.Multiline = $true
$script:txtModeSummary.ReadOnly = $true
$script:txtModeSummary.ScrollBars = 'Vertical'
$script:txtModeSummary.MinimumSize = New-Object System.Drawing.Size(260, 150)

$script:lblStatus = New-Object System.Windows.Forms.Label
$script:lblStatus.Location = New-Object System.Drawing.Point(15, 330)
$script:lblStatus.AutoSize = $true
$script:lblStatusValue = New-Object System.Windows.Forms.Label
$script:lblStatusValue.Location = New-Object System.Drawing.Point(150, 330)
$script:lblStatusValue.Width = 140

$script:lblPid = New-Object System.Windows.Forms.Label
$script:lblPid.Location = New-Object System.Drawing.Point(300, 330)
$script:lblPid.AutoSize = $true
$script:txtPid = New-Object System.Windows.Forms.TextBox
$script:txtPid.Location = New-Object System.Drawing.Point(340, 326)
$script:txtPid.Width = 75
$script:txtPid.ReadOnly = $true

$script:progress = New-Object System.Windows.Forms.ProgressBar
$script:progress.Location = New-Object System.Drawing.Point(425, 326)
$script:progress.Size = New-Object System.Drawing.Size(445, 22)

$script:lblLastMessage = New-Object System.Windows.Forms.Label
$script:lblLastMessage.Location = New-Object System.Drawing.Point(15, 365)
$script:lblLastMessage.AutoSize = $true
$script:txtLastMessage = New-Object System.Windows.Forms.TextBox
$script:txtLastMessage.Location = New-Object System.Drawing.Point(150, 362)
$script:txtLastMessage.Width = 720
$script:txtLastMessage.ReadOnly = $true

$metricsY = 405
$metricW = 118
$labels = @()

$script:lblAnalyzed = New-Object System.Windows.Forms.Label
$script:lblAnalyzed.Location = New-Object System.Drawing.Point(15, $metricsY)
$script:lblAnalyzed.Width = $metricW
$script:txtAnalyzed = New-Object System.Windows.Forms.TextBox
$script:txtAnalyzed.Location = New-Object System.Drawing.Point(15, ($metricsY + 23))
$script:txtAnalyzed.Width = $metricW
$script:txtAnalyzed.ReadOnly = $true
$script:txtAnalyzed.Text = '0'

$script:lblDuplicates = New-Object System.Windows.Forms.Label
$script:lblDuplicates.Location = New-Object System.Drawing.Point(150, $metricsY)
$script:lblDuplicates.Width = $metricW
$script:txtDuplicates = New-Object System.Windows.Forms.TextBox
$script:txtDuplicates.Location = New-Object System.Drawing.Point(150, ($metricsY + 23))
$script:txtDuplicates.Width = $metricW
$script:txtDuplicates.ReadOnly = $true
$script:txtDuplicates.Text = '0'

$script:lblExif = New-Object System.Windows.Forms.Label
$script:lblExif.Location = New-Object System.Drawing.Point(285, $metricsY)
$script:lblExif.Width = $metricW
$script:txtExif = New-Object System.Windows.Forms.TextBox
$script:txtExif.Location = New-Object System.Drawing.Point(285, ($metricsY + 23))
$script:txtExif.Width = $metricW
$script:txtExif.ReadOnly = $true
$script:txtExif.Text = '0'

$script:lblMoved = New-Object System.Windows.Forms.Label
$script:lblMoved.Location = New-Object System.Drawing.Point(420, $metricsY)
$script:lblMoved.Width = $metricW
$script:txtMoved = New-Object System.Windows.Forms.TextBox
$script:txtMoved.Location = New-Object System.Drawing.Point(420, ($metricsY + 23))
$script:txtMoved.Width = $metricW
$script:txtMoved.ReadOnly = $true
$script:txtMoved.Text = '0'

$script:lblCopied = New-Object System.Windows.Forms.Label
$script:lblCopied.Location = New-Object System.Drawing.Point(555, $metricsY)
$script:lblCopied.Width = $metricW
$script:txtCopied = New-Object System.Windows.Forms.TextBox
$script:txtCopied.Location = New-Object System.Drawing.Point(555, ($metricsY + 23))
$script:txtCopied.Width = $metricW
$script:txtCopied.ReadOnly = $true
$script:txtCopied.Text = '0'

$script:lblNeedsReview = New-Object System.Windows.Forms.Label
$script:lblNeedsReview.Location = New-Object System.Drawing.Point(690, $metricsY)
$script:lblNeedsReview.Width = 180
$script:txtNeedsReview = New-Object System.Windows.Forms.TextBox
$script:txtNeedsReview.Location = New-Object System.Drawing.Point(690, ($metricsY + 23))
$script:txtNeedsReview.Width = 180
$script:txtNeedsReview.ReadOnly = $true
$script:txtNeedsReview.Text = '0'

$script:lblPhase = New-Object System.Windows.Forms.Label
$script:lblPhase.Location = New-Object System.Drawing.Point(15, 465)
$script:lblPhase.Width = 95
$script:txtPhase = New-Object System.Windows.Forms.TextBox
$script:txtPhase.Location = New-Object System.Drawing.Point(110, 462)
$script:txtPhase.Width = 120
$script:txtPhase.ReadOnly = $true

$script:lblFps = New-Object System.Windows.Forms.Label
$script:lblFps.Location = New-Object System.Drawing.Point(245, 465)
$script:lblFps.Width = 95
$script:txtFilesPerSecond = New-Object System.Windows.Forms.TextBox
$script:txtFilesPerSecond.Location = New-Object System.Drawing.Point(340, 462)
$script:txtFilesPerSecond.Width = 70
$script:txtFilesPerSecond.ReadOnly = $true

$script:lblEta = New-Object System.Windows.Forms.Label
$script:lblEta.Location = New-Object System.Drawing.Point(420, 465)
$script:lblEta.Width = 45
$script:txtEta = New-Object System.Windows.Forms.TextBox
$script:txtEta.Location = New-Object System.Drawing.Point(465, 462)
$script:txtEta.Width = 95
$script:txtEta.ReadOnly = $true

$script:lblElapsed = New-Object System.Windows.Forms.Label
$script:lblElapsed.Location = New-Object System.Drawing.Point(570, 465)
$script:lblElapsed.Width = 80
$script:txtElapsed = New-Object System.Windows.Forms.TextBox
$script:txtElapsed.Location = New-Object System.Drawing.Point(650, 462)
$script:txtElapsed.Width = 95
$script:txtElapsed.ReadOnly = $true

$script:lblBatch = New-Object System.Windows.Forms.Label
$script:lblBatch.Location = New-Object System.Drawing.Point(755, 465)
$script:lblBatch.Width = 50
$script:txtBatch = New-Object System.Windows.Forms.TextBox
$script:txtBatch.Location = New-Object System.Drawing.Point(805, 462)
$script:txtBatch.Width = 53
$script:txtBatch.ReadOnly = $true

$script:lblSkipped = New-Object System.Windows.Forms.Label
$script:lblSkipped.Location = New-Object System.Drawing.Point(15, 497)
$script:lblSkipped.Width = 105
$script:txtSkipped = New-Object System.Windows.Forms.TextBox
$script:txtSkipped.Location = New-Object System.Drawing.Point(120, 494)
$script:txtSkipped.Width = 135
$script:txtSkipped.ReadOnly = $true
$script:txtSkipped.Text = '0'

$script:lblInaccessible = New-Object System.Windows.Forms.Label
$script:lblInaccessible.Location = New-Object System.Drawing.Point(275, 497)
$script:lblInaccessible.Width = 120
$script:txtInaccessible = New-Object System.Windows.Forms.TextBox
$script:txtInaccessible.Location = New-Object System.Drawing.Point(400, 494)
$script:txtInaccessible.Width = 135
$script:txtInaccessible.ReadOnly = $true
$script:txtInaccessible.Text = '0'

$script:lblEmptyRemoved = New-Object System.Windows.Forms.Label
$script:lblEmptyRemoved.Location = New-Object System.Drawing.Point(555, 497)
$script:lblEmptyRemoved.Width = 170
$script:txtEmptyRemoved = New-Object System.Windows.Forms.TextBox
$script:txtEmptyRemoved.Location = New-Object System.Drawing.Point(730, 494)
$script:txtEmptyRemoved.Width = 128
$script:txtEmptyRemoved.ReadOnly = $true
$script:txtEmptyRemoved.Text = '0'

$lockedY = 532
$script:lblLocked = New-Object System.Windows.Forms.Label
$script:lblLocked.Location = New-Object System.Drawing.Point(15, $lockedY)
$script:lblLocked.Width = 105
$script:txtLocked = New-Object System.Windows.Forms.TextBox
$script:txtLocked.Location = New-Object System.Drawing.Point(120, ($lockedY - 3))
$script:txtLocked.Width = 135
$script:txtLocked.ReadOnly = $true
$script:txtLocked.Text = '0'

$script:lblRetries = New-Object System.Windows.Forms.Label
$script:lblRetries.Location = New-Object System.Drawing.Point(200, $lockedY)
$script:lblRetries.Width = 90
$script:txtRetries = New-Object System.Windows.Forms.TextBox
$script:txtRetries.Location = New-Object System.Drawing.Point(295, ($lockedY - 3))
$script:txtRetries.Width = 80
$script:txtRetries.ReadOnly = $true
$script:txtRetries.Text = '0'

$script:lblMetadataCorrupt = New-Object System.Windows.Forms.Label
$script:lblMetadataCorrupt.Location = New-Object System.Drawing.Point(390, $lockedY)
$script:lblMetadataCorrupt.Width = 155
$script:txtMetadataCorrupt = New-Object System.Windows.Forms.TextBox
$script:txtMetadataCorrupt.Location = New-Object System.Drawing.Point(550, ($lockedY - 3))
$script:txtMetadataCorrupt.Width = 80
$script:txtMetadataCorrupt.ReadOnly = $true
$script:txtMetadataCorrupt.Text = '0'

$script:lblMetadataBackupSize = New-Object System.Windows.Forms.Label
$script:lblMetadataBackupSize.Location = New-Object System.Drawing.Point(645, $lockedY)
$script:lblMetadataBackupSize.Width = 90
$script:txtMetadataBackupSize = New-Object System.Windows.Forms.TextBox
$script:txtMetadataBackupSize.Location = New-Object System.Drawing.Point(740, ($lockedY - 3))
$script:txtMetadataBackupSize.Width = 118
$script:txtMetadataBackupSize.ReadOnly = $true
$script:txtMetadataBackupSize.Text = '0.000'

$resourceY = 570
$script:lblCpu = New-Object System.Windows.Forms.Label
$script:lblCpu.Location = New-Object System.Drawing.Point(15, $resourceY)
$script:lblCpu.Width = 60
$script:txtCpu = New-Object System.Windows.Forms.TextBox
$script:txtCpu.Location = New-Object System.Drawing.Point(80, ($resourceY - 3))
$script:txtCpu.Width = 80
$script:txtCpu.ReadOnly = $true

$script:lblRam = New-Object System.Windows.Forms.Label
$script:lblRam.Location = New-Object System.Drawing.Point(175, $resourceY)
$script:lblRam.Width = 60
$script:txtRam = New-Object System.Windows.Forms.TextBox
$script:txtRam.Location = New-Object System.Drawing.Point(240, ($resourceY - 3))
$script:txtRam.Width = 80
$script:txtRam.ReadOnly = $true

$script:lblWorkers = New-Object System.Windows.Forms.Label
$script:lblWorkers.Location = New-Object System.Drawing.Point(345, $resourceY)
$script:lblWorkers.Width = 80
$script:txtWorkers = New-Object System.Windows.Forms.TextBox
$script:txtWorkers.Location = New-Object System.Drawing.Point(430, ($resourceY - 3))
$script:txtWorkers.Width = 105
$script:txtWorkers.ReadOnly = $true

$script:lblQueue = New-Object System.Windows.Forms.Label
$script:lblQueue.Location = New-Object System.Drawing.Point(555, $resourceY)
$script:lblQueue.Width = 80
$script:txtQueue = New-Object System.Windows.Forms.TextBox
$script:txtQueue.Location = New-Object System.Drawing.Point(640, ($resourceY - 3))
$script:txtQueue.Width = 218
$script:txtQueue.ReadOnly = $true

$script:lblLogPath = New-Object System.Windows.Forms.Label
$script:lblLogPath.Location = New-Object System.Drawing.Point(15, 608)
$script:lblLogPath.AutoSize = $true
$script:txtLogPath = New-Object System.Windows.Forms.TextBox
$script:txtLogPath.Location = New-Object System.Drawing.Point(150, 605)
$script:txtLogPath.Width = 720
$script:txtLogPath.ReadOnly = $true

$script:grpSummary = New-Object System.Windows.Forms.GroupBox
$script:grpSummary.Location = New-Object System.Drawing.Point(15, 642)
$script:grpSummary.Size = New-Object System.Drawing.Size(855, 135)
$script:txtSummary = New-Object System.Windows.Forms.TextBox
Set-LogTextBoxStyle -TextBox $script:txtSummary
$script:txtSummary.Location = New-Object System.Drawing.Point(10, 22)
$script:txtSummary.Size = New-Object System.Drawing.Size(835, 100)
$script:grpSummary.Controls.Add($script:txtSummary)

$script:statusStripInfo = New-Object System.Windows.Forms.StatusStrip
$script:statusStripInfo.Dock = [System.Windows.Forms.DockStyle]::Bottom
$script:statusStripInfo.SizingGrip = $false

$script:lblEngineRuntime = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:lblEngineRuntime.Text = Get-EngineRuntimeMessage
$script:lblEngineRuntime.ToolTipText = $script:lblEngineRuntime.Text

$script:lblFirstRunRecommendation = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:lblFirstRunRecommendation.Text = UT 'first_run'
$script:lblFirstRunRecommendation.ToolTipText = UT 'first_run_tip'
$script:lblFirstRunRecommendation.Spring = $true

$script:lblDownloadPowerShell = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:lblDownloadPowerShell.Text = UT 'download_ps7'
$script:lblDownloadPowerShell.IsLink = $true
$script:lblDownloadPowerShell.Visible = -not (Test-EngineHostIsPowerShell7 -EngineHost (Get-EngineHost))
$script:lblDownloadPowerShell.ToolTipText = UT 'download_ps7_tip'
$script:lblDownloadPowerShell.Add_Click({ Open-PowerShell7Download })

$script:statusStripInfo.Items.AddRange([System.Windows.Forms.ToolStripItem[]]@(
    $script:lblEngineRuntime,
    $script:lblFirstRunRecommendation,
    $script:lblDownloadPowerShell
))

$script:form.Controls.AddRange(@(
    $script:lblSource, $script:txtSource, $script:btnBrowseSource,
    $script:lblDestination, $script:txtDestination, $script:btnBrowseDestination,
    $script:grpOptions,
    $script:btnTestScan, $script:btnReconcile, $script:btnPurgeMissing, $script:btnDryRun, $script:btnRun, $script:btnRunExif, $script:btnCancel, $script:btnLogs, $script:btnHelp, $script:btnSettings, $script:btnExit,
    $script:lblStatus, $script:lblStatusValue, $script:lblPid, $script:txtPid, $script:progress,
    $script:lblLastMessage, $script:txtLastMessage,
    $script:lblAnalyzed, $script:txtAnalyzed, $script:lblDuplicates, $script:txtDuplicates,
    $script:lblExif, $script:txtExif, $script:lblMoved, $script:txtMoved,
    $script:lblCopied, $script:txtCopied, $script:lblNeedsReview, $script:txtNeedsReview,
    $script:lblPhase, $script:txtPhase, $script:lblFps, $script:txtFilesPerSecond,
    $script:lblEta, $script:txtEta, $script:lblElapsed, $script:txtElapsed,
    $script:lblBatch, $script:txtBatch,
    $script:lblSkipped, $script:txtSkipped, $script:lblInaccessible, $script:txtInaccessible,
    $script:lblEmptyRemoved, $script:txtEmptyRemoved,
    $script:lblLocked, $script:txtLocked, $script:lblRetries, $script:txtRetries,
    $script:lblMetadataCorrupt, $script:txtMetadataCorrupt,
    $script:lblMetadataBackupSize, $script:txtMetadataBackupSize,
    $script:lblCpu, $script:txtCpu, $script:lblRam, $script:txtRam,
    $script:lblWorkers, $script:txtWorkers, $script:lblQueue, $script:txtQueue,
    $script:lblLogPath, $script:txtLogPath, $script:grpSummary, $script:statusStripInfo
))

Build-ResponsiveLayout

if ($settings) {
    if ($settings.PerformanceMode -and @('Safe', 'Balanced', 'HighPerformance') -contains [string]$settings.PerformanceMode) {
        $script:cmbPerformance.SelectedItem = [string]$settings.PerformanceMode
    }
    if ($settings.MaxParallelJobs) {
        $script:txtMaxParallel.Text = [string]$settings.MaxParallelJobs
    }
    if ($settings.PSObject.Properties['RenameInternalFolders']) {
        $script:chkRenameInternal.Checked = [bool]$settings.RenameInternalFolders
    }
}

if ([string]::IsNullOrWhiteSpace($script:txtMaxParallel.Text)) {
    Update-MaxWorkerDisplay
}

Apply-DashboardWindowSettings -WindowSettings $dashboardSettings

$script:cmbPerformance.Add_SelectedIndexChanged({
    Update-MaxWorkerDisplay
    Save-Settings
})
$script:txtMaxParallel.Add_Leave({ Save-Settings })

Update-Texts
$script:lblStatusValue.Text = T 'ready'
$script:txtLastMessage.Text = T 'ready_message'
Update-ModeSummary

$script:form.Add_FormClosing({
    Save-DashboardWindowSettings
    if ($script:CurrentProcess -and -not $script:CurrentProcess.HasExited) {
        $answer = [System.Windows.Forms.MessageBox]::Show((T 'confirm_exit_running'), (T 'warning_title'), 'YesNo', 'Warning')
        if ($answer -ne 'Yes') {
            $_.Cancel = $true
            return
        }
        else {
            Stop-OrganizerRun
        }
    }
    Stop-TechnicalConsoleSession
})

[System.Windows.Forms.Application]::Run($script:form)








