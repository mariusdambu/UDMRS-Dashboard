# UDMRS Dashboard

Universal Digital Memory Recovery System

Recover · Repair · Organize · Preserve

UDMRS Dashboard is a Windows desktop tool for organizing, repairing and preserving photo/video libraries. It is designed for real-world personal archives: mixed phones, WhatsApp media, Google Takeout exports, cloud-synced folders, duplicates, missing metadata, manual moves, and old folder structures.

## Read This First

This project can move, rename, copy, quarantine and repair media files when you explicitly run it in Apply mode.

That means it can affect irreplaceable memories. Use it carefully.

Before using it on a real library:

1. Make a full backup of your photos/videos.
2. Run DryRun first.
3. Read the generated logs/reports.
4. Only then run Apply.
5. Do not run two UDMRS operations against the same library at the same time.

UDMRS is intentionally conservative, but no media organizer can guarantee safety if used without backups.

## What It Does

- Organizes photos and videos into a stable `Year\Quarter` structure.
- Detects exact duplicates by SHA256.
- Uses EXIF, media metadata, reliable filename dates, provider sidecars and fallback evidence to decide dates.
- Sends uncertain files to review folders instead of guessing aggressively.
- Repairs missing EXIF dates only when explicit repair conditions are met.
- Tracks processed files in `ProcessedFiles.json`.
- Reconciles the index after manual moves or renamed folders.
- Handles cloud placeholders conservatively: visible cloud-only files are skipped, not hydrated automatically.
- Supports Google Photos / Google Takeout imports.
- Supports Apple Photos / iCloud Photos imports.
- Supports generic XMP / JSON / YAML sidecar imports.

## What It Does Not Do

- It does not replace a backup strategy.
- It does not guarantee perfect metadata recovery.
- It does not automatically download cloud-only files from OneDrive, Dropbox, iCloud, Google Drive or similar providers.
- It does not delete photos silently.
- It does not treat probable visual duplicates as exact duplicates.
- It does not clean `_Duplicados_Para_Revisar` / `_Duplicates_To_Review` automatically.

## Current Status

Public release package prepared from:

```text
UDMRS Build 2026.05.30-SH-RC1
```

with later provider-import hardening for:

- Google Photos / Google Takeout
- Apple Photos / iCloud Photos
- XMP / Sidecar Library

The classic organizing flow remains the primary stable path. Provider import is available for exports with sidecars or provider metadata.

## Requirements

- Windows 10/11.
- Windows PowerShell 5.1 or PowerShell 7+.
- Optional but recommended: ExifTool for metadata reading/repair.

ExifTool is not bundled in this public source package. See [Tools/ExifTool/README.md](Tools/ExifTool/README.md).

## Quick Start

1. Download or clone this repository.
2. Place it anywhere, for example:

   ```text
   C:\Tools\UDMRS
   ```

3. If you want EXIF repair support, install ExifTool under:

   ```text
   Tools\ExifTool
   ```

4. Run:

   ```text
   Start-PhotoOrganizer.cmd
   ```

5. Select:

   ```text
   SourcePath
   DestinationPath
   Language
   DryRun
   ```

6. Review logs/reports before Apply.

## Main Workflows

### Organize Library

Use this for normal folders containing photos/videos.

```text
Source folder
↓
Organize
↓
Destination\Year\Quarter
```

### Import Gallery

Use this for provider exports or sidecar-rich libraries.

Available providers:

- Google Photos / Takeout
- Apple Photos / iCloud
- XMP / Sidecar Library

Planned / sample-gated providers:

- Samsung Gallery
- Immich

### Maintenance

Advanced operations include:

- ReconcileProcessedDatabase
- PurgeMissingFromProcessedDatabase
- NormalizeExistingFolders
- DedupeCleanup
- RepairOnlyExistingOrganizedLibrary
- MetadataAudit
- MetadataRepair
- RetentionCleanup

Use these only after reading the manuals.

## Cloud-Aware Behavior

UDMRS can work inside cloud-backed folders, but it follows this rule:

```text
Process local verified files.
Skip cloud-only placeholders.
Never hydrate large cloud libraries automatically.
```

If files are online-only, make them available offline first if you want deep EXIF/hash/dedupe validation.

## User Data Location

The shared application folder contains code and documentation.

Per-user state lives under:

```text
%APPDATA%\PhotoOrganizer
```

This may include:

- `ProcessedFiles.json`
- user exclusions
- dashboard settings
- logs
- runtime files
- index backups

Do not publish `%APPDATA%\PhotoOrganizer` contents to GitHub.

## Documentation

- [Safe Usage Guide](SAFE_USAGE.md)
- [Manual ES](Docs/Manuals/Manual_ES.md)
- [Manual RO](Docs/Manuals/Manual_RO.md)
- [Command Reference](Docs/CommandReference.html)
- [Visual Map](Docs/VisualMaps/UDMRS-VisualMap.html)
- [Branding](Branding/UDMRS-Branding.md)
- [Security Policy](SECURITY.md)
- [Third-party Notices](THIRD-PARTY-NOTICES.md)

## GitHub Publishing Notes

This repository should not include:

- personal photos or videos
- real `ProcessedFiles.json`
- logs
- progress files
- runtime files
- migration packages
- EXIF backups
- duplicate quarantines
- Google Takeout samples
- private OneDrive paths

The included `.gitignore` is intentionally strict.

## License

MIT License. See [LICENSE](LICENSE).


