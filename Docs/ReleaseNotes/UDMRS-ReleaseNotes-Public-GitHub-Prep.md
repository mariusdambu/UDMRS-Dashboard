# UDMRS GitHub Publication Prep

Build baseline: `UDMRS Build 2026.05.30-SH-RC1`

Publication prep date: `2026-06-04`

Status: `Public source package prepared`

## Summary

This package prepares UDMRS Dashboard for GitHub publication with the stable classic organizing core and the current Import Gallery work.

## Included

- Classic Organize workflow.
- Official `Year\Quarter` organization model.
- Dashboard with normal mode, Import Gallery, maintenance and advanced tools.
- Google Photos / Google Takeout ImportProvider.
- XMP / Sidecar Library ImportProvider.
- Reconcile, Purge, Normalize, DedupeCleanup, RepairOnly and RetentionCleanup workflows.
- Cloud-aware placeholder skipping.
- User excluded/protected folders.
- Per-user state under `%APPDATA%\PhotoOrganizer`.
- Documentation, visual map, command reference and safety guidance.

## Not Included

- Personal galleries.
- Real logs.
- `ProcessedFiles.json`.
- Runtime state.
- Migration packages.
- Google Takeout samples.
- Bundled ExifTool binaries.

## Known Limitations

- ExifTool must be installed separately for full metadata repair workflows.
- Provider imports depend on export quality and sidecar metadata.
- Cloud-only files are intentionally skipped until local.
- Apple Photos / iCloud is available. Samsung Gallery and Immich providers are planned but not active.
- A public license still needs to be chosen before formal open-source publication.

## Validation Performed

- Syntax validation of `App\PhotoOrganizer.ps1`.
- Read-only Google Takeout audit on a larger sample.
- Synthetic XMP provider localization smoke tests.
- Publication package scan for private path patterns outside ignored runtime/temp areas.
