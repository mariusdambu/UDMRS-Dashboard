# Project Status

## Current Public Package

```text
UDMRS Dashboard
Build baseline: UDMRS Build 2026.05.30-SH-RC1
Prepared for GitHub publication: 2026-06-04
```

## Stable Core

- Organize classic library flow.
- Year\Quarter destination structure.
- Reconcile / Purge / Normalize / Dedupe / RepairOnly maintenance modes.
- Cloud-aware local-only processing.
- User excluded folders.
- Runtime/user state separated under `%APPDATA%\PhotoOrganizer`.

## Import Gallery

Available:

- Google Photos / Takeout
- XMP / Sidecar Library

Planned:

- Apple Photos / iCloud
- Samsung Gallery
- Immich

## Known Limits

- Provider support depends on export structure and sidecar quality.
- Cloud-only files are skipped until available locally.
- Visual/near duplicate detection is not the same as exact duplicate cleanup.
- A public license still needs to be chosen before formal open-source publication.

