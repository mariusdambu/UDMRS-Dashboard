# Security Policy

## Supported Version

This public package is prepared from the current UDMRS Dashboard source tree.

## Reporting Issues

When reporting a bug, do not attach real photo/video files, full logs, `ProcessedFiles.json`, or screenshots containing private paths unless you have reviewed and sanitized them.

Useful safe information:

- UDMRS build/version.
- Windows version.
- PowerShell version.
- Operation mode, for example DryRun or Apply.
- Sanitized command line.
- Sanitized summary counts.
- A minimal synthetic reproduction if possible.

## Sensitive Files

Never publish:

- `%APPDATA%\PhotoOrganizer\ProcessedFiles.json`
- `%APPDATA%\PhotoOrganizer\Logs`
- `%APPDATA%\PhotoOrganizer\IndexBackups`
- `%APPDATA%\PhotoOrganizer\Runtime`
- EXIF metadata backups
- duplicate quarantine folders
- Google Takeout exports
- personal media

## Operational Safety

UDMRS may manipulate local files in Apply mode. Always test with copies and maintain backups.

