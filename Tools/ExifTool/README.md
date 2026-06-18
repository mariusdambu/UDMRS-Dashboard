# ExifTool

UDMRS can use ExifTool for metadata reading and repair, but the public source package does not bundle third-party binaries.

Install ExifTool manually if you want full metadata support:

1. Download ExifTool for Windows from the official ExifTool site.
2. Extract it under this folder or configure the dashboard to point to your ExifTool executable if supported by your local setup.
3. Keep the folder portable with the UDMRS installation when moving to another PC.

Expected portable layout example:

```text
Tools\ExifTool\exiftool-<version>_64\exiftool.exe
```

Without ExifTool, UDMRS can still run safe scans and some workflows, but deep EXIF/QuickTime reading and metadata repair will be limited.
