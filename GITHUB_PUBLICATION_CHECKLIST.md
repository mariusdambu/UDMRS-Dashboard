# GitHub Publication Checklist

Before making the repository public:

- [ ] Choose and add a `LICENSE` file.
- [ ] Confirm no personal photos/videos are present.
- [ ] Confirm no `ProcessedFiles.json` is present.
- [ ] Confirm no logs/progress/runtime files are present.
- [ ] Confirm no Google Takeout/iCloud/Samsung/Immich samples are present.
- [ ] Confirm no private paths are present.
- [ ] Confirm `.gitignore` is active before committing.
- [ ] Run syntax validation:

```powershell
$errors = $null
[System.Management.Automation.Language.Parser]::ParseFile(
  ".\App\PhotoOrganizer.ps1",
  [ref]$null,
  [ref]$errors
) > $null
$errors
```

- [ ] Run the dashboard from a disposable test folder.
- [ ] Run DryRun on synthetic test media.

