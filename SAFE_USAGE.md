# Safe Usage Guide

UDMRS works with personal memories. Treat every Apply operation as serious.

## Golden Rules

1. Keep a full backup outside the library being processed.
2. Run DryRun before Apply.
3. Read the report, especially conflicts, NeedsReview, duplicates and warnings.
4. Do not run two operations against the same library at once.
5. Do not use Apply on cloud-only libraries unless you understand what is local and what is placeholder.

## Recommended First Test

Use a temporary copy with 20-100 files:

```text
TestPhotos
↓
DryRun
↓
Review report
↓
Apply on the copy
```

Only move to your real library after you understand the result.

## Modes With Real File Impact

These can move, rename, copy, repair, quarantine or delete operational artifacts when Apply is used:

- Organize Apply
- Organize Apply + RepairExif
- ImportProvider Apply
- NormalizeExistingFolders Apply
- DedupeCleanup Apply
- RepairOnlyExistingOrganizedLibrary Apply
- PurgeMissingFromProcessedDatabase Apply
- RetentionCleanup Apply
- Rename folder language modes with Apply

## Safe by Default

DryRun does not intentionally modify your media library. Use it heavily.

## What Needs Manual Review

Anything placed in NeedsReview or Duplicates_To_Review should be inspected by a person. UDMRS deliberately avoids guessing when evidence is weak.

