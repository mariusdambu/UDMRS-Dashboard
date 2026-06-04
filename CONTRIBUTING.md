# Contributing

Thank you for helping improve UDMRS.

## Priorities

The project priorities are:

1. Preserve user memories.
2. Avoid destructive surprises.
3. Prefer DryRun/reporting before Apply behavior.
4. Keep cloud-only files from being hydrated accidentally.
5. Maintain clear logs and documentation.

## Contribution Rules

- Do not add behavior that deletes, overwrites or repairs media without explicit Apply semantics.
- Do not weaken duplicate detection from SHA256 exact matches without a separate review path.
- Do not add provider import logic based only on guesses; prefer real export samples or documented sidecar formats.
- Keep user-specific paths out of docs and tests.
- Use synthetic test data whenever possible.

## Documentation

Any behavioral change should update:

- README.md
- relevant manuals
- CommandReference.html if commands changed
- Visual map if workflows changed

