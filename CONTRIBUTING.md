# Contributing

Thanks for improving Sonder Next Up.

## Before opening a pull request

1. Create a focused branch from `main`.
2. Keep PowerShell compatible with Windows PowerShell 5.1.
3. Run:

   ```powershell
   .\Tests\validate_repo.ps1
   .\Tests\test_fetcher.ps1
   .\Scripts\pack.ps1
   ```

4. Include a synthetic before/after preview for visual changes.
5. Explain the user impact and any parser edge case covered.

## Privacy rules

Never commit or paste:

- `secrets.ini` or a real ICS URL;
- fetched `.ics` content;
- `gcal_events.txt` or `gcal_log.txt`;
- event titles, descriptions, attendees, meeting links, UIDs, locations, or phone numbers; or
- screenshots containing a real calendar.

Use `example.invalid`, synthetic names, and generated dates in tests. If a diagnostic log is necessary, inspect and redact it locally before attaching it.

## Scope

This repository is an overlay. Do not add fonts, binaries, images, plugins, or unrelated modules from the full Sonder distribution. See [NOTICE.md](NOTICE.md).

Preserve the attribution, change notice, and CC BY-NC-SA 3.0 metadata in `overlay/Calendar/Calendar.ini`. New Clock, PowerShell, test, installer, and documentation contributions are accepted under the repository's MIT license.
