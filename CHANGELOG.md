# Changelog

## Unreleased

### Fixed

- Recurring instances removed with `EXDATE` no longer remain in the desktop agenda.
- Each exclusion now keeps its own timezone/value parameters and is compared with a normalized recurrence key.
- Recurrence date construction no longer inherits nondeterministic fractional seconds from the current clock.

### Tests

- Added offline regressions for excluded, cancelled, and moved recurring instances.

## v2.0.0 — 2026-08-06

### Added

- Non-admin overlay installer with timestamped backups.
- Offline multi-feed fixture test and repository privacy validation.
- Synthetic SVG preview and explicit upstream attribution.
- Automatic calendar fetch on skin load, wake, five-minute interval, and manual **SYNC**.
- Explicit MIT/CC BY-NC-SA 3.0 license mapping for every distributed file.

### Changed

- Redesigned the Calendar card with a continuous adaptive vector surface, stronger hierarchy, and four compact agenda rows.
- Redesigned the Clock as a transparent editorial stack aligned to the Windows mark.
- Events from every feed now form one strictly chronological queue; feed order is only a tie-breaker.
- Fetch output is written atomically so Rainmeter cannot observe a truncated file.
- Logs now contain operational metadata instead of event titles, UIDs, local paths, email addresses, or calendar URLs.
- Repository distribution is now a focused overlay that requires official Sonder instead of repackaging upstream assets.
- Public sources now live under `overlay/`; the installer applies them to the user's local Sonder copy.
- Replaced the inherited Clock configuration with a clean-room, system-font implementation under MIT.
- Made the Calendar derivative self-contained while retaining its upstream CC BY-NC-SA 3.0 license, attribution, and change notice.

### Removed

- Elevated Task Scheduler workflow and hardcoded-path task XML.
- Generated calendar caches, event output, logs, private screenshots, third-party binaries, fonts, and unrelated Sonder modules.
- Historical copies of the old `Sonder/` tree, including third-party binaries and unresolved Clock provenance.

### Security

- Purged previously tracked private ICS credentials and generated calendar data from branches and tags.
- Hardened ignore rules and release packaging against future credential or runtime-data commits.

## v1.1.0

- Added multiple ICS feeds and broader recurrence parsing.
- Added initial documentation and packaging scripts.

## v1.0.0

- Initial public release.
