# Changelog

## v1.1.0
- Multi-ICS support (merge feeds; primary fills first, then secondary).
- Robust ICS parsing:
  - Unfold folded lines (RFC 5545).
  - DTSTART parser for TZID (IANA→Windows), UTC `Z`, VALUE=DATE, floating local.
  - RRULE expansion (DAILY/WEEKLY with BYDAY/INTERVAL/UNTIL/COUNT) + combined EXDATE.
  - Per-instance edits/cancels via RECURRENCE-ID, including `RANGE=THISANDFUTURE`.
  - Tolerant fallbacks for vendor quirks; optional local filters in `secrets.ini`.
- Logging improvements; same output format for Rainmeter.
- README/docs and ignore rules updated.
