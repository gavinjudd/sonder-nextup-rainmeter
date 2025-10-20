# Security & Privacy

This skin reads **private Google Calendar ICS URLs**. Treat them like passwords.

- **Never** commit or share your `secrets.ini` or private ICS links.
- The repo ships a `secrets.example.ini` template. Copy → `secrets.ini` locally.
- `.gitignore` excludes `secrets.ini`, `gcal_log.txt`, and `gcal_events.txt`.
- If you open an issue, redact personal data and ICS URLs from any logs.

If you think you’ve leaked a private ICS URL, rotate it from Google Calendar settings immediately.
