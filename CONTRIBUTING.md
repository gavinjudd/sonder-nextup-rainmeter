# Contributing

Thanks for helping improve **sonder-nextup-rainmeter**!

## How to contribute
1. **Fork** the repo and create a feature branch: `feat/short-name`.
2. Make changes and run the fetch once locally to verify `gcal_events.txt` and `gcal_log.txt` look right.
3. **Do not** commit `secrets.ini`, `gcal_events.txt`, or `gcal_log.txt`.
4. Open a **pull request** with a clear summary of changes + screenshots if UI is affected.

## Good PRs include
- What the change does and why.
- Before/after screenshots if visuals are touched.
- A short tail of `gcal_log.txt` to confirm parsing/selection.

## Code style
- PowerShell 5.1 compatible.
- Keep the script side-effect free outside `Sonder/@Resources/`.
- Log helpful, concise lines to `gcal_log.txt`.

## Security / privacy
- Never post private ICS URLs.
- Redact personal details from logs before attaching.
