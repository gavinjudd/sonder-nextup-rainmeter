# Security and privacy

Private ICS addresses are bearer credentials: possession of a URL is enough to read its calendar.

## If an ICS address is exposed

1. Reset or rotate the calendar's **Secret address in iCal format** immediately.
2. Update only your local, ignored `Sonder/@Resources/secrets.ini`.
3. Remove the credential and any fetched calendar data from every reachable branch and tag.
4. If it entered Git history, follow GitHub's sensitive-data removal process and request cached-object garbage collection.

Deleting a file or making a repository private does not revoke an exposed URL.

## Reporting a vulnerability

Use [GitHub private vulnerability reporting](https://github.com/gavinjudd/sonder-nextup-rainmeter/security/advisories/new) for credential exposure or another security problem. Do not place credentials, calendar content, raw logs, or personal screenshots in a public issue.

For ordinary bug reports, provide versions, reproduction steps, and metadata-only diagnostics. Review every attachment for event titles, attendees, UIDs, email addresses, meeting links, phone numbers, local paths, and URLs.

## Repository controls

The validation and packaging scripts reject known runtime paths and private Google Calendar URL patterns. Local credentials, event output, logs, raw ICS files, and atomic-write scratch files are ignored by Git.
