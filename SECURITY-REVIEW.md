# Security Review

**Date:** 2026-10-26  
**Scope:** Static review of the admin login flow in `index.html`, `portal.html`, `index-live-backup.html`, and the related Google Apps Script authorization checks.

## Finding

| Severity | File | Lines | Finding | Impact |
|---|---|---:|---|---|
| Low | `index.html`, `portal.html`, `index-live-backup.html` | Login handler, around lines 1120-1180 | The dashboard shell was displayed while credentials were still being checked. | Admin controls could briefly appear before rejection. The review found no access to report data or successful admin actions during this interval because the Apps Script checks the session before serving admin data or processing protected changes. |

## Remediation

The local HTML copies now keep the login dialog visible and hide the dashboard until the authenticated report-data request succeeds. Failed login leaves the dashboard hidden. The updated scripts passed syntax and login-flow ordering checks.

The HTML changes are local and still need to be published to GitHub/Vercel. No Apps Script update is required for this UI change.
