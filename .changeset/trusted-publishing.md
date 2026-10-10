---
"highcharts-editable-excel": patch
---

Releases are now published from GitHub Actions through npm trusted publishing (OIDC): no npm token is stored anywhere, and every version from this one on carries a provenance attestation that links the tarball to the commit and workflow run that built it (`npm audit signatures` verifies it). No code changes.
