# Final portability review

A read-only reviewer confirmed that the historical absolute report path defines
one root shared by report, catalog and time-v command arguments. Relocated
artifacts are checked by byte length and SHA-256, and the rerun verifier checks
the retained summary digest. The original binary is never opened during replay.

The first review identified that the portable driver copied shared validators
from the current checkout without pinning them. The driver now uses four
bundled validators, checks their exact manifest name set and hashes before
export, and records a failed receipt on mismatch. Their bytes were also checked
against source revision `dc7cb687da2fc629695510b9610c97a0233f9df1`.

A second read-only review approved this change. The standalone bundle test
passes all three replay checks without repository layout, then rejects a
modified validator before replay. The reviewer requested a refreshed inventory;
the final SHA256SUMS includes all bundled modules, manifest and new checks.
No Cargo, tests or workloads were run by the reviewer.
