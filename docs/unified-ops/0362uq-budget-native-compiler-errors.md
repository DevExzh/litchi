# 0362uq — native budget compiler repairs

The budget-validation native check at `c9bf9c8c40` reported the final
integration compiler batch in PPTX, XLSX, DOCX, and XLSB. The source repair
removes the duplicate PPTX durable size helper, routes topology sizing through
the associated helper, retains the `Arc` order representation, and migrates
metadata commit/reset to the current shared properties API. XLSX now uses the
current relationship planner and read-only source projections, and its tests
use their in-scope ledger types. DOCX uses `SourceXmlPart::bytes`, maps OPC
size errors through the crate error, keeps the final snapshot budget check
after compaction without borrowing a partially moved edit, and updates the
publication-proof test helpers for the explicit budget argument.

All existing policy values, assertions, budget admissions, immutable-owner
boundaries, and source-backed behavior are preserved. No compatibility alias,
lint allowance, new target, or fixture was added. Validation was source-only:
rustfmt and `git diff --check`; Cargo was not run because the coordinator is
rerunning the combined native check after this leaf is merged.
