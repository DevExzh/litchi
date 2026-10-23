# 0362uqc — ODF changed-only families use the shared commit triad

This slice starts from `173dd5330476713714871007316794b9a3a1bdaa` on
`uo2/docx-runtime`. It completes the changed-only ODF family migration required
by record 0351 §2.2; richer ODF metadata and publication diagnostics remain
their existing crate-owned types.

## API boundary

The primary changed-only families in ODC, packaged ODI, OTH and ODM now alias
`litchi_core::edit::Commit<Package, Patch>` (or the corresponding flat ODI
`Image`/`Patch` pair) and re-export the core `Diagnostics`. Their construction
records diagnostics in the triad, including replay and exact no-op commits,
without a second wrapper or a deferred recomputation. Existing package, patch,
diagnostics, `into_package`, and `into_parts` consumers keep the same contract.

ODI flat semantic planning is now a capability of `flat::Patch` rather than a
derived helper on `flat::Commit`. The source and target already retained by the
patch are used directly, and package semantic/history consumers read the
source through `Commit::patch().source()`. This preserves exact source bytes,
security-policy limits, transfer planning, and the existing semantic tests
without adding stored state or scans.

ODM transaction history and patch merging were migrated from private commit
fields to the shared accessors. The transaction computes and stores the core
diagnostic at construction, so its previous changed/no-op behavior is retained.

All affected source consumers and ODI semantic-planning tests were migrated in
this commit. No deprecated spelling or compatibility alias remains.

## Validation

This was a static-only leaf change while the coordinator's build slot was
reserved. The changed Rust files pass Rustfmt (`rustfmt --edition 2024 --check`)
and `git diff --check`; no Cargo command, benchmark, fixture, or evidence
artifact was added. The source base and resulting immutable commit SHA are
reported by the worker handoff.
