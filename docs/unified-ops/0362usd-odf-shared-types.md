# 0362usd — ODF changed-only families use the shared commit triad

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

The source validated here is `15ab9324ef4f8c927acd6ecfc19d9707c3f886ee`,
the amended successor of this slice's original base
`173dd5330476713714871007316794b9a3a1bdaa`. With the memory-capped cargo
shim, one build job, `CARGO_INCREMENTAL=0`, debug information disabled and
`TMPDIR=/tmp/claude-1001/uo-t/odf-shared-types`, these commands passed:

```text
cargo check --offline --locked --all-targets --all-features \
  -p litchi-odc -p litchi-odi -p litchi-odm -p litchi-oth
cargo clippy --offline --locked --all-targets --all-features \
  -p litchi-odc -p litchi-odi -p litchi-odm -p litchi-oth -- -D warnings
```

Both used `CARGO_TARGET_DIR=(deleted-coord-artifact)`,
which measured 4,330,035,351 bytes (4.2G) after validation. The concise
external receipts are retained outside the checkout:

| Command | Receipt | Size | SHA-256 |
|---|---|---:|---|
| `cargo check` | `(deleted-coord-artifact)/check.log` | 820 bytes | `250d42c6c9859f1ab84646ba2bf0473ad4dac84b900585b6a94a3bc043abdfda` |
| strict `cargo clippy` | `(deleted-coord-artifact)/clippy.log` | 836 bytes | `6ccfea6a4d032e972f40ee61879faf0cde9ebd0608de371c873f76b0d2fd1759` |

Rustfmt (`rustfmt --edition 2024 --check`), `git diff --check`, and the
repository artifact guards also passed. The ODM transaction now evaluates
the exact-byte changed flag once while constructing the commit, rather than
computing it lazily in `diagnostics()`. Final native qualification must cover
that moved comparison; this record makes no global scan-free performance
claim. No benchmark, fixture, binary, or large repository evidence artifact
was added.
