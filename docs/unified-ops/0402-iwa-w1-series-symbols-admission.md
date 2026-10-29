# 0402 — W1: bounded shared chart-series symbol admission

- Status: implementation slice on `uo2/iwa-series-symbols-admission`, cut from
  coordinator source `bc9a37fd75977b01ab3f24cabfdc4b2fe977c4bc`.
- Commits: `e57d0df0c7677ba6986a49dcf165eb6d4fd3f722` initial implementation;
  `212ba356d7b5904e988c4f636f2cc5ceedc50c10` closes the independent
  admission-review findings; `c0dfabd11678fb90bb2aea8c333e8ee4df6915f4` adds the exact codec/registry
  requirements repair after coordinator trunk
  `59a501492cb58a538085280e383007ca4d39dad7`.
- Final clean WP tip after formatting: `5385fd632c8a28d4217422dcedae1c3b74cc81ac`.
- Adoption repair: `ed50c20a25528f8a0a9aecb2a1ca32dfb778b02b` fixes the styled-parent reference framing and
  residual wire-field/work hard-ceiling clamp; the final tip includes this
  repair.
- Final hard-axis repair: `1e1e56651c4fa3e5795e31267c7e6c58980afae9` clamps residual input, nesting, codec
  bytes, fields, and wire work to hard ceilings after zero-exhaustion checks;
  the large-profile regression proves no hidden quota headroom.
- Adoption follow-up: `5888a2b4675e52aebe8e1182d49d20ea052b9de1` caches exact prepared registry inspection
  requirements and the inspected registry root, invalidating both caches on
  metadata changes. The final direct-engine fix does not alter facade budgets
  or add default headroom. Commit `085eb46d55` adds the cache-accounting
  repair: it charges the
  compatibility work getter only for the next uncached inspection pass, while
  metadata rewrites invalidate both caches.
- Scope: shared chart-series symbol read/write admission, Numbers mechanical
  caller adoption, and the direct shared-engine parity caller. Pages and
  Keynote facade adoption remain owned by their work packages.

## Public shared seam

`litchi-iwa-style::series` now exposes explicit `Limits`, typed
`SymbolError`, actual `Report`, and source-derived `MutationScope` values.
The symbol entry points take one caller-supplied profile and return the
semantic result together with the report. Mark, size, parent traversal,
codec, graph, registry, allocation, and readback work share one residual
budget; no default-budget compatibility route is used.

The report scope is source-relative and names selected objects, members,
message types, known fields, and planned object/member additions/removals.
It is intended for facade locality verification rather than deriving an
allowlist from observed post-edit differences.

The repair pass adds fallible reservations before every bounded collection
growth, residual accounting for repeated wire preparse/codec passes, explicit
registry inspection and cross-component-reference debits, and a bounded
container writer instead of delegating to the legacy infallible helper.

The second review replaces the remaining container-size estimate with exact
varint sizing, exposes exact codec decode/reference requirements and reports,
and uses package-metadata inspection requirements rather than registry-byte
multipliers. Existing chart-font consumers retain a compatibility work getter
backed by the same exact inspection requirements.

## Bounded workspace finalization seam

`Workspace::finish_requirements` now exposes source-relative counts for the
changed/deleted entries, owned member-name copy bytes, final vectors,
retained bytes, and traversal work without allocating or rescanning member
names. `Workspace::finish_bounded` checks a caller residual
`FinishLimits`, uses fallible exact reservations for both `Changes` vectors
and owned names, moves existing archives/insertions without copying, and
returns a `FinishReport`. The historical `finish` method delegates through
an unlimited profile for unchanged consumers.

The seam deliberately does not absorb facade-owned `LeafComponents` or
workspace index construction: those source/index costs remain an explicit
pre-admission responsibility of each format caller and must not be charged
again as finalization work.

Implementation commits: `bd2d138a07` plus the overflow-checked follow-up
`c50fcaf8db` (on top of coordinator freeze merge `2f1f7a8a78`).

## Correctness and evidence boundary

The bounded path preserves unknown fields, exact no-op behavior, shared-owner
copying, selected-chart locality, and immutable source snapshots. Errors after
staging leave only the private copy-on-write workspace changed; callers must
discard that failed workspace, and no source bytes are published.

This slice makes no native-acceptance or performance claim. Numbers currently
derives a finite physical profile and validates the shared report, but its
legacy series transaction still lacks a reusable residual caller ledger. That
is an explicit follow-up gap, not a full Numbers facade-ledger qualification.

## Verification

- `cargo test --offline -p litchi-iwa-style --lib`: 32 passed.
- `cargo test --offline -p litchi-iwa --test parity chart_font_cases -- --nocapture`:
  29 passed, 13 ignored; Numbers, Pages, and Keynote chart-font parity and
  low-budget controls pass.
- `cargo test --offline -p litchi-numbers --test sheet_chart_font -- --nocapture`:
  4 passed.
- `cargo clippy --offline -p litchi-iwa-style --lib --all-features -- -D warnings`: pass.
- `cargo check --offline -p litchi-numbers --lib`: pass.
- `cargo clippy --offline -p litchi-numbers --lib --all-features -- -D warnings`: pass.
- `cargo check --offline -p litchi-iwa --tests`: pass.
- `cargo fmt --all -- --check`: pass.
- `cargo doc --offline -p litchi-iwa-style --no-deps`: pass.
- `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py --target-dir target`: pass.
- `python3 tools/iwa_exit_ledger.py --check`: pass.
- `python3 tools/check_crate_boundaries.py`: pass.
- `python3 -m unittest tools.test_no_3rdparty_references`: 4 passed.
- The post-review focused symbol suite includes typed zero-field,
  zero-reference, and zero-allocation refusals plus shared-owner copying and
  planned-scope assertions.
- Codec requirement/report regression: reference and decode preflight reports
  match their follow-on codec passes for fields and work.
- Direct engine ABBA parity probe across Keynote, Numbers, and Pages:
  `cargo test --offline -p litchi-iwa --test parity
  series_graph_write_cost_abba -- --ignored --nocapture`: pass.

No ledger, queue, progress, fixture, performance, native evidence, or
generated binary changes are included.
