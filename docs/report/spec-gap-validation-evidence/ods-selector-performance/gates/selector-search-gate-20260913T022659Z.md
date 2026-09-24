# Selector-search integration gate

Run date: 2026-09-13 UTC. The candidate baseline commit was
`d5b5afdb4032b718c0f94c113a3b1b09f17ae6ba`. The candidate was tested with an
isolated Cargo target at
`/var/tmp/litchi-ods-selector-search-20260913T022659Z/target`.

The focused target
`crates/litchi-ods/tests/sheet_metadata_selector_search.rs` passed 7/7. It
covers repeated row and cell interval boundaries, supported row containers,
implicit merge coverage, explicit covered runs, duplicate worksheet names,
missing coordinates, cancellation and failed-edit atomicity, and a dense
last-cell lookup under a bounded work budget.

The full gate passed **711 tests across 43 targets**, with 0 failures and 0
ignored:

* `full-tests-selector-search-20260913T022659Z.log` records
  `cargo test --locked --offline -p litchi-ods --all-features --all-targets
  --no-fail-fast` (log SHA-256
  `aef6032ce4e45e26afb2bcf11c03b0d1d8deca6c7a34f6e02e69458fd82aa6ee`).
* `clippy-selector-search-20260913T022659Z.log` records strict all-features,
  all-targets Clippy with `-D warnings` and exit 0 (log SHA-256
  `a2e30e5376daffca70d2f7e9361aa34401ceb35bb827feff776071a43a5f1363`).

The source identity used by both gates is in
`selector-search-source-hashes-20260913T022659Z.sha256`: `index.rs` SHA-256
`6f08a6488414c1eb357fa6c877ffe41d3d7c27eed9682ee4f62b3e815b10aa2b` and this
test target SHA-256
`691045cfea0894ccf20073e07053bb07b82b77c521be6465c6474893b88e498d`. A
post-Clippy check matched both hashes.

No source implementation file was changed by this tester and no commit was
made.
