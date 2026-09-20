# Shared calendar preparation

This preparatory change extracts the existing TEXT Gregorian day conversion,
serial epoch, seconds-per-day constant and English month names into the private
`evaluation/calendar.rs` module. TEXT imports the same code and data; its accepted
date domain and output behavior remain unchanged. The upcoming VALUE parser can
share the deterministic 1899-12-30 epoch without adding ambient locale or clock
dependencies. This change does not implement any new formula function.

Validation on the root checkout on 2026-09-20:

- `cargo test --locked --offline -p litchi-ods --test ods_formula_text_evaluation --test ods_formula_text_oracle -- --quiet`: 39 passed, zero failed or ignored.
- `cargo check --locked --offline -p litchi-ods`: passed.
- `rustfmt --edition 2024 --check` for evaluation.rs, calendar.rs and text/format.rs: passed.
- `git diff --check`: passed.

These checks use the existing root Cargo.lock, not a frozen isolated gate copy.
No performance improvement is claimed: this is a private code extraction with
unchanged arithmetic, allocation and traversal behavior. Full inspection-batch
semantic, resource, native and performance evidence remains pending.
