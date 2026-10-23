# 0362ui — Repair stable PPTX refusal consumers

This source-only repair starts at `a814baa1de47c0844eb1ea092805b73e7d7bb4ae`,
the v6 API assembly head. The compiler-error capture contained 58 nested
experimental guard patterns in tracked `litchi-pptx` consumers, followed by
cascaded unused `refusal` bindings. The media-transfer test also retained an
unused `DurablePatchFormat` import.

Every affected consumer now matches the outer `Err(Error::Refused(...))`
variant on stable Rust and asserts the same `refusal.kind()` value in the
match arm. The one source cross-copy match arm keeps its existing reason
assertion. The remaining top-level `Error::Refused(refusal) if ...` patterns
in other tests were left unchanged because those are ordinary stable match
guards, not nested guard patterns.

The repair covers the source-primary order/removal tests, opened copy/removal
tests, media transfer, source order, source edit, source cross-copy and the
source engine cross-copy test. No format behavior, refusal kind, or test
fixture changed.

## Validation

- `rustfmt --edition 2024` passed on all eight changed Rust files.
- `git diff --check` passed.
- A repository-wide PPTX search reports no remaining nested
  `Err(Error::Refused(refusal) if ...)` guard.
- Cargo was not run by design; this lane was limited to source and static
  checks while the coordinator's capped build environment was occupied.
