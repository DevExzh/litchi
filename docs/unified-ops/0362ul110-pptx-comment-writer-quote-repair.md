# 0362ul110 — PPTX comment writer quote repair

Status: source-ready runtime follow-up based on `e130751fc5`. The supplied
PPTX full-feature runtime packet had 1,386 passing tests, 10 failures, and one
ignored test. Five comments-package writer failures and two opened
serialization failures (seven writer-related failures total) surfaced the same
malformed XML. Two collaboration MCE failures and one projection error-order
failure were separate. Cargo, rustdoc, native fixtures, and the full gate were
not run on this leaf.

## Diagnosis and correction

The writer refactor routes the root namespace declaration through
`BoundedOutput::attribute`, which already emits the closing quote. Both root
writer paths still appended the old `b"\">"` suffix, producing an extra quote
(`...xmlns:p="..."">`) and the parser's “attribute value not closed” error.
Both suffixes now emit only `b">"`.

This is a serialization-only correction: the selected aggregate string
counter, exact escaped physical-attribute preflight, output bound, namespace
profiles, and source-admission behavior are unchanged. Existing strict writer,
raw exact/one-under, numeric-whitespace, package round-trip, and opened
transaction tests cover the repaired paths. The two collaboration MCE failures
and the projection error-order failure in the packet remain assigned to their
separate owners.

## Source-only checks

```text
rustfmt --edition 2024 --check on changed Rust files  PASS
git diff --check                                      PASS
Cargo/native/full gate/rustdoc/ABBA                 pending coordinator slot
```

No generated evidence, fixture, executable, or `3rdparty/` path is part of
this record.
