# 0362ul89ard — PPTX comment writer contract correction

Status: source-ready follow-up to `0362ul89arc`, based on source commit
`68576a5283`. The original candidate basis remains
`ef9470f7da08a131f31600041f9d562203d62bfa`; the coordinator must compose this
source patch with the previously reviewed 89ara, 89arb, and 89arc fixes before
the full validation gate. Cargo, rustdoc, native fixtures, and ABBA were not
run on this leaf.

## Corrections

- The classic-comment `BoundedOutput` now carries the selected aggregate
  string budget. Every emitted XML attribute charges its name and decoded
  value, including the root namespace declaration; text charges its decoded
  bytes; date-time attributes therefore use the same aggregate budget as the
  reader. This is a streaming counter in the existing writer path, not a
  second model traversal or a full preflight rescan.
- `BoundedOutput::attribute` computes the exact escaped physical value length
  (`&amp;`, `<`, `>`, quotes, and numeric whitespace references included) and
  rejects `max_attribute_bytes` before writing or reserving that attribute's
  bytes. The writer and raw reader therefore agree on the physical wire
  ceiling, while `max_string_bytes` remains the selected decoded aggregate.
- Empty and populated author/comment writers have exact and one-byte-under
  aggregate-limit coverage. An escaped author value has exact and one-byte-
  under physical-attribute coverage and round-trips through the reader at the
  exact bound. The existing date-time boundary and inactive-MCE raw-attribute
  regressions remain in the same package tests.

The earlier source-admission corrections remain part of this return: raw
attribute values are checked before MCE can discard an inactive branch, and
collaboration presence/threading reuse one admitted `MceSource` for semantic
validation and offset translation. The selected parser still checks aggregate
decoded names, values, text, references, and date-time data before ownership.

## Source-only checks

```text
git diff --check                                      PASS
rustfmt --edition 2024 on changed Rust files       PASS before final source-only cleanup
Cargo/native/full gate/rustdoc/ABBA                 pending coordinator slot
```

No generated evidence, fixture, executable, or `3rdparty/` path is part of
this record.
