# 0362ul89are — PPTX comment whitespace wire-length correction

Status: source-only follow-up to `0362ul89ard`, based on source commit
`9badfd8acf`. The previous 89ard writer counters, physical admission ordering,
MCE/source reuse, and exact/one-under coverage remain unchanged. Cargo,
rustdoc, native fixtures, and ABBA were not run on this leaf.

## Correction

The writer's exact escaped-attribute preflight now counts both `&#xA;` and
`&#xD;` as five wire bytes, matching the existing escape sink. The prior six-byte
count could refuse a valid output at an exact `max_attribute_bytes` boundary.
The regression writes an author value containing both newline and carriage
return references, round-trips it through the reader at the exact bound, and
checks the one-byte-under refusal with the typed `XmlAttributeBytes` dimension.

This keeps the physical check before attribute output/reservation and does not
change the selected aggregate string counter. Per-attribute length preflight
is intentional for exact physical admission; its qualification cost remains a
later performance-review item and is not claimed measured here.

## Source-only checks

```text
git diff --check                                      PASS
Cargo/native/full gate/rustdoc/ABBA                 pending coordinator slot
```

No generated evidence, fixture, executable, or `3rdparty/` path is part of
this record.
