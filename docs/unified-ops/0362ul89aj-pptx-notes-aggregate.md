# 0362ul89aj — PPTX selected notes aggregate limits

Status: source-ready repair based on `93070d72e3f261496b9d57d0d41b75bfc47b4f7d`.
The leaf was kept source-only; the coordinator owns the combined build and
runtime validation.

## Repair

The targeted source-backed and eager per-slide notes readers now enforce the
selected notes-family aggregate over exactly the visited notes slide, notes
master, and theme payloads.  The owning slide and presentation root retain
their independent related-root ceilings and are not charged to this notes
aggregate, and unrelated notes parts remain unvisited.

The charge uses the original raw part payload length.  The source loader reads
that length from `SourcePart::source_data()` before exposing any
markup-compatibility replacement through `Part::blob()`; the eager loader
captures the package part length before processing.  This keeps the selected
aggregate aligned with `SourceLimits::max_total_part_bytes` and with the
existing eager notes graph's raw-byte accounting, even when MCE changes the
processed view.  A limit refusal is typed as `CodecLimitKind::XmlBytes` with
the exact observed and configured values, and each visited part is charged
before the next closure part is loaded.

## Focused regressions

`tests/source_related.rs` checks exact and one-under aggregate limits for both
source-backed and eager notes reads, including the observed/maximum pair.  A
separate case adds a malformed unrelated notes master and verifies that both
readers still succeed at the selected aggregate boundary.

## Source-only checks

`rustfmt --edition 2024 --check` and `git diff --check` pass for the changed
Rust files.  No Cargo, native, rustdoc, full gate, generated evidence,
fixture, executable, or `3rdparty/` path was used or added; combined validation
and performance review remain coordinator work.
