# 0362ul82a — PPTX metadata history bounded admission

- Date: 2026-10-25.
- Status: source-ready for independent review; native Cargo validation remains
  with the coordinator's isolated build slot.
- Base: `57fecc892ec8d48cc7db30964af8a90392ad4cd9`
  (`0362ul82` metadata-history policy slice).

This correction closes the admission and consumer-propagation findings from
the 82 review. Changes Information, Revision Information and slide
synchronization now map each selected family profile into a bounded MCE
processor before AlternateContent expansion. The selected semantic policy is
intersected with the package's physical source policy in the modern-comment
raw command reader and writer; the modern-comment edit passes its retained
Changes profile through both discovery and update, and the final XML write
uses that same profile.

Slide synchronization counts escaped `serverSldId` bytes before constructing
the escaped value or output buffer. Slide-show event storage computes its
fragment length with checked arithmetic, reserves the fragment fallibly, and
passes the selected output ceiling through the slide splice. Every splice
branch preflights the exact replacement size, uses `try_reserve_exact`, and
asserts the emitted length before returning. Boundary tests cover selected
MCE dimensions, escaped identifiers, retained raw-helper profiles and output
budgets.

Source checks performed here:

- `rustfmt --edition 2024 --check` over all changed Rust files;
- `git diff --check`;
- focused searches confirming no default MCE call remains in these three
  metadata parsers and no fixed event fragment capacity remains;
- glossary and record text reviewed for the selected-profile behavior.

Cargo tests, rustdoc, clippy, native fixture processing and ABBA measurements
remain intentionally outstanding for the coordinator's independent build
slot. No binary, fixture, or large performance evidence is committed.
