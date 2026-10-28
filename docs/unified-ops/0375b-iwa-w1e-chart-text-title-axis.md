# 0375b — W1-E slice: shared chart title and axis-title payload engine

- Status: implementation slice on uo2/iwa-w1e; coordinator review and
  runtime gates are pending.
- Date: 2026-10-28.
- Parent: 0375a (0375a-iwa-w1e-style-variation-write-engine.md), plan
  0363 (0363-iwa-monolith-exit-plan.md), parity harness and ledger
  0369 (0369-iwa-exit-parity-harness-and-ledger.md), and the native queue
  0363-mac-queue.md.
- Scope: title and category/value-axis title only. Caption authoring remains
  host-only and is explicitly deferred; no Keynote ledger row is flipped.

This is a thin vertical slice. litchi-iwa-style::chart_text owns the bounded
generated-extension algorithm. Numbers and Pages retain selector resolution,
reference census, graph ownership, archive publication and candidate reopen.
No native object identifier, UUID, archive name, ZIP path or generated
protobuf value crosses a supported leaf signature.

## 1. Shared engine

ChartTextKind::{Title, Axis(Axis)} selects the generated scalar family.
decode borrows the visible UTF-8 value from a complete chart payload; rewrite
applies a presence-preserving set/clear to field 10,000 while retaining
unknown outer and generated-extension spans. The engine validates a unique
canonical length-delimited extension before dispatching to the strict
title/axis codecs.

Limits is finite for input bytes, fields, output bytes, nesting, work,
selected text bytes, allocations, retained bytes and scratch bytes.
Report records those costs and whether the selected payload changed.
Error::{InvalidSource, Limit, Allocation} is terminal and is mapped by each
leaf into its own Error/LimitKind; codec/archive failures are not flattened
into a successful host fallback. Axis rewrites use the generated prepared
rewrite path, including candidate readback before publication.

The engine tests cover:

- title set, exact no-op, unknown outer/extension preservation and readback;
- category-axis locality with value-axis preservation;
- duplicate-extension refusal and a typed low text-byte refusal.

## 2. Leaf facades and contract

The Numbers destinations are:

- Package::{sheet_chart_title, edit_sheet_chart_title,
  apply_sheet_chart_title};
- Package::{sheet_chart_axis_title, edit_sheet_chart_axis_title,
  apply_sheet_chart_axis_title}.

The Pages destinations are:

- Package::{body_chart_title, edit_body_chart_title,
  apply_body_chart_title};
- Package::{body_chart_axis_title, edit_body_chart_axis_title,
  apply_body_chart_axis_title}.

Each route has immutable source snapshots, exact-source patches,
Patch::inverse, exact no-op behavior, bounded rewrite limits, selector-first
refusals, candidate reopen/readback and untouched-object locality. Caption
facades from the earlier donor are removed from this slice rather than
published as a second orchestration path. Existing Keynote title, axis-title
and caption implementations remain unchanged.

## 3. Host differential vectors

crates/litchi-iwa/tests/parity/chart_text_cases.rs is the host-backed
differential oracle. It covers title and both axis selections for Numbers and
Pages, set/read/remove sequences, exact inverse restoration, strict leaf
reopen, and semantic host-vs-leaf reads. The host remains authoritative; the
cases are declared Evidence::Queued until native application evidence is
recorded.

The intentional boundary is caption authoring: the host rows remain
host-only until a later shared caption graph slice supplies the actual
bounded graph algorithm and native evidence.

## 4. Performance qualification

crates/litchi-iwa/tests/parity/chart_text_cost.rs contains the source-bound
host/leaf ABBA probe for Numbers and Pages title/category/value routes. It
includes open, selector resolution, mutation, commit, candidate reopen and
serialization on both routes. It has not been run on this branch while the
coordinator quiet hold is active; no performance claim is made here.

The coordinator should run the final integrated candidate with release
debug-info, CPU pinning and callgrind, collecting
chart_text_host_{numbers,pages} and chart_text_leaf_{numbers,pages}. The
qualification must report each full facade route separately from this
engine-only source file.

## 5. Native evidence and verification

No native-authored read receipt proves a Litchi-mutated title or axis title
survives native open/save/close/reopen. The six Numbers rows and six Pages
rows are therefore queued in 0363, with pending results. The ledger and
queue must not be treated as W3-ready until the coordinator records native
evidence, final integrated parity and final performance.

Scoped rustfmt, JSON validation, the API glossary registration, ledger rows,
queue rows and git diff --check are source-level checks only. Cargo, Clippy,
tests, rustdoc, boundary, fuzz locks and the full 53-row gate remain pending
on the integrated trunk commit.
