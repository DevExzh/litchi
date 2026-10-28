# 0375b — W1-E slice: shared chart title and axis-title payload engine

- Status: implementation slice on uo2/iwa-w1e; scoped branch gates pass,
  while coordinator review, native evidence and final integrated gates are
  pending.
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
reopen, and semantic host-vs-leaf reads. Numbers uses
test-data/iwork/numbers/chart-data-native.numbers and adds the Column2d chart
through the host before serializing the differential input. Pages uses a
source-built Draft with a host-added body chart: the tracked native Pages
chart-data fixture currently contains a malformed cross-component chart
reference for host discovery, so it is retained for the native queue but not
used as a false differential vector. The host remains authoritative; the
cases are declared Evidence::Queued until native application evidence is
recorded.

The intentional boundary is caption authoring: the host rows remain
host-only until a later shared caption graph slice supplies the actual
bounded graph algorithm and native evidence.

## 4. Performance qualification

crates/litchi-iwa/tests/parity/chart_text_cost.rs contains the source-bound
host/leaf ABBA probe for Numbers and Pages title/category/value routes. It
includes open, selector resolution, mutation, commit, candidate reopen and
serialization on both routes. The ignored probe passed as a functional smoke
check, and a pre-integration release callgrind run recorded these instruction
counts (six calls per route; route totals, not an engine-only extrapolation):

| Route | Host Ir | Leaf Ir | Difference |
|---|---:|---:|---:|
| Numbers title/category/value | 453,131,860 | 217,590,505 | -52.0% |
| Pages title/category/value | 326,433,860 | 14,267,859 | -95.6% |

The raw callgrind files remain outside the checkout under `/tmp`:
`litchi-w1e-chart-text-{host-numbers,host-pages,leaf-numbers,leaf-pages}.callgrind`
(117,751 / 121,749 / 106,267 / 92,341 bytes; SHA-256 respectively
`22f03fbd3c95ac8da61cf7544f9daf12db532afde3b64865e93efef8ac2c5448`,
`cf40c745b4a3b00dc015e25a8028c6fc1a5e743da3beeb545e1626b2ac185fe0`,
`93b8c74e53dc84400eff742be10a762318744ff0fb99cfde01e185e7a2135973`,
`599ea051b8f04c745e8db05fd8fba1e173d9630f89901d35656125a10c4f50ce`).
The release probe binary was built from ae062046d5 with
`CARGO_BUILD_JOBS=4 CARGO_INCREMENTAL=0 CARGO_PROFILE_RELEASE_DEBUG=0 cargo
test --release --offline -p litchi-iwa --test parity chart_text_cost_abba
--no-run`; its 39,643,816-byte SHA-256 is
`8adc4f35bf5a254663efe3108ee99b09fd88db101e0c8a7d09fb81cf74117cbe`.
These counts are branch diagnostics only: later merges invalidate them, no
wall-clock result is claimed, and final integrated release-mode, pinned-CPU
qualification remains required.

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

Scoped verification on this worktree includes rustfmt, targeted Clippy,
targeted rustdoc, the shared/Numbers/Pages unit tests, six host differential
tests, the ignored full-facade cost-probe smoke test, JSON validation, API
glossary registration, ledger and boundary checks, no-3rdparty checks,
affected fuzz-lock cargo checks and git diff --check. These checks are
source/branch evidence only: native queue rows remain pending, and the
complete 53-row gate plus final performance qualification must run on the
integrated trunk commit. No performance qualification is claimed from this
slice.
