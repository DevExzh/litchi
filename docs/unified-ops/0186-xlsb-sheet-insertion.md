# 0186 — XLSB worksheet insertion

## Contract

The native primary `Edit::insert_sheet(Position, String, sheet::State)` authors
one blank worksheet at a zero-based slot in the complete tab catalog. Chart
tabs participate in positions. Repeated insertions interpret positions against
the previously staged insertions. Name collisions, invalid positions and unsafe
dependency rewrites fail before staging the new operation. Other native edit
families remain structural barriers. Durable operation tag 9 stores the checked
position, visibility state and name; existing tags keep their semantics.

Insertion retains existing tab identifiers, relationship IDs and physical
payloads. It authors a blank worksheet and its binary index through native
writers, allocates collision-free parts and identifiers, and uses workbook-
relative relationship targets. Worksheet parsing now resolves the main workbook
part through its relationship as well, allowing the existing reader to validate
custom workbook locations. The new tab is not implicitly selected. Existing
workbook-view positions, scoped names, internal XTI endpoints and modeled
volatile references shift to preserve their original sheet identities.

A blank worksheet does not need a singleton XTI. Absent Externals stays absent;
existing XTI entry order and cardinality stay unchanged. Insertion through an
internal 3-D range receives typed refusal because its old referenced set would
otherwise include the new sheet. The existing opaque dependency guards apply,
including VBA, calculation chains, shared-workbook revisions, models,
connections, slicers, timelines and unmodeled workbook records.

The shared lexical title helper inserts a catalog title and updates its vector
size and matching heading count. It validates ordered Sheet/Worksheets and
optional trailing named-range categories, preserves unrelated XML bytes, and
handles strict/transitional namespace bindings. New title text preserves carriage
returns through character references and rejects XML-forbidden characters. Missing optional title metadata
needs no rewrite. Ambiguous, inconsistent or unsupported metadata receives a
typed error; valid but unmodeled categories/layouts map to unsupported capability.
No save-time repair is performed.

The facade exposes `File::add` and transactional addition under the workbook,
with start/end, positional and relative placements. It admits name, hidden and
visibility through the property registry, chooses a collision-free default name,
and preserves source fences, policy, quotas and atomic mixed-batch rollback.
Capabilities enumerate the remaining structural and dependency gaps.

## Scope and remaining work

Worksheet removal and broader XLSB structural CRUD remain open. Repeated native
insertions currently revalidate their staged prefix before admission, and the
facade commits each insertion across its structural barrier. Homogeneous
insertion coalescing and linear batch scaling are not claimed by this record.
Tier 2/3 writes and Python/CLI/full-goal acceptance remain open.

## Verification

The shared OOXML suite passes 477 tests/doctests and the native XLSB suite
passes 1,196, including the 17 new focused insertion/package checks. Fixtures
cover mixed chart catalogs, start/middle/end slots, staged duplicate admission,
exact inverse replay, malformed tag 9 fields, internal/external references,
scoped names, volatile owners, absent Externals, opaque dependency refusal,
custom workbook locations, strict relationship flavor and selected flags.

Review corrected native package assembly and preserved typed binary-index
errors, raw BIFF12 headers and workbook-relative targets. It tightened shared
title category ordering/count validation and preflight operation limits. Facade
capability review separated addition from removal so unsupported worksheet
removal is not accidentally advertised. The
new native allocator avoids repeated full namespace scans for ordinary occupied
worksheet names. Repository lockfiles remain unchanged; workspace Cargo checks
use an external copied lockfile with cached cc 1.6.0.

The combined DOCX/XLSX/PPTX/XLSB/encryption facade suite passes 1,250 tests
and doctests. Its seven new insertion tests exercise eager/source ownership,
all placements, collision-free defaults, visibility, quota refusal, sequential
mixed set/add/move transactions and rollback, plus create/add/cell-edit/reopen.
Full/minimal all-target Clippy and affected-crate rustdoc pass with warnings
denied. Both XLSB glossary profiles and the generated OfficeCLI parity check
pass. Final focused checks pass 27 shared title tests, 17 native insertion tests,
and 16 facade capability/add tests. The boundary audit passes for 66 packages
and 268 internal declarations, with 11 existing debt items. All five delegated
worktrees were removed after integration. No base-only commits were pending.
This record does not claim Python, CLI or parser-fuzz acceptance.

## Paired measurements

The new standalone harness compares one native insertion with one unified
`File::add` on workbooks with 1, 10, 100 and 1,000 existing worksheets. Each
existing worksheet has a distinct numeric marker; the new blank visible sheet
is inserted at `floor(N/2)`. Caller inputs are prepared outside timing, including
the native owned name and the facade parent/node. Source promotion and facade
property coercion remain inside their respective timed calls.

Both eager/source matrices use nine ABBA rounds on CPU 0 with no concurrent
builds, tests or heavy audits. The [after](results/0186-xlsb-sheet-insert-after.json)
and [repeat](results/0186-xlsb-sheet-insert-repeat.json) receipts have empty
failure matrices and byte-exact native/facade ZIPs in all eight cases. Checks
also verify old cell markers, the inserted blank tab and its binary index,
visibility, the returned canonical path and preserved active-tab identity.

| Owner / existing tabs | After native / facade µs | Repeat native / facade µs |
|---|---:|---:|
| eager / 1 | 783.704 / 803.333 | 788.613 / 801.434 |
| source / 1 | 931.095 / 947.885 | 931.634 / 951.954 |
| eager / 10 | 2074.720 / 2118.320 | 2083.419 / 2128.289 |
| source / 10 | 2510.803 / 2539.672 | 2513.362 / 2542.781 |
| eager / 100 | 13187.054 / 13383.645 | 13436.630 / 13672.341 |
| source / 100 | 15789.596 / 15952.217 | 15832.811 / 15999.371 |
| eager / 1000 | 126514.629 / 128155.056 | 126278.861 / 127806.588 |
| source / 1000 | 150397.789 / 152126.616 | 151744.284 / 153767.453 |

Measured facade overhead is 1.0–2.5%, within the scoped 10% single-operation
bar. Both receipts identify the same binary and source-file digest, with raw
samples, corpus hashes and build provenance. No equivalent primary insertion
API existed before this change, so these are paired after/repeat measurements,
not a before/after speedup claim. They do not establish insertion batch scaling
or full acceptance. Temporary build and measurement artifacts are removed after
retaining these bounded receipts.
