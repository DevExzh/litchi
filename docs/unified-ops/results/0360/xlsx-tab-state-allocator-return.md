# XLSX tab-state allocator return

This is a returned allocator result for the W4 source-backed tab-state
history. It does not accept the b1a5 candidate and it does not close the W4
performance gate. The measured pair was built from the reproducible pre-W4
integration state `e63a1fe3534bcaaae2ce367339f7f8275877baef` and the W4
candidate `b1a5b7b13f7355ffd6f374a11ad75c819c45aa60`. Both builds used
`cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml`
with the allocator driver and the matched 204-package registry projection
`b8cd800aee2eecbaa3003601f68586c8233236741fd201fb309f9c5d2502559`.

The protected allocator binaries are outside the repository:

| revision | binary SHA-256 | target | receipt SHA-256 |
| --- | --- | --- | --- |
| e63 | `c9efeaf4c70f9e91bc0052abdda2f31cc434cfc1e25b4252b2cc5ef3f46a7fd0` | `final-e63-alloc` | `509453270b20147fbc034a02e0d1842abf5f60b1bc9d61e00c9c050f37cdc098` |
| b1a5 | `10b195ffc06527eb258c42671401e5c0d506b603cd211e15783cdb1c6464e110` | `final-b1a5-alloc` | `6ca73a16dfa8b0229d9db30177de11123a2d9883cdd9cc4ebdc7d497b154f18f` |

The receipt directory is
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/allocator-final-b1a5-v2`.
The fixture has two defined names, 64 retained history steps, two warmups,
and exact forward and inverse checks of 64/64 for each revision.

| interval | e63 | b1a5 | change |
| --- | ---: | ---: | ---: |
| retained history allocated bytes | 4,834,958 | 5,033,235 | +198,277 (+4.10%) |
| retained history allocation calls | 46,730 | 47,847 | +1,117 (+2.39%) |
| retained history live bytes after | 167,697 | 371,809 | +204,112 |
| retained history peak live after | 427,929 | 427,929 | unchanged |
| independent replay allocated bytes | 103,381,056 | 103,028,268 | -352,788 (-0.34%) |
| independent replay live bytes after | 167,953 | 375,977 | +208,024 |
| independent replay peak live after | 606,577 | 831,267 | +224,690 (+37.05%) |
| source-owner admission allocated bytes | 23,346,752 | 24,363,616 | +1,016,864 (+4.36%) |

The source audit identifies the returned candidate's retention hotspot in the
catalog reuse change (`9c91b18f14`). `Snapshot` now owns a deep
`Box<[raw::DefinedName]>` (`crates/litchi-xlsx/src/tab_state/snapshot.rs:56-63`).
Both source capture and rewrite allocate that box (`:80-88` and `:174-201`).
`Snapshot::deferred_catalog` then clones the box (`:234-238`). The deferred
owner stores a catalog, and `RestoreEndpoint::from_workbook` clones that
catalog into the exact after endpoint (`crates/litchi-xlsx/src/workbook/source/deferred.rs:50-67,
:142-158, :246-251`). A retained changed patch therefore keeps independent
before/after name trees and an additional deferred/endpoint catalog tree.
`DefinedName` contains multiple owned strings and optional owned strings, so
these are deep copies rather than slice-header copies.

The admission interval has a second, distinct source. `SourceBackedWorkbook`
already retains parsed names in `SourceInner` during open
(`crates/litchi-xlsx/src/workbook/source.rs:359-459`), while
`edit_tab_state_with_options` reparses the workbook and retains another name
box through `Snapshot::load_source_backed`. The pre-W4 adapter opens the
smaller `SourceBackedEditor`, whose snapshot dropped the catalog names. This
receipt therefore attributes the +1,016,864 bytes to both the new read-only
owner admission and the duplicate snapshot catalog; it must not be presented
as an isolated cost of the deferred publisher.

The repair is now implemented in the W4 candidate, but this returned receipt
remains unaccepted until the candidate is rebuilt and the paired allocator
route is rerun. `SourceInner`, the defined-name and tab-state snapshots, the
deferred `Catalog`, and both restore endpoints now share an immutable
`Arc<[DefinedName]>`; a changed defined-name edit creates a new Arc. Deferred
materialization passes the known Arc into the normal source-backed open path.
That path still parses and validates the reopened workbook under the owner
options and compares its parsed names with the shared catalog before retaining
the Arc. A non-workbook family reuses the base Arc because its change contract
does not alter workbook XML. Catalog replay with different `SourceLimits`
continues through the publisher's recipient-policy parse, so this sharing does
not bypass stricter admission.

The source validation after this repair passed the focused defined-name suite
(7 tests), focused tab-state suite (14 tests), all-feature checks for the W4
consumers, and warning-denied Clippy for the same consumer set. These checks
establish correctness and policy preservation only; they do not replace the
required paired allocator rerun. The original admission and retained-history
numbers above remain the baseline evidence, and the b1a5 result remains a
returned performance failure until the new binary is measured.

The native source-family attempt was not a measurement: a wrapper path typo
failed before any child launch. Its host/error files are retained under
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/final-source-family-20261021-v1/page_margin`.
No native timing was run after the allocator return, and no performance result
from this candidate is accepted.
