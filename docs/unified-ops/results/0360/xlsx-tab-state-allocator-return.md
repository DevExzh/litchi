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
`(deleted-coord-artifact)`.
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

The repair is implemented in W4 candidate `14370c12c9` (the stale deferred
catalog documentation was corrected in the same follow-up). `SourceInner`,
the defined-name and tab-state snapshots, the deferred `Catalog`, and both
restore endpoints now share an immutable `Arc<[DefinedName]>`; a changed
defined-name edit creates a new Arc. Deferred materialization passes the known
Arc into the normal source-backed open path.
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
paired allocator evidence below.

The repaired allocator pair was rerun sequentially in the fresh directory
`(deleted-coord-artifact)`, with
the same 64-step, two-name fixture and two warmups. The e63 JSON was unchanged
from the protected baseline; its SHA-256 is
`509453270b20147fbc034a02e0d1842abf5f60b1bc9d61e00c9c050f37cdc098`.
The repaired binary is outside the repository at
`(deleted-coord-artifact)`,
SHA-256 `87973a079fc1bb5e3b87032c8b58702b320bde817bbfdf5096517ac530935173`;
its JSON receipt SHA-256 is
`539a1c9c6643f15a4c7988a7e3663d5f26065797c060535b9b53fab8e41ccac8`.

| interval | e63 | 14370c12c9 | change vs e63 |
| --- | ---: | ---: | ---: |
| retained history allocated bytes | 4,834,958 | 4,942,259 | +107,301 (+2.22%) |
| retained history allocation calls | 46,730 | 46,859 | +129 (+0.28%) |
| retained history live bytes after | 167,697 | 283,681 | +115,984 |
| retained history peak live after | 427,929 | 427,929 | unchanged |
| independent replay allocated bytes | 103,381,056 | 102,907,436 | -473,620 (-0.46%) |
| independent replay live bytes after | 167,953 | 287,849 | +119,896 |
| independent replay peak live after | 606,577 | 741,779 | +135,202 (+22.29%) |
| source-owner admission allocated bytes | 23,346,752 | 24,392,288 | +1,045,536 (+4.48%) |

Exact fixture and output hashes match between the two revisions, with 64/64
forward and inverse checks. The Arc repair materially reduces the b1a5
candidate's retained live bytes (371,809 to 283,681), replay live bytes
(375,977 to 287,849), and replay peak (831,267 to 741,779), but the repaired
candidate remains above the e63 allocation baseline. The retained names and
read-only `SourceBackedWorkbook` admission are mandatory capability costs, so
the result is retained as a returned performance result pending the required
equivalent-workload decision; no elapsed-time, RSS, or native source-family
acceptance is claimed.

## Deferred-policy storage repair (ac929d4957)

The next returned candidate removes a duplicated `SourceLimits` value from
each `Deferred` owner and each deferred `RestoreEndpoint`. `SourceLimits` has
24 scalar ceilings (eight `u64` and sixteen `usize`, 192 bytes on the
64-bit build). The retained endpoint's `Arc<SourceBackedPackage>` already
owns the exact policy under which its catalog was parsed, so `catalog_for`
compares the recipient package's `read_limits()` with that retained package's
policy. Equal policies reuse the catalog; unequal policies reparse through the
publisher under the recipient policy. Snapshot limits remain by value for the
tab-state XML audit and `EditOptions` path, so the edit and stricter-recipient
policy semantics are unchanged.

The candidate binary is outside the repository at
`(deleted-coord-artifact)`,
SHA-256
`3e21996eedbbe71ee98daee5ed6184b0e67bc55e6f348cd595cfbdf0db7aaec4`.
The sequential count-only receipt is outside the repository at
`(deleted-coord-artifact)`.
The e63 JSON SHA-256 is
`509453270b20147fbc034a02e0d1842abf5f60b1bc9d61e00c9c050f37cdc098`; the
candidate JSON SHA-256 is
`05ecfdfa74142e036b3fe4be09327dd9e7648e9648c87ab709f17f24e7504837`.
Both used the same two-defined-name fixture, 64 retained steps, two warmups,
and exact 64/64 forward and inverse checks. This was an allocator-count run;
it makes no latency or RSS claim.

| interval | e63 | ac929d4957 | change vs e63 |
| --- | ---: | ---: | ---: |
| retained history allocated bytes | 4,834,958 | 4,911,667 | +76,709 (+1.59%) |
| retained history allocation calls | 46,730 | 46,859 | +129 (+0.28%) |
| retained history live bytes after | 167,697 | 253,473 | +85,776 |
| retained history peak live after | 427,929 | 427,929 | unchanged |
| independent replay allocated bytes | 103,381,056 | 102,882,860 | -498,196 (-0.48%) |
| independent replay live bytes after | 167,953 | 257,641 | +89,688 |
| independent replay peak live after | 606,577 | 711,571 | +104,994 (+17.31%) |
| source-owner admission allocated bytes | 23,346,752 | 24,392,288 | +1,045,536 (+4.48%) |

Compared with 14370c12c9, the duplicated policy removal reduces retained live
bytes by 30,208, replay live bytes by 30,208, and replay peak by 30,208. The
candidate still has positive retained/replay live and peak deltas versus the
reproducible pre-W4 baseline, so this remains returned evidence and does not
close the W4 performance gate.

The native source-family attempt was not a measurement: a wrapper path typo
failed before any child launch. Its host/error files are retained under
`(deleted-coord-artifact)`.
No native timing was run after the allocator return, and no performance result
from this candidate is accepted.

## Deferred catalog-policy proof sharing (3dddbaf8cb)

This is a returned allocator and heap-attribution result, not a performance
acceptance. It was measured with the same two-defined-name fixture, 64 retained
steps, two warmups, and exact 64/64 forward and inverse checks. The count-only
run used two protected release binaries sequentially; it makes no latency, RSS,
or native source-family claim.

| item | e63 baseline | 3dddbaf8cb candidate |
| --- | --- | --- |
| source | `e63a1fe3534bcaaae2ce367339f7f8275877baef` | `3dddbaf8cb8ec3b6844109cce0f1b568061fc357` |
| protected binary | `(deleted-coord-artifact)` | `(deleted-coord-artifact)` |
| binary SHA-256 | `c9efeaf4c70f9e91bc0052abdda2f31cc434cfc1e25b4252b2cc5ef3f46a7fd0` | `6d52f150fbd1b6fb5ea91ad84293b34d7091181986b21555a873c6a923297a5c` |
| JSON receipt SHA-256 | `509453270b20147fbc034a02e0d1842abf5f60b1bc9d61e00c9c050f37cdc098` | `08ad27c2c1c3e82ba61d67598318dcbeebb8f197c7c0b774522a8ed3c7a686eb` |
| receipt directory | `(deleted-coord-artifact)` | `(deleted-coord-artifact)` |

Both binaries were built with `cargo build --locked --release
--features allocator-metrics --manifest-path tools/perf-baseline/Cargo.toml
--bin xlsx_tab_state_history_alloc`, `CARGO_BUILD_JOBS=2`, and the capped W4
environment. The shared `Cargo.lock` SHA-256 is
`aa4300aff7f76cb467a2a6bc5896e706b3f5b506feed1d16b45f7422a0a9ba23`; the
benchmark manifest SHA-256 is
`52e349415b12b80765e810f47b4f6ac2e76cc092566eb257efc463e07d25fc2f`.

| interval | e63 | 3dddbaf8cb | change vs e63 |
| --- | ---: | ---: | ---: |
| retained history allocated bytes | 4,834,958 | 4,885,827 | +50,869 (+1.05%) |
| retained history allocation calls | 46,730 | 46,863 | +133 (+0.28%) |
| retained history live bytes after | 167,697 | 228,433 | +60,736 |
| retained history peak live after | 427,929 | 427,929 | unchanged |
| independent replay allocated bytes | 103,381,056 | 102,861,356 | -519,700 (-0.50%) |
| independent replay live bytes after | 167,953 | 232,601 | +64,648 |
| independent replay peak live after | 606,577 | 686,195 | +79,618 (+13.13%) |
| source-owner admission allocated bytes | 23,346,752 | 24,410,208 | +1,063,456 (+4.56%) |

The matched heap profiles are outside the repository at
`(deleted-coord-artifact)` and
`(deleted-coord-artifact)`. They
attribute the retained difference to the deferred family endpoint rather than
to the OPC rewrite itself: the `raw::catalog_edit::package::rewrite` stacks are
identical (17.57 KiB over 32 hide calls and 17.09 KiB over 32 show calls),
while 3dd adds 9.22 KiB over 32 calls for each hide/show route through
`tab_state::source::commit_change -> Edit::commit`. The extra live state is
therefore in publication/deferred-owner bookkeeping; removing that state would
remove exact endpoint authentication or recipient policy unless a replacement
proof is demonstrated.

The replay peak has a separate route attribution. Both revisions allocate the
same 380.03 KiB of zlib workspace over 64 writes. e63 reaches it through
`PackageWriter::to_bytes` and its direct writer; 3dd reaches it through
`RestoreEndpoint::output_artifact`, `SourceBackedPackage`'s preserve overlay
writer, and `soapberry_zip::preserve::generated_entry`. The 32.77 KiB zeroed
allocation group is also present in both routes. The 3dd route keeps the
preserve index/overlay and the zlib workspace live while the deferred endpoint
is authenticated and hashed, accounting for the +79,618-byte verification peak.
This is a lifetime difference in the exact-output path, not evidence that the
zlib buffer can be removed; source freshness, content binding, and recipient
limits must remain enforced.

The owner-admission difference is likewise concrete. e63 admits a
`SourceBackedEditor` from `SourceBackedPackage::source_catalog`; 3dd admits
the public `SourceBackedWorkbook` projection after the same OPC catalog read.
The latter additionally parses the workbook catalog, validates the sheet graph,
shared strings and styles, and constructs `SourceSheetData`, sheet-name order,
and `SourceInner` caches. The current measured sizes are
`SourceBackedPackage=384`, `SourceBackedWorkbook=16`, `SourceInner=1784`,
`SourceCatalogProof=24`, `Catalog=32`, `Deferred=128`,
`RestoreEndpoint=160`, exact endpoint pair `Option=328`, `Patch=336`, and
`SourceArtifact=152` bytes. These sizes show why the observed roughly-288-byte
per-endpoint delta cannot be assigned to one standalone allocation: the
endpoint is a small inline handle over Arc-backed package, publisher, catalog,
and source state, and the remaining difference is in their live ownership
graph. The profile did not show an unbounded duplicate catalog proof.

The facts needed before the first edit differ by family. A tab-state edit must
have the source execution/version check, workbook main part and content type,
recipient `read_limits`, parsed sheets/active position/defined names, the
workbook and sheet relationship/content-type closure, semantic tab projection,
and the exact touched `PartState`. A defined-name edit needs only the source
execution/version check, workbook main part/XML, and `raw::parse_catalog` under
the recipient policy; it does not need sheet bodies, sheet handles, or the
semantic sheet projection. The public `SourceBackedWorkbook` read-only API
currently promises sheet names, active position, selectors, and validated graph,
shared-string, and style behavior at open, so its broader `SourceInner`
projection cannot be deleted as an optimization without changing that API.
The admitted OPC package already shares content-types, package relationships,
parts, and source limits, but it does not contain parsed XLSX sheet metadata.

A future repair can therefore defer the unused workbook projection for an
edit-only path, or share it from an already admitted source owner, but it must
keep the current guard-before-parse ordering. Shared catalog proofs may be
reused only when the recipient `SourceLimits` are equal; a stricter recipient
must reparse and revalidate under its own limits. Names changed by an edit must
produce a new proof while retaining the lineage policy. These constraints are
also covered by the focused stricter-recipient test and the deferred lineage
policy test; no policy or capability is waived here.

The source validation for 3dd passed the focused defined-name and tab-state
suites, the private deferred-policy test, full `litchi-xlsx` library tests,
all-feature warning-denied Clippy, and formatting. The allocator receipt has
exact forward/inverse output checks, but the positive live/peak and admission
deltas remain returned work for a later repair and native acceptance run.

## Catalog-endpoint deduplication follow-up (913ed0f3bd)

Commit `913ed0f3bd` removes the duplicate `Catalog` value from each deferred
restore endpoint. Equal-policy replay now borrows the still-live pending
owner's catalog proof; if that owner has been dropped, the immutable publisher
reconstructs the catalog under the recipient package's limits. A stricter
recipient always takes the latter path. The focused owner-drop replay test
passed, preserving exact forward and inverse bytes and recipient-policy
validation.

The protected release allocator binary is outside the repository at
`(deleted-coord-artifact)`
(10,470,200 bytes, SHA-256
`fc7e3184e6dccbca758ae069810d692664c082479d820ada232975bd3e5424ac`). It was
built from `913ed0f3bd` with the locked perf-baseline manifest, allocator
features, two jobs, and a distinct target directory. The matching e63 binary
and the 913e binary ran the same 64-history/2-warmup count-only driver against
the same fixture; the receipt is outside the repository at
`(deleted-coord-artifact)`.
The e63 JSON SHA-256 is
`509453270b20147fbc034a02e0d1842abf5f60b1bc9d61e00c9c050f37cdc098`; the 913e
JSON SHA-256 is
`08ad27c2c1c3e82ba61d67598318dcbeebb8f197c7c0b774522a8ed3c7a686eb`.
Both runs reported 64 exact forward and 64 exact inverse checks, with equal
fixture and output hashes. The deduplication did not change the measured
retained-history or admission deltas:

| interval | e63 | 913e | change vs e63 |
| --- | ---: | ---: | ---: |
| retained history live bytes after | 167,697 | 228,433 | +60,736 |
| retained history peak live after | 427,929 | 427,929 | unchanged |
| independent replay live bytes after | 167,953 | 232,601 | +64,648 |
| independent replay peak live after | 606,577 | 686,195 | +79,618 (+13.13%) |
| source-owner admission allocated bytes | 23,346,752 | 24,410,208 | +1,063,456 (+4.56%) |

This candidate remains returned. The catalog handle removal is safe and
correct, but it does not explain the retained endpoint/source-owner graph or
the preserve-overlay and zlib workspace lifetime seen during replay. A further
repair must identify and reduce those live objects while retaining exact
artifact authentication, recipient limits, and the public projection's
read-only capabilities.

## Deferred endpoint and snapshot sharing (5e0672b53c)

Commit `5e0672b53c870eb567bc9cd39d4ab0ed7959c56b` shares one exact restore
endpoint between adjacent deferred patch records and caches the immutable
tab-state source closure for repeated edits on one source owner. The endpoint
cache preserves the existing weak owner fast path, exact before/after artifact
authentication, source-version checks, recipient limits, and deferred
publisher streaming. The snapshot cache is read-only and is populated from
the already admitted catalog proof; no output or source ceiling is changed.
The source validation passed `source_backed_tab_state` (15 tests),
`source_backed_defined_names` (7 tests), and all-target/all-feature
warning-denied Clippy for `litchi-xlsx`.

The protected candidate binary is outside the repository at
`(deleted-coord-artifact)` (10,425,736 bytes, SHA-256
`f6dcafed3330cafac233a38f9678ac976320586efb5d8b652bc8c27888b708d0`). It
was built with the capped W4 environment using
`cargo build --locked --release --features allocator-metrics --manifest-path tools/perf-baseline/Cargo.toml --bin xlsx_tab_state_history_alloc`, with `CARGO_BUILD_JOBS=2` and a distinct target directory. The shared `Cargo.lock` SHA-256 is
`aa4300aff7f76cb467a2a6bc5896e706b3f5b506feed1d16b45f7422a0a9ba23`; the
benchmark manifest SHA-256 is
`52e349415b12b80765e810f47b4f6ac2e76cc092566eb257efc463e07d25fc2f`.

The count-only pair used the same two-defined-name fixture, 64 retained
steps, two warmups, and exact 64/64 forward and inverse checks. The receipt is
outside the repository at
`(deleted-coord-artifact)`.
The e63 JSON SHA-256 is
`509453270b20147fbc034a02e0d1842abf5f60b1bc9d61e00c9c050f37cdc098`; the
candidate JSON SHA-256 is
`f6b8eaa5222c7092c92615d5888c630b83e0199fbb5e1c120b71f5ed0e54f598`.
Fixture and forward/inverse output hashes are identical between the pair.

| interval | e63 | 5e0672b53c | change vs e63 |
| --- | ---: | ---: | ---: |
| retained history allocated bytes | 4,834,958 | 3,887,975 | -946,983 (-19.59%) |
| retained history allocation calls | 46,730 | 35,129 | -11,601 (-24.83%) |
| retained history live bytes after | 167,697 | 136,291 | -31,406 |
| retained history peak live after | 427,929 | 427,929 | unchanged |
| independent replay allocated bytes | 103,381,056 | 102,867,500 | -513,556 (-0.50%) |
| independent replay allocation calls | 104,576 | 73,582 | -30,994 (-29.64%) |
| independent replay live bytes after | 167,953 | 140,459 | -27,494 |
| independent replay peak live after | 606,577 | 594,117 | -12,460 (-2.05%) |
| source-owner admission allocated bytes | 23,346,752 | 24,427,744 | +1,080,992 (+4.63%) |
| source-owner admission allocation calls | 27,808 | 38,592 | +10,784 (+38.78%) |

The endpoint and snapshot changes materially improve retained history and
replay memory while preserving exact output. The owner-admission interval is
still returned: it creates a fresh source owner for each one-shot edit, so the
first version of the snapshot cache retained a full parsed closure after that
edit was dropped. A follow-up is being prepared to arm the cache only after a
second successful use, preserving the repeated-history benefit without
retaining a one-shot admission snapshot. No native latency, RSS, or final W4
performance acceptance is claimed from this count-only result.
