# 0532 current-head CFB ownership source review

`scope: read-only OLE2/CFB source and backlog review`

`source: 70e04d90181847ffe24c2ec8b8d5e1cd3b47981f`

`performance_claim: none`

This review reads the current [`file.rs`](../../../../crates/litchi-cfb/src/file.rs)
(`sha256 62f7b357d35b5bb2921336bb6a983a8d7e9ec58c30ae694a3be4c89e9c251788`),
the frozen [0532 plan](plan.json), and the retained [0523 CFB attribution](../../changes/0523-cfb-open-allocation-attribution.md)
and [0524 visited-bit decision](../../changes/0524-cfb-visited-bit-evaluation.md).
The generated 0532 baseline source manifest is SHA-256
`c2fbb838229f7fcb40d51d7d8bf90721d7ff2e9658601db60e28342ccd6bd514`.
The queue context is recorded in the current [hotspot inventory](../../HOTSPOTS.md)
and [phase report](../../REPORT.md).
It does not adopt an optimization or claim a latency gain.

## Decision

The next action is the planned current-head constructor attribution plus
disassembly of `OleFile::claim_sector`, with separate owners for
`validate_stream_allocations` and `validate_physical_sector_layout`. The
highest-justified source hypothesis after that baseline is a mechanical
success-path layout change: move the three formatted error constructions into
private `#[cold] #[inline(never)]` helpers, keep the exact current checked body,
and give `claim_sector` an ordinary `#[inline]` hint so release code can inline
the successful `Result` path. This does not fuse chain collection and claiming,
skip a conversion or bounds check, or alter any error text.

The hint is conditional. The standalone benchmark declares its own workspace and has no explicit
LTO release profile; the repository root LTO setting does not apply to it.
Disassembly must establish that a call, `Result` return sequence, or hot-block
formatting setup remains in this measured configuration. If the current binary already has the desired layout, this
hypothesis is closed without a source change. `#[inline(always)]`, unsafe
indexing, and error construction on the success path are outside the review.

The fresh 0532 normal-build assembly now gives the required code-shape
evidence: all six `claim_sector` monomorphizations are **363 bytes** and have
the same `push %rbx; sub $0x70,%rsp` prologue, checked bounds/ownership
branches, and distant success stores/`Ok` return; the `format!` setup is in the
error branches ([assembly index](assembly-index.json),
[`claim_sector-0` dump](baseline/assembly-claim_sector-0.stdout)). This is
static layout evidence only. It does not establish dynamic counts, latency, or
a benefit in other compiler or consumer build configurations. The proposed cold-helper /
ordinary-inline comparison remains conditional on the parent-constructor
profile and disassembly of the resulting caller code. The current
`assembly-validate_stream_allocations-0.stdout` already contains explicit calls
to `claim_sector`, so the helper is not universally inlined in this binary.

The residual-owner evidence justifies measuring this boundary, not a gain
estimate. The final 0523 XLS constructor profile reported disjoint exclusive
shares of `collect_exact` **40.08%**, `claim_sector` **17.81–17.82%**,
`validate_physical_sector_layout` **14.25–14.26%**,
`validate_stream_allocations` **14.16–14.17%**, and `load_fat`
**1.79–1.80%** ([0523 profile analysis](../change-0523/profile-analysis.json)).
The 0524 visited-bit candidate lowered constructor instructions only
**1.17–1.19%**, left allocation vectors unchanged, and regressed the direct
CFB few-large p50 by **2.35–2.41%**; that candidate is closed. The 0531 OOXML
candidate was separately rejected at dense-sparse repeat 2 (**0.8937%** p50,
**0.6056%** mean; [0531 final comparison](../change-0531/final-native-comparison.json)),
so it supplies no OLE2 result.

## Current call and validation boundary

`OleFile::open_with_limits` allocates the physical role map and runs the
pipeline in this order: `load_fat`, `load_directory`, optional
`load_minifat`, `validate_stream_allocations`, then
`validate_physical_sector_layout` ([`file.rs#L539`](../../../../crates/litchi-cfb/src/file.rs#L539)
and [`file.rs#L676`](../../../../crates/litchi-cfb/src/file.rs#L676)). The
role map is a fallibly allocated dense `Vec<PhysicalSectorRole>` indexed only
by physical sectors ([`file.rs#L133`](../../../../crates/litchi-cfb/src/file.rs#L133)
and [`file.rs#L676`](../../../../crates/litchi-cfb/src/file.rs#L676)).

The shared positional ingress invokes that parser once before retaining the
validated index ([`shared.rs#L472`](../../../../crates/litchi-cfb/src/shared.rs#L472)
and [`shared.rs#L504`](../../../../crates/litchi-cfb/src/shared.rs#L504)).
The measured XLS source owner reaches it through
`SourceBackedWorkbook::from_read_at_with_limits`
([`source.rs#L726`](../../../../crates/litchi-xls/src/workbook/source.rs#L726));
XLS validation uses the same shared ingress
([`validation.rs#L304`](../../../../crates/litchi-xls/src/validation.rs#L304)).
PPT package opening and PPT validation also use `SharedOleFile`
([`package/source.rs#L54`](../../../../crates/litchi-ppt/src/package/source.rs#L54),
[`validation.rs#L402`](../../../../crates/litchi-ppt/src/validation.rs#L402)).
Legacy DOC and direct XLS/PPT package readers call `OleFile::open` and thus
the same CFB pipeline ([`doc/package/codec.rs#L85`](../../../../crates/litchi-doc/src/package/codec.rs#L85),
[`xls/workbook/package.rs#L80`](../../../../crates/litchi-xls/src/workbook/package.rs#L80),
[`ppt/package/codec.rs#L77`](../../../../crates/litchi-ppt/src/package/codec.rs#L77)).
These are caller boundaries, not pooled performance evidence; DOCX OPC and
other OOXML clocks remain separate.

## Source facts and proof obligations

`claim_sector` converts the hostile `u32` sector to `usize`, uses checked
`get_mut` on the physical role map, rejects an existing role, writes the new
role, and returns `Ok(())` ([`file.rs#L1025`](../../../../crates/litchi-cfb/src/file.rs#L1025)).
Its exact error sequence is:

1. `"{role} sector {sector} does not fit usize"` on conversion failure;
2. `"{role} sector {sector} is outside the file"` on a missing physical slot;
3. `"Sector {sector} is claimed by both {old} and {new}"` on overlap; and
4. the role write on success.

The direct callers claim header/DIFAT FAT locations before storing them
([`file.rs#L762`](../../../../crates/litchi-cfb/src/file.rs#L762)), and
`claim_chain` applies the same check to MiniFAT, directory, root mini-stream,
and regular-stream chains ([`file.rs#L1051`](../../../../crates/litchi-cfb/src/file.rs#L1051)).
The chain collectors validate against a FAT or MiniFAT table, not necessarily
the physical role-map length, so that earlier table check is not a proof that
`claim_sector`'s physical bounds check can be removed.

`validate_stream_allocations` first fully collects and validates the root
chain, then claims its physical sectors and retains the vector. For each
regular stream it fully walks the reusable scratch chain and then calls
`claim_chain`; MiniFAT streams instead perform mini-sector bounds and duplicate
ownership checks in a separate bit set
([`file.rs#L1058`](../../../../crates/litchi-cfb/src/file.rs#L1058)).
`collect_exact` reserves its chain vector and visited map fallibly, checks
conversion, table bounds, cycles, terminal markers, and invalid markers in
that order, and resets scratch state on errors
([`file.rs#L2778`](../../../../crates/litchi-cfb/src/file.rs#L2778)).

The final physical pass checks `fat.get(sector)` **before** testing whether the
role is unclaimed, then rejects an unclaimed physical sector whose marker is
not `FREESECT` ([`file.rs#L1157`](../../../../crates/litchi-cfb/src/file.rs#L1157)).
The lookup order is observable when the FAT is short or an earlier unclaimed
sector has a bad marker.

## Ranked, proof-gated hypotheses

### 1. `claim_sector` cold errors and checked success inlining

The exact candidate is to extract only these existing `OleError::CorruptedFile`
formatters into cold, never-inlined private helpers: conversion failure,
missing physical slot, and overlap. The wrapper keeps `try_from`, `get_mut`,
the overlap comparison, the role assignment, and `Result<(), OleError>`
unchanged. A normal `#[inline]` hint may then let `load_fat` and `claim_chain`
specialize their constant role arguments and remove a private call/return
boundary on success.

This is fallibility-preserving: no allocation occurs on the successful path,
and each malformed path still executes the same `format!` with the same role
labels and text. It is error-order preserving because no check moves. A
candidate must be rejected if disassembly shows no remaining boundary, if the
compiler duplicates the error formatting into hot blocks, or if any semantic,
allocation, or malformed-input oracle changes. The 0532 parent-constructor
profiles should retain positive incoming scope and distinct exclusive owners;
per-sector dumps are unnecessary.

### 2. Pair the physical role and FAT walks while retaining the missing-entry tail

`validate_physical_sector_layout` can be investigated with a no-allocation
iterator layout that visits exactly `0..min(sector_roles.len(), fat.len())`
using paired role/FAT references, then emits the existing missing-entry error
at `sector == fat.len()` only when `fat.len() < sector_roles.len()`.

The proof is direct: for every common index the current `get` returns `Some`
and performs the same unclaimed-marker test; if the role map is longer, the
current loop reaches the first missing index only after all common indexes
have passed, which is exactly when the explicit tail error would run. FAT
entries beyond the physical role map remain ignored. The transformation adds
no reservation, unsafe access, or new fallible operation.

Do not replace this with `if role != Unclaimed { continue }` before the FAT
lookup or with an upfront length error. Either form can suppress a missing-FAT
error or move it ahead of an earlier unclaimed bad marker. Disassembly and
constructor attribution must establish whether the `Option` bounds branch is
material before staging this secondary candidate.

### 3. Measure the duplicate chain walk; do not fuse it

The sequence `collect_exact`/`collect_sector_chain_exact` followed by
`claim_chain` visibly walks a validated chain vector twice for physical
ownership ([`file.rs#L1067`](../../../../crates/litchi-cfb/src/file.rs#L1067),
[`file.rs#L1133`](../../../../crates/litchi-cfb/src/file.rs#L1133)). This is a
valid attribution question. It is not a safe fusion candidate from this
source: claiming during collection could report an overlap before a later
cycle, short-chain, or marker failure and could leave partial role mutations.
The current full validation also preserves the scratch reset and exact
fallible-reservation boundaries.

Root-chain storage is retained for positional reads, and the MiniFAT ownership
bit set is a different logical namespace. Any future design would need a
separate proof of error precedence, rollback or commit behavior, and allocation
accounting before it could alter this order. No such design is selected here.

### 4. Keep the remaining branch/layout observations as review leads

The `contains` plus `insert` pair for mini-sector ownership
([`file.rs#L1109`](../../../../crates/litchi-cfb/src/file.rs#L1109)) resembles
the closed 0524 visited-bit experiment. It is not a reason to revive that
generic fusion: 0524's native result and few-large regression already reject
it as the next action.

The count checks in `open_with_limits` and `load_fat` are intentionally
duplicated ([`file.rs#L650`](../../../../crates/litchi-cfb/src/file.rs#L650),
[`file.rs#L744`](../../../../crates/litchi-cfb/src/file.rs#L744)). Removing the
earlier checks would permit a fallible role-map allocation before a malformed
count is rejected; removing the loader checks would weaken its private-call
defense. This is a maintenance question, not a performance candidate.

`validated_directory_entries` and `build_storage_tree_iterative` each use a
directory ownership/visited bit set ([`file.rs#L1217`](../../../../crates/litchi-cfb/src/file.rs#L1217),
[`file.rs#L1500`](../../../../crates/litchi-cfb/src/file.rs#L1500)), but they
enforce different ordering, reachability, and public-tree construction
contracts. Retaining a merged representation would add state and fallible
allocation without a current owner measurement. Likewise, pre-filtering
stream directory entries or changing the enum/role-map representation is not
justified by the present evidence.

## Capture boundary and exclusions

Use the frozen 0532 matrix: two native repeats with 20 warmups and 1,000
samples, CFB tiny/many-small/few-large guards, two 30-sample operation-local
allocator repeats, and two five-dump parent-constructor profile repeats. Keep
allocator elapsed time and whole-child hardware counters out of native or
operation-local claims. Bind source, binary, corpus, positive incoming owner
edges, and exact CFB/XLS error, source, overlap, marker, and output oracles.

The current 0523/0524 protocols require an end-to-end native result, not a
Callgrind reduction alone. The CFB few-large profile remains a sensitivity
guard, while tiny and many-small protect against extrapolating from one shape.
No OLE2 speedup, allocation reduction, RSS, cold-cache, physical-I/O,
producer, or scaling claim follows from this source review. OLE2 and OOXML
remain active; ODF is deferred until that goal completes, and iWork is outside
scope.
