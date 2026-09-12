# 0534 CFB physical role/FAT paired-prefix source review

`scope: read-only OLE2/CFB source and contract review before implementation`

`revision: 7c1d3911da286b8a9dd0fde15475bbbeed2f84c9`

`baseline source: crates/litchi-cfb/src/file.rs SHA-256
bb1928970dc3d652c5091d521cad86c8873ab8352eee5d2eed18a6aa4dd4c20b`

`performance_claim: none`

This review covers the frozen [0534 plan](plan.json), the current
[`validate_physical_sector_layout`](../../../../crates/litchi-cfb/src/file.rs#L1178),
the existing physical-role and FAT callers, and the accepted ADR bindings in
[`adr-manifest.json`](adr-manifest.json). The 0534 manifest is byte-identical
to the 0533 manifest; all 30 accepted ADR/index entries remain unchanged. No
Rust source, test, build, benchmark, or profiler was run or changed for this
review.

## Candidate boundary

The proposed runtime delta is limited to the final physical reconciliation
loop. The current method enumerates every `sector_roles` entry, obtains the
same-index FAT entry with checked `get`, and then checks the unclaimed marker.
The candidate should enumerate the common prefix with a paired iterator:

```text
for (sector, (role, entry)) in sector_roles.iter().zip(fat).enumerate() {
    if role == Unclaimed && entry != FREESECT {
        return the existing marker error
    }
}
if fat.len() < sector_roles.len() {
    return the existing missing-FAT error at sector = fat.len()
}
```

The pseudocode is a behavior description. The implementation must retain the
reference/value dereferences needed by the actual iterator types and must not
collect either slice, copy the FAT, or replace the checked method with an
unchecked index. The candidate has no public API, dependency, unsafe-code,
resource-policy, ownership, or validation-policy scope.

## Exact behavior obligations

For every index `i` in the common prefix, the candidate must perform the same
check in ascending sector order:

1. Read the role and the same-index FAT marker from the two borrowed slices.
2. If the role is `PhysicalSectorRole::Unclaimed`, accept only `FREESECT`.
3. On another marker, return `OleError::CorruptedFile` with the exact text
   `unclaimed physical sector {i} has FAT marker 0x{marker:08X}`.

After the prefix is exhausted, a shorter FAT is the only remaining error
case. If `fat.len() < sector_roles.len()`, the candidate must return
`OleError::CorruptedFile` with exactly:

```text
FAT does not contain an entry for physical sector {fat.len()}
```

That index is the first missing physical sector. It must be formatted as the
same decimal `usize` index that the current `get`-based loop formats. A
claimed role at that missing index does not change the result: the missing-FAT
error still wins after all available prefix entries have passed.

The length cases are part of compatibility:

| `sector_roles.len()` versus `fat.len()` | Required result |
| --- | --- |
| both zero | `Ok(())`; no iteration and no error allocation |
| equal | Check every role/marker pair in order; only unclaimed non-`FREESECT` markers fail |
| roles longer | Check the complete common prefix first, then report the missing entry at `sector = fat.len()` |
| FAT longer | Check every role entry and ignore all extra FAT padding, regardless of its marker |

If an earlier common-prefix unclaimed entry has a bad marker, that marker
error must be returned before a later short-FAT error. Conversely, a valid
prefix followed by a missing FAT entry must return the missing error rather
than treating the missing role as unclaimed or inventing a marker. An empty
role map with a nonempty FAT remains successful because all FAT entries are
outside the physical file and are padding from this method's point of view.

Only `Unclaimed` is interpreted by this pass. `Fat`, `Difat`, `Directory`,
`MiniFat`, `MiniStream`, and `RegularStream` entries must skip the
`FREESECT` comparison, exactly as they do today. The FAT/DIFAT marker checks,
chain checks, and stream ownership checks remain in their existing callers.
The candidate must not turn this final reconciliation pass into a check that
every claimed role has a particular marker.

The existing exact messages are therefore a source contract:

```text
unclaimed physical sector 0 has FAT marker 0xFFFFFFFE
FAT does not contain an entry for physical sector 1
```

The first example uses `ENDOFCHAIN`; the second assumes two role entries and
one FAT entry. The marker must retain eight uppercase hexadecimal digits.

## State, safety, and resource boundaries

The method takes `&self`, so the candidate must remain observational. It must
not mutate `sector_roles`, `fat`, or any other field on either success or
failure. The role map has already been populated by `load_fat`,
`load_directory`, `load_minifat`, and `validate_stream_allocations`; this
method only reconciles it with the final FAT table. There is no rollback or
new partial-publication behavior to introduce.

`iter().zip(...).enumerate()` performs no allocation and does not form a
physical byte offset. `Vec::len()` is already a checked `usize` value, and
the loop index is produced by safe slice iteration. No `u32` to pointer or
`usize` arithmetic is needed here, so the candidate must not add `sector *
sector_size`, unchecked indexing, pointer casts, or wrapping arithmetic. A
`u32::MAX` FAT marker is ordinary data (`FREESECT`) and must not be confused
with a physical index. There is no meaningful test that constructs a
`usize::MAX`-length vector; the no-pointer-arithmetic proof and normal
malformed-input gates cover that boundary.

The success path must remain allocation-free. Formatting an existing typed
error on a failure path may allocate as it does in the baseline, but the
candidate must not add an intermediate vector, iterator collection, string,
or error object on success. The paired borrow must end before the length
comparison if the compiler requires an explicit scope; no mutable alias or
interior mutation is permitted.

## Callers and compatibility

`OleFile::open_with_limits` calls this method only after the FAT, directory,
MiniFAT, and stream allocations have been validated and their physical roles
claimed. The candidate must retain that order. Moving the pass earlier would
observe an incomplete role map and could reject valid files; moving it later
would weaken the final safety fence.

`load_fat` builds a table whose length is based on complete FAT sectors, so the
FAT can contain padding entries beyond the number of physical file sectors.
The existing `tolerates_nonfree_fat_padding_beyond_the_physical_file` test
changes such padding to `ENDOFCHAIN` and expects a successful open. The
paired-prefix candidate must preserve that compatibility and must not inspect
the extra FAT tail. This is also why checking `fat.len() != sector_roles.len()`
or requiring extra entries to be `FREESECT` would be a semantic regression.

The final pass must continue to reject an unclaimed physical sector whose
FAT entry is not `FREESECT`, even when that marker is `ENDOFCHAIN`, `FATSECT`,
`DIFSECT`, `MAXREGSECT`, zero, or another value. It must not infer that a
claimed role is invalid from one of those values. Chain traversal and the
dedicated FAT/DIFAT marker checks remain the authority for those other cases.

No `claim_sector` or `claim_chain` behavior changes are implied. In
particular, the accepted 0533 checked conversion, bounds, conflict order and
role publication remain intact, and `validate_stream_allocations` retains
its collect-then-claim sequencing. The candidate does not change the mini-
sector namespace, directory metadata binding, source ownership, or public
OLE2/XLS/DOC/PPT behavior.

## Accepted ADR constraints

The frozen 0534 manifest records the same accepted hashes as 0533. The
relevant obligations are:

| ADR | obligation for this candidate |
| --- | --- |
| 0001, 0002, 0024 | Keep the change in the existing `litchi-cfb` physical container owner; add no public layer or peer dependency. |
| 0003 | Preserve the opened snapshot's validation state and deterministic failure behavior; the reconciliation pass cannot mutate the snapshot. |
| 0005 | Treat paired iteration as a performance hypothesis only; measure native latency, allocations, profiles, and RSS under the frozen plan. |
| 0006 | Retain complete structural validation, typed deterministic errors, marker precedence, and observed producer padding compatibility. |
| 0008 | Bind source, binaries, corpora, raw vectors, quality results, and any generated-code evidence to the exact measured source. |
| 0010, 0011 | Leave archive and OOXML physical ownership boundaries unchanged; no package handoff or preservation rule is part of this slice. |
| 0026 | Keep directory metadata and physical-role validation sequencing intact; no directory binding is inferred from this loop rewrite. |

The accepted program priority remains OLE2/OOXML, with ODF deferred until
that goal completes and iWork outside scope.

## Measurement and assembly guard implications

The paired-prefix loop is a narrow hot-loop hypothesis, not proof of a
speedup. The frozen plan still requires the nine XLS workflows and three CFB
shapes in its native ABBA order, the four primary XLS p50 improvements in both
repeats, the constructor and physical-reconciliation exclusive instruction
gates, unchanged allocation calls/bytes/incremental peak, and all applicable
quality checks. Every matched latency/RSS change over five percent and every
same-build variation over five percent remains reviewable. Allocator elapsed
time is separate instrumentation and cannot replace native timing.

Assembly review should confirm that the measured candidate uses the paired
common-prefix loop, emits one post-loop missing-length branch, does not read
past either slice, and retains the unclaimed-marker branch and exact error
paths. It should compare the final measured binary, not infer behavior from
the `zip` spelling or from root-workspace LTO settings. A changed rebuild or
test-only source state requires the fresh full protocol in the plan.

## Source disposition

The candidate is safe to stage for differential measurement if it preserves
the paired-prefix order, exact errors, extra-FAT compatibility, and
allocation-free `&self` behavior above. This review does not approve
production retention. The final decision still requires the focused private
contract test, the complete CFB/XLS/DOC/PPT quality matrix, exact source and
binary custody, and the frozen native, allocation, profile, and assembly
gates. If admission fails, the runtime change should be restored while tests
that exercise this unchanged private contract may remain.
