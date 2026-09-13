# 0552 OLE2 next target: validated directory-record handoff

`status: measurement target only`

`performance_claim: none`

`scope: OLE2/OOXML active; ODF deferred; iWork excluded`

The next OLE2 experiment should measure a private handoff from the validated
CFB directory records to the public `DirectoryEntry` graph in
`OleFile::load_directory`. The current open path validates every 128-byte
directory record and then decodes the names of the reachable records again
while building the public graph. The handoff can remove that repeated name
decode, and potentially the repeated public-record field extraction, while
leaving the validation pass and graph checks in place. This document selects
the mechanism for a fresh experiment; it records no runtime improvement.

## Current path and measured reason to inspect it

The current owner loads the FAT, loads the directory, loads the MiniFAT, and
then validates stream allocation and physical layout in that order
([`file.rs#L731-L750`](../../../../crates/litchi-cfb/src/file.rs#L731)). Within
the directory phase, `validated_directory_entries` parses and validates every
record ([`file.rs#L1199-L1327`](../../../../crates/litchi-cfb/src/file.rs#L1199)),
including its name and `DirectoryNameData`. After that pass,
`load_directory` calls `parse_directory_entry` for the root and reachable
records, and `build_storage_tree_iterative` calls it for each queued SID
([`file.rs#L997-L1017`](../../../../crates/litchi-cfb/src/file.rs#L997),
[`file.rs#L1459-L1579`](../../../../crates/litchi-cfb/src/file.rs#L1459)). The
public parser calls `decode_utf16le` a second time
([`file.rs#L1466-L1474`](../../../../crates/litchi-cfb/src/file.rs#L1466)).

The retained 0547 baseline contains mechanism evidence for this path. The
numbers below are one positive timed constructor dump; the five positive
dumps for each CFB job repeat the same named function values. `summary_ir` is
the constructor's collected Ir for that dump. Function self Ir is a disjoint
function-local diagnostic; it is not a hardware-cycle or latency measurement.

| Retained workload | Directory shape | `summary_ir` | `validated_directory_entries` self Ir | `parse_directory_entry` self Ir | `decode_utf16le` self Ir | Direct decoder calls seen |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `cfb-many-small` | 256 1-KiB streams plus root | 2,794,366 | 206,982 (7.407%) | 29,298 (1.048%) | 213,344 (7.635%) | 257 |
| `cfb-few-large` | 4 4-MiB streams plus root | 2,052,759 | 4,100 (0.200%) | 570 (0.028%) | 3,680 (0.179%) | 5 |
| `cfb-tiny` | 3 512-byte streams plus root | 48,983 | 3,003 (6.131%) | 456 (0.931%) | 2,848 (5.814%) | 4 |

The many-small row is the useful signal: the accepted synthetic corpus has
archive SHA-256
`dca7a96c4548b37a7fd971835afffd0529433b34e07a192e3b9964e4803c634c`, and
the profile observes 257 public name decodes for the 256 stream entries and
root. The retained validator's `directory_name_data` self Ir is 573,640 in
that dump. That work creates the validated comparison key and remains
necessary for MS-CFB ordering and lookup; it is not counted as removable
duplicate work here.

The source-backed XLS owner also carries the same CFB directory shape in the
`xls-comments-opaque-heavy` corpus (256 comments, 257 archive entries). Its
positive 0547 dumps show `parse_directory_entry` self Ir 1,368 and
`decode_utf16le` self Ir 4,404. The validator is inlined or absent as a named
function in that owner profile, so the XLS row corroborates the public decode
surface while a fresh profile must establish its full directory attribution.

This is enough evidence to select a measured target, with two limits. 0547
was a diagnostic-only single-source baseline, and its operation-local
allocation, live-peak, native-latency, and hardware-cycle fields were
unavailable. The profile also cannot establish that every instruction inside
the public parser is removable. Those facts require a candidate profile and
matched end-to-end evidence.

## Candidate mechanism

The experiment should keep `validated_directory_entries` as the authority and
make its validated record available while the public graph is built. The
minimum candidate outcome is a move of the validated normal-entry name into
the corresponding `DirectoryEntry`, avoiding the second UTF-16 decode and a
second name construction. If the ownership proof requires it, the candidate
may carry the already-read public scalar fields and CLSID formatting result in
the same private validated seed, then move those values into the public entry.
That wider seed is conditional on the memory measurement below; adding a
larger retained cache without a peak bound is not an optimization result.

The candidate must build only the reachable public entries in the existing
SID-indexed shape. It must preserve the existing fallible reservations and
resource labels for `directory entries`, `directory name comparison data`,
directory traversal state, and any new private seed. The final graph and name
cache must still be installed only after the complete directory phase succeeds.

The root needs an explicit two-view rule. Validation canonicalizes the
supported classic-Mac SID-0 encoding to `Root Entry`, while the historical
public parser retains its decoded public name. A handoff must retain the
canonical validation name for ordering and `dir_name_data` and the historical
public value for `DirectoryEntry.name`. Normal accepted records may share the
validated name; the root exception cannot be folded away.

The candidate must retain these semantics:

- raw-record extent checks, name-length checks, UTF-16 and forbidden-name
  validation, node-color validation, version-3 stream-size masking, and all
  existing `OleError` text and precedence;
- SID alignment, root-at-zero rules, sibling ordering, cycle and repeated-SID
  rejection, storage ownership, and the iterative graph traversal;
- public `DirectoryEntry` fields (`name`, type, sibling and child SIDs, CLSID,
  start sector, size, MiniFAT classification, and empty `children` vectors);
- the comparison cache's SID alignment and the final internal agreement check,
  or a documented equivalent guard that preserves the same corruption
  rejection and source-index proof; and
- the open phase order, directory byte limit, zero-filled short-sector tail,
  source reads, physical claims, and later stream-allocation and physical
  reconciliation passes.

Moving a `String` or `DirectoryNameData` can lower duplicate ownership, but
retaining the whole validated seed until graph publication can raise the live
peak. A candidate that clones the validated name, eagerly materializes all
entries, changes an allocation's failure point, or relies on an infallible
reserve should be closed unless the matched measurements and error proof
justify it.

## Fresh measurement gate

Before applying code, capture a current-source baseline and a candidate with
the same positive timed constructor scopes. The CFB matrix should include the
0547 tiny, many-small, and few-large corpora; the XLS owner should include the
source-backed `xls-comments-opaque-heavy` open-and-one-cell path. Record the
directory entry count, nonempty count, reachable count, total UTF-16 name
units, root encoding form, and directory byte count for every case.

The profile must retain separate self and inclusive attribution for
`load_directory`, `validated_directory_entries`,
`parse_validated_directory_entry`, `parse_directory_entry`,
`decode_utf16le`, `directory_name_data`, `format_clsid`,
`build_storage_tree_iterative`, and relevant drop paths. Assembly or an
equivalent emitted-work record must show whether the candidate removes the
second decode/field work; a lower aggregate symbol count alone is
insufficient. The XLS owner profile must classify inlined validation work
explicitly.

Pair that attribution with operation-local allocation calls and bytes,
deallocation calls, and incremental live peak. Check the existing public
output and error oracles and the end-to-end XLS primary gate only after the
candidate has a valid private ownership design. The candidate remains a
measurement artifact if directory work is below profile noise, if name moves
turn into extra allocation or peak growth, or if any primary workflow
regresses. No Callgrind Ir delta, allocation delta, or output equality result
alone establishes a latency or speedup claim.

## Nearby mechanisms deliberately excluded

The selected target is higher in the open workflow and removes work that the
current source visibly performs twice. The following retained evidence closes
nearby alternatives for this experiment:

| Mechanism | Retained disposition |
| --- | --- |
| Brent/checkpoint collector state | 0548 rejected. Its decision records XLS-owned constructor Ir increases of 2.865% and 2.891% and collector self Ir increases of 5.855% (XLS) and 5.871% (CFB few-large); one primary p50 row regressed. A directory handoff does not add checkpoint state. |
| Checked `test_and_mark` / fused membership | 0549 rejected. All eight primary XLS p50 rows regressed by 2.080%–10.876% despite collector self Ir falling 2.939%–2.940%. The selected target does not rename or revisit that loop. |
| FAT-sector batching | Already retained by 0511; the old 9.30% `load_fat` figure predates the retained exact decode path, whose current attribution was 1.79% in the earlier owner review. |
| Chain scratch reservation/reuse | Already retained by 0190 and 0413. Reusing `SectorChainScratch` or changing exact reservation does not select a new mechanism. |
| Collect-and-claim fusion | 0523 records the error-order and partial-role-mutation hazards. It requires a different representation and rollback proof before measurement. |
| Paired physical reconciliation walk | 0534 was rejected after primary XLS regressions. The final unclaimed-sector check remains required for the physical ownership contract. |
| Per-stream chain cache in `SharedOleFile` | The retained cursor already resolves a chain ordinal once and advances without restarting traversal ([`shared.rs#L579-L645`](../../../../crates/litchi-cfb/src/shared.rs#L579)). A catalog-wide chain cache has no current consumer attribution and carries a clear retained-memory cost, so it is deferred. |

The accepted ADR constraints remain the private CFB owner boundary, fallible
bounded allocations, typed validation refusals, lossless behavior, matched
measurement, and physical package/directory ownership. They are recorded in
[ADR 0001](../../../adr/0001-priorities-and-api-layers.md), [ADR 0002](../../../adr/0002-crate-topology.md), [ADR 0003](../../../adr/0003-snapshots-edits-and-patches.md), [ADR 0005](../../../adr/0005-io-memory-and-performance.md), [ADR 0006](../../../adr/0006-validation-security-and-compatibility.md), [ADR 0008](../../../adr/0008-migration-and-verification.md), [ADR 0010](../../../adr/0010-facade-archive-ownership.md), [ADR 0011](../../../adr/0011-ooxml-physical-package-ownership.md), [ADR 0024](../../../adr/0024-current-topology.md), and [ADR 0026](../../../adr/0026-ole-directory-metadata-binding.md).

## Evidence custody

The selected profile and source identity are reproducible from these retained
hashes:

| Artifact or source | SHA-256 |
| --- | --- |
| [`change-0547/profile-analysis.json`](../change-0547/profile-analysis.json) | `6cd52ecef7af77671270f95983a6dff6c3cc7d414cf94667a2f23f020b1756bd` |
| [`change-0547/analysis.json`](../change-0547/analysis.json) | `5fea2d98ed11d1aa166e45327c99218d8f9f5a2ad3db81197ed323f1f0dedb18` |
| 0547 baseline source manifest | `db197f5ea518082646669db5ba229cc2ea4081f5d2b6a51357e07f9cc4090511` |
| [`profile-r1-cfb-many-small.json`](../change-0547/baseline/profile-r1-cfb-many-small.json) | `7c4b820d553ee6f542f5953d418217a8eed35949885f9927e2311a1c2c773a63` |
| [`profile-r1-cfb-few-large.json`](../change-0547/baseline/profile-r1-cfb-few-large.json) | `a6461ecfee3a1171085ca4eff9a9afe27a9bb9846d5758ab556ae20e0e2849c2` |
| current `crates/litchi-cfb/src/file.rs` (revision `bb5bfaa7be4fdffeae3fddaee3ed1266c3417520`) | `72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b` |
| current `crates/litchi-cfb/src/shared.rs` | `2aef59e5e7d20a8984b0c2f07d561865a87a74d927cb088d4a08ff85532c0624` |
| current `crates/litchi-xls/src/workbook/source.rs` (revision `98ea87658ce1ca2697a1dfc6901d72d8888cf5d4`) | `3e7dbde9f197771ee7766a597f3f27f6a0143b79f1cf2575dd52a9f8493f6d92` |
| [`change-0548/decision.json`](../change-0548/decision.json) | `b57ab8c9962935b8a5e48bf21227c5f3b7f558a3a9e7137e20b7b95fac637e33` |
| [`change-0549/decision.json`](../change-0549/decision.json) | `f08dacc2dc4791e9c7c916050b58369b111207ec2a6d19fc6ba38dab66c719fe` |

The underlying 0547 sub-operation table, which is limited to collector
diagnostics, has SHA-256
`ebe7466d7decbdf94f5e15e64a6e7c5659932b0abf860dc457871c5a7178ecff`. The
0548 and 0549 profile comparison artifacts have SHA-256
`37c8a8eaca6147217b85f8c1f186d178869b86adef48c6237e1b6e60f60d1c37` and
`b4560164ecbe02e98acbfb47946c9229f8bea73d67e9ea643f1eb48051762956`,
respectively. These records establish rejection and source continuity; they
do not establish a result for this directory handoff.

This audit read the retained files and current source only. It made no Rust,
driver, analyzer, build, test, or capture change.
