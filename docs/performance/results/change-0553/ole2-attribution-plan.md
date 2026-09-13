# 0553 OLE2 directory handoff attribution plan

`status: attribution plan only`

`performance_claim: none`

`scope: OLE2/OOXML active; ODF deferred; iWork excluded`

This is a small, independent campaign for the next OLE2 question. It is not
part of the change-0553 XLSX verdict, and it does not depend on that verdict.
This task changes only this document: it makes no Rust, driver, analyzer,
build, test, or capture change.

## Question and boundary

The committed 0552 target selects one fresh mechanism: let the private
validated CFB directory records seed the public directory graph. The current
path reads every 128-byte directory record in
[`validated_directory_entries`](../../../../crates/litchi-cfb/src/file.rs#L1199-L1327),
including its UTF-16 name and `DirectoryNameData`, and then parses the root and
reachable records again through
[`parse_directory_entry`](../../../../crates/litchi-cfb/src/file.rs#L1459-L1498)
while
[`build_storage_tree_iterative`](../../../../crates/litchi-cfb/src/file.rs#L1500-L1579)
builds the public graph. The public parser calls
[`decode_utf16le`](../../../../crates/litchi-cfb/src/file.rs#L2908-L2938)
again. The candidate hypothesis is that moving the already validated name and,
where justified, already extracted public fields into the reachable public
entry can remove that repeated work without retaining a second unbounded
directory representation.

The validation pass remains authoritative. The candidate must retain the
existing graph traversal, SID-indexed public entries, comparison cache,
physical claims, fallible reservations, error order, and staged final
publication in
[`load_directory`](../../../../crates/litchi-cfb/src/file.rs#L966-L1055).
The campaign decides whether emitted work and operation resources support this
handoff. It does not assume that a source-level move, a missing symbol, or a
lower Callgrind Ir count is a performance result.

Measure the name and scalar-field parts as separate private arms when the
isolated candidate is prepared. **N** moves the validated normal-entry name
into the public entry and leaves scalar extraction as an independently
attributed operation. **NF** extends that handoff with the public scalar
values (`CLSID`, masked size, start sector, MiniFAT classification, and the
already validated type/SID fields). NF is worth measuring only after N shows a
real name-work signal; adding a wider private seed is itself a memory cost to
measure. These arms are measurement variants, not production changes in this
task. A candidate must not claim field reuse merely because a raw record is
read once: the profile must show whether `format_clsid`, size masking, and
other public field construction actually disappear or simply move.

This split follows the current data shapes: the private
[`ValidatedDirectoryEntry`](../../../../crates/litchi-cfb/src/file.rs#L121-L131)
already carries the name, name data, type, color, and SIDs, while
[`RawDirectoryEntry`](../../../../crates/litchi-cfb/src/file.rs#L65-L98) and
the public parser additionally supply the CLSID, start sector, masked size,
and MiniFAT classification. The NF arm must prove that extending the private
record to carry those values removes enough public work to justify its live
storage.

This plan does not rename or revisit either rejected nearby experiment:

| Excluded mechanism | Retained disposition |
| --- | --- |
| Brent/checkpoint collector state | Change-0548 decision `b57ab8c9962935b8a5e48bf21227c5f3b7f558a3a9e7137e20b7b95fac637e33` rejected it after constructor and collector Ir increases and a primary p50 regression. |
| Checked `test_and_mark` / fused membership | Change-0549 decision `f08dacc2dc4791e9c7c916050b58369b111207ec2a6d19fc6ba38dab66c719fe` rejected it after all eight primary XLS p50 rows regressed. |

The selected mechanism comes from the committed
[`0552 target`](../change-0552/ole2-next-target.md), SHA-256
`24bbaa616075016fd4c682d158c469643e97a7d5753b4fe62561f364714b9897`.

## What is retained evidence

The 0547 profile is mechanism evidence from one repeat, five positive timed
dumps per job, warmup zero, and no paired candidate. Its `summary_ir` and
function self Ir values are Callgrind diagnostics. They do not establish
native latency, allocation, live peak, RSS, hardware cycles, or a speedup.

| Timed owner workload | Directory shape | `summary_ir` | `validated_directory_entries` self Ir | `parse_directory_entry` self Ir | `decode_utf16le` self Ir | `parse_directory_entry` → `decode_utf16le` calls |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| `cfb-many-small` | 256 1-KiB streams plus root | 2,794,366 | 206,982 | 29,298 | 213,344 | 257 |
| `cfb-few-large` | 4 4-MiB streams plus root | 2,052,759 | 4,100 | 570 | 3,680 | 5 |
| `cfb-tiny` | 3 512-byte streams plus root | 48,983 | 3,003 | 456 | 2,848 | 4 |
| `xls-owned` | Fixed XLS/CFB comments corpus | 2,262,324 | validator inlined/absent as a named function | 1,368 | 4,404 | 12 |

The many-small row is the clearest retained duplicate-name signal: the
profile saw 257 public decoder calls for the root and 256 stream records, and
`directory_name_data` self Ir was 573,640 in that dump. The validator's name
data is required for CFB ordering and lookup and is therefore not counted as
removable merely because the public graph may reuse a name. The few-large and
tiny rows provide scale and non-regression controls; their directory work is
small relative to their payload-shaped archives. The XLS row confirms that the
actual source-backed owner reaches the same public CFB parsing surface, while
its inlined validation requires fresh attribution.

The XLS number is deliberately scoped to the incoming edge from
`parse_directory_entry`: in retained raw dump
`baseline/profile-r1-xls-owned.callgrind.1`, that edge records
`calls=12` ([raw lines 2941-2943](../change-0547/baseline/profile-r1-xls-owned.callgrind.1#L2941-L2943)).
The separate `calls=126` at raw lines 11095-11096
([raw lines 11095-11096](../change-0547/baseline/profile-r1-xls-owned.callgrind.1#L11095-L11096))
is the outgoing edge from `decode_utf16le` to
`<alloc::string::String>::try_reserve`; it is an allocator-call edge in that
decoder body, not 126 decoder invocations and not a count of CFB directory
nodes. The current CFB source has only one non-test call site for
`decode_utf16le`, at [`parse_directory_entry` line 1473](../../../../crates/litchi-cfb/src/file.rs#L1473), so this correction does
not reclassify 126 mixed BIFF or other XLS text-decoder calls; it corrects the
edge direction and call-count meaning in the retained raw profile. The fresh
campaign must preserve these caller/callee boundaries when reporting
name-decoder work.

The actual retained corpus identities are:

| Harness identity | Archive bytes | Members / logical entry count | Per-entry bytes | Archive SHA-256 | Target and target SHA-256 |
| --- | ---: | ---: | ---: | --- | --- |
| `cfb-tiny-incompressible` | 3,584 | 3 / 3 | 512 | `186750b66895472e6c4b61bdd6e89d7cfd066baec1000cc0e9aab5d86457c0e8` | `benchmark_stream_00002.bin`, `df355b60021f82a84ec1ca06edcf7aea64a5272388c4eed239562fd63d3fceb3` |
| `cfb-many-small-incompressible` | 314,880 | 256 / 256 | 1,024 | `dca7a96c4548b37a7fd971835afffd0529433b34e07a192e3b9964e4803c634c` | `benchmark_stream_00255.bin`, `d558142d0dea1760aa463990aeb05298248820d4b06e509eed32ae6dc19fbce4` |
| `cfb-few-large-incompressible` | 16,912,384 | 4 / 4 | 4,194,304 | `4b732058fb9f06fa9166208b207dff2339e2c9caf6e6091a06648a27b3ac4cfa` | `benchmark_stream_00003.bin`, `57d0fc8d7b94f2ef821acdb06ac16a3cd450572b825d4e26e0d400df5e838706` |
| `xls-comments-opaque-heavy` | 16,995,840 | 10 archive members / 257 manifest entries | 2,097,152 opaque payload bytes | `6a57231ba681bc7bdd38d447ebd5348ef3b1fefedeefb1e61c97f22faa074e53` | `Workbook`, 80,946 bytes, `c78e03ba3743132e04b08bf6f4579ceb1c112a22c441c1e036381d3e06c6d041` |

The CFB rows use the exact harness generator
`litchi-cfb-synthetic-v1`, no compression, the deterministic incompressible
payload, and stream names `benchmark_stream_{index:05}.bin`. The XLS row uses
the exact generator `litchi-xls-comments-opaque-heavy-v1`: 256 comments,
eight incompressible 2-MiB opaque streams, the fixed `Workbook` target, and
the selected worksheet/cell oracle below. In the XLS manifest,
`entry_count: 257` is corpus identity for the comments-plus-extra content; it
is not a claim that the CFB directory has 257 nodes. The current 0547 XLS
profile's `parse_directory_entry` edge count is 12.

The available corpus builders and shape constants are in
[`tools/perf-baseline/src/lib.rs`](../../../../tools/perf-baseline/src/lib.rs#L151-L151)
and its `CorpusShape` table
([`lib.rs#L332-L369`](../../../../tools/perf-baseline/src/lib.rs#L332-L369)). The
XLS constants and generator are in
[`lib.rs#L242-L260`](../../../../tools/perf-baseline/src/lib.rs#L242-L260) and
[`lib.rs#L26025-L26106`](../../../../tools/perf-baseline/src/lib.rs#L26025-L26106).
The existing CLI selectors are `--case`, `--shape`, `--payload`,
`--corpus-manifest`, `--warmup`, and `--samples`
([`lib.rs#L11252-L11280`](../../../../tools/perf-baseline/src/lib.rs#L11252-L11280));
the campaign uses those selectors with the values above and adds no new
driver script.

## Minimal harness matrix

The first pass is a diagnostic pilot. It uses only existing harness case
values and existing generated corpora; it adds no selector or generator.

| Pilot job | Existing case value | Corpus selection | Why it is required |
| --- | --- | --- | --- |
| CFB many-small | `cfb_open`, shape `many-small`, payload `incompressible` | `litchi-cfb-synthetic-v1` | Highest retained repeated-name count and the primary scale signal. |
| CFB few-large | `cfb_open`, shape `few-large`, payload `incompressible` | `litchi-cfb-synthetic-v1` | Checks whether a name/field handoff regresses when payload work dominates. |
| CFB tiny | `cfb_open`, shape `tiny`, payload `incompressible` | `litchi-cfb-synthetic-v1` | Small-directory crossover and setup-sensitive control. |
| Actual XLS owner | `xls_owned_source_open_one_cell` | Fixed `xls-comments-opaque-heavy` manifest from `litchi-xls-comments-opaque-heavy-v1` | Exercises `SourceBackedWorkbook` through the production owned-source path and preserves 0547 owner continuity. |

The case parser and dispatch already accept
`cfb_open`, `cfb_list_streams`, `cfb_read_one`,
`xls_source_backed_open`, `xls_source_backed_open_list_worksheets`,
`xls_source_backed_open_one_cell`, `xls_owned_source_open`,
`xls_owned_source_open_list_worksheets`, and
`xls_owned_source_open_one_cell`
([case parsing](../../../../tools/perf-baseline/src/lib.rs#L11720-L11758),
[dispatch](../../../../tools/perf-baseline/src/lib.rs#L23102-L23203)). The
pilot needs only the four rows above. The following existing rows are cheap
controls when the pilot has a clear attribution signal:

* `cfb_list_streams` checks the public name set and count after an untimed
  open; `cfb_read_one` checks the selected stream bytes and digest.
* `xls_owned_source_open` isolates the actual owner with its worksheet-count
  query, and `xls_owned_source_open_list_worksheets` checks the public
  worksheet-name result.
* `xls_source_backed_open` and `xls_source_backed_open_one_cell` exercise the
  same directory owner through `InstrumentedSource`, including the source
  locality/version controls. They are paired controls, not substitutes for
  the actual owned-source owner row.

The fixed XLS layout requires archive size at least 16 MiB, `Workbook` size at
least 64 KiB, at least two worksheets, worksheet index 1, row 20, column 4,
and value `42.0`. The source-backed implementation enters
[`SharedOleFile::open_source_with_limits`](../../../../crates/litchi-cfb/src/shared.rs#L472-L526)
and then the same CFB open path; its cursor resolves each stream chain once
([`shared.rs#L579-L645`](../../../../crates/litchi-cfb/src/shared.rs#L579)).
The XLS source constructor and worksheet scan are in
[`source.rs#L726-L743`](../../../../crates/litchi-xls/src/workbook/source.rs#L726-L743),
[`source.rs#L1470-L1515`](../../../../crates/litchi-xls/src/workbook/source.rs#L1470-L1515),
and [`source.rs#L1856-L1887`](../../../../crates/litchi-xls/src/workbook/source.rs#L1856-L1887).

## Timed owners and Callgrind attribution

Use the existing owner scopes and classify every dump by its positive incoming
benchmark edge. Do not identify an operation by dump ordinal. CFB corpus
construction performs an untimed validating `OleFile::open` in
`build_cfb_corpus`; that setup call can produce the first numbered dump. The
setup ancestry is `build_cfb_corpus` or the harness run closure, while the
operation ancestry is `run_cfb_open`.

| Workload | Benchmark driver | Selected timed owner | Timed region |
| --- | --- | --- | --- |
| CFB shapes | `litchi_perf_baseline::run_cfb_open` | `litchi_cfb::file::OleFile<R>::open` | `OleFile::open(Cursor::new(...))`; expected file size, output checks, and object drop are outside the timer. |
| Actual XLS owner | `litchi_perf_baseline::run_xls_owned_source_case` | `litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits` | `SourceBackedWorkbook::from_read_at(...)` plus the selected open/list/one-cell query; corpus layout, `OwnedSource` construction, version probe, post-timer oracles, and drop are outside. |

The CFB timing and allocator brackets are visible at
[`lib.rs#L54620-L54653`](../../../../tools/perf-baseline/src/lib.rs#L54620-L54653).
The source-backed and owned-source XLS brackets are at
[`lib.rs#L26734-L26820`](../../../../tools/perf-baseline/src/lib.rs#L26734-L26820)
and [`lib.rs#L26987-L27059`](../../../../tools/perf-baseline/src/lib.rs#L26987-L27059).
The XLS owner is reached through the positive ancestry

```text
run_xls_owned_source_case
  -> SourceBackedWorkbook::from_read_at
  -> SourceBackedWorkbook::from_read_at_with_limits [owner]
  -> SharedOleFile::open_with_limits
  -> OleFile::open_with_limits
  -> load_directory
```

The CFB positive ancestry is

```text
run_cfb_open
  -> OleFile<R>::open [owner]
  -> (possibly inlined open_with_limits)
  -> load_directory
```

The pilot must retain these edges, with direct and inclusive Ir kept as
separate fields:

```text
load_directory
  -> validated_directory_entries                 [one validation pass]
  -> parse_directory_entry                       [root and reachable entries]
  -> build_storage_tree_iterative
validated_directory_entries
  -> parse_validated_directory_entry             [may be inlined]
  -> directory_name_data                         [validated records]
build_storage_tree_iterative
  -> parse_directory_entry
parse_directory_entry
  -> decode_utf16le
  -> format_clsid
```

Also retain the drop paths for validated-entry `String` values, validated
`DirectoryNameData`, the `Vec<Option<DirectoryNameData>>`, and public
`DirectoryEntry` values. A candidate that merely makes one symbol disappear
must be classified as inlining or code motion until emitted instructions or
line attribution shows which work moved or vanished. A missing
`parse_validated_directory_entry` symbol in the XLS owner profile is not zero
work.

Use the retained Callgrind collection contract: `--dump-instr=yes`,
`--dump-line=no`, `--compress-pos=no`, `--collect-jumps=yes`, `--vgdb=no`,
`--collect-atstart=no`, `--toggle-collect=<owner>`,
`--zero-before=<owner>`, and `--dump-after=<owner>`. The pilot uses one repeat,
warmup zero, and five samples per selected job, serially on CPU 2. If a
candidate passes attribution triage, repeat the same owner profiles twice with
five samples and preserve the setup dump, five positive timed dumps, and the
zero-Ir termination dump. The profile report must bind the binary, owner,
source revision, exact corpus manifest, command, and every raw dump hash.

The positive signal is specific: on many-small and actual XLS, the candidate
must show the public `parse_directory_entry` → `decode_utf16le` work reduced or
eliminated while the one required validator-side `directory_name_data` work
remains attributable. Any `format_clsid` or public scalar-field reduction
must be shown separately. On few-large and tiny, absence of a large signal is
expected as a possible scale result; it is a control against added work, not a
reason to infer a zero from an inlined symbol.

Do not add inclusive parent and child Ir. Inclusive values answer ownership;
self Ir values answer disjoint local work. Collection-off call labels do not
provide operation-local dynamic counts.

## Allocation and latency controls

If the pilot establishes a real emitted-work opportunity, use the existing
matched resource controls before considering adoption.

The allocator lane uses `litchi-perf-baseline-alloc`, whose wrapper installs
`CountingSystemAllocator(std::alloc::System)`. Use two repeats, three warmups,
and 30 measured samples, serially. The existing
`OperationGlobalSystemAllocator` region starts immediately before the timed
operation and finishes immediately after it. It reports checked deltas for
allocation, deallocation, reallocation, and failed-allocation calls;
allocated and deallocated bytes; live bytes before and after; absolute peak
live bytes before and after; and `region_peak_live_bytes`. The counter revision
is `serialized_region_peak_v3`.

`Unavailable` and `Overflow` are states, not zero measurements. Region peaks
include callbacks from other process threads, exclude allocator-internal
realloc overlap and physical RSS, and can perturb allocator scheduling. The
allocator result therefore supplies operation-resource evidence only. The
candidate must not retain a duplicate private cache merely to lower a profile
edge; compare call/byte deltas and incremental live/region peak for every
pilot and follow-up row.

The native lane uses `litchi-perf-baseline`, CPU 2, serial children, two
repeats, 20 warmups, and 1,000 samples in this ABBA order:

```text
baseline r1: XLS primary rows, then CFB tiny/many-small/few-large
candidate r1: XLS primary rows, then CFB tiny/many-small/few-large
candidate r2: CFB tiny/many-small/few-large, then XLS primary rows
baseline r2: CFB tiny/many-small/few-large, then XLS primary rows
```

The primary XLS rows are exactly `xls_source_backed_open`,
`xls_source_backed_open_one_cell`, `xls_owned_source_open`, and
`xls_owned_source_open_one_cell`. CFB `cfb_open` rows use the three fixed
shapes. Preserve the harness timer boundaries: CFB measures open only, and
XLS measures the owner constructor plus the selected query. The native p50,
p95, and p99 policy thresholds are 5%, 10%, and 15%; the repository's
operation-resource and work review threshold is 5%. The prior 0548/0549
admission rule requiring at least a 3% p50 improvement in all four XLS rows in
both repeats applies only if this candidate is later proposed for retention;
this plan makes no such result claim.

Callgrind Ir is not a latency proxy. Whole-child RSS or hardware counters, if
the normal measurement lane records them, cannot be used as operation-local
proof without a separate scope and identity review.

## Preservation and error controls

The public valid controls compare exact behavior after the timed region:

* `cfb_list_streams` compares the complete stream-name set, count, ordering,
  and digest; `cfb_read_one` compares the selected target bytes and digest.
* The XLS owned-source controls compare worksheet count or names, the selected
  worksheet 1 / row 20 / column 4 value `42.0`, the `Workbook` digest, and the
  fixed archive identity. Source-backed controls additionally compare source
  version stability, logical read/locality evidence, and the selected source
  ranges.
* A focused CFB fixture compares every public `DirectoryEntry` field,
  SID-indexed occupancy, child vectors, and `dir_name_data` alignment. The
  existing checks
  [`parsed_name_cache_is_sid_aligned_with_directory_entries`](../../../../crates/litchi-cfb/src/file.rs#L3057-L3070),
  [`lookup_refuses_a_missing_validated_name_cache_entry`](../../../../crates/litchi-cfb/src/file.rs#L3072-L3082),
  and the wide sibling lookup
  ([`file.rs#L3046-L3054`](../../../../crates/litchi-cfb/src/file.rs#L3046-L3054))
  define the relevant public/cache contract.

Before adoption, run a differential malformed-directory fixture set through
the baseline and candidate and compare the exact `OleError` variant and
`Display`/`Debug` text, failure precedence, resource label, and partial-state
postcondition. The set must cover:

* empty and partial 128-byte records; invalid entry types and node colors;
  invalid name lengths, odd lengths, missing terminators, internal NULs,
  invalid UTF-16, forbidden names, and name-encoding mismatches;
* invalid root SID/type/name/links, sibling ordering, repeated SID and cycle
  detection, cross-storage ownership, invalid child/sibling SIDs, and stream
  versus storage field rules;
* version-3 stream-size masking, invalid allocation markers and indexes,
  directory byte limits, short final-sector zero fill, source I/O or source
  instability, and fallible reservation failures; and
* the final directory/public-entry agreement check and the no-partial-
  publication rule. Reservations must retain their current order and labels,
  including `directory entries`, `directory name comparison data`, directory
  traversal state, and any new private handoff storage.

The existing `perf_chain_guard` cases (`valid`, `shortselfcycle`,
`prefixcycle`, `latercycle`, `earlyend`, `invalidmarker`, `invalidindex`, and
`lateexcess`) at sizes 128 and 16,384 remain a chain-validation regression
control with its exact error oracles and 4x same-invalid / 2x baseline-valid
envelopes. That guard measures the FAT/MiniFAT allocation-chain mechanism; it
does not cover directory names, directory graph ordering, or the two-view root
encoding and cannot replace the fixture set above.

## Classic-Mac root two-view guard

The current harness has no corpus selector for the classic-Mac root encoding,
so no current synthetic profile covers it. A dedicated future fixture is
required before retaining a candidate that moves directory names.

For SID 0, entry type root, `name_len == 2`, bytes `00 52`, and zero in the
remaining name array, validation explicitly accepts the classic-Mac form and
canonicalizes its private validated name to `Root Entry` in
[`parse_validated_directory_entry`](../../../../crates/litchi-cfb/src/file.rs#L1362-L1379).
The historical public parser instead decodes
`raw.name[..name_len - 2]` in
[`parse_directory_entry`](../../../../crates/litchi-cfb/src/file.rs#L1466-L1474),
which yields its existing public value (empty for this compact form). The
final handoff check intentionally exempts SID 0 while still storing the
canonical validated `DirectoryNameData`
([`file.rs#L1020-L1037`](../../../../crates/litchi-cfb/src/file.rs#L1020-L1037)).

The candidate must carry both views: canonical `Root Entry` for validation,
ordering, and the SID-aligned comparison cache; the historical public value
for `DirectoryEntry.name`. It must not make public-name equality the
canonicalization rule, and it must preserve the same acceptance, rejection,
error text, and publication behavior for this fixture and for ordinary roots.

## Campaign sequence and decision

The future campaign should proceed in this order:

1. Freeze the baseline and candidate source manifests, compiler/build
   identity, exact corpus catalog, and owner/case matrix. Keep this campaign
   independent of the change-0553 XLSX source and verdict.
2. Capture the four-job Callgrind pilot with the exact owner scopes and
   positive-edge classification above. Inspect emitted line/instruction work
   and drop paths before spending the full latency matrix.
3. If the pilot demonstrates the specific second-parse/name/field work
   reduction, run the matched allocator and native controls, then the public
   valid and malformed/error preservation controls. Repeat profiles at the
   retained two-repeat setting for the candidate comparison.
4. Select the candidate for a later retention review only if the emitted-work
   attribution is positive across many-small and actual XLS, few-large/tiny
   have no unexplained added work, operation allocations/live peak are bounded,
   the native and policy gates pass, and all public/error/root controls match.

The campaign must stop with no adoption recommendation if any of these
attributions is missing:

* a positive owner edge for each timed dump, especially when
  `open_with_limits`, `parse_validated_directory_entry`, or scalar extraction
  is inlined;
* an emitted line/instruction account showing which second decode, name
  construction, CLSID formatting, or scalar extraction disappeared;
* measured operation-local allocation calls/bytes and live/region peak with a
  measured status;
* the classic-Mac root fixture's two-view result and exact malformed error
  comparison; or
* graph, public-field, name-cache, XLS-cell, source-locality, and no-partial-
  publication differentials.

If the pilot shows only symbol renaming/inlining, a lower aggregate Ir count,
or an unbounded retained seed, it does not select a valid mechanism. No
speedup, regression, or adoption claim is made by this plan.

## Evidence custody

The retained measurement and source artifacts that this plan reads are bound
by these hashes:

| Artifact | SHA-256 |
| --- | --- |
| 0547 `profile-analysis.json` | `6cd52ecef7af77671270f95983a6dff6c3cc7d414cf94667a2f23f020b1756bd` |
| 0547 `analysis.json` | `5fea2d98ed11d1aa166e45327c99218d8f9f5a2ad3db81197ed323f1f0dedb18` |
| 0547 baseline source manifest | `db197f5ea518082646669db5ba229cc2ea4081f5d2b6a51357e07f9cc4090511` |
| 0547 `profile-r1-cfb-tiny.json` | `98c7d662b065920a8245df3bd6327a3c096d5a223a7c5fe95b7aad0f90222b90` |
| 0547 `profile-r1-cfb-many-small.json` | `7c4b820d553ee6f542f5953d418217a8eed35949885f9927e2311a1c2c773a63` |
| 0547 `profile-r1-cfb-few-large.json` | `a6461ecfee3a1171085ca4eff9a9afe27a9bb9846d5758ab556ae20e0e2849c2` |
| 0547 `profile-r1-xls-owned.json` | `5ee0e6c43925a608f490139fd47765ca4d1e5feb90a91b549e20df9376498913` |
| 0547 positive raw `profile-r1-xls-owned.callgrind.1` | `f84dbc7174423202223ad592158e6de7d96737dc45bb4cd55af2c8f1f149c308` |
| 0547 `suboperations.json` | `ebe7466d7decbdf94f5e15e64a6e7c5659932b0abf860dc457871c5a7178ecff` |
| 0548 `decision.json` | `b57ab8c9962935b8a5e48bf21227c5f3b7f558a3a9e7137e20b7b95fac637e33` |
| 0549 `decision.json` | `f08dacc2dc4791e9c7c916050b58369b111207ec2a6d19fc6ba38dab66c719fe` |
| current `crates/litchi-cfb/src/file.rs` at revision `bb5bfaa7be4fdffeae3fddaee3ed1266c3417520` | `72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b` |
| current `crates/litchi-cfb/src/shared.rs` | `2aef59e5e7d20a8984b0c2f07d561865a87a74d927cb088d4a08ff85532c0624` |
| current `crates/litchi-cfb/src/shared_bulk.rs` | `3ea1cccdbdd4b45f3801983bfa8df67b94defe5476834ebc6175011760bb5b48` |
| current `crates/litchi-xls/src/workbook/source.rs` at revision `98ea87658ce1ca2697a1dfc6901d72d8888cf5d4` | `3e7dbde9f197771ee7766a597f3f27f6a0143b79f1cf2575dd52a9f8493f6d92` |
| current `tools/perf-baseline/src/lib.rs` | `b71922e53c507842e4d8fb4983f52872a690940e3f7e9bf1b0f2d03fdbcd525f` |
| current allocator wrapper `tools/perf-baseline/src/bin/litchi-perf-baseline-alloc.rs` | `d895986f85e12bec5a30d728bfda2985efd2b0b897e97970e9fa45b2568ae3db` |

The relevant accepted constraints are recorded in
[ADR 0005](../../../adr/0005-io-memory-and-performance.md),
[ADR 0006](../../../adr/0006-validation-security-and-compatibility.md), and
[ADR 0026](../../../adr/0026-ole-directory-metadata-binding.md): bounded
fallible memory, validation and compatibility preservation, and CFB/shared
OLE directory metadata ownership remain at their current boundaries.

This document records a read-only audit of retained documents and current
source. It records no new measurement and makes no performance claim.
