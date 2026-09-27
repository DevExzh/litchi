# 0777 — bounded XML attribute checks across the format readers

Status: source integration, quality, paired native capture, differential
replay, raw instruction/allocation observations, and offline validation are
complete for candidate `1e334ffba9efb89e5badb8aa7a3e33052e96635c`, based on
`87e926fcc6`. The independent review and owned-target cleanup are complete;
the packet is sealed with a complete SHA-256 file inventory. The observations
below remain scoped to their recorded probes and
whole-process measurements; they are not a claim of program-wide performance
completion.

The change carries the bounded quick-xml attribute iterator from the 0770
work into the current XML readers. It moves fail-fast readers away from
quick-xml's unkeyed post-32-name duplicate check, preserves the first error
observable by those readers, and keeps lenient readers on their explicit
first-wins or counting policies. It also wires the common MCE codec and stream
paths without applying the same iterator twice, and charges the new DOCX MCE
scratch in the existing requested-storage envelope.

The evidence packet is [results/change-0777/README.md](results/change-0777/README.md).

## Why this integration exists

quick-xml 0.41 checks attribute names linearly through its 32-name boundary.
After that boundary it uses an unkeyed hash pre-filter and scans earlier names
on a pre-filter hit. A fail-fast reader stops after the first error and does not
pay for later iteration, but it can still pay the duplicate-check cost needed
to reach a late first error. A reader that deliberately continues after errors
needs a different semantic policy. The integration gives each reader the
policy it already requires instead of changing malformed-input behavior
implicitly.

The bounded fail-fast adapter keeps quick-xml's checked behavior for the first
32 successful names. Before quick-xml's later hash path is reached, it turns
duplicate checking off and records borrowed names and source positions in an
ordered index. It yields the same items and the same first error through the
first error, including duplicate-before-value precedence, then stops. Later
name comparisons are ordered-map comparisons; names and values are not copied
by the index. The adapter does not promise quick-xml's post-error recovery API.

The existing lenient paths remain separate. `first_wins` parses every value,
reports duplicate occurrences as the caller expects, and uses its bounded
linear-to-ordered transition. `count_up_to` counts without duplicate checking
when its caller only needs a bounded reservation/count. The MCE stream uses an
unchecked iterator because it validates lexical names, limits, and expanded
attributes in its own bounded path. This avoids running two duplicate checks
over the same event.

## Architecture and integration

The canonical `BytesStartExt` implementation lives in the OPC XML-attribute
owner. `litchi-ooxml-common` re-exports the adapter through its XML module.
Four crates that cannot depend on the OPC owner under the accepted crate
topology retain synchronized copies: `litchi-ole-common`, `litchi-sign`,
`litchi-xldm`, and `xml-minifier`. The copies compile the shared adapter tests;
the workspace lint and copy checks prevent an accidental return to raw checked
iteration. No archive implementation type, executor, global cache, ambient
I/O, or parallel runtime is introduced.

The migration survey found 581 attribute sites across 15 production crates.
`litchi-imgconv` is the sixteenth lint-enforced owner and has no XML attribute
sites. The current classification is:

| Path class | Current treatment |
| --- | --- |
| 546 fail-fast sites | 545 use `checked_attributes`; one dynamic site retains its explicit checked selection and exits before the hash-filter path can be reached |
| 26 unchecked sites | 25 use the unchecked adapter, plus the MCE stream's own unchecked-and-validate path |
| 4 lenient sites | explicit first-wins behavior |
| 2 bounded count sites | explicit count-up-to behavior, including the reservation count |
| 3 helpers | adapter or shared helper ownership rather than reader iteration |

All 545 checked migrations have matching per-crate counts in the current
source. Four additional checked calls are adapter tests. An independent review
inspected all 26 exceptional checked callers: each stops at the first iterator
error through `?`, an explicit return, `all`, or a single `next`; none filters
or flattens errors and continues accidentally. The four first-wins migrations
retain their first-occurrence policy and the two count sites retain their
bounded count semantics.

The integration sequence is deliberately narrow:

1. The fail-fast adapter and synchronized copies establish the common iterator
   contract and its tests.
2. The MCE stream selects the unchecked adapter because its own code performs
   name, byte, event, and expanded-name checks. The tree MCE codec selects the
   checked adapter and retains the 0776 expanded-attribute duplicate check
   after namespace declarations are installed and before opaque, skipped, or
   selected-branch returns.
3. The DOCX tail-append MCE profile charges the ordered attribute index,
   iterator state, and owner conservatively when the tag can cross the
   32-name boundary. This is requested-storage accounting under the pinned
   standard-library layout assumptions. It is not a claim that standard
   `BTreeMap` or `Box` allocations are globally fallible, nor a portable RSS
   formula.

The expanded-name check in the tree MCE path uses borrowed local names. It
compares local names first, resolves namespace identities only for colliding
local-name groups, keeps up to eight candidates in fixed storage, and uses a
fallibly reserved sorted spill vector above that size. The scratch is released
before children are processed. This check is specifically for duplicate
expanded attributes; it does not turn MCE preprocessing into a general XML
validator and does not claim full QName parity on every opaque or skipped
malformed path.

The stale broad stream lint exemption was removed. Existing 0775 namespace
sharing and 0776 tree duplicate detection remain in place. The work does not
change the ordinary CRUD API, snapshot/edit/commit model, preservation policy,
or the dependency direction enforced by the accepted ADRs.

## Semantic and breaking-change scope

Readers that stop at their first attribute error now use a bounded duplicate
check while retaining the error and item sequence through that point. A
caller that intentionally consumes quick-xml's recovery after an error is
outside this adapter's contract; the migration audit found no production
fail-fast caller that does so. Lenient and counting readers retain their
separate semantics.

MCE inputs continue to be bounded and typed-refusal paths. The stream's own
duplicate and lexical checks remain responsible for stream events. The tree
MCE path still refuses duplicate expanded attributes before branch selection,
attribute removal, or opaque-child return. Requested-storage accounting can
make a DOCX operation near its configured ceiling refuse earlier when the
ordered scratch is charged. This is a deliberate safety boundary.

No claim is made here about all XML malformed-input recovery, all possible
quick-xml hash collisions, all public XML validation, or all format CRUD
workflows. The separate upstream note in the packet is an unsent technical
draft, not an upstream submission or compatibility promise.

## Quality and source custody

The current source-bound quality run has nine passing gates over 16 owners and
records 10,519 passing tests, 80 ignored tests, and 407 suites. The gates are
format checks, all-target/all-feature checks, owner tests, library Clippy with
warnings denied, rustdoc with warnings denied, the OOXML-enabled facade check,
the standalone performance-harness check, and crate-boundary validation. The
facade check retains one pre-existing unused-helper warning for
`missing_ooxml_catalog_part_error`; it exits successfully and is not presented
as a warning-free facade build. Owner Clippy and documentation warning-denied
gates pass.

The quality receipt and all source hashes are tied to the candidate source.
`final-source.json` records 305 source/config hashes, with candidate
`1e334ffba9efb89e5badb8aa7a3e33052e96635c` and base `87e926fcc6`. The packet
also retains the architecture-input digest set, the workspace lock, the
standalone probe lock, the synchronized probe source, the migration review,
and the independent source review. No native result is inferred from these
quality gates.

The contract-only CRUD coverage validation also passes: 15 non-iWork
categories and 33 mapped selectors are indexed and bound to their reviewed
inputs. It supplies no scenario promotion or full-run timing report and does
not convert this packet into a CRUD-completion claim.

## Paired capture and native results

The coordinator's capture runner uses separate release target directories,
the same standalone probe source and lock for both legs, release LTO with
panic abort, CPU 12, and serial order:

```text
before, after, after, before, before, after
```

There are 19 native cases and six processes per case, for 114 timed processes:

| Family | Cases | Timed operation |
| --- | --- | --- |
| MCE | benign worksheet; benign document; stream-count worksheet; stream-count document; prefixed controls with 1, 2, 8, 9, and 32 attributes | tree or stream MCE processing with the configured observer/digest |
| OPC | `n = 0, 8, 29, 30, 32, 33, 256, 1024, 4096, 16384` extra namespace declarations | `OpcPackage::from_bytes` followed by a deterministic name-sorted part digest |

The OPC probe's relationship start tag has the extra namespace declarations
and three ordinary attributes. The `n=29/30` pair straddles the reviewed
32-name transition once the fixed declarations are counted. The larger values
exercise bounded source limits and the ordered-map path.

Each process uses nine measured samples after two warmups. MCE timing begins
after the input has been constructed and covers the selected parser plus its
configured observer and digest. OPC timing begins after the synthetic package
has been built and covers the OPC read plus deterministic digest. Input
construction, process startup, and JSON report writing are outside both timing
intervals. `/usr/bin/time` records whole-process peak RSS, including startup,
input, parser, observer, and report-related process state; it is not an
allocator-only or portable memory measurement.

The same capture runs one candidate iterator-equivalence probe and one
before/after differential per leg. The differential mutates bounded package
members with repeated names, late repeated names, malformed values, and
continuation controls, then compares reader outcomes. The equivalence probe
compares items and exact errors through the first error on parser-reached tags.
Neither is an exhaustive proof of every malformed-reader recovery path.

The completed native capture reports each case's median of the nine measured
sample p50 values. The table retains absolute before-to-after values and marks
every positive change above five percent in either latency or RSS. `flag` means
that the corresponding metric crossed that threshold.

| Case | p50 before → after (µs) | Δ p50 | RSS before → after (KiB) | Δ RSS | Outcome | Flags |
| --- | ---: | ---: | ---: | ---: | --- | --- |
| `mce_benign_worksheet` | 13001.607 → 13048.457 | +0.36% | 7168 → 7056 | −1.56% | accepted fixture | — |
| `mce_benign_document` | 1768.387 → 1791.998 | +1.34% | 3600 → 3528 | −2.00% | accepted fixture | — |
| `mce_stream_count_worksheet` | 40218.956 → 40735.599 | +1.28% | 4564 → 4524 | −0.88% | accepted stream count | — |
| `mce_stream_count_document` | 5167.763 → 5251.613 | +1.62% | 3344 → 3368 | +0.72% | accepted stream count | — |
| `mce_prefixed_1` | 795.234 → 816.294 | +2.65% | 3012 → 3104 | +3.05% | accepted synthetic control | — |
| `mce_prefixed_2` | 1243.686 → 1277.395 | +2.71% | 3268 → 3336 | +2.08% | accepted synthetic control | — |
| `mce_prefixed_8` | 4343.419 → 4544.980 | +4.64% | 3812 → 3864 | +1.36% | accepted synthetic control | — |
| `mce_prefixed_9` | 5154.542 → 5377.454 | +4.32% | 3864 → 3772 | −2.38% | accepted synthetic control | — |
| `mce_prefixed_32` | 18777.112 → 19371.936 | +3.17% | 6332 → 6496 | +2.59% | accepted synthetic control | — |
| `opc_relationship_declarations n=0` | 6.510 → 6.690 | +2.76% | 3232 → 3068 | −5.07% | accepted package | — |
| `opc_relationship_declarations n=8` | 7.070 → 7.440 | **+5.23%** | 3244 → 3092 | −4.69% | accepted package | latency |
| `opc_relationship_declarations n=29` | 8.730 → 8.860 | +1.49% | 3028 → 3132 | +3.43% | accepted package | — |
| `opc_relationship_declarations n=30` | 9.190 → 9.750 | **+6.09%** | 3044 → 3108 | +2.10% | accepted package | latency |
| `opc_relationship_declarations n=32` | 9.480 → 9.950 | +4.96% | 3024 → 3292 | **+8.86%** | accepted package | RSS |
| `opc_relationship_declarations n=33` | 9.520 → 10.120 | **+6.30%** | 3188 → 3260 | +2.26% | accepted package | latency |
| `opc_relationship_declarations n=256` | 22.811 → 35.151 | **+54.10%** | 3264 → 3204 | −1.84% | accepted package | latency |
| `opc_relationship_declarations n=1024` | 15.240 → 15.600 | +2.36% | 3092 → 3276 | **+5.95%** | exact existing namespace-limit refusal in both legs | RSS |
| `opc_relationship_declarations n=4096` | 43.030 → 43.480 | +1.05% | 3540 → 3520 | −0.56% | exact existing namespace-limit refusal in both legs | — |
| `opc_relationship_declarations n=16384` | 153.091 → 154.491 | +0.91% | 3780 → 3716 | −1.69% | exact existing namespace-limit refusal in both legs | — |

The six flagged observations are the OPC `n=8`, `n=30`, `n=32`, `n=33`,
`n=256`, and `n=1024` rows. The `n=256` input is accepted and rises from
22.811 µs to 35.151 µs (+54.10%). The `n=1024`, `n=4096`, and `n=16384`
inputs all return the same existing quick-xml namespace-binding-limit refusal
in both legs; their timings measure refusal handling, not successful parsing.
The `n=1024` RSS flag is retained even though it is a refusal control.

The three-process-per-leg design and nine samples per process do not establish
stable p95/p99 tails. A process RSS value includes startup, input generation,
parsing, observers, and report overhead. The table therefore records observed
whole-process behavior for these named cases and does not claim a portable
allocator or live-byte change.

## Differential and equivalence results

The completed before/after differential covered 2,088 packages: 1,567 DOCX,
275 XLSX, 126 PPTX, and 120 other OPC packages. It mutated 2,037 packages in
113,830 bounded mutations, producing 228,716 reader outcomes in each source
leg comparison. The before and after outcome maps are identical: 228,716
comparisons are unchanged and there are zero changed outcomes or mismatches.

The candidate iterator-equivalence run visited 13,764 package members,
7,598,414 real start tags, and 2,332,430 mutated start tags. Of those, 105
real tags and 1,015,360 mutated tags were over the 32-name boundary. The
comparison covered the canonical OPC adapter and its synchronized OLE copy
against quick-xml; the five synchronized production copies are also covered
by their shared quality tests, but this probe is not a claim that all five
runtime copies were independently compared over the corpus. It reported zero
mismatches. The probe checks item and error identity through the first parser
error on reader-reached tags; it does not prove every reader's malformed-input
recovery after an error.

These results support the migration's fail-fast contract for the exercised
corpus and mutations. They do not establish standards validity for generated
malformed inputs, nor do they demonstrate a practical collision family for
quick-xml's unkeyed pre-filter. The deterministic ordered fallback gives a
comparison bound; it is not a claim that every possible hash collision has
been constructed or measured.

## Instruction and allocation observations

The raw instruction lane completed 56 primary perf rows plus eight primary
iterator rows in `instructions-0/`, and eight secondary `n=256` perf rows in
`instructions-256/`: 72 instrumented process measurements total, plus two
supported qualification runs. It uses whole-process `perf stat` user
instruction counters and a marginal 3-versus-23 run design; its instrumented
elapsed time is excluded from the native table. The qualification receipts
use the `running_percent` field, with a 100.0% runtime/enable fraction. The
per-operation instruction counts are:

| Case | Repeat 0 before → after | Repeat 1 before → after |
| --- | ---: | ---: |
| MCE benign document | 35,362,239.65 → 35,633,171.05 (+0.766%) | 35,387,501.55 → 35,646,627.30 (+0.732%) |
| MCE benign worksheet | 272,240,669.95 → 274,909,748.75 (+0.980%) | 272,901,941.35 → 274,909,887.15 (+0.736%) |
| MCE prefixed 32 | 411,432,831.75 → 415,268,946.25 (+0.932%) | 411,432,844.25 → 415,268,987.25 (+0.932%) |
| OPC `n=8` | 110,584.35 → 112,261.35 (+1.516%) | 110,505.70 → 113,467.45 (+2.680%) |
| OPC `n=30` | 152,755.50 → 172,933.65 (+13.209%) | 152,823.25 → 172,194.10 (+12.675%) |
| OPC `n=4096` refusal | 847,255.00 → 847,516.95 (+0.031%) | 847,185.50 → 847,383.45 (+0.023%) |
| OPC `n=16384` refusal | 2,947,928.45 → 2,947,823.95 (−0.004%) | 2,948,155.80 → 2,948,147.30 (−0.0003%) |
| OPC `n=256` follow-up | 466,442.55 → 619,088.75 (+32.726%) | 466,918.15 → 620,586.70 (+32.911%) |

The n=30 and n=256 instruction observations cross five percent. They are
marginal whole-process estimates, including process/report work, rather than
parser-only attribution. The separate iterator sublane repeated 440,371 tags
per 3/23-operation comparison. Its checked path used 259,474,057.50 versus
quick-xml's 244,257,564.05 instructions in repeat 0 (+6.23%), and
259,570,557.45 versus 244,652,690.65 in repeat 1 (+6.10%). This comparator
covers the reviewed adapter probe scope and is not a runtime claim for every
synchronized copy.

The raw allocation lane completed ten observations in `allocations-0/` and two
additional `n=256` observations in `allocations-256/`. Each is one whole-
process instrumented sample per case and leg, including startup, input
generation, parsing, observation, and report serialization.

| Case | Allocated bytes before → after | Allocation calls before → after |
| --- | ---: | ---: |
| MCE benign worksheet | 12,455,831 → 12,455,829 (−2) | 129,512 → 129,512 (0) |
| MCE benign document | 2,539,502 → 2,539,500 (−2) | 16,056 → 16,056 (0) |
| OPC `n=29` | 100,165 → 100,163 (−2) | 233 → 233 (0) |
| OPC `n=30` | 104,420 → 104,034 (−386) | 237 → 242 (+5) |
| OPC `n=4096` refusal | 519,451 → 519,449 (−2) | 136 → 136 (0) |
| OPC `n=256` follow-up | 184,997 → 173,331 (−11,666) | 478 → 512 (+34) |

The ordinary worksheet and document −2-byte differences are whole-process
effects and do not establish a parser allocation reduction. The extra `n=256`
allocation row's −11,666-byte change, accompanied by 34 additional allocation
calls, is also a whole-process result and does not establish a parser
allocation reduction. The extra `n=256` instruction and allocation
observations are retained as follow-ups to the native +54.10% row; they did
not rerun or select the native result. The consolidated `observations.json`
receipt binds these raw rows, source/binary hashes, and unchanged outcomes.
Any failed attempt must remain in the packet with its raw logs, command
receipts, and source identity. Historical 0770 timing or equivalence
observations remain stale and are not adopted as 0777 evidence.

Offline validation, independent review, and owned-target cleanup pass. The
cleanup receipt records removal of the three owned build targets after source,
fixture, package-inventory, binary, and unrelated-file checks. The independent
packet-path replay also passes without native binaries or a source checkout;
its receipt retains the same 19 cases, 114 native processes, 72 instruction
measurements, 12 allocation measurements, and zero differential/equivalence
changes. The packet seal records the complete file-hash inventory, and final offline
validation verifies every sealed byte.

## Limits and follow-up

This batch addresses the XML attribute integration and its measured probe
design. It does not complete the program-level goal. In particular, it does
not establish CRUD coverage, physical cold-cache behavior, caller-supplied
range-source behavior, sequential non-seek output, concurrent scaling,
parallel execution efficiency, or native Office interoperability. Those audits
remain active after this batch, and iWork work is outside this report's scope.

## Integration and cleanup follow-through

Main was fast-forwarded to `8ac869e5f8` after review, complete packet validation,
and byte-for-byte verification of all 732 staged evidence files, including the
seal itself. All 305 frozen candidate source/config hashes match main. The
quality source is `4884539345`; the capture candidate `1e334ffba9` adds only
its capture harness documentation and probe archive.

The three owned build targets were removed after source, fixture, package
inventory, and executable checks. The owned 0777 integration worktree, its
copied root lock, three reference symlinks, and merged temporary branch were
then removed. Offline validation passes from main with that worktree and the
executables absent. The sealed packet was not changed during this follow-through.

The original 0770 worktree and its six reviewed uncommitted edits remain
unchanged; all other pre-existing worktrees are retained. The unrelated
`FORMAT_IMPLEMENTATION_REVIEW.md`, `UNIFIED_OPS_API_DESIGN.md`, and
`matrix-analysis.json` keep their original hashes. The main checkout's own
lock and references remain in place. The broader non-iWork goal remains active.
