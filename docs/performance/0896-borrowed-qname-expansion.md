# 0896 — borrowed QName expansion evaluation

This batch evaluates eliminating temporary owned `QualifiedName` construction
inside MCE stream expansion. The candidate reuses `xml_name::is_qualified_name`
and borrows prefix/local slices; namespace resolution, owned expanded local
names, typed errors, callback order, and resource checks retain their existing
behavior. The rejected 0894 metadata route is excluded.

The native paired evaluation fails the frozen adoption gate: the primary p50
ratio is **0.999894**, with a 95% paired bootstrap interval of
**0.989597–1.007445**. The candidate is rejected for production. Direct-stream
allocation savings remain diagnostic evidence.

## Source and correctness

The build source is the exact retained 506aa2 snapshot, unchanged in selected
build paths at 3ff493b638. Two existing files are replaced: `mce/stream.rs` and
`mce/tests.rs`. The isolated source census has 10,771 selected files plus the
owned marker, generated workspace lock, and four boundary-support files.
The standalone benchmark uses its tracked lock; the shared workspace's ignored
lock is not consumed. Source overlays, patch, source census, compiler commands,
and quality logs remain in [the packet](results/change-0896/).

All eight final quality gates pass: common all-features tests (445 total),
common all-targets Clippy with warnings denied, common documentation with
warnings denied, XLSX library tests (1,390), diagnostic harness tests (24),
producer tests (22), scoped formatting, and crate boundaries. The new stream
tests cover valid plain/prefixed/Unicode names, element versus attribute default
namespaces, malformed QName structure, Unicode-invalid attributes, unbound
prefixes, and name-limit/refusal precedence.

The formal baseline and candidate MCE differential reports are byte-identical.
Each covers 20,000 generated cases, 20,000 aliasing cases, three packages and
67 selected real members. The complete populations contain 200,201 oracle
records and 200,405 legacy result records, with ten controls, zero incomplete
oracles, and zero control failures. Generated/aliasing populations include
13,811/19,418 distinct inputs; counts are not claims of 40,000 unique files.
Deterministic compressed reports retain exact decompressed hashes.

Two preliminary qualification failures remain visible: the first runner used
the same path for its report and receipt; the corrected formal baseline uses
a fresh labeled run. All common tests then passed, but Clippy caught an
unnecessary `format!` in a new Unicode test. The literal-only correction was
refrozen, and all final quality gates passed. Neither failure occurred during
a performance capture; the original helpers, logs, and source relocation map
are retained.

## Direct stream allocation diagnostic

A separate frozen diagnostic uses the baseline/candidate `xml_attribute_bounds`
debug binaries (opt-level 0, debug info 0, locked/offline) and Heaptrack 1.5.0.
For each real worksheet/document part, four fresh processes run in ABBA order,
CPU 12, no warmups, and 20 scans per process: eight children / 160 scans.
All successful input hashes, event/attribute counts, and stream reports match
within and across roles. Both repeats yield identical allocation counts.

| Whole-process allocation calls, 20 scans | Baseline | Candidate | Reduction |
| --- | ---: | ---: | ---: |
| Worksheet part | 20,152,391 | 15,361,491 | 23.77% |
| Document part | 2,650,247 | 1,619,687 | 38.89% |

These totals include setup, archive/member loading, result serialization, and
teardown. They are not per-scan timed counters, release latency results, peak
heap measurements, or a production adoption gate. Profiler overhead and debug
code generation preclude inferring native latency from captured durations.
Raw compressed profiles and command/report receipts are retained.

The current producer full-cell case refuses the selected stream at
`sheetViews` and then uses its authoritative fallback. It exercises only the
stream prefix, so a large direct-stream allocation reduction does not imply a
large end-to-end XLSX improvement. The native primary gate remains the same
2% p50 improvement with a bootstrap upper ratio below 1.0; direct diagnostics
cannot override it.

## Native paired protocol

The trial uses the exact retained 0894 baseline native/allocator binaries
and the candidate's matching Rust 1.95 release configuration: opt-level 3,
debug info 1, fat LTO, one codegen unit, two build jobs. The fixed schedule is
64 native and eight allocator children, 3,320 measured samples, CPU 12, warm
filesystem, and alternating AB/BA pairs. The primary producer full-cell case
retains forced post-endpoint per-worksheet tolerance guards; controls cover
marker-free XLSX, DOCX, and PPTX. The shared host is not exclusive.

The frozen bootstrap uses seed 896, 10,000 draws of whole paired child ratios,
and interval indices 249/9749. Tail, RSS, and spread flags remain review
evidence; executable residency prevents a causal heap interpretation of RSS.

All 72 children / 3,320 measured samples pass the strict semantic guards and
offline verifier, without retries or replacement samples. Paired p50 results:

| Scenario | Candidate/baseline ratio | Bootstrap 95% interval |
| --- | ---: | ---: |
| Producer XLSX full-cell (primary) | 0.999894 | 0.989597–1.007445 |
| Marker-free XLSX | 1.000150 | 0.986432–1.009721 |
| DOCX selected paragraph | 1.006040 | 0.989288–1.033022 |
| PPTX selected slide | 0.997590 | 0.989869–1.003223 |

Primary p95/p99 ratios are 0.997301/0.999042. The primary timed allocator
region repeats exactly across four pairs: calls fall from 25,933 to 25,865
(68 calls, 0.2622%); allocated bytes from 37,618,900 to 37,618,388 (512 bytes);
region peak live bytes from 9,252,430 to 9,252,427 (3 bytes). These small
prefix savings explain why the full direct-stream diagnostic does not predict
the current fallback workflow's gain.

All review flags remain visible: one baseline DOCX p50 spread flag, two
primary metric flags, and seven control metric flags. Candidate RSS is about
144 MiB (147,288–147,472 KiB across control children), far above
the baseline readings; process/executable residency confounds a causal heap
interpretation. No RSS improvement or regression is attributed to QName
allocation changes. Tail and spread observations do not rescue the failed
primary gate or justify discarding samples.

## Decision and cleanup

The candidate is rejected because both primary criteria fail: median ratio
must be at most 0.98 and the confidence upper bound must be below 1.0. The
bounded QName hypothesis removes direct-stream allocation work, but the
current primary route reaches only a short stream prefix before fallback.
Future work should profile the restored end-to-end route's dominant work;
revisiting this candidate requires a separately frozen scenario and gate.
The rejected 0894 metadata route remains excluded.

Independent native replay reproduces every ratio and bootstrap endpoint from
all raw samples, confirms every semantic receipt, and recommends rejection.
Both owned production files are restored byte-for-byte to 3ff493b638. The
candidate patch, two overlay versions, source inventories, quality logs,
release/diagnostic binaries, compressed corpus reports, profiles, and all
native capture receipts remain available.

The marked isolated source/build trees were removed after review, reclaiming
3,302,917,127 bytes, in addition to 2,387,941,938 bytes of prior debug/doc
outputs and 678,588,171 bytes of redundant raw reports. Exact source-to-repo
fixture hash mappings preserve evidence after source cleanup. No worktrees
were created in this batch; unrelated worktrees and the existing untracked
`bench_word_parse.rs` are untouched. Retained local binaries are intentionally
ignored by Git; the source archive, locks, overlays, and build receipts allow
reconstruction. The broader performance goal remains active, and iWork is
excluded from implementation.
