# 0481: borrow DOCX scanner event names

`performance_claim: none; scoped control/candidate scanner comparison`

`claim_authorized: false`

The source-backed plain-paragraph scanner allocated a separate byte vector for
all start, empty and end element names. Each name is consumed before its event
is dropped. The seven-line candidate retains the parser's `LocalName` wrapper
and borrows its bytes. It changes no namespace resolution, scope transition,
attribute check, source fence, event/depth/paragraph limit, range construction,
scan count, patch format or publication path. The optional name-copy allocation
failure disappears with the unnecessary allocation; required validation remains.

The previous sampled profile assigned 68.037% of weighted lifecycle periods to
the source-backed scanner. This experiment targets its temporary allocation
work; it does not skip validation or claim that all scanner work is removed.
Control is change 0480's source; both arms have fresh normal and separately
instrumented allocator builds, with 7,048-file source manifests and eight
unchanged embedded templates. The source difference is exactly one file.

## Protocol and scope

The frozen A1/B1/B2/A2 order crosses three source paragraph counts,
normal/allocator binaries and total/six-phase modes. All 48 reports have 30
measured samples after three warmups, pinned to CPU 2: 1,440 formal samples.
Eight pilots contain 24 excluded operations; two separate large normal PMU
processes contain 60 excluded operations. Commands, source and binary hashes,
corpus oracles, driver hashes and raw receipts are retained in change-0481.

The lifecycle includes source adapter/package construction, snapshot, staging,
commit, sequential publication, digest finalization and all owner destruction.
Corpus and independent oracle storage preexist this operation. GNU `time` RSS
and PMU counters include the whole process, setup, warmups and report teardown.
Phase measurements are separate executions, and their peaks are not summed.
Allocated bytes include full realloc `new_size`, not physical copying.

## Normal total lifecycle observations

Means and nearest-rank percentiles are milliseconds. Parentheses give the mean
95% Student interval using t(29)=2.045. RSS is one whole-process maximum in KiB,
not a 30-sample operation-local RSS statistic.

| Paragraphs | Repeat | Arm | Mean (95% interval) | p50 | p95 | p99 | RSS KiB |
| ---: | ---: | --- | ---: | ---: | ---: | ---: | ---: |
| 64 | 1 | control | 0.168671 (0.166877–0.170466) | 0.166951 | 0.180891 | 0.183601 | 4,808 |
| 64 | 1 | candidate | 0.155033 (0.153317–0.156748) | 0.153471 | 0.164351 | 0.174640 | 4,800 |
| 64 | 2 | control | 0.169057 (0.168332–0.169783) | 0.168831 | 0.172771 | 0.177640 | 4,804 |
| 64 | 2 | candidate | 0.155796 (0.154082–0.157510) | 0.154071 | 0.168821 | 0.170571 | 4,588 |
| 8,192 | 1 | control | 16.597530 (16.531845–16.663214) | 16.557378 | 17.027420 | 17.132290 | 10,728 |
| 8,192 | 1 | candidate | 14.786540 (14.756554–14.816525) | 14.767391 | 14.921771 | 14.924561 | 10,980 |
| 8,192 | 2 | control | 16.545505 (16.495074–16.595937) | 16.571138 | 16.734568 | 16.948739 | 10,720 |
| 8,192 | 2 | candidate | 14.624062 (14.588164–14.659960) | 14.609741 | 14.727752 | 15.023713 | 11,020 |
| 131,072 | 1 | control | 264.665767 (263.839045–265.492489) | 264.071112 | 267.857768 | 273.712351 | 107,200 |
| 131,072 | 1 | candidate | 236.834936 (236.625099–237.044774) | 236.738972 | 237.896047 | 237.990327 | 107,348 |
| 131,072 | 2 | control | 270.711302 (270.199647–271.222958) | 270.448847 | 272.990398 | 273.869812 | 107,052 |
| 131,072 | 2 | candidate | 238.225314 (237.995850–238.454778) | 238.031517 | 239.282742 | 239.493653 | 107,200 |

Candidate-versus-control mean changes are 64/r1: -8.086%, 64/r2: -7.844%, 8,192/r1: -10.911%, 8,192/r2: -11.613%, 131,072/r1: -10.515%, 131,072/r2: -12.000%. These are observations for this corpus and lifecycle, not a general DOCX speedup claim.

## Allocation and preservation

The predicted removal is `24*N+28` allocation callbacks and `24*N+108`
requested bytes: two scans of the N-paragraph source and two scans of the
(N+1)-paragraph candidate. The measured total rows below agree exactly in both
repeats and all 30 samples. The incremental region peak subtracts entry live
bytes; it is distinct from the process allocator high-water mark.

| Paragraphs | Allocation calls control → candidate | Calls removed | Requested bytes removed | Incremental peak control → candidate |
| ---: | ---: | ---: | ---: | ---: |
| 64 | 1,886 → 322 | 1,564 | 1,644 | 506,598 → 506,598 |
| 8,192 | 213,344 → 16,708 | 196,636 | 196,716 | 2,669,662 → 2,669,662 |
| 131,072 | 3,653,984 → 508,228 | 3,145,756 | 3,145,836 | 35,371,102 → 35,371,102 |

Reallocation callbacks remain unchanged. The very large cumulative requested
bytes remain dominated by exact-one range growth, as analyzed in change 0479;
they do not prove equivalent physical copy traffic. The XML and paragraph index
remain document-sized. The operation releases its owners back to entry live
bytes. Every formal report preserves the corpus, source-read and sink oracle;
independent review retains all latency/RSS pair and repeat flags above 5%.

## Wider CPU context

These PMU values belong to separate whole processes, including corpus/oracle
preparation, warmups, measured runs and teardown. They do not isolate operation
CPU work. All six event running fractions are 100%. Generic cache misses do not
provide an exact L1/LLC breakdown. IPC is derived as instructions divided by
cycles over the same full process.

| Event | Control | Candidate |
| --- | ---: | ---: |
| cycles | 45,009,650,140 | 39,796,687,162 |
| instructions | 190,372,337,409 | 171,900,349,557 |
| branches | 38,033,276,983 | 33,375,930,885 |
| branch-misses | 106,422,675 | 36,637,401 |
| cache-misses | 8,363,440 | 9,494,031 |
| page-faults | 41,926 | 42,437 |
| IPC | 4.2296 | 4.3195 |

Generic cache misses rise about 13.52% in this single whole-process pair,
despite fewer cycles, instructions and branch misses. This adverse observation
is retained rather than folded into a combined score. The keep decision rests
on the repeated normal lifecycle results and exact allocation removal; this
profile does not establish an isolated operation cache-locality improvement.

## Decision and validation

Keep the seven-line candidate, committed as `90f47ea41`, for its repeated
normal lifecycle improvement and exact allocation reduction. There are no
positive total-lifecycle latency or whole-process RSS pair flags above 5%.
Individual phase and repeat flags remain in the independent measurement
review and generated summary; they are not hidden in an aggregate score.

All 22 required final gates pass. The ledger contains 24 successful receipts,
including the two optional PMU processes. The scoped suite passes 926 DOCX
library tests, 12 paragraph-copy tests, six removal tests, nine OPC shared
tests, seven harness tests and 14 evidence tests. DOCX formatting, all-feature
Clippy and rustdoc with warnings denied, crate boundaries and strict registry
checks also pass. The existing unrelated workspace-format/default-feature
Clippy issues documented in 0480 are outside this unchanged source delta.

The data-only verifier and final gate-ledger/cleanup checks pass. Source and
eight embedded templates match their frozen identities. All four authenticated
runtime binaries were removed; shared Cargo caches and user-owned untracked
documents remain. `setup-attempts.json` retains the missing-directory setup
failure before Cargo or a build gate started. No formal capture was excluded,
replaced or retried.

Full sealed verification and a fresh copied baseline pass without the original
runtime binaries. All eight independently resealed corruption probes are
rejected, including arithmetic changes and gate omissions/exclusions. The
portable results and their input seal are retained in the bundle.

## Remaining work

The full non-iWork goal remains open. Removing scanner name copies helps this
materialized lifecycle but does not supply an explicit-window existing-document
append. The separate `window-contract.md` names the missing OPC decoded-splice
replay substrate, DOCX source/candidate validation and section placement, and
exact inverse replay requirements. Repeated append, large appended streams,
native/cold/range sources, atomic replacement and explicit worker scaling remain
separate required evidence. No iWork source is changed.
