# 0713 — exact MCE substring search

Status: retained private-helper optimization, with explicit baseline test-lint debt.
Performance scope: the two named DOCX edit workloads in this packet only.

The candidate replaces the private shared MCE `find_bytes` scalar byte-window
search with `memchr::memmem::find`, using the existing direct dependency
(locked at 2.8.3). It preserves first-match offsets for arbitrary bytes, empty
needles and overlapping patterns. The four callers still perform namespace
detection, marker search, comment-close search and marker-collision checks.
Validation, source offset inputs, typed refusals and publication remain in place.

This is a local implementation change under ADRs 0005 and 0006. It adds no
API, dependency, unsafe code, cache, parallelism or ambient I/O. The accepted
ADR and GOAL hashes are retained in `constraints.json`; no architectural ADR
change is required.

## Frozen experiment

The experiment uses the ordinary-save harness at revision
`1d65a12dfa17c417632bd1d5c1b29d0bc7bc76ee`: a deterministic 200-paragraph
generated DOCX and the admitted 55,519-byte `NumberedList.docx` at
`test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx`.
Its SHA-256 is `ebb078d791b6deb4d0a2dae69a15a5274baf2ade815f07b336b394bb43af4c88`.

A1/B1/B2/A2 stages retain 16 native children (100 samples, 10 warmups) and
16 allocator children (three samples, no warmup), pinned to CPU 12. B2 and A2
reverse corpus and phase order. Each report binds exact source, executable,
fixture, argv, capture script and plan identities. Machine/toolchain details
are in [host.json](results/change-0713/host.json). Cargo and native measurements
were serialized in the coordinator lane. A reviewer ran unplanned default-target
Cargo checks before native captures; their unbound results are disclosed in
[review-preflight.json](results/change-0713/review-preflight.json) and do not
replace the coordinator’s source-bound checks.

The release baseline already includes two cfg(test) differential guards; the
candidate adds the helper replacement and a third cfg(test) marker-exhaustion
guard. Test additions do not enter the release benchmark binary. The original
HEAD, measured baseline and candidate source snapshots and both patches are
retained with hashes in `source-preparation.json`.

## Native paired observations

Times are nanoseconds; deltas are candidate relative to baseline. Edit measures
the documented `Owner::edit` admission and append. Lifecycle measures open, edit
and save-to-path; these separately sampled intervals are not additive.

| Pair | Corpus | Phase | p50 baseline → candidate | Delta | Mean baseline → candidate | Delta |
| --- | --- | --- | ---: | ---: | ---: | ---: |
| pair-1 | generated | edit | 377476 → 350321 | -7.193835% | 380475.98 → 352407.85 | -7.377110% |
| pair-1 | generated | lifecycle | 5664843 → 5677818 | +0.229044% | 5697634.83 → 5675813.66 | -0.382986% |
| pair-1 | numbered-list | edit | 159311 → 143736 | -9.776475% | 161213.94 → 145217.15 | -9.922709% |
| pair-1 | numbered-list | lifecycle | 5686573 → 5736109 | +0.871105% | 5702406.76 → 5796878.4 | +1.656698% |
| pair-2 | generated | edit | 375571 → 347567 | -7.456380% | 380586.61 → 356505.51 | -6.327364% |
| pair-2 | generated | lifecycle | 5691216 → 5726444 | +0.618989% | 5702183.69 → 5743505.79 | +0.724671% |
| pair-2 | numbered-list | edit | 157451 → 144616 | -8.151742% | 159349.22 → 146849.13 | -7.844463% |
| pair-2 | numbered-list | lifecycle | 5737128 → 5735094 | -0.035453% | 5736882.01 → 5727427.36 | -0.164805% |

All 48 frozen hard gates pass: edit p50 and mean improve at least 3% for
both corpora in both pairs; lifecycle p50/mean regress at most 3%; allocation
requests and requested bytes regress at most 3%. All 32 normalized deterministic
output comparisons pass. Lifecycle is only a non-regression gate, not an
end-to-end speedup claim.

## Allocation and live-byte observations

Allocation request counts and requested bytes are identical across baseline and
candidate in both pairs, at both p50 and mean. Values below apply to all three
samples of every corresponding allocator child.

| Corpus | Phase | Allocation requests | Requested bytes | Net live bytes | Peak above operation start |
| --- | --- | ---: | ---: | ---: | ---: |
| generated | edit | 8277 | 1768329 | 388061 | 390629 |
| generated | lifecycle | 15616 | 3199408 | 461654 | 1049866 |
| numbered-list | edit | 2556 | 436173 | 18976 | 21310 |
| numbered-list | lifecycle | 4674 | 2661485 | 122437 | 710663 |

Net-live and peak-above-start values are also unchanged in the matched samples.
These per-operation values are never summed across phases or repeats. Allocator
elapsed values are instrumented diagnostics and do not establish native latency.

## Every tail and repeat flag

The two native tail regressions remain explicit:

| Pair | Corpus | Phase | Metric | Candidate delta |
| --- | --- | --- | --- | ---: |
| pair-1 | numbered-list | lifecycle | p99 | +17.113426% |
| pair-2 | generated | lifecycle | p99 | +8.151842% |

The 14 same-source repeat-drift flags over 5% are:

| Lane | Corpus | Phase | Source | Metric | Spread |
| --- | --- | --- | --- | --- | ---: |
| native | generated | edit | baseline | p99 | 29.018430% |
| native | generated | edit | candidate | p99 | 40.189436% |
| native | generated | lifecycle | candidate | p99 | 10.989378% |
| native | numbered-list | lifecycle | candidate | p99 | 18.377026% |
| allocator | generated | lifecycle | baseline | mean | 15.143323% |
| allocator | generated | lifecycle | baseline | p95 | 36.229748% |
| allocator | generated | lifecycle | baseline | p99 | 36.229748% |
| allocator | numbered-list | edit | baseline | p50 | 8.902765% |
| allocator | numbered-list | edit | baseline | p95 | 6.657433% |
| allocator | numbered-list | edit | baseline | p99 | 6.657433% |
| allocator | numbered-list | lifecycle | candidate | p50 | 15.477722% |
| allocator | numbered-list | lifecycle | candidate | mean | 15.288258% |
| allocator | numbered-list | lifecycle | candidate | p95 | 17.632507% |
| allocator | numbered-list | lifecycle | candidate | p99 | 17.632507% |

These flags limit the claim: no stable lifecycle-tail improvement is established.
Edit p50 and mean improve in all four comparisons and none of the native p50
or mean same-source repeats crosses 5%, but that does not establish the cause
of tail variability. No samples were removed or rerun. The two p99 regressions
are accepted as disclosed review limitations under the frozen p50/mean gate;
they are not erased by an aggregate or attributed to unmeasured system noise.

## Correctness and limits

The two differential-search tests compare the helper against the former scalar
implementation, including all 130,305 pairs of short binary words, overlaps,
empty and oversized needles, longer arbitrary bytes, namespace and marker
near-matches. The third guard verifies the exact refusal after all 256 marker
salts collide. All three candidate-focused tests pass.

The unchanged 0712 public active-offset probe is compiled against both source
states. Its two 11-case JSON reports are byte-identical and also identical to
0712. It covers strict/transitional namespaces, Choice/fallback selection,
nested outer-range suppression, empty input, exact errors, and the 1,000,001
offset-count refusal. Two no-anchor malformed-MCE inputs still refuse at
`document_mut`. The package name/schema retain 0712 deliberately because the
probe itself is unchanged. This is differential evidence, not exhaustive native
Office compatibility or whole-format certification.

## Resource diagnostics and final validation

All 4,995 all-feature/all-target tests across common MCE, DrawingML,
spreadsheet-drawing, DOCX, XLSX, PPTX and XLSB pass with no ignored tests.
Formatting and all six repository evidence gates pass. Warning-denied Clippy
passes for all seven libraries and all targets of the other five crates.
Doctests report 92 passed and 46 existing ignored examples; warning-denied
rustdoc passes. The exact scope and commands are retained in `quality.json`.

Full warning-denied all-target Clippy exposes seven preexisting test-only lint
violations: four XLSB `expect_used` cases and three PPTX `err_expect` cases.
The exact files/lines and both failed candidate attempts are retained in
[quality-exceptions.json](results/change-0713/quality-exceptions.json).
Separate baseline runs reproduce the same four and three errors, respectively;
those files are byte-identical between baseline and candidate. No warning
allowance or unrelated source edit is used. This batch does not claim clean
all-target Clippy for PPTX or XLSB.


Eight Callgrind children collect only the exact edit owner, with one sample
and no warmup. Every child retains five numbered parts (one setup build, three
reference publications, one measured call from `ordinary_save::run_case`) and
the zero-Ir terminal part. Two retained independent raw-edge parsers validate
the owner/caller selection. Guest Ir is a mechanism diagnostic, not hardware
instructions, native elapsed time or allocation counts.

| Pair | Corpus | Owner Ir baseline → candidate | Delta |
| --- | --- | ---: | ---: |
| pair-1 | generated | 7190936 → 6712693 | -6.650636% |
| pair-1 | numbered-list | 2684609 → 2490323 | -7.237032% |
| pair-2 | generated | 7219671 → 6732795 | -6.743742% |
| pair-2 | numbered-list | 2685708 → 2488150 | -7.355900% |

All four Ir comparisons improve; same-source Ir repeat spreads are below 0.4%.
The post-capture [exact-edge attribution](results/change-0713/search-attribution.json)
shows generated A1 → B1 `active_offsets` inclusive Ir falling 516,556 → 13,760,
and numbered-list 1,105,302 → 908,772. The generated baseline inlines search,
so absence of a standalone `find_bytes` edge does not mean no search occurred.
These inclusive rows overlap their parents and cannot be added to owner totals.

Sixteen `/usr/bin/time -v` children use ten samples and two warmups. Each cell
below is that child’s whole-process peak RSS in KiB; no peaks are summed.

| Corpus | Phase | A1 | B1 | B2 | A2 | Pair 1 delta | Pair 2 delta |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |
| generated | edit | 40052 | 39868 | 40088 | 40124 | -0.459403% | -0.089722% |
| generated | lifecycle | 40164 | 40028 | 39976 | 40052 | -0.338612% | -0.189753% |
| numbered-list | edit | 40060 | 40036 | 40028 | 40092 | -0.059910% | -0.159633% |
| numbered-list | lifecycle | 39952 | 40004 | 40040 | 40000 | +0.130156% | +0.100000% |

No paired RSS comparison or same-source RSS repeat crosses the 5% review
threshold. The analyzer also retains its lower-of-two-reference diagnostic;
the complete individual peaks above are the basis for this bounded review.
The small differences do not establish a general memory-reduction claim.
All 24 diagnostic outputs match their native pilot identities.

The first diagnostic attempt stopped after its first profile subprocess and
before a receipt because of an undefined bookkeeping variable. Its 13 raw
artifacts and original scripts/freeze are preserved under
`mechanism-incomplete-attempt/` and excluded from comparisons. The corrected
driver was frozen before the complete 24-child diagnostic matrix. The native
pilot was not rerun; no samples or thresholds were selected by outcome.
`mechanism-recovery.json` also records locale custody and GNU-time whitespace
parser fixes. This is an execution repair, not a performance adjustment.

Both analyzers and the post-capture attribution replay exactly after cleanup.
The three owned scratch roots are removed, with six exact binary identity
witnesses retained. The four native analyzer corruption checks pass. The packet
audit, documentation gate and artifact seal bind the final retained state.

## Scope and next work

The native measurements are warm repeated observations; the allocator lane has
no warmup. All measurements are specific to the named workloads, CPU,
source and binaries. No hardware-counter, cold-cache, cross-architecture,
throughput, scaling, other-format or universal substring-speedup claim follows.
No geometric mean is used to conceal individual results. The shared helper’s
other OOXML consumers require quality coverage but have no fresh timing claim.

The broader non-iWork goal remains active. Fresh candidate profiles still show generated `alt::scan` work and
numbered-list active-block/MCE work. These are the next attribution boundaries; MCE validation must not be bypassed and
nested source anchors must not be treated as independent writer ranges.

[Evidence packet](results/change-0713/README.md).
