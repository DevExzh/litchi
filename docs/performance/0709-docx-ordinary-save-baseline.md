# 0709: current ordinary DOCX save baseline refresh

Status: current-source descriptive baseline only. `performance_claim: none`.
This record refreshes the ordinary DOCX open, edit, and save route at revision
`9b6a9bbfc7bcc2fbf103419eb6968e1025c529c7`. It does not compare a production
candidate and it does not select an optimization. The only source change in
this batch is a harness metadata correction in
[`tools/perf-baseline/src/ordinary_save.rs`](../../tools/perf-baseline/src/ordinary_save.rs);
the DOCX production crates are unchanged.

The refresh is needed because the older ordinary-save observation predates
subsequent DOCX and OPC changes, and its real fixture refused the semantic
edit. This batch measures an admitted real file beside that refusal and a
deterministic generated control, while keeping edit, publication, and
serialization boundaries visible.

## Workload

The generated corpus uses the harness's deterministic medium semantic DOCX:
200 paragraphs, with the fixed edit marker
`litchi-perf-0638-ordinary-save` appended through
`Package::document_mut().add_paragraph_with_text(...)`. The two caller-named
files are bound by path, byte count, and SHA-256:

| Corpus | Source | Bytes | SHA-256 | Edit outcome |
| --- | --- | ---: | --- | --- |
| generated | deterministic harness corpus, medium | generated before each untimed corpus proof | recorded in each child manifest | admitted |
| numbered-list | `test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx` | 55,519 | `ebb078d791b6deb4d0a2dae69a15a5274baf2ade815f07b336b394bb43af4c88` | admitted |
| alt-chunk-header | `test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx` | 77,621 | `ee33b932a6c31f430bc8a6e450592970419b0bcc155e5c9884065602b9c733b8` | refused |

`NumberedList.docx` is a checked-in third-party fixture with self-reported
application metadata; the packet makes no native Office producer or
compatibility claim. The `alt-chunk-header` result is an intentional typed
refusal. The current structural reason is that body-final section properties
are not the final body child. A refusal is retained as part of the ordinary
route's behavior, rather than being silently converted into an edit.

## Timed boundaries

Each phase runs in its own fresh benchmark child. The native lane uses three
repeats, 100 measured samples, and 10 warmups. The allocator lane uses two
repeats, three samples, and no warmup. Children are pinned to CPU 12. Repeat
two reverses both corpus and phase order; this provides same-source drift
evidence, not a before/after comparison.

The phase clocks are deliberately separate:

| Phase | Timed interval |
| --- | --- |
| lifecycle | documented open, one semantic edit, and save to a path; destination preparation, readback, digest, and cleanup are outside the clock |
| edit | `Owner::edit`: `document_mut` admission and marker append only; documented open and verification are outside the clock |
| atomic publish | documented save-to-path, including temporary sibling creation, publication write, permission preservation, temporary `sync_all`, replacement rename, and parent-directory sync |
| counting publish | documented sequential serialization into a bounded counting sink; open, edit, and byte accounting are outside the clock |

The phase distributions are descriptive and cannot be added or subtracted to
reconstruct a matched lifecycle. The atomic interval includes filesystem sync
work, so it is not a physical cold-cache or hardware-device latency claim.
The allocator lane is operation instrumentation evidence; its elapsed values
are not latency-comparable to the native lane.

## Descriptive measurements

The formal native lane completed all 36 children and the independent analysis
recomputed every sample statistic. The table reports each repeat's p50 in
milliseconds; it is a current-source observation, not a before/after result:

| Corpus | Phase | Repeat 1 | Repeat 2 | Repeat 3 |
| --- | --- | ---: | ---: | ---: |
| generated | lifecycle | 5.716 | 5.741 | 5.774 |
| generated | edit | 0.379 | 0.376 | 0.378 |
| generated | atomic publish | 5.236 | 5.205 | 5.247 |
| generated | counting publish | 0.323 | 0.327 | 0.344 |
| numbered-list | lifecycle | 5.747 | 5.707 | 5.704 |
| numbered-list | edit | 0.160 | 0.159 | 0.160 |
| numbered-list | atomic publish | 5.423 | 5.333 | 5.399 |
| numbered-list | counting publish | 0.124 | 0.124 | 0.123 |
| alt-chunk-header | lifecycle | 5.975 | 5.957 | 5.947 |
| alt-chunk-header | edit | 0.170 | 0.173 | 0.173 |
| alt-chunk-header | atomic publish | 5.475 | 5.492 | 5.534 |
| alt-chunk-header | counting publish | 0.092 | 0.092 | 0.130 |

The save-to-path interval is the largest measured phase on every corpus in
this run, while the counting-sink and edit intervals are much smaller. This
describes the selected intervals only; it does not assign the difference to a
particular device or permit phase summation. The native repeat review retains
18 over-5-percent spread flags. The largest are generated atomic p95/p99
spreads of 49.675%/77.344% and refused-fixture counting p50 spread of
40.815%; these are retained as repeat variability, not discarded as outliers.

The allocator lane completed 24 children. Its six samples per corpus/phase
are summarized below with operation-relative p50 values. `net live` and `peak
above start` are region-relative byte observations; absolute process gauges
and phase peaks are not substituted or summed. Its instrumented elapsed review
retains 24 over-5-percent repeat-spread flags; those flags describe the
allocator lane's elapsed instrumentation and do not become native latency or
allocation-count variability claims.

| Corpus | Phase | Allocation calls | Net live (B) | Peak above start (B) |
| --- | --- | ---: | ---: | ---: |
| generated | lifecycle | 15,617 | 461,654 | 1,049,866 |
| generated | edit | 8,277 | 388,061 | 390,629 |
| generated | atomic publish | 6,528 | 21,575 | 609,787 |
| generated | counting publish | 6,523 | 21,575 | 544,147 |
| numbered-list | lifecycle | 4,675 | 122,437 | 710,663 |
| numbered-list | edit | 2,556 | 18,976 | 21,310 |
| numbered-list | atomic publish | 758 | 6,487 | 594,713 |
| numbered-list | counting publish | 753 | 6,487 | 529,073 |
| alt-chunk-header | lifecycle | 5,885 | 147,518 | 749,204 |
| alt-chunk-header | edit | 2,453 | 87 | 26,698 |
| alt-chunk-header | atomic publish | 971 | 1,310 | 602,996 |
| alt-chunk-header | counting publish | 966 | 1,310 | 537,356 |

These allocator values are descriptive instrumentation results from the
current source. The separate profile lane is reported below as instruction
attribution only.

## Callgrind attribution

The independent profile analysis passes all eight profiles: two repeats for
`edit` and `counting_publish` on the generated and admitted real corpora. The
measured owner rows are guest-instruction (`Ir`) counts:

| Corpus | Owner | Repeat 1 | Repeat 2 |
| --- | --- | ---: | ---: |
| generated | `Owner::edit` | 7,224,625 | 7,268,758 |
| numbered-list | `Owner::edit` | 2,685,230 | 2,684,009 |
| generated | `write_plain` | 7,248,464 | 7,232,750 |
| numbered-list | `write_plain` | 3,047,180 | 3,046,831 |

The edit owner is the documented `Owner::edit` boundary, including
`document_mut` admission and marker append. `document_mut` accounts for more
than 99.95% of each edit owner row; the direct append child ranges from 924 to
2,328 Ir. Each measured part is selected from the positive raw incoming edge
whose caller is `ordinary_save::run_case`: edit uses part 5, generated
publication uses part 6, and the admitted real publication uses part 5. Setup
parts and the zero-instruction termination part remain retained but excluded
from the measured rows. Immediate direct children reconcile with each owner;
the nested `document_mut` diagnostic is excluded from that disjoint partition.

The profile binary remains bound to source revision
`9b6a9bbfc7bcc2fbf103419eb6968e1025c529c7`; only packet profile scripts were
checked out at `1de70e7c702081c2d7a46ff58b2bdc27339fe2ec`, as recorded in the
[revision transition](results/change-0709/profile-revision-transition.json).
Callgrind Ir is guest-instruction attribution, not native latency, hardware
cycles, allocation counts, RSS, cache counters, or evidence for a production
optimization or causal hotspot.

## Metadata correction and semantic checks

The first preflight build exposed a harness identity error before formal
measurement. The real-file manifest had labeled the compressed ZIP member
payload and hash as the decoded main-part size and hash, and had not represented
the aggregate uncompressed member total. The preflight binaries, source
receipts, and one-sample admission artifacts are archived under
[`results/change-0709/metadata-preflight`](results/change-0709/metadata-preflight)
and are excluded from formal timing evidence.

The corrected harness reads `word/document.xml` through bounded
`ArchiveReader` limits for the decoded target identity and sums each member's
declared uncompressed size with checked arithmetic. Two focused tests cover the
decoded target/hash distinction and overflow refusal. The formal baseline was
rebuilt after this correction; raw preflight evidence is preserved and never
rewritten.

The independent public-API oracle supplements the harness determinism checks,
but its strict result currently fails. For the admitted fixture it checks the
existing paragraph projection as an exact prefix, appends the marker exactly
once as the final paragraph, and reopens outputs from both `to_stream` and
`save`; those scoped semantic checks pass. They do not certify whole-package
relationship or unknown-markup preservation: the admitted rewrite also changes
decoded `word/_rels/document.xml.rels`. For the refused fixture it checks the
typed refusal and no-marker semantic state, but both publication doors rewrite
the decoded `docProps/custom.xml` payload even though no custom-property edit
was requested. Source investigation identifies the cause: DOCX `write_plain`
unconditionally calls
[`custom_props.write_for`](../../crates/litchi-docx/src/package/codec.rs:1194)
when the custom-properties state is not dirty, and the canonical encoder
changes the fixture's BOM and `op:` prefix form. The decoded
custom-properties member changes from 632 to
602 bytes, and the decoded content is not equal after ASCII-whitespace
removal. Both no-edit controls produce the same archive SHA-256 as the
refused-edit outputs for both `to_stream` and `save`, proving that this
rewrite predates the attempted edit. The main document member and semantic
text remain intact. The strict failure is preserved in
[`oracle/first-semantic`](results/change-0709/oracle/first-semantic) and is
not weakened into a correctness pass. The v2 decoded/no-edit controls and
retained output artifacts are in `oracle/report.json` and `oracle/artifacts/`.
The compressed-member comparison remains a payload identity diagnostic; it
does not prove production compressed passthrough or a native Office producer
contract.

## Evidence state

The frozen plan, source/build receipts, fixture census, corrected metadata
patch, oracle, child capture receipts, and independent analysis live in
[`results/change-0709`](results/change-0709). `analysis.json` verifies custody,
native/allocator semantic parity at the harness-output level, deterministic
corpus evidence, metric cardinality, and recomputed statistics for all 60
children. That parity is not a substitute for the failed strict preservation
oracle, so the batch promotes no workflow-correctness result and no
performance claim. The profile lane passes as scoped attribution only. The
packet records no native Office reopening, no physical cold-cache campaign, no
cross-platform result, and no producer compatibility result.

Harness formatting, six focused ordinary-save tests, warning-denied Clippy,
source-policy checks, and all six repository evidence gates pass. Independent
review accepts the native/allocator custody and raw profile selection. Owned
build, binary, and filesystem scratch roots are removed with retained identity
witnesses; derived analysis replays byte-identically after cleanup. These checks
do not change the strict preservation failure above.

The broader non-iWork performance goal remains active. This baseline refresh
does not by itself justify a production change. The next correctness handoff is
to guard the clean custom-properties path in `write_plain`, then add focused
coverage for clean saves, dirty custom-property edits, retry after a failed
write, and the raw OPC route before repeating the strict oracle.
