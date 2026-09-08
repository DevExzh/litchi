# 0478 PPTX metadata-spool measurement review

This is an independent, read-only review of the final corrected-source
matrix, its public PPTX writer lifecycle, shared corpus and oracle, frozen
capture protocol, and retained arithmetic. The matrix contains 24 reports and
720 measured operations. The results remain descriptive observations; they do
not become a registered latency or whole-process memory claim.

The final matrix is internally complete. Every capture receipt exits with
status 0, every report has 30 samples indexed 0 through 29, and every measured
operation has the expected `37 + 2N` physical members. Each report contains
both the control and explicit-file-spool oracle cases. The six case proofs
`byte_exact_control_match`, `every_physical_member_verified`,
`every_slide_semantic_verified`, `presentation_graph_verified`,
`slide_geometry_verified`, and `text_digest_verified` are true in every case.
All 720 operations report `output_matches_oracle=true`.

The source and environment custody is bounded by more than the corpus and
output hashes. The recorded starting revision is
`4dd9b5cf66ad6463539d2c1a09be502fff4d6b70`. The final normal and allocator
builds each bind a 7,046-entry source manifest with SHA-256
`38d17523810931fd9e41d3ea6566f2542869b32e32a09148b58fa35f9fd3f080`; their
source-before and source-after records are identical, and all 24 receipts
carry this same manifest hash. The source manifest does not enumerate the
`include_str!` XML resources, so `embedded-inputs.json` supplies that missing
closure: it inventories and hashes all 19 XML inputs and their copied build
assets in a 4,980-byte manifest with SHA-256
`6d85f8b4cc16cbd21257265d7933b6f6378aae4d2585da92eed1b6abb859fcee`.
Both build records report the embedded manifest unchanged. Together these
records bind the exact source and embedded resource inputs; corpus and output
hashes alone would not establish that closure.

The captured executable identities are:

| Lane | Bytes | Binary SHA-256 | Source manifest | Embedded-input manifest |
| --- | ---: | --- | --- | --- |
| Normal | 414,496,768 | `6b2f543cdbd11868209deed7d0b265496bab9fc3bfbd9d4e49d1357135f0c969` | `38d175…f080` | `6d85f8…fcee` |
| Allocator | 414,504,696 | `6279bccb61337e572d215b7f788a670e95ea839a6d6dfe9a83e5f61fae52b632` | `38d175…f080` | `6d85f8…fcee` |

The allocator executable copy initially hit errno 122 after the unchanged
Cargo build. The retained failure and recovery receipts show the recovered
destination hash equal to the origin hash, with source and embedded-input
custody unchanged. The formal captures use the recovered final executable;
the superseded preliminary executables were removed by the recorded cleanup
receipt.

The environment record has SHA-256
`27469836a0b72136b4ad7ce0f85121a2dee3036b066b8c1ac28a3c7103ee4289`.
Its filesystem observation reports `/tmp` as `tmpfs`, and every protocol
spool path is under `/tmp/litchi-goal-0478/spools/`. The final capture
directories contain no remaining spool files.

The historical corpus identities are constant across all reports and match
the retained 0474/0476 evidence:

| Slides | Members | Archive bytes | Archive SHA-256 | Semantic SHA-256 | Full-text SHA-256 |
| ---: | ---: | ---: | --- | --- | --- |
| 8 | 53 | 36,259 | `951505889af106f032241c30f07b5d237e54822dade768c911aca2d0f68c22c5` | `f3444404ef5130757c79ae624a161a38bc7c2f7ed053b47315ccf31067573a4e` | `bbc9f0e6cf3b3c48cd991dca9dded766cec7c93f762dd74b1559a34cdc05a966` |
| 256 | 549 | 274,398 | `1f33f8b2c36a2a51abc62d827e4c915dd3b4e52859b300323f250ab561942cf2` | `147697c24b54b91e37e9906330802f40a2c4f9da1af6b564ca92eca6faad4f3b` | `4c4a1a185cd33c9a3362bed00bd9e77cd64f898976a57b56220ea62bb3a614ee` |
| 8,192 | 16,421 | 7,940,406 | `c7b08da644e651046d368b1baaff9a12c6d7218c4f96914dacb1033722e4b527` | `1bf460386d6f8962d4a04444a5e2b3971e6548fc1273cc933a6beefff9bf1417` | `521f638a72f55371d0d40ddf01dac1577f1e584f58ae535bb0cc01f90b76e1f9` |

For every count, control and spool oracle output bytes and SHA-256 values are
identical to the corresponding source archive values in the table. The
physical gate opens one `PhysPkgReader`, checks the exact unique member-name
set, and reads every member payload before the shared per-slide text,
geometry, and presentation relationship-graph oracle runs. This covers all
fixed and indexed members rather than only selected payloads.

The spool extent has an independent central-directory identity. For this
ZIP32 corpus, with no extra fields or name comments, it is
`Σ(46 + UTF-8 filename length)` over the 37 fixed names and the two names for
each slide. The fixed names contribute 2,914 bytes, so the equivalent
calculation is `2914 + 143N + 2Σ digits(i)` for `i = 1..N`:

| Slides | Members | Expected and observed scratch bytes |
| ---: | ---: | ---: |
| 8 | 53 | 4,074 |
| 256 | 549 | 40,842 |
| 8,192 | 16,421 | 1,237,692 |

Every spool oracle and every spool sample in all 24 reports has the matching
value. Control lanes have no scratch extent. The analyzer independently
derives this formula, so a self-consistent but malformed reported extent does
not pass the analysis gate.

The timed lifecycle is correctly scoped. The timer starts immediately before
the public writer is constructed. The timed operation constructs the writer
and its generated name/metadata plan, creates deterministic text strings,
opens the caller-selected `create_new` spool file in the spool lane, writes
all slides and relationships, runs `finish` through central-directory
publication, updates the scalar SHA-256 sink, and flushes and closes the
spool file. Corpus preparation, limits and expected output, allocator/process
endpoint observations, spool metadata length, sink digest finalization,
oracle reopening, and unlink are outside the clock. `HashingSink` retains only
scalar counters and digest state; it does not retain the generated archive.
The generated OPC plan retains bounded descriptors and a cursor, without a
completed-name index or proportional name collection.

The current error path propagates oracle and timed-operation failures before
cleanup. Successful endpoint observations are followed by unlink; failed
`create_new` or operation paths retain their scratch files for inspection, so
a caller-owned pre-existing file is not removed. The caller-selected
directory remains the exclusive custody boundary. Formal paths were fresh and
successful, and the final spool directories are empty.

The normal timing observations and spool-versus-control mean changes were:

| Slides | Repeat 1 control / spool (ms) | Repeat 1 mean change | Repeat 1 RSS change | Repeat 2 control / spool (ms) | Repeat 2 mean change | Repeat 2 RSS change |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 8 | 0.905319 / 0.933505 | +3.113% | 0.000% | 0.901271 / 0.931109 | +3.311% | -0.295% |
| 256 | 6.344202 / 6.320574 | -0.372% | +0.353% | 6.471067 / 6.295405 | -2.715% | -3.624% |
| 8,192 | 185.536251 / 182.077140 | -1.864% | -3.232% | 185.602817 / 181.198961 | -2.373% | -2.026% |

The analyzer’s nearest-rank p50, p95, and p99 comparisons also stay below the
5% review threshold; the largest is +4.794% at p99 for the 8-slide repeat-2
pair. All 12 pair rows have zero positive review flags, and all 12 repeat
drift rows have zero absolute review flags. External `/usr/bin/time -v` RSS
is a broad process observation that includes setup, oracle work, capture, and
teardown; its largest normal pair change is 3.624% in magnitude.

Allocator observations are identical across the two external repeats:

| Slides | Control → spool allocation calls | Control → spool requested bytes | Control → spool incremental peak live bytes |
| ---: | ---: | ---: | ---: |
| 8 | 850 → 617 (-27.412%) | 499,462 → 475,312 (-4.835%) | 435,701 → 432,436 (-0.749%) |
| 256 | 9,049 → 5,825 (-35.628%) | 1,415,242 → 861,904 (-39.098%) | 681,819 → 432,436 (-36.576%) |
| 8,192 | 278,153 → 179,677 (-35.404%) | 31,428,173 → 13,379,518 (-57.428%) | 8,875,252 → 432,436 (-95.128%) |

Every allocator operation has zero failed allocation calls, equal live bytes
before and after the operation, and a region peak at least as large as its
starting live bytes. The spool incremental peak is 432,436 bytes for every
count and both repeats, giving a 0% range against the 8-slide baseline and a
pass under the 1% memory gate. This is an operation-scoped generated-route
spool result. It does not bound total RSS, control-lane metadata growth,
native allocations outside the measured allocator, or all PPTX creation
memory.

The analyzer summary reports 24 captures, 720 samples, no positive or repeat
regression flags, and a passing allocator memory gate. The evidence and build
receipts also pass their recorded validation checks. These results support
matched public control and explicit-file-spool output for the deterministically
bound corpus, the stated operation timings, exact scratch extents,
physical and semantic oracle coverage, and the allocator spool peak gate.
They do not support durability, cold-storage or physical-I/O latency, total
process memory, or a broad PPTX writer performance claim. The final
ledger/portable seal remains a root-owned artifact-binding step and must bind
the source, embedded-input, binary, environment, protocol, report, and
summary records named above.
