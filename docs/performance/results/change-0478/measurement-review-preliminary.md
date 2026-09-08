# 0478 PPTX metadata-spool measurement review

This is an independent, read-only review of the public PPTX
metadata-spool benchmark, its shared corpus and oracle, the frozen capture
protocol, and the retained arithmetic. The review covers the completed
24-report matrix and 720 measured operations. It does not turn the results
into a registered latency or whole-process memory claim.

The timed lifecycle is correctly scoped. `Instant` starts immediately before
the public writer is constructed. The timed operation creates the writer and
its metadata/name plan, creates fresh deterministic text strings, opens the
caller-selected `create_new` spool file in the spool lane, writes all slides
and relationships, runs `finish` including central-directory publication,
updates the scalar-only SHA-256 sink, and flushes and closes the spool file.
The endpoint allocator/process snapshots, spool metadata length, sink digest
finalization, materialized oracle reopening, and unlink are outside the
clock. The timed sink retains only counters and a digest state; it does not
retain the generated archive or a proportional output/name collection.

The `/usr/bin/time -v` RSS values cover process setup, corpus and oracle
preflight, capture, and teardown. They are contextual process observations,
whereas allocator records are operation-scoped. Writer finalization calls
`flush`, but the benchmark does not call `sync_all`, so the timing describes
logical file activity and does not claim storage durability latency.

The source and environment custody is bounded. The build manifest covers the
Rust sources plus all 19 production XML inputs used by `include_str!` (the 18
generated presentation resources and `notesMaster.xml`), copies and hashes
each input, and verifies the same manifest before and after both builds. Both
captured binaries bind the same source manifest
`d911e79d8e3ebeb795ab2dc00dd55963e017e3fc41c56f386a0e1d12b8c3a645` and the
same embedded-input manifest. The portable verifier authenticates
`environment.json` (`27469836a0b72136b4ad7ce0f85121a2dee3036b066b8c1ac28a3c7103ee4289`),
requires its clean `df -T` observation to report `/tmp` as `tmpfs`, and binds
every protocol spool path under `/tmp/litchi-goal-0478/spools/`.

The historical corpus identities are bound in the analyzer and agree in
every retained report with the sealed 0474/0476 corpus:

| Slides | Members | Archive bytes | Archive SHA-256 | Semantic SHA-256 | Full-text SHA-256 |
| ---: | ---: | ---: | --- | --- | --- |
| 8 | 53 | 36,259 | `951505889af106f032241c30f07b5d237e54822dade768c911aca2d0f68c22c5` | `f3444404ef5130757c79ae624a161a38bc7c2f7ed053b47315ccf31067573a4e` | `bbc9f0e6cf3b3c48cd991dca9dded766cec7c93f762dd74b1559a34cdc05a966` |
| 256 | 549 | 274,398 | `1f33f8b2c36a2a51abc62d827e4c915dd3b4e52859b300323f250ab561942cf2` | `147697c24b54b91e37e9906330802f40a2c4f9da1af6b564ca92eca6faad4f3b` | `4c4a1a185cd33c9a3362bed00bd9e77cd64f898976a57b56220ea62bb3a614ee` |
| 8,192 | 16,421 | 7,940,406 | `c7b08da644e651046d368b1baaff9a12c6d7218c4f96914dacb1033722e4b527` | `1bf460386d6f8962d4a04444a5e2b3971e6548fc1273cc933a6beefff9bf1417` | `521f638a72f55371d0d40ddf01dac1577f1e584f58ae535bb0cc01f90b76e1f9` |

Each of the 24 reports contains both control and spool oracle cases. Their
output byte counts and SHA-256 values match exactly. All cases report
`byte_exact_control_match`, `every_physical_member_verified`,
`every_slide_semantic_verified`, `presentation_graph_verified`,
`slide_geometry_verified`, and `text_digest_verified` as true. The physical
pass opens one `PhysPkgReader`, checks the exact `37 + 2N` unique member names,
and reads every member payload before the shared slide, geometry, text, and
relationship-graph oracle runs.

The spool extent has an independent arithmetic identity. For this corpus it
is the ZIP central-directory sum `Σ(46 + UTF-8 filename length)` over the 37
fixed names and the two names per slide. The expected values are:

| Slides | Members | Expected and observed scratch bytes |
| ---: | ---: | ---: |
| 8 | 53 | 4,074 |
| 256 | 549 | 40,842 |
| 8,192 | 16,421 | 1,237,692 |

Every spool oracle and every spool operation in all 24 reports has the
corresponding value. Control lanes have no scratch extent. This direct check
also guards against accepting a self-consistent but malformed reported file
length.

All reports have 30 samples with indices 0 through 29, positive elapsed
times, and the expected internal repeat value of zero. The independent audit
found no report identity, output, oracle, sample-index, scratch, or allocator
invariant failures. The normal mean elapsed times and spool-versus-control
mean changes were:

| Slides | Repeat 1 control / spool (ms) | Repeat 1 change | Repeat 2 control / spool (ms) | Repeat 2 change |
| ---: | ---: | ---: | ---: | ---: |
| 8 | 0.906590 / 0.936994 | +3.354% | 0.905735 / 0.935587 | +3.296% |
| 256 | 6.342694 / 6.278200 | -1.017% | 6.386109 / 6.239078 | -2.302% |
| 8,192 | 188.333871 / 181.617815 | -3.566% | 184.315187 / 180.776141 | -1.920% |

The analyzer’s p50, p95, and p99 pair comparisons produce the same outcome:
no positive review flag exceeds the 5% threshold. Its repeat-drift checks
also produce no absolute flag above 5% for the normal timing or allocator
RSS observations.

Allocator samples have zero failed allocation calls and zero live-byte exit
delta. They also satisfy `live_before <= region_peak_live_bytes <=
peak_live_bytes_after` in every operation. The generated-route spool
incremental peak is exactly 432,436 bytes in every count and both external
repeats. Therefore the memory gate has baseline minimum 432,436, global
maximum 432,436, a 0% range, and passes its 1% threshold. This is an
operation-scoped generated-route spool result; it does not bound total RSS,
control-lane metadata growth, native allocations outside the measured
allocator, or all PPTX creation memory.

The source review found an error-path cleanup hazard before the formal run:
if `open_spool` failed because a caller-supplied path already existed, an
unconditional cleanup could unlink that preexisting file. The current source
now propagates `oracle_archive` and timed-operation errors before cleanup and
retains scratch for caller inspection, which removes that deletion hazard at
the cost of leaving a partial file after an error. The formal capture paths
were fresh and successfully created, so the original behavior does not
invalidate the retained 24 reports. Because this source fix was made after
the frozen binaries were captured, final binary/source-manifest custody must
bind the fixed source before making a final executable reproducibility claim.

The results therefore support the bounded statements above: matched public
control and explicit-file-spool output for the sealed deterministic corpus,
the stated operation timings, exact central-directory scratch extents, and
the allocator spool peak gate. They do not support durability, cold-storage
I/O, total-process memory, or a broad PPTX writer performance claim.
