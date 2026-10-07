# 0944 — XLSX publication after the attribute-summary change

Source-compatible XML auditing remains the largest named nested attribution on
the current source: **53.48%** of medium and **56.88%** of dense-sparse guest
instructions. Preservation writing accounts for **45.89–45.90%** and
**42.64%**, respectively. These are Callgrind instruction shares, not native
time shares or speedup estimates. No production change is adopted. iWork is
excluded.

The earlier [0901 profile](0901-xlsx-publication-instruction-profile.md)
predates the accepted [0905 attribute summary](0905-xml-attribute-summary.md).
This batch refreshes that attribution with a newly built executable and the
same workload and corpus shapes. It does not perform a matched before/after
experiment or requalify 0905's native performance result.

## Source, build, and workload

The source is commit `933ee45cbb5d2ef9dc29583f7ec890ad234e04fb`.
`crates/xml-minifier/src/audit.rs` is 126,817 bytes, SHA-256
`eddc437156c976c55ce36b0bf8f0ef47cc4943d296735a1ac8d6ef120b8754e7`,
identical to accepted 0905 commit
`8267929233736e3f4f4612eee94e47016a74685e`.

A fresh Git export initially included the whole tracked tree. A selected
10,788-file build closure, totaling 259,872,266 bytes, was then derived for
compilation and independently checked. The export excludes unrelated
working-tree changes. Full source listings and raw evidence remain outside
the repository under `/home/zhuhe/litchi-performance-0944`.

The standalone `tools/perf-baseline` manifest and its tracked lockfile were
built offline with Rust 1.95.0, release optimization, fat LTO, one codegen unit,
debug information level 1, incremental compilation disabled, and two Cargo
jobs. No `RUSTFLAGS` or compiler wrapper was supplied. The executable is
265,794,648 bytes, SHA-256
`63cbffd8c17299a4e2d1096e4ffa2a2d34a549cc452dfdf8ba5691b8ab7e75a3`.
The build passed within its 900-second deadline.

The case is `xlsx_source_backed_cell_values_one_percent_edit_save`, using
the marker-free `litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1`
generator. Medium updates 93 of 9,216 cells; dense-sparse updates 178 of
17,792 cells. Each package has four worksheets, seventeen ZIP members, and
eight 512 KiB media members. This is the existing in-memory source-backed
workflow, not a physical cold-cache, filesystem, or remote-provider test.
It covers the [CRUD taxonomy](../CRUD_Scenario_Checklist.md) scenario of
updating approximately one percent of logical objects and saving.

Four fresh processes ran serially on CPU 12 of the shared AMD EPYC 9R45 host:
medium, dense-sparse, dense-sparse, medium. Each used zero warmups and one
sample. The first child passed semantic and raw-profile admission before the
remaining three ran. No workload was retried or replaced.

## Scope and reconciled counts

Collection is toggled, reset, and dumped around
`SourceBackedEditor::publish_multi_commit_to_stream`. Opening, planning,
staging, commit, output reopen, and correctness checks are outside this owner.
The returned snapshot's destruction is also outside it, although destruction
is included in the harness's native publication timer. Instrumented elapsed
durations are retained but excluded from performance conclusions.

| Shape / repeat | Publication Ir | Topology preparation Ir | Nested XML audit Ir | Preservation writer Ir |
| --- | ---: | ---: | ---: | ---: |
| Medium / 0 | 92,990,939 | 50,025,084 | 49,735,942 | 42,677,347 |
| Dense-sparse / 0 | 167,532,948 | 95,587,073 | 95,298,865 | 71,430,887 |
| Dense-sparse / 1 | 167,533,924 | 95,587,667 | 95,298,831 | 71,431,602 |
| Medium / 1 | 92,989,937 | 50,023,578 | 49,735,426 | 42,678,081 |

Preparation and writing are separate immediate children of
`SourceBackedPackage::write_topology_to_stream`. XML auditing is nested
inside preparation and must not be added to it. The remaining publication
work is 288,508 / 514,988 / 514,655 / 288,278 Ir in schedule order; it includes
topology-plan construction, other children, drops, and self costs. The raw
analysis retains each direct edge rather than hiding this residual.
The XML and writer columns name `validate_source_part_replacement_xml` and
`write_changed_overlays_with_appended_inner`, respectively.

Five XML validation calls cover the four changed worksheets and workbook.
The writer branch includes unchanged-member transfer, regenerated-member
compression, framing, and accounting. Nested `PreservationIndex` and Deflate
rows overlap that branch and are reported separately, without addition.

The leading exclusive function is
`zlib_rs::deflate::longest_match::longest_match`: 18,584,283 Ir for each medium
repeat and 35,695,609 Ir for each dense-sparse repeat. XML event reading and
memory operations also remain visible. Exclusive self costs reconcile to
each measured dump total. Two repeats per shape do not establish confidence intervals,
and instruction shares cannot be multiplied by historical native phase shares
to predict a gain.

Each process contains one measured publication, five separately classified
lifecycle publications, and a zero-cost termination dump. The positive
incoming edge from `run_xlsx_cell_values_edit_save` identifies the measured
call; lifecycle parents are retained separately. Five lifecycle calls were a
prediction, not permission to discard unexpected calls or parents.

## Correctness and limitations

All four reports pass the harness's output reopen, updated-cell, semantic
hash, package topology/relationship, untouched-member, source/cache, and sink
checks. Their corpus, output, semantic and untouched-member hashes match the
retained 0900 context. Twelve untouched members survive; maximum writes are
65,536 bytes. Logical source reads are 4,233,170 / 4,258,604 bytes for medium /
dense-sparse, respectively, each in 183 calls.
These archive-range reads are not physical I/O or decoded replacement-copy
sizes.

Every run retains Valgrind's `brk segment overflow` warning. Exit status,
semantic checks, owner classification, and instruction conservation pass,
but the warning remains a profiling limitation. No native latency, hardware
cycle, allocation, peak-memory, RSS, concurrency, or universal performance
claim follows from these captures.

## Follow-up boundary

XML auditing and preservation writing remain the main attribution targets.
The prior 0890/0896/0898/0899/0902 rejections remain in force. This profile
does not justify weakening original-first XML validation, limits, offsets,
source freshness, compression preservation, or durable inverse restoration.

A distinct ownership hypothesis exists during topology-plan construction:
`append_worksheet_replacement` and `append_owner_topology` copy generated
payloads before passing replacements to OPC. These calls occur inside the
publication owner, after commit. The entire `MultiSnapshot::topology_plan_from`
branch is only 274,680–274,734 / 501,145–501,173 Ir here, about 0.30% of the
owner. Sharing already-owned replacement bytes therefore needs separate
allocation and native qualification; this profile does not establish useful
latency savings. Any candidate must retain managed/budgeted behavior, typed
refusals, full XML audit, source isolation, and exact preservation. Exact
replacement-copy bytes are not recorded by this cohort.
The conditional calculation-chain addition path is not exercised here.

## Verification and retained evidence

The [bounded packet](results/change-0944/README.md) contains the frozen plan,
build receipts, replay tools, reduced results, controls, and independent audit.
Preflight replay found and fixed malformed-dump and terminal-numbering checks
before the workload started. The original findings and corrected controls are
retained. No capture was rerun to fix analysis tooling.

Independent audit and root replay pass for all four children. Root replay
checks exact commands, environments, descriptors, artifact sets, terminal
status, chronological order, semantics, and raw analysis equality. Relocation
replay copied 11,003 files, hid the original root, and reproduced the 52,227-byte
reduction exactly (SHA-256
`7156b72d6b36e310d13f0365a154de9a1fa16bd413f4db27f32fd523f2f3ddea`).
Demangled function names and parent edges establish the reported scope; no
individual machine-code instantiation claim is made.

The independent audit's original receipt had two presentation errors: an
extra-field descriptor comparison labeled the binary a mismatch, and the
top-twenty self-cost sum was labeled as the full self-cost total. The retained
v2 supplement corrects both, checks the required conditions, and passes.
Neither correction changes a capture or a reported profile total.

The external archive is
`/home/zhuhe/litchi-performance-0944-archive/litchi-performance-0944-v1.tar.zst`,
226,729,433 bytes, SHA-256
`3f43604846b87843bbc71fcd4c8f075d57219022a1c737c472973186e85fb592`.
Extraction verified all 11,006 files and 995,046,474 logical bytes. The original
v1 audit is in that archive; its v2 correction is separately retained and
committed. The [archive summary](results/change-0944/archive-summary.json)
binds both and records regeneration commands through the retained archive tool.

Cleanup removed 2,112 build-cache paths, the redundant whole-repository export,
bytecode caches, and the verified extraction. It preserved both bound copies
of the executable and all 10,788 selected source files, rechecked after cleanup.
No production code or tests changed. Unrelated checkout changes remain intact;
the broader non-iWork goal remains active.
