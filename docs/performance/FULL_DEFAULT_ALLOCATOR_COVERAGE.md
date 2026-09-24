# Full default allocator coverage

The allocator binary has a fail-closed coverage contract for the canonical
default selection. When `Case::DEFAULT` is selected with each default corpus
shape, payload, writer-shape, XLSX-shape, and semantic-shape exactly once, and
only the plain RTF variant, the report must carry
`allocation_coverage.contract = "default_case_matrix_213_rows"` and contain
213 in-process rows:

| workload | rows |
| --- | ---: |
| ZIP/OPC and CFB/OLE2 substrate | 144 |
| fresh DOC/XLS/PPT writers | 9 |
| XLSX snapshot/edit/save matrix | 45 |
| existing ODP append lifecycle | 3 |
| plain RTF/ODT/ODS/ODP semantic text export | 12 |
| **total** | **213** |

The contract was `default_case_matrix_201_rows` until the spec-gap branch
merge ([0759](0759-spec-gap-branch-merge.md)) combined it with the four
semantic text-export cases that
[0508](changes/0508-default-semantic-text-export.md) added to `Case::DEFAULT`;
those cases now open the same operation-scoped allocator region around
`write_text_to`. Reports captured under the 201-row contract remain valid
historical evidence for that matrix.

Every row has the same retained sample count as `configuration.samples_per_case`,
an operation-scoped allocator region, and measured values for every allocation
vector. Missing operation metrics, absent vectors, unavailable or overflowed
regions, duplicate sample identities, an unexpected case, duplicate
`(case, corpus)` identities, or an incomplete row count causes the allocator
run to fail before it writes a report. The coverage envelope also carries a
SHA-256 digest of the sorted canonical identities. The
`operation_metrics.latency_claim` for these rows is
`allocator_instrumented_elapsed_not_latency_claim`; allocator elapsed samples
are evidence only and are not latency measurements.

The normal binary uses the same region boundaries with instrumentation disabled.
Its elapsed timing path is unchanged and it does not emit the allocator
coverage contract. Metrics that are not available remain absent rather than
being represented as zero values.

## Build-lock provenance

The JSON report records the source revision, worktree state, Rust version,
Cargo target, flags, and allocator identity. Those fields identify the source
and executable that produced the report; they do not identify a temporary lock
file overlay used while building it.

Capture tooling that applies a temporary dependency-lock overlay must record the
base lock-file digest, the overlay patch or digest, the exact locked build
command and toolchain, and the post-capture restored-tree state in its capture
manifest. A clean-tree value observed after restoring the overlay is post-build
state and must not be reported as proof that the executable was built from the
unmodified lock file. The harness deliberately leaves this provenance to the
capture manifest instead of inferring it from `git_worktree_dirty`.
