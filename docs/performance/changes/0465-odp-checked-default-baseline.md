# Change 0465: checked-default ODP append baseline

0465 adds the existing `odp_existing_append_lifecycle` scenario to the checked
default matrix. It changes benchmark selection and checked corpus/coverage
metadata; production code, generated fixture bytes, timing boundaries and
preservation oracles are unchanged. Preflight passes with all prior 198 row
identities unchanged. The exact Rust/Python checked identity is 37 default
cases, 201 rows and 31 deterministic corpora, with catalog SHA-256
`d2c35126ee4e862ada539944ddb6cc2c654b82fe1e1465f505034fd1a9f7a84f`.

The frozen protocol has two normal full-matrix runs and two ODP-only allocator
runs, each with three warmups and 15 samples per row. It retains 6,030 normal
samples, 90 allocator samples and 180 ODP samples. Normal ODP p50 values are:

| Shape | Normal p50 R1 / R2 | Allocated bytes per iteration | Allocation calls | Reallocation calls |
|---|---:|---:|---:|---:|
| Tiny, 64 slides | 1.691887 / 1.687518 ms | 8,521,059 | 11,020 | 1,854 |
| Medium, 4,096 slides | 67.786424 / 67.745553 ms | 99,149,357 | 490,881 | 90,611 |
| Large, 8,192 slides | 136.334131 / 136.843521 ms | 191,235,475 | 978,314 | 180,732 |

The allocator values match between repeats. Region peaks are absolute
process-live observations including baseline and are not net working-memory
measurements. GNU time reports 161,524/152,224 KiB for the full normal runs
and 82,680/82,568 KiB for the ODP-only allocator runs; the different case
sets mean these RSS values are not compared across instruments.

The representative taxonomy remains 15 categories and 33 mappings, with 11
measured and 22 correctness-only mappings. The mixed append-incremental
category contains exactly one measured ODP row and two correctness-only fresh
streaming rows; both full-report validators pass this invariant. The four lane
receipts pass. The harness has 353 passing library tests with one ignored;
warning-denied all-feature/all-target Clippy, rustdoc, scoped formatting, 167
latest Python tests and boundary checks pass. The initial two stale
hash/count-pin Python failures remain retained as historical receipts.

The sealed precleanup verifier, five resealed negative probes and finalize
precleanup pass. Fresh-copy flagless portable verification passes with an
unchanged seal and absent temporary directory. Cleanup removed two owned
binaries totaling 116,542,728 bytes; `/tmp/litchi-goal-0465` and its temporary
audits are absent. This is descriptive fully materialized ODP lifecycle evidence only;
it makes no regression, speedup, independent native-producer, bounded-memory
streaming or scaling claim. See the [0465 result bundle](../results/change-0465/README.md),
[source review](../results/change-0465/source-review.md), [protocol](../results/change-0465/protocol.json),
and [checked catalog](../results/change-0465/checked/perf-corpus-manifest-v2.json).
