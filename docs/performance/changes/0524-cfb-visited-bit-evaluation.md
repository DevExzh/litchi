# 0524: reject CFB visited-bit fusion and retain validation guards

The private CFB visited-bit fusion does not meet the frozen native admission
rule and is reverted. It removes duplicate checked bit access but lowers XLS
constructor instructions only 1.17–1.19%; none of the four primary workflows
improves by at least 3% in both repeats. Allocation-call, allocated-byte and
incremental-peak vectors remain identical. The retained source adds a collector
differential test and stabilizes an existing temporary-substitution test that
failed during quality checks. Production behavior is unchanged.

The [evidence bundle](../results/change-0524/README.md) retains the complete
candidate, source manifests, patches, native/allocator reports, scoped profiles,
hardware counters, failures, reviews and replay tools. The base is
`477281a2f83c3256bf3cc06fbc5c57724c9b6bbc`; 0523 motivates the experiment but
does not substitute for this batch's fresh control captures.

## Matched result

Each stage retains 24,000 native samples and 720 separate allocation samples.
Native blocks follow A1/B1/B2/A2 with two children per group, 20 warmups and
1,000 samples per row. Input is already materialized in each child. The final
control block uses the retained control binary under the frozen candidate
workspace; receipts distinguish executable source from execution-workspace
source. Allocator timings are excluded from native latency summaries.

| Primary workflow | Repeat 1 p50, µs | Repeat 2 p50, µs |
| --- | ---: | ---: |
| `xls_source_backed_open` | 135.391 → 141.980 (+4.87%) | 135.555 → 134.871 (-0.50%) |
| `xls_source_backed_open_one_cell` | 139.865 → 143.275 (+2.44%) | 136.195 → 136.581 (+0.28%) |
| `xls_owned_source_open` | 129.035 → 130.916 (+1.46%) | 127.611 → 126.601 (-0.79%) |
| `xls_owned_source_open_one_cell` | 133.095 → 135.961 (+2.15%) | 129.955 → 126.870 (-2.37%) |

CFB few-large p50 regresses 2.41%/2.35%, tiny improves 1.28%/1.70%, and
many-small is mixed (+0.79%/−0.94%). The gate fails independently of any
instruction or allocation result. No production speedup is claimed.

Across 16 profile children, 80 measured constructor dumps are separated from
12 CFB setup dumps. All final process dumps have zero collected instructions.
XLS constructor Ir changes 13,973,331 → 13,807,165 and
13,971,462 → 13,808,665. Its collector self Ir changes
5,601,140 → 5,436,525 per repeat, about −2.94%. CFB few-large constructor
Ir falls 1.25%. A large owner share is not a measurement of removable work.
The profiler ends at the constructor, before the selected-cell query; nested
call metadata is not a timed-call or allocation count.

All 24 matched allocation rows have identical allocation-call, allocated-byte
and incremental-region-peak vectors. Live-byte balances reconcile. Absolute
live values, deallocations and whole-child RSS remain separate. Four grouped
hardware captures retain whole-process counters; they include setup, queries,
oracles, drops and reporting and establish no operation-local counter claim.

The [adverse review](../results/change-0524/adverse-review.json) retains all
43 matched flag records and 60 same-build variations. The matched records
include repeated whole-child system-time diagnostics for each row of a shared
child, spread/extreme statistics and a +5.19% tracked-source list p99. They
are not 43 independent latency regressions. The second candidate tracked-source
list maximum rises 236.35%; that extreme is retained without trimming. Two
children cannot establish a cause or stable tails. Rejection does not depend
on explaining away any flag.

## Retained tests and quality corrections

The candidate's four new tests compare bit operations with existing insertion
semantics and scratch chain collection with the unchanged owned-result helper.
The collector guard checks exact errors, valid output, malformed/empty/cyclic
chains, reset after failures and buffer reuse. Only that independent collector
guard remains after rejection; the three candidate-API tests remain in the
retained candidate patch.

The first preflight passes 273 library tests, including all new guards, but
fails five existing filesystem tests. Three report shared `/tmp` disk quota
errors; two fail their expected filesystem-state assertions. Repeating the
same source with a disk-backed owned `TMPDIR` passes all 278 library tests.
Both runs and cleanup of the first run's five identified artifacts are kept.

A later full CFB run passes 277 library tests but fails the existing temporary
substitution guard. That guard unlinks a staged file and immediately recreates
its pathname, which can reuse the same native identity. A filesystem-only
probe observes identity reuse 100/100 times for unlink/recreate and 0/100
when retaining the displaced original. This supports the failure mechanism;
it does not retrospectively recover the original failed test's inode values.

The final test renames the original to a separate test-owned path, writes the
replacement, explicitly checks distinct identities, and retains the original
until assertions finish. It then removes both files. The production best-effort
identity/cleanup contract and trusted-parent requirement are unchanged. All
final quality gates run against restored production plus these two test changes.
The interrupted candidate quality run and its failure are retained explicitly.

All fourteen final quality gates pass with 4,374 test executions, in addition
to the valid 278-test candidate preflight. Post-cleanup verification passes
three source-stage replays, 59 serial intervals, eight analysis/comparison
reports, 160 annotations and four semantic negative vectors. Owned build paths
and Python caches are absent.

## Limits and next work

No public API, dependency, unsafe code, provider, allocation policy or production
validation change is adopted. All 30 previously read accepted ADR/index hashes
remain unchanged. This synthetic in-memory comparison does not close native
Office producer, fuzz, physical-provider, cold/range, concurrency scaling or
broad CRUD requirements. Fuzz tooling remains unavailable in this environment.

The [next-priority review](../results/change-0524/next-priority-review.md) follows
larger OOXML reconstruction/rewrite work, requiring a concrete removable pass
or representation before another candidate. Mandatory CFB ownership and
physical reconciliation remain required. OLE2/OOXML stays active; ODF is
deferred until that optimization goal completes, and iWork is excluded.
