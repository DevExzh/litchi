# XLSX direct action-plan flattening was rejected

Date: 2026-08-11

Production base: `0fbabeba2841a6a1b09fd2c02a9666f7b8a67b88`

Scope: private XLSX worksheet emission. OLE2, RTF and ODF production code were
unchanged, and iWork/IWA was explicitly excluded.

## Hypothesis

Public workbook transactions retain effective cell actions in row-major
`BTreeMap<Address, Action>` order. The lossless worksheet writer then moves
those actions into an outer row-number `BTreeMap` and a second cell
`BTreeMap` per touched row before emitting the worksheet. Consuming the
already ordered plan directly should remove tree lookup, comparison and node
allocation work without changing transaction semantics or XML output.

## Measured prototype

The final prototype consumed the owned cell and row maps in one forward pass,
grouping only the current row in one reusable `Vec<(Address, Action)>`. It did
not change any public API, dependency, operation order, semantic filtering,
conflict or inverse construction, exact untouched-span copying, validation,
candidate parse/readback, style/shared-string checks, calculation invalidation
or bounded validated-store handoff.

An earlier borrowed-iterator prototype was discarded before the formal run
because its 100-sample screen regressed medium commit/save by about 1%. The
owned-stream version passed the existing 20-test raw worksheet edit slice and
prototype-only checks for interleaved row/cell actions, sparse materialization,
clear/remove/style effects, address order, exact untouched spans and semantic
readback. Those prototype-only tests and all production changes were removed
with the rejected implementation.

## Corpus and experiment

Both executables use the unchanged deterministic public harness and prepare
cell mutations outside timing. The timed cases are public transaction commit
and commit plus deterministic save; both retain patch-count checks, final
semantic verification, exact expected changed bytes and reopen validation.

- medium: four 32 by 32 worksheets, 4,096 cells, 41 updates, 15,254-byte
  archive, SHA-256
  `9574867b4f1ab4d30ce150de32d2a0b01267d15399ec9edd2c0d57b4bc60fab6`;
- dense-wide: two 256 by 256 worksheets, 131,072 cells, 1,311 updates,
  384,525-byte archive, SHA-256
  `5dd3ad701eb686f6d2d14e9f177a4e9433445728b57b484d53f663b2f87a7714`.

Environment: release Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD EPYC
9575F VM, Rust system allocator and CPU 2 pinned with `taskset`.

The formal order was before A, after A, after B, before B. Medium used 50
warmups and 500 samples per leg, pooling 1,000 samples per state. Dense-wide
used five warmups and 50 samples per leg, pooling 100 samples per state.

The common-harness baseline executable SHA-256 is
`ae3f016de5d4f7c3826eba6be3f60943ad43b61c899a2c2d79bfa841b1051bc1`;
its `.text` SHA-256 is
`c214309a986c1edc7a5c84f564c56c60dce101e98e31066a8a714aafe0e7c3f0`.
The prototype executable SHA-256 is
`d1bad9b113b86e48b0bd68c09ec73a640feb7f941793365678b8e1f9d3ab232b`;
its `.text` SHA-256 is
`88e6b572c63734e20b6087a528f81a0d419c2013020c2487e716dce5f6f92987`.

## Formal latency result

| Workload | Before p50 | Prototype p50 | p50 delta | p95 delta | p99 delta | Mean delta | Approximate 95% interval for mean delta |
|---|---:|---:|---:|---:|---:|---:|---:|
| Medium 1% commit | 13.701 ms | 13.490 ms | -1.54% | +0.14% | +4.33% | -0.73% | [-1.20%, -0.27%] |
| Medium 1% commit + save | 15.235 ms | 14.990 ms | -1.61% | -0.50% | +0.18% | -1.26% | [-1.45%, -1.08%] |
| Dense-wide 1% commit | 431.879 ms | 430.719 ms | -0.27% | -0.13% | -1.80% | -0.33% | [-0.84%, +0.19%] |
| Dense-wide 1% commit + save | 514.926 ms | 511.407 ms | -0.68% | -0.93% | -0.95% | -0.66% | [-1.18%, -0.15%] |

Both prototype legs were below both baseline legs at p50 for the two medium
cases. That direction did not translate into practical end-to-end materiality:
the largest formal p50 improvement was 1.61%, dense commit was statistically
indistinguishable on the approximate mean interval, and medium commit p99
moved 4.33% in the wrong direction.

## Allocation and memory attribution

Matched one-sample dense-wide commit processes used Heaptrack. The complete
process includes deterministic corpus generation and final verification, so
these are process-wide attribution numbers rather than timed-region allocation
claims.

| Metric | Before | Prototype | Delta |
|---|---:|---:|---:|
| Allocation calls | 35,577,673 | 35,555,505 | -0.0623% |
| Temporary allocations | 5,877,353 | 5,877,356 | effectively flat |
| Peak heap | 91.79 MiB | 91.79 MiB | flat |
| Heaptrack RSS | 94.61 MiB | 94.57 MiB | -0.04% |
| Leaked bytes | 1.78 KiB | 1.78 KiB | unchanged |

The profile confirms that removing the nested regrouping is real but too small
relative to complete worksheet scan, XML emission, compaction, semantic parse
and publication readback. It does not justify retaining a second streaming
state machine in this safety-sensitive writer.

## Decision

Rejected and fully reverted. No XLSX production code, public API, dependency,
test-only helper or new retained allocation remains. The current nested map
regrouping stays until a larger measured design can remove a material pass,
such as coalescing validated semantic planning with lossless emission while
retaining every existing error and readback boundary. Direct regrouping alone
should not be revived from this record.

The eight raw ABBA reports, matched Heaptrack summaries and their harness JSON
are under `docs/performance/results/`; digests are in
`xlsx-action-plan-sha256.txt`.

## Final verification

- rebuilding the release harness after the revert reproduced the baseline
  executable SHA-256 exactly;
- all-feature/all-target `litchi-xlsx` tests pass, including 730 library tests,
  integration suites, security/encryption cases and examples;
- warning-denied all-feature/all-target XLSX Clippy passes with only the two
  named existing test-lint allowances (`needless_question_mark` and
  `module_inception`);
- all 23 harness tests and warning-denied harness Clippy pass; unchanged ODF
  dependencies still print the recorded GenericArray deprecation warnings;
- all retained JSON shapes/sample counts, evidence digests, formatting,
  `git diff --check`, production-source identity and staged scope are commit
  gates.

## Next non-iWork work

1. ODF: add a media-rich ODS publication corpus and attribute unchanged member
   inflate/deflate/copy work before changing transport ownership.
2. OOXML: profile a broader semantic-planning/emission coalescing boundary only
   if it can eliminate full-sheet work, not merely replace the rejected row
   regrouping.
3. RTF: measure the new CP-1252, LZFu and producer-watermark variants before
   another parser specialization.
4. OLE2: continue final-publication attribution without reviving the rejected
   XLS terminal-render handoff.
