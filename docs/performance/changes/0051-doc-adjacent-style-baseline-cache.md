# Change 0051: native DOC adjacent style-baseline cache

Date: 2026-08-11

Production control: `4f2dca7d5f`

Scope: private native DOC paragraph-property parsing only. iWork/IWA crates
were explicitly excluded.

## Hypothesis and change

The accepted post-0050 profile still attributed 4.44% of large DOC-open self
cycles to `validate_style_sprms`, 2.50% to
`resolve_paragraph_style_sprms`, and 7.11% to `PapBinTable::parse`. The fixed
large writer corpus has 512 source-ordered paragraphs with the same initial
paragraph style. Every PAPX run independently reconstructed and revalidated
that identical inherited style before applying its own direct properties.

`PapBinTable::parse` now keeps one private `(style index, resolved paragraph
baseline)` pair. A matching subsequent initial style clones that immutable
baseline and applies the run's direct PAPX plus piece modifier. A different
style replaces the one-entry cache only after its baseline resolves
successfully.

The direct-property parser remains authoritative for every run. In particular,
direct `sprmPIstd` and style permutations still resolve through the stylesheet
from the document baseline, table/revision state still follows the existing
ordered rules, huge/Data-indirected PAPX still expands and validates before
the cascade, and the final properties, direct `grpprl`, and initial style index
are retained as before. Inputs without a stylesheet or initial style keep the
existing scalar path.

This is constant extra memory scoped to one `PapBinTable::parse`. It introduces
no public type, retained snapshot state, lock, global cache, runtime,
dependency, unsafe code, output change, or weakened owner/public-reader
readback.

## Matched latency evidence

The frozen release binaries have SHA-256:

- control: `fc19181881e5d92479ecee39a4f5f9c9a56d09aa9538fc6682adc8ab81d6343f`;
- candidate: `44bfb89818856aaea4dfe96e7471c25b5b65d038509c210c486554f4078d7538`.

Both use the unchanged standalone harness, release profile, Rust 1.95.0,
Linux 6.8.0-101-generic, the Rust system allocator, and CPU 2 pinned with
`taskset`. The generated large DOC is 97,792 bytes with 512 paragraphs and
SHA-256 `3d96764fe48e213b972ff5921df183dab9e8bfc8c8e751bcf3bf20190de4fec6`.
Its 81,920-byte `WordDocument` stream has SHA-256
`33e6cd70a45181c28d4a3e7bfa4e7817bd82d7b2e89e39437a589243abdc38eb`.

The primary `doc_semantic_open` measurement used 50 warmups and 500 samples in
each of five control/candidate and five candidate/control pairs. Pooling raw
samples gives 5,000 observations per state while balancing binary order.
Corpus generation and complete semantic verification remain outside timing.

| Large DOC open | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 343.503 us | 304.199 us | **-11.44%** |
| mean | 350.684 us | 309.058 us | **-11.87%** |
| p95 | 395.906 us | 345.376 us | **-12.76%** |
| p99 | 490.476 us | 380.625 us | **-22.40%** |

The approximate independent-sample 95% interval for the mean delta is
`[-12.15%, -11.59%]`. All ten matched pair comparisons improve: p50 deltas
range from -13.57% to -9.17%, and mean deltas range from -15.33% to -8.76%.

The secondary large `doc_semantic_one_edit_save` ABBA used 30 warmups and 500
samples per leg, or 1,000 pooled observations per state. It includes the
changed commit and output materialization; exact output comparison, forward
patch, inverse restoration, strict snapshot reopen, independent public DOC
reopen, and complete semantic verification remain outside timing.

| Large DOC one paragraph edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 912.288 us | 875.736 us | **-4.01%** |
| mean | 920.905 us | 881.994 us | **-4.23%** |
| p95 | 988.022 us | 950.730 us | **-3.77%** |
| p99 | 1.047 ms | 0.995 ms | **-5.00%** |

Primary raw reports use the `doc-style-cache-open-{forward,reverse}-*` prefix;
secondary and guard reports use `doc-style-cache-{edit,guards,tiny}-*`. Their
hashes are indexed in
[`doc-style-cache-sha256.txt`](../results/doc-style-cache-sha256.txt).

## Attribution and resources

Matched 3,000-sample `perf record` processes directly confirm the intended
owner. Kernel symbols are restricted on this host, but userspace DOC frames
are resolved and both reports have zero lost samples.

| Self-cycle frame | Before | After |
|---|---:|---:|
| `validate_style_sprms` | 4.44% | 0.83% |
| `resolve_paragraph_style_sprms` | 2.50% | 0.58% |
| `PapBinTable::parse` | 7.11% | 6.76% |
| resolved-baseline clone | 1.29% | 2.67% |

The remaining clone is the intended constant-space tradeoff: it copies the
resolved typed baseline instead of rebuilding and revalidating the stylesheet
inheritance chain.

Matched process-wide `perf stat` A/B/B/A used 50 warmups and 1,000 measured
opens per leg. Pooling the two legs per state gives:

| Counter | Before | After | Delta |
|---|---:|---:|---:|
| task clock | 4,249.01 ms | 4,010.44 ms | -5.62% |
| cycles | 20,751,080,962 | 19,707,954,315 | -5.03% |
| instructions | 85,952,073,084 | 81,087,578,596 | -5.66% |
| branches | 21,015,951,026 | 20,304,625,920 | -3.39% |
| branch misses | 42,427,374 | 42,410,342 | -0.04% |
| cache references | 1,931,205,944 | 1,943,677,197 | +0.65% |
| cache misses | 151,142,249 | 156,871,807 | +3.79% |
| page faults | 20,680 | 20,679 | -0.005% |
| CPU migrations | 0 | 0 | unchanged |

Heaptrack used two warmups and 20 measured opens per state:

| Whole-process metric | Before | After | Delta |
|---|---:|---:|---:|
| allocation calls | 724,827 | 589,923 | **-18.61%** |
| temporary allocations | 400,098 | 298,920 | **-25.29%** |
| peak heap | 5.67 MiB | 5.67 MiB | unchanged |
| Heaptrack RSS | 17.49 MiB | 17.54 MiB | +0.29% |
| leaked bytes | 544 B | 544 B | unchanged |

Uninstrumented GNU Time ABBA reports a 30,976 KiB worst-case maximum RSS for
both states. All runs have zero major faults.

## Guardrails and correctness

Large guards pool 600 samples per state; tiny open and changed edit/save pool
2,000 per state.

| Guard | p50 delta | Mean delta | p95 delta | Disposition |
|---|---:|---:|---:|---|
| large list paragraphs | -1.06% | -1.50% | -1.28% | neutral/better |
| large one paragraph | -3.72% | -3.80% | -5.37% | better |
| large full text | -8.49% | -4.92% | -8.57% | sub-microsecond/better |
| large exact no-op edit/save | -0.76% | -1.15% | -3.07% | neutral/better |
| tiny open | -1.37% | -1.99% | -10.62% | neutral/better |
| tiny one edit/save | -0.38% | -0.05% | +2.89% | neutral |

The new focused test constructs base and derived paragraph styles, compares a
fresh resolved path with cached reuse under different direct and piece
modifiers, proves direct mid-run style switching remains authoritative, and
proves the one-entry cache rekeys when the initial style changes. Existing
tests retain inheritance, permutation, huge/Data-indirected PAPX,
table/revision, malformed-style, protection/refusal, preservation, patch,
inverse, and real Word/LibreOffice coverage.

Verification completed:

- `litchi-doc --all-targets --all-features`: 959 unit tests passed, two
  fixture-dependent tests remained ignored, and all integration/example
  targets passed;
- warning-denied all-target/all-feature DOC Clippy passed, including the
  previously requested deprecation cleanup;
- all 32 standalone harness tests and warning-denied all-target Clippy passed;
- the DOC libFuzzer target compiles;
- formatting, JSON parsing, artifact hashes, and `git diff --check` pass.

Warning-denied DOC rustdoc remains blocked by three pre-existing private links
in unchanged `section/columns` and `shape` module documentation. No changed
file contributes a rustdoc warning.

## Remaining work

This cache removes adjacent/repeated paragraph-style inheritance resolution.
It does not change CFB publication, character-style resolution, table-style
cascading, exact patches, security policy, streaming output, or the mandatory
strict owner and independent public-reader reopens. The remaining DOC work
must be attributed to a distinct owner; neither validation boundary is a
candidate for removal.
