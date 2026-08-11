# Change 0052: ODT final changed-result byte handoff

Date: 2026-08-11

Production control: `58dbdcc055`

Scope: private packaged-ODT transaction finalization only. iWork/IWA crates
were explicitly excluded.

## Hypothesis and change

After every changed ODT transaction had completed operation publication and
per-operation compact-XML auditing, `Edit::commit` still copied the complete
candidate archive into a new `Vec`, parsed that copy through
`Snapshot::from_bytes`, and then independently reopened the resulting snapshot
once more. Heaptrack attributed exactly one 16.79 MB allocation per changed
iteration to this `copy_bytes` call on the fixed media-rich corpus.

The finalization path now distinguishes only exact no-op from changed bytes:

- exact no-op still returns the original immutable `Snapshot` clone;
- a changed, already validated `Document` creates a byte-only snapshot through
  the existing `Snapshot::from_document` path, which repeats the 64 MiB package
  bound and clones the package's private immutable `Arc<Vec<u8>>`;
- `after.document()?` remains mandatory and performs one fresh, independent
  complete package/document parse before commit publication.

This removes one archive-sized copy and one redundant parse. It does not adopt
or retain the parsed final `Document`, and it does not remove the final
independent readback boundary.

That distinction is important. The rejected change-0020 prototype retained an
already parsed final document in the snapshot and regressed the dedicated
medium one-paragraph read guard by 6.33% mean and 17.64% p95. This change keeps
`Snapshot` byte-only and records a new pointer-identity regression proving the
changed package bytes are shared while semantic readback comes from a fresh
reopen.

No public API, archive type, patch vocabulary, dependency, runtime, lock,
global cache, unsafe code or output format changes.

## Matched latency evidence

The frozen release binaries have SHA-256:

- control: `44bfb89818856aaea4dfe96e7471c25b5b65d038509c210c486554f4078d7538`;
- candidate: `6a73ba2ca107d877e612bf484257d132ae003b54363ab3a43f9da717eabe5e50`.

Both use the unchanged standalone harness, release profile, Rust 1.95.0,
Linux 6.8.0-101-generic, the Rust system allocator, CPU 11 pinned with
`taskset`, and the established ODT allocator policy:
`MALLOC_MMAP_THRESHOLD_=33554432` and `MALLOC_TRIM_THRESHOLD_=-1`.

The fixed package is 16,786,287 bytes, contains 200 paragraphs and eight exact
2 MiB deterministic incompressible media resources, and has SHA-256
`097b17ebc8fe811888eba6a9a7d118bc94bb99651880782ccecedf93c003b5ed`.
Every timed iteration replaces paragraph 100. Complete paragraph/media,
manifest, deterministic publication, forward patch, inverse restoration and
stale-source verification remain outside the timer.

Two balanced execution cycles used 100 warmups and 500 measured samples per
leg. Pooling four legs per state gives 2,000 observations per state:

| Media-rich ODT paragraph edit/save | Before | After | Delta |
|---|---:|---:|---:|
| p50 | 5.216 ms | 4.030 ms | **-22.74%** |
| mean | 5.261 ms | 4.074 ms | **-22.56%** |
| p95 | 5.700 ms | 4.475 ms | **-21.48%** |
| p99 | 6.322 ms | 4.982 ms | **-21.19%** |

The approximate independent-sample 95% interval for the mean delta is
`[-22.89%, -22.23%]`. All eight individual legs are stable: control p50 spans
5.141-5.266 ms and candidate p50 spans 4.001-4.080 ms. Every matched
control/candidate comparison improves materially.

Primary raw reports use the `odt-final-handoff-cycle*` prefix. Their hashes and
the frozen binary hashes are indexed in
[`odt-final-handoff-sha256.txt`](../results/odt-final-handoff-sha256.txt).

## Attribution and resources

Matched 30-sample whole-process profiles recorded zero lost samples. Kernel
symbols were restricted, but the relevant userspace frames resolved:

| Exclusive profile frame | Before | After |
|---|---:|---:|
| `__memmove_avx512_unaligned_erms` | 13.62% | 11.98% |
| `transaction::copy_bytes` under `Edit::commit` | 1.27% | absent |

The approximate sampled event count falls 2.56%. The concise profile record is
[`odt-final-handoff-profile.txt`](../results/odt-final-handoff-profile.txt).

Matched whole-process `perf stat` A/B/B/A used 20 warmups and 200 samples per
leg. Pooling both legs per state gives:

| Counter | Before | After | Delta |
|---|---:|---:|---:|
| task clock | 18,290.40 ms | 17,956.40 ms | -1.83% |
| cycles | 89,827,344,964 | 87,883,419,386 | -2.16% |
| instructions | 186,756,632,693 | 186,066,945,287 | -0.37% |
| branches | 18,827,463,806 | 18,732,226,597 | -0.51% |
| branch misses | 39,205,887 | 38,612,985 | -1.51% |
| cache references | 8,128,461,010 | 7,534,357,812 | -7.31% |
| cache misses | 511,869,404 | 462,543,329 | **-9.64%** |
| page faults | 68,504 | 68,491 | -0.02% |
| CPU migrations | 0 | 0 | unchanged |

Heaptrack over ten changed iterations confirms the removed owner:

| Whole-process metric | Before | After | Delta |
|---|---:|---:|---:|
| final `copy_bytes` calls / peak attribution | 10 / 16.79 MB | 0 / 0 | removed |
| allocation calls | 128,535 | 124,085 | **-3.46%** |
| temporary allocations | 22,821 | 20,651 | **-9.51%** |
| peak heap | 106.03 MiB | 107.02 MiB | +0.93% (flat) |
| Heaptrack RSS | 121.65 MiB | 121.41 MiB | -0.20% (flat) |
| leaked bytes | 1.78 KiB | 1.78 KiB | unchanged |

The peak remains dominated by simultaneously required package/publication and
verification owners, so removing a short-lived 16.79 MB allocation reduces
allocation traffic without lowering that global peak. Uninstrumented GNU Time
A/B/B/A reports 112,640 KiB worst-case maximum RSS before and 112,512 KiB
after (-0.11%, flat), with zero major faults.

## Guardrails and correctness

The allocator-matched ordinary ODT guard matrix pools 600 samples per state:

| Guard | p50 delta | Mean delta | p95 delta | Disposition |
|---|---:|---:|---:|---|
| medium open | -0.96% | -0.75% | -4.95% | neutral/better |
| medium one paragraph | +2.77% | +1.29% | -4.24% | within 3% / p95 better |
| medium exact no-op edit/save | -9.38% | -11.12% | -22.40% | sub-microsecond/better |
| medium one edit/save | -7.92% | -8.66% | -12.66% | better |
| medium 1% edit/save | -6.87% | -5.71% | -3.77% | better |
| large open | -3.39% | -3.19% | -2.54% | better |
| large one paragraph | -0.77% | -0.97% | -4.28% | neutral/better |
| large exact no-op edit/save | +60 ns / +2.70% | -1.43% | -4.86% | sub-microsecond/neutral |
| large one edit/save | -6.71% | -5.99% | -5.10% | better |
| large 1% edit/save | -6.76% | -6.46% | -5.78% | better |

The new focused regression constructs a genuinely changed final `Document`,
proves the resulting snapshot is backed by the identical package `Arc`, then
reopens that byte-only snapshot and verifies semantic text. Existing tests
retain exact no-op sharing, raw unchanged media/metadata, deterministic durable
patch/apply/inverse, stale-source refusal, signed/encrypted refusal,
malformed/oversize rejection, compact XML auditing, and structural/resource
operations.

Verification completed:

- all-feature/all-target ODT tests passed: 529 unit tests plus every integration
  and example target;
- warning-denied all-target/all-feature ODT Clippy passed;
- warning-denied ODT rustdoc passed;
- all 32 standalone harness tests and warning-denied all-target Clippy passed;
- the ODF `parse_odt` libFuzzer binary built from its locked manifest (the
  `cargo-fuzz` wrapper is not installed on this host);
- formatting, JSON parsing, artifact hashes and `git diff --check` pass.

The earlier deprecation cleanup remains enforced by these warning-denied gates;
this tranche introduces no deprecated call.

## Remaining work

This handoff removes the final changed-result copy and one redundant parse. It
does not retain a parsed final graph, remove the independent final reopen,
change raw-member publication, add source-backed reads, broaden structural or
resource-edit coverage, or complete real-producer/security matrices. ODS and
ODP ownership paths are unchanged. The next ODF optimization still requires a
distinct measured owner rather than weakening validation.
