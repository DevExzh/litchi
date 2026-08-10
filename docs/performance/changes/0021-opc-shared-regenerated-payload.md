# OPC shared regenerated payload

Date: 2026-08-11

Production base: `5d043f0070d1e5100a1d252b4abfc61262ef413a`

Scope: shared ZIP/OPC publication used by OOXML packages. OLE2, RTF and ODF
production code are unchanged, and iWork/IWA crates were explicitly excluded.

## Hypothesis

The targeted same-topology OPC path already raw-copies every semantically
unchanged ZIP member. Heaptrack nevertheless attributed one 4.19 MiB
allocation to `litchi_opc::pkgwriter::regenerated_action` when the single
changed 4 MiB Part was handed to the ZIP regeneration layer. The immutable
Part already owned those exact bytes, so this complete logical-payload copy
was publication scaffolding rather than required semantic or framing work.

## Change

`soapberry-zip::RegeneratedEntry` can now retain either its existing owned
`Vec<u8>` payload or an additive shared `Arc<Vec<u8>>` payload. Both variants
remain private behind the low-level entry type and expose the same immutable
slice to stored or Deflate publication. Existing callers and generated
content-type/relationship XML continue to use the owned constructor.

Only a changed ordinary OPC Part uses the shared constructor. The package
still builds and audits the same `PublicationPlan`, compares the source and
candidate Part, applies the same dirty-closure rules, compresses the changed
bytes once, preserves unchanged local and central ZIP records, and falls back
before output for topology changes or unsupported layouts. No format facade,
dependency, executor, cache, lock, unsafe code, resource limit, signature
policy or failure-atomicity boundary changes.

## Matched latency measurement

The before release executable SHA-256 is
`953a0a45a34e9b527de222418468bf5dac81695fd70aee72b8512619cc117c77`;
the after SHA-256 is
`e306ca0a56252511abf6fd265581c25f3d0c97663be1855f228db5271afd831b`.
The environment is release Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD
EPYC 9575F VM, Rust system allocator, `perf_event_paranoid=1`, and CPU 2 pinned
with `taskset`.

The primary run used 20 warmups and 200 samples per leg in before-A, after-A,
after-B, before-B order. The table pools 400 raw samples per state. Every
operation changes the middle Part, publishes through the bounded sequential
sink, and verifies the resulting package and target payload outside timing.

| Targeted OPC mutation | Before p50 | After p50 | p50 delta | p95 delta | Mean delta |
|---|---:|---:|---:|---:|---:|
| Few-large, compressible | 1.342 ms | 1.063 ms | **-20.73%** | -6.01% | **-18.49%** |
| Few-large, incompressible | 60.123 ms | 59.569 ms | -0.92% | -0.51% | -0.86% |
| Many-small, compressible | 177.736 us | 180.537 us | +1.58% | +1.22% | +0.99% |
| Many-small, incompressible | 204.521 us | 201.837 us | -1.31% | +3.00% | -0.64% |

The approximate independent-sample 95% interval for the few-large,
compressible mean delta is `[-20.16%, -16.82%]` of the before mean. The
few-large incompressible interval is `[-1.11%, -0.61%]`. Output bytes and sink
summaries match within each corpus.

Primary reports and SHA-256 digests:

- `abba-opc-shared-regeneration-primary-before-a.json`:
  `7160565cc8eeed97ba3f4e0cd6cf7eadac8f2294bb15d8adfb0a41adce4376f3`
- `abba-opc-shared-regeneration-primary-after-a.json`:
  `0060e57c230be73e67b707e01441a32611d827559d500491e04d0f5b067fc0a9`
- `abba-opc-shared-regeneration-primary-after-b.json`:
  `81e004ca33f1203e5bcb5e34026599ff61fd612648830ef83822e4c43c247f0b`
- `abba-opc-shared-regeneration-primary-before-b.json`:
  `4952af11ad39be1845ea0c9e8d84a9f5cf9591314faffb2d2eaa5707713c21c3`

## Guardrails

A four-leg 30-warmup/500-sample targeted-save matrix over tiny and 2,048-Part
wide-root shapes is clean: pooled p50 deltas range from -2.23% to -0.83%, mean
from -4.79% to -1.11%, and p95 from -5.68% to +2.07% across both payload
families.

The unchanged exact-source path cannot reach the new constructor. A broad
20-warmup/300-sample no-op matrix stayed below the 5% p50 review threshold;
one many-small mean moved +5.85%, which triggered a dedicated run. That run
used 200 warmups and 10,000 samples per leg. Pooled many-small p50 is unchanged
for compressible input and +1.44% for incompressible input; means are -1.85%
and +3.00%. Tiny no-op means move by +3.45 ns and +8.76 ns, with large relative
percentages at the 30-40 ns timer floor. Because the changed publication path
is unreachable and absolute movement is below 10 ns, this is disclosed as
non-actionable timer/code-layout noise rather than hidden or generalized.

The raw guard reports are the
`abba-opc-shared-regeneration-{edge,noop,noop-small}-*.json` files in
`results/`.

## Profile, counters and memory

Matched one-shot Heaptrack on the few-large incompressible case removes the
single 4.19 MiB allocation under OPC `regenerated_action`. Allocation calls
move from 876 to 875 and peak heap from 122.49 to 118.30 MB (-4.19 MiB,
-3.42%). Temporary allocations move from 164 to 165. The remaining 4.20 MiB
profile bucket belongs to the separate generated local-entry/compression
buffer and is not claimed as removed.

Instrumented RSS is noisy at 106.06/126.32 MB. Uninstrumented GNU Time ABBA
maximum RSS is 116,236/116,236 KiB before and 116,496/116,236 KiB after; the
maximum-to-maximum delta is +0.22%, treated as flat.

Matched `perf stat` ABBA processes used 20 warmups and 500 few-large,
compressible samples per leg:

| Process-wide counter | Before A+B | After A+B | Delta |
|---|---:|---:|---:|
| task clock | 1,578.240 ms | 1,245.490 ms | -21.08% |
| cycles | 7,594,283,017 | 6,120,506,986 | -19.41% |
| instructions | 19,392,504,615 | 19,183,792,398 | -1.08% |
| branches | 2,466,746,330 | 2,448,931,035 | -0.72% |
| branch misses | 2,227,714 | 2,219,097 | -0.39% |
| cache references | 2,521,949,717 | 2,193,278,928 | -13.03% |
| cache misses | 79,620,188 | 54,844,317 | -31.12% |

Matched sampled profiles also reduce the relative `memmove` share from 22.76%
to 11.01%; compression and comparison frames consequently occupy a larger
fraction of the smaller after process. Heaptrack provides the direct
allocation-site attribution.

Raw evidence is in `opc-shared-regeneration-*-heaptrack.txt`,
`opc-shared-regeneration-*-perf-report.txt`,
`opc-shared-regeneration-perf-stat-*.csv`, and
`opc-shared-regeneration-time-*.txt`.

## Correctness verification

- all-feature `soapberry-zip` passed 183 unit tests and 30 doctests (one
  doctest ignored), including shared-payload retention;
- all-feature `litchi-opc` passed 125 unit tests, 13 integration tests and five
  doctests, including exact framing, dirty closure, fallback, bounded sink and
  complete reopen/readback coverage;
- both fuzz targets and their production dependency graphs compile offline;
- warning-denied all-target/all-feature Clippy and warning-denied rustdoc pass
  for both touched crates;
- the unchanged benchmark harness passes 23 tests and warning-denied Clippy;
  and
- formatting, JSON parsing and `git diff --check` are final commit gates.

The repository-wide all-target/all-feature check reached unrelated packages
but exhausted the build volume while writing incremental artifacts. It is not
counted as a passing umbrella gate; the focused touched-crate checks above are
the applicable result.

## Next non-iWork audits

1. OLE2: measure reuse of the validated PPT Presentation document bytes when
   capturing root slide order, avoiding a second CFB ingress only if exact
   patch/inverse and final public-reader proofs stay intact.
2. ODF: profile moving parser-created block strings into full-text output;
   do not revive either rejected package-adoption candidate.
3. RTF: extend the benchmark to compressed, legacy-code-page and real-producer
   inputs before considering another parser specialization.
4. OOXML: separately attribute the remaining generated-local-entry buffer;
   do not split publication ownership without a matched profile and the same
   framing/fallback gates.

iWork remains deferred while the `iwa-*` crates are modified independently.
