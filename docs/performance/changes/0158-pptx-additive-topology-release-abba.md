# Change 0158: PPTX additive-topology release ABBA

Date: 2026-08-16

Status: accepted for deterministic canonical generated PPTX cross-presentation
slide copy at the prepared in-memory plan/commit/publication boundary. No
source-backed, cold-I/O, decompression/recompression-byte, general OPC/PPTX,
real-producer, or iWork claim is accepted.

## Scope

Candidate `d900ae63328aef0e58678cc3e2d55aec28612e34` extends the owned-source
ZIP/OPC publication path used by PPTX cross-presentation slide copy. It can
retain complete source topology while appending generated Parts and can omit
only an exact physical suffix when applying the inverse. Unchanged physical
members are raw-copied; changed and appended members are generated. Control
`e8a67b19e5be950a77431ef664b46e130b2db90f` uses the prior complete OPC
rewrite for this topology-changing operation.

The existing opt-in selectors are unchanged:

- `pptx_cross_copy_plain`; and
- `pptx_cross_copy_media_rich`.

The harness and its locked dependency graph are byte-identical between the
two revisions. The matrix therefore remains 301 selectable names and the
historical default remains 36 cases / 198 records.

Each retained sample starts from an already constructed owned source
presentation, destination presentation, and copy selection. Top-level
`elapsed_ns` is the checked arithmetic sum of plan, atomic commit, and
publication. The harness records those three phases separately and retains
diagnostic reopen timing outside the total. Corpus generation and
source/destination setup, reopen, semantic/topology/dependency checks, durable
forward/inverse replay, stale and foreign-source refusals, and output
verification are outside the accepted operation interval.

The plain corpus is a 30,539-byte, 41-member package. The media-rich corpus is
a 16,814,664-byte, 49-member package containing eight deterministic 2 MiB
incompressible media payloads. Both copy source slide 2 into destination slide
1 at position 1.

## Correctness boundary

Every leg retained identical corpus hashes and passed:

- semantic output and package-topology reopen;
- finite dependency closure and collision remapping;
- source immutability;
- durable patch forward/inverse round trip;
- stale source, stale destination, foreign source, and borrowed-provenance
  refusal; and
- exact sink byte count and deterministic per-revision output digest.

Control and candidate output digests intentionally differ because the
candidate preserves source member framing/order while the control rebuilds
the complete package. Semantic content, member topology, and dependency gates
match. The separate production tests at the candidate revision cover EOCD
comments, data descriptors, unknown physical members, raw untouched local and
central records, inverse restoration, partial-sink progress, and typed
unsupported-layout refusal. Those adversarial raw-preservation tests are not
properties of this canonical generated benchmark corpus.

The complete `soapberry-zip`, `litchi-opc`, and `litchi-pptx` suites passed at
the candidate revision, as did warnings/deprecations-denied Clippy and
rustdoc. Two independent read-only implementation reviews returned SAFE for
the bounded additive/suffix-removal contract.

## Clean release ABBA

The release binaries were built from clean detached worktrees. Control binary
SHA-256 is
`445cd47a3e8d1b8d64cdde29543543106676bb6a99466afd6c10f318bcf1b2ef`;
candidate binary SHA-256 is
`f399911eaf09ff07c8903cba1ec2c8375fd2d2c9dc5e19a300b6415a315ec16f`.
Both record harness SHA-256
`a95ed354b5c3bef06291147e6f674eb118bf832f3c2ff019850b2ce5df54f2dd`
and lockfile SHA-256
`8a4f02cdf936c7c3456984406b71d49b6461de3a15811618fdc0fdb2b5430344`.

The strict order was `A1 control, B1 candidate, B2 candidate, A2 control`,
pinned to CPU 2. Each leg used 20 warmups and 200 retained samples per case,
for 1,600 retained total observations. The host reported Linux
6.8.0-101-generic, AMD EPYC 9575F, Rust 1.95.0, and the system allocator.

### Total operation

| Corpus | A1 control p50 | B1 candidate p50 | B2 candidate p50 | A2 control p50 | A1 -> B1 | A2 -> B2 |
|---|---:|---:|---:|---:|---:|---:|
| Plain | 10.874 ms | 7.651 ms | 7.805 ms | 10.576 ms | 29.643% | 26.196% |
| Media-rich | 1,933.821 ms | 1,096.585 ms | 1,093.719 ms | 1,939.342 ms | 43.294% | 43.604% |

The total-operation distribution agrees in both directions:

| Corpus | P95 A1 -> B1 / A2 -> B2 | P99 A1 -> B1 / A2 -> B2 | Mean A1 -> B1 / A2 -> B2 |
|---|---:|---:|---:|
| Plain | 33.973% / 24.938% | 32.357% / 26.511% | 30.328% / 26.072% |
| Media-rich | 43.668% / 43.440% | 44.165% / 43.312% | 43.397% / 43.615% |

Same-implementation total p50 drift is 2.744% for the plain control, -2.022%
for the plain candidate, -0.286% for the media-rich control, and 0.261% for
the media-rich candidate. All total-operation review thresholds pass.

### Phase attribution

Media-rich publication p50 improves 49.321% and 49.680% in the two pair
directions; p95, p99, and mean each improve by 48.631%-49.840%. Media-rich
plan p50 improves 40.115%/40.456%, and commit p50 improves
42.499%/42.577%. Diagnostic reopen is neutral to slightly slower
(-0.598%/-0.214% p50), so no reopen improvement is claimed.

Plain publication p50 improves 82.798%/82.304% and mean improves
82.765%/81.784%. Candidate same-implementation publication drift reached
10.443% at p95 and 26.670% at p99, however. Plain publication p95/p99 claims
are therefore withheld. The total-operation result remains accepted because
its paired directions agree and its own drift checks stay below threshold.

## Process-wide resource sidecars

The following observations are descriptive whole-process measurements, not
operation-local attribution.

- Matched `/usr/bin/time -v` media-rich runs used 3 warmups and 30 retained
  samples. Control wall time was 73.04/73.06 seconds and candidate wall time
  41.35/41.18 seconds. Candidate maximum RSS was 0.486%/0.480% higher
  (839,244/839,372 KiB versus 835,184/835,360 KiB), within the 5% neutral
  boundary. Filesystem inputs were zero and major faults were zero; this is
  not physical-I/O evidence.
- Matched `perf stat` media-rich runs used 3 warmups and 10 retained samples.
  Candidate task-clock falls 42.399%/43.122%, cycles 42.583%/43.116%, and
  instructions 46.686%/46.775%. Branches and branch misses fall about 50%; the
  process-wide cache-miss reduction is only 8.788%/8.693%.
- One clean-process Heaptrack capture per revision records 463,704 -> 452,589
  allocation calls (-2.397%), 115,530 -> 113,777 temporary allocations
  (-1.517%), and effectively unchanged peak heap (850.25 -> 850.09 MiB).
  Heaptrack peak RSS is 0.465% higher for the candidate. These single captures
  are descriptive and cannot establish operation-only allocation or memory
  savings.

## Artifacts

The [machine-readable summary](../results/pptx-additive-topology-abba-0158-summary.json)
contains raw samples, exact statistics, confidence intervals, per-phase
drift/paired deltas, gates, output/sink evidence, profiler summaries, source
identities, and claim boundaries. Its SHA-256 is
`5a838728b66246cc1f0e08df42cf46c046f2ac0c6e4a2c72b58c05fbdd8a7e8f`.

The [artifact manifest](../results/pptx-additive-topology-0158.sha256) binds
the summary, four compressed raw reports, matched time/perf raw and parsed
sidecars, and the two valid compressed Heaptrack captures. The mistakenly
profiled wrapper-only Heaptrack attempts are excluded.

## Accepted boundary and remaining gaps

For these two deterministic canonical generated owned-source PPTX slide-copy
workloads, the candidate reduces the named total prepared operation by
24.9%-44.2% across p50/p95/p99/mean in both ABBA directions (the p50 range is
26.2%-43.6%). For the
approximately 16 MiB media-rich workload, prepared publication improves by
48.6%-49.8% across the same statistics. Plain publication p50/mean is accepted
but its tail claim is withheld.

This does not establish source-backed or borrowed-source topology mutation,
end-to-end filesystem save, cold or remote I/O, physical bytes read,
decompressed/recompressed bytes, memory-copy volume, energy, concurrency or
scaling, charts/notes/themes/macros/signatures, real-producer behavior, general
OPC/PPTX CRUD, other formats, or iWork. The implementation still reads/copies
the complete owned source and the ordinary OPC model remains eager. Those are
separate optimization and evidence tranches.

## Reproduction

Build locked release binaries from the two named clean revisions, pin to one
CPU, and execute both selectors in strict four-leg order:

```sh
taskset -c 2 litchi-perf-baseline-control \
  --warmup 20 --samples 200 \
  --case pptx_cross_copy_plain,pptx_cross_copy_media_rich \
  --json pptx-additive-topology-a1-control.json

taskset -c 2 litchi-perf-baseline-candidate \
  --warmup 20 --samples 200 \
  --case pptx_cross_copy_plain,pptx_cross_copy_media_rich \
  --json pptx-additive-topology-b1-candidate.json

# Repeat candidate for B2, then control for A2.
```
