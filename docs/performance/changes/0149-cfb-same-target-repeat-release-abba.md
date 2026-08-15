# Change 0149: CFB same-target repeat release ABBA

Date: 2026-08-16

Status: scoped configured-simulator aggregate repeat result accepted; local,
per-invocation, bulk, concurrent, and resource claims withheld.

## Compared revisions and matrix

This record measures the target-aware MiniFAT repeat policy from
[change 0148](0148-cfb-same-target-repeat-policy.md) against its immediate
pre-change production state while keeping the harness source identical. The
clean candidate is `4332f87cd68048473f886f1b50ef35c629fd8806`. The clean
control is temporary measurement commit
`643a8997b8b3689330631dd242de4f106cefe82a`, made from the candidate by
restoring only `crates/litchi-cfb/src/shared.rs` and
`crates/litchi-cfb/src/shared_bulk.rs` from
`7aaf81a8488dd16fa93d9f653fc1ec0e341878c9`. The shared harness has SHA-256
`9646a0506f7917046badb88589768d43a350e557b1e252684ffdc60617f30b50`.

The candidate `shared.rs` / `shared_bulk.rs` hashes are
`872bc0159d39ccf2ceb4a3a5c1267aadf55d493f8330996e5afe4bbc3d6f0a62` /
`3ea1cccdbdd4b45f3801983bfa8df67b94defe5476834ebc6175011760bb5b48`.
The control hashes are
`ae74089e4abdf672b9ec1c3010db007cbc9ba9a3736de1ac66e21b62781661b9` /
`90ff6465d904cf80bac93ebff42320724f27d329555beabd946595fb28af4ef8`.
The Rust 1.95.0 release binaries are:

- candidate: 40,286,592 bytes, SHA-256
  `e8e026b67a088b15fb31338008f5bdfcd072254a8427fcc6de127daeb86684f0`;
- control: 40,278,840 bytes, SHA-256
  `ae555c8d1b819e8ea8c8535cf854b19b8089545a50c7ef3263a0290d54b59943`.

Four fresh CPU-2 release processes ran in strict
`A1 control, B1 candidate, B2 candidate, A2 control` order. Every leg used 20
warmups and 200 samples for each of 36 records: 18 selectors across
`many-small` and `wide-root`. The retained matrix therefore contains 28,800
samples. All four reports identify clean worktrees, CPU affinity `2`, one
affinity-visible CPU, the same deterministic corpus identities, exact output
hashes, stable source versions, and typed missing-stream refusal.

The configured simulator uses 100 us fixed latency, 25 us request overhead,
50 MiB/s bandwidth, and a 4 KiB maximum physical request. It is a deterministic
harness model, not a network, storage device, syscall, cold-cache, or physical
I/O observation.

## Exact source-work policy

Let `L` be the selected target size, `R` the declared root Mini Stream size,
`D = [target_start,L,L]`, and `C = [512,R,R]` before simulator chunking. The
raw reports preserve every actual `(offset, requested, returned)` event.

```text
same-target repeat-3/8: control   [L,R,0...]
                        candidate [L,L,...]

different-SID A-B-A:   both      [L,R,0]
public bulk A-B-A:      control   aggregate {D,C}
                        candidate aggregate {C}
overlapping same target: control  aggregate {D,C}
                         candidate permitted aggregate {D,D} or {D,C}
one-shot:               both      [L]
```

For the measured `n`-call same-target repeats where `n` is 3 or 8, the control
reads `L + R` logical range bytes and the candidate reads `n * L`; the avoided work is
`R - (n - 1) * L`, excluding mandatory parser-open metadata. The candidate
therefore avoids root Mini Stream materialization for sequential
same-SID repeats. It does not turn later calls into zero-source-work cache hits:
each remains one exact target-sized positional read. Different targets and a
multi-MiniFAT public batch retain cache takeover. An overlapping caller that
observes an in-flight direct read also requests takeover; scheduler ordering can
instead complete two direct reads, producing the documented `{D,D}` outcome. The
root cache allocated by the public bulk path remains outside the batch-local
`Resource::Memory` reservation, so this record makes no bounded-resident-memory
claim.

The resulting logical per-operation range-byte reductions are:

| Shape / target | Repeat-3 | Repeat-8 |
|---|---:|---:|
| many-small / 36 | 99.959% | 99.890% |
| many-small / 4,095 | 95.438% | 87.836% |
| wide-root / 36 | 99.995% | 99.986% |
| wide-root / 4,095 | 99.416% | 98.443% |

These ratios exclude mandatory parser-open work and are not physical-I/O
reductions.

The exact aggregate direct/cache examples are:

| Shape / target | `D` | `C` |
|---|---|---|
| many-small / 36 | `[261632,36,36]` | `[512,261184,261184]` |
| many-small / 4,095 | `[261632,4095,4095]` | `[512,265216,265216]` |
| wide-root / 36 | `[2096640,36,36]` | `[512,2096192,2096192]` |
| wide-root / 4,095 | `[2096640,4095,4095]` | `[512,2100224,2100224]` |

These are logical positional-source events. The simulator may split `C` at its
4 KiB configured ceiling; the logical requested/returned totals remain exact.

## Accepted configured-simulator aggregate result

The table reports candidate improvement percentages for checked aggregate
`total_ns = open_ns + operation_ns` in the two adjacent ABBA directions.
Positive is faster. Percentiles use midpoint p50 and nearest-rank p95/p99.

| Operation / shape / target | Total p50, A1->B1 / B2->A2 | Total p95 | Total p99 | Total mean |
|---|---:|---:|---:|---:|
| repeat-3 / many / 36 | 61.47% / 61.55% | 61.52% / 61.78% | 61.74% / 62.01% | 61.49% / 61.59% |
| repeat-3 / many / 4,095 | 60.70% / 60.70% | 60.90% / 60.63% | 61.24% / 60.40% | 60.73% / 60.67% |
| repeat-3 / wide / 36 | 64.09% / 64.01% | 64.02% / 64.11% | 64.09% / 64.11% | 64.10% / 64.02% |
| repeat-3 / wide / 4,095 | 63.85% / 63.69% | 63.89% / 63.68% | 63.90% / 63.66% | 63.85% / 63.69% |
| repeat-8 / many / 36 | 58.19% / 58.15% | 58.42% / 58.17% | 58.34% / 58.03% | 58.23% / 58.14% |
| repeat-8 / many / 4,095 | 55.92% / 55.86% | 56.12% / 55.90% | 55.83% / 56.16% | 55.93% / 55.87% |
| repeat-8 / wide / 36 | 63.67% / 63.57% | 63.76% / 63.52% | 63.80% / 63.25% | 63.67% / 63.56% |
| repeat-8 / wide / 4,095 | 63.16% / 63.16% | 63.14% / 62.87% | 63.42% / 62.74% | 63.16% / 63.12% |

Both directions agree for all eight named aggregate repeat cells. The isolated
aggregate operation interval improves about 95.4-99.6% for repeat-3 and
87.2-98.9% for repeat-8. The accepted claim is limited to this simulator
configuration, corpus, host, build, and aggregate operation shape.

The four configured-simulator one-shot controls remain near neutral: total p50
changes range from -0.37% to +0.25% and total mean from -0.35% to +0.20% in the
two directions. That is the required no-one-shot-regression control; it is not
a new one-shot speedup claim.

## Tradeoffs and withheld local results

Aggregate improvement must not be described as improvement for every
invocation. The control's third and later calls are zero-source cache hits,
whereas the candidate performs an exact `L` read on every call. The later
per-invocation simulator intervals consequently regress by orders of magnitude
in some cells. This is the intended policy tradeoff and is retained in the raw
reports; no per-invocation speed claim is accepted.

Local in-memory repeat totals are positive at p50 in both directions, ranging
from about 2.1% to 40.9%, but their tails and same-implementation controls are
too unstable for a generic local claim. The special workloads also retain
explicit review triggers:

- wide-root / 4,095 different-SID A-B-A has a first-direction total p99
  regression of 8.19% despite mostly near-neutral aggregate totals;
- wide-root / 4,095 public bulk reports first/paired-direction total changes of
  -7.86% / -0.20% p50, -15.55% / +5.42% p95, -21.08% / +16.46% p99, and
  -8.83% / +1.06% mean; same-candidate drift is also material;
- overlapping local calls include wide-root / 4,095 first-direction total p95
  23.75% and p99 29.94% regressions, plus a many-small / 36 paired-direction
  p99 regression of 16.01%;
- local bulk and concurrent same-implementation total drift reaches 20.19% and
  23.20%, respectively.

These results are not averaged away. They withhold local wall-clock, bulk, and
concurrent acceptance and require a more isolated scheduler/resource study
before any such claim.

## Artifacts and claim boundary

The [machine-readable summary](../results/cfb-repeat-abba-0149-summary.json),
SHA-256
`70ff7c1057ee4fb592ddb6397a51e4c8741486a2b6b8707747a402a680397613`,
contains all 36 record statistics per leg, both paired improvements, source
vectors, corpus/environment/configuration identities, revisions, binary
identities, review triggers, and claim decisions. Complete raw reports are:

- [A1 control](../results/cfb-repeat-a1-control-0149.json.zst), compressed
  SHA-256 `032767adabd49f215db3a269c8123fb2ab78e842486d6397f10f7f82c1b2492a`,
  518,860,689 bytes after decompression, raw SHA-256
  `e407e3e235c5361407b3574571077ee269750ec21d3ac0107228ee6a97081a9d`;
- [B1 candidate](../results/cfb-repeat-b1-candidate-0149.json.zst), compressed
  SHA-256 `612d21401033161fc4755cd152d56309d7bd9792f96ec74a26b3d595b41e989b`,
  258,082,768 bytes after decompression, raw SHA-256
  `61cf5336cd33ec53c56b6b9212264a19bd593f8d0bd2143b17c2e68e47511dee`;
- [B2 candidate](../results/cfb-repeat-b2-candidate-0149.json.zst), compressed
  SHA-256 `d855bdf3bcc38aef1f5ebe5f3325d56cf89c047dd3c86241335261c87dfb26c8`,
  258,083,843 bytes after decompression, raw SHA-256
  `44770db7c8928e5fd366981cb082b6b4e597c365a136e2415b484a389a927b26`;
- [A2 control](../results/cfb-repeat-a2-control-0149.json.zst), compressed
  SHA-256 `a42103face385439ec94ce8a21ac7213518bb8f7c776218e0c0f32d7a1152c5b`,
  518,861,031 bytes after decompression, raw SHA-256
  `f3e8297e58ee746c590eb0c8332c8d4e2c2a5ea69930b4c51a6b3e5b18ec3555`.

No generic/local wall-clock, per-invocation, bulk, concurrent, allocation,
allocated-byte, RSS, peak-memory, physical-I/O, cold-cache, remote/network,
device, decompression, native DOC/XLS/PPT, OOXML, ODF, RTF, or iWork result is
claimed. Failure/retry, ineligible-root, FAT, native semantic, and complete
resource-accounting matrices remain open.
