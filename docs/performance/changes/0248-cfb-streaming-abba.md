# Change 0248: CFB streaming release ABBA evidence

Date: 2026-08-21

Status: the performance claim is held because the retained ABBA workload is
mis-scoped for the production candidate and cannot adjudicate it. The
unmeasured candidate was rolled back in `67a37235c` pending applicable
direct-payload-read or native-consumer evidence.

Claim registry ID: `claim-0248-cfb-streaming`

## Evidence-package identity

The independently audited package is tracked with its
[`manifest`](../results/0248-cfb-streaming-20260821/0248-cfb-streaming-abba-manifest.json).
The historical external staging path was
`/home/zhuhe/CodeProjects/litchi-perf-evidence/0248-cfb-streaming-20260821/`;
it is provenance only and is not required to verify the repository copy. The
manifest SHA-256 is
`7ffd58b3252879713697244972284a44a40244f47090d8a5aa0ac07a709314ee`.
The `summary.json` file SHA-256 is
`e652f3923d923579c86eab764958fb708ab60c1a9a732d66168cfbcf81e32f72`; its
canonical SHA-256 is
`ef2ce9ab577b959bd4bcc650e20d011a4f6e9a877cf2feefe71782d5196dfd9c`.
The package is schema version 1, has change ID
`0248-cfb-streaming-abba`, and contains 24 result rows from
`litchi-perf-abba-summary` 0.1.0.

The retained package members have these identities:

| Member | File SHA-256 | Canonical SHA-256 |
|---|---|---|
| `0248-cfb-streaming-abba-manifest.json` | `7ffd58b3252879713697244972284a44a40244f47090d8a5aa0ac07a709314ee` | — |
| `summary.json` | `e652f3923d923579c86eab764958fb708ab60c1a9a732d66168cfbcf81e32f72` | `ef2ce9ab577b959bd4bcc650e20d011a4f6e9a877cf2feefe71782d5196dfd9c` |
| `a1.json.zst` | `bf3620c11d8511c4a74660f633fcafcc43d2d87336c7ca61c25d55ecd292fbc7` | `0885267e653a22b7224bbfa34c50cb8342f7dd0fa5b8360e32dc6ace27ffe071` |
| `a2.json.zst` | `3236c61e7eeed8119c253695fd41bfc9aef7a697de955508304423251ebd4ca7` | `3d9291e3257acfd6b021a9fcbe73289ad5559a9b72a771fc0a13f5bdb170e497` |
| `b1.json.zst` | `90a41b8e25a22fc7724d24cecf612a784caf14f7f641585d8b97edb0cb884d05` | `d0d43af91c8d16f46ace40fe6e4ce51029db36d16877309549a65ce88f90384b` |
| `b2.json.zst` | `7af813790b1bbf446f893876ed3eed270d50e0dc077c45330d6e037930cdd9d0` | `c8ce0b1b827e7972e9f668b86119d28c686ce01b9b07dc9e5f234580c0d102e9` |

The control revision is `55ff6fea74456c2ef6a16861cf58e41dc927b499`
(A1/A2); the candidate revision is
`97361c8505f3b7e3d02cf15f89770f571bb858da` (B1/B2). The package records
clean worktrees, CPU affinity 2, one visible logical CPU, Rust 1.95.0, and
the AMD EPYC 9575F host for every leg.

## Protocol and workload

The order is A1 control, B1 candidate, B2 candidate, A2 control, with 30
warmups and 500 retained samples per result row. The 24 rows cover the 12
public CFB MiniFAT streaming selectors over the `many-small` and `wide-root`
corpora:

```text
cfb_open_stream_mini_shared_{one_shot,repeat,repeat8,different_sid,bulk,concurrent}
cfb_open_stream_mini_4095_shared_{one_shot,repeat,repeat8,different_sid,bulk,concurrent}
```

These selectors exercise `SharedOleFile::open_stream` and the positional
read paths in `shared.rs`. The measured candidate instead changed direct
`OleFile` stream replay in `file.rs`. `SharedOleFile` uses `OleFile` only for
structural parsing/index conversion in this workload, so the measured rows do
not exercise the changed FAT/MiniFAT payload-read implementation. This scope
mismatch was discovered during the current-code claim audit and makes the
latency classification below descriptive of the frozen binaries/workload,
not evidence for or against that production candidate.

Drift ceilings are 5%/5%/10%/15% for p50/mean/p95/p99. The summary recomputes
statistics from samples and verifies case/corpus, configuration, and stable
environment identity. Positive readings would mean lower candidate elapsed
time; they are not treated as a claim unless the paired-direction and drift
gates pass.

## Latency classification

Across the 24 result rows, the exact accepted and adverse-both counts are:

| Statistic | Accepted rows | Adverse-both rows |
|---|---:|---:|
| p50 | 1 | 14 |
| mean | 2 | 10 |
| p95 | 2 | 7 |
| p99 | 2 | 4 |

The seven accepted cells are limited to these rows and paired readings
(A1→B1 / A2→B2, percent lower):

| Selector / shape | Statistic | A1→B1 | A2→B2 |
|---|---|---:|---:|
| `cfb_open_stream_mini_shared_different_sid` / wide-root | p50 | 2.5967795683374484% | 0.4440299132528637% |
| same | mean | 3.3016804461648195% | 0.5290179478580318% |
| same | p95 | 7.056654133189526% | 1.3493939393939394% |
| `cfb_open_stream_mini_shared_one_shot` / wide-root | p99 | 14.224565651869762% | 20.27806870084924% |
| `cfb_open_stream_mini_shared_repeat` / wide-root | p95 | 1.682570929486742% | 1.8317263698642239% |
| `cfb_open_stream_mini_shared_repeat` / many-small | mean | 1.4146271275085123% | 3.708567390902115% |
| `cfb_open_stream_mini_shared_repeat8` / many-small | p99 | 4.754913445268776% | 4.887215751356035% |

All remaining cells are rejected by adverse paired directions, disagreement,
or drift. These classifications remain an exact description of the retained
package, but the workload mismatch means neither the accepted cells nor the
adverse-both population can adjudicate the direct-`OleFile` optimization.

## Decision and landed oracle boundary

Do not land the production CFB streaming optimization on this latency run.
The run does not exercise the changed payload-read path, so it is neither a
latency approval nor a valid latency rejection for that code. Because no
applicable evidence for this candidate's direct payload-read/native-consumer
path was retained, policy rollback `67a37235c` restores the validated control
implementation. A future candidate requires an applicable workload before it
can land; no aggregate CFB speedup or regression is claimed here.

The independent actual-output oracle did land as
`66bb83abbc4e7259ff66e83f5b911d94dca4fd40` (`66bb83abb`)
(`perf(cfb): hash actual bulk and concurrent outputs`). That oracle computes
hashes from the actual returned buffers while retaining independent expected
payload checks. Its landing is a correctness/evidence-harness result, not a
decision to land the latency candidate measured here.

## Source-identity measurement projection

The summary uses the source-identity measurement projection fix
`cd21f7670e811493c813b18320157219823fa2e8` (`cd21f7670`,
`fix(perf): separate CFB source metrics from identity`). For
`source.cfb_open_stream`, the summarizer removes timing and measurement
vectors—logical/open/operation/total times, read counts and bytes, range
vectors, and root-cache bytes—before comparing identity. It retains source
identity fields such as the expected payload hash, stream/corpus identity, and
the public `source_version` fence. This prevents measured per-leg counters
from falsely changing source identity; it does not turn measurements into
identity or create output identity.

Accordingly, the package reports source identity verified equal for all 24
rows, while operation metrics, sink identity, and top-level output identity
remain consistently absent. The independent actual-output oracle above must
not be read as a retroactive output-identity claim for this package.

## Claim boundary

This note is generic `SharedOleFile` CFB/OLE2 substrate evidence only. It makes
no claim about direct `OleFile` payload reads, native DOC/XLS/PPT semantics or
CRUD, source-backed filesystem behavior, physical I/O, decompression,
cold-cache behavior, allocations, RSS, or a general CFB performance
improvement. Applicable direct-path allocation and latency evidence remains a
separate prerequisite for any future version of the rolled-back candidate.

The evidence correction is checked by Python package/hash/count and link
validation plus `git diff --check`. The production rollback separately passed
the complete `litchi-cfb` test suite, focused strict Clippy/rustdoc, formatting,
and independent review; no new benchmark result is introduced by the rollback.
