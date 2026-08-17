# Change 0172: immutable CFB numeric-plan publication

Date: 2026-08-17

## Decision

Retain the narrow immutable-provenance publication path for native XLS
plan-only fixed-width Number/RK/MulRK edits. `SharedOleFile::open_owned` accepts
an `Arc<[u8]>` and retains it in a private CFB adapter. That constructor is the
only production way to mark a source immutable; arbitrary `ReadAt` adapters
remain on the generic path.

Only direct sequential `ValidatedOverlayPlan::write_to` is specialized. An
owned plan skips the redundant complete pre-emission and post-emission
fingerprint scans. It still reads the complete source in 64 KiB publication
chunks, hashes both source and target during emission, checks exact read and
write progress, reports partial output, and flushes the sink. Checked composed
views retain their preflight. Atomic `save` retains its initial and pre-rename
complete fingerprint fences plus emission hashing.

The sole format opt-in is `SourceBackedPlanCommit` for XLS fixed-width numeric
plans. Ordinary source-backed commits, eager commits, XLS comments and
visibility, DOC/PPT consumers, and every generic CFB positional source keep
their previous behavior.

## Deterministic work reduction

Direct publication formerly performed:

1. a complete source/target fingerprint preflight;
2. one 64 KiB emission pass with source and target SHA-256; and
3. a complete source/target fingerprint postflight.

Sealed immutable ownership makes the two outer source scans redundant while
the emission proof remains. Per effective publication this removes exactly two
complete logical source scans, `2 * ceil(artifact_bytes / 1,048,576)` logical
`ReadAt` calls, and two source/target SHA-256 pairs.

- Number, 16,995,840-byte artifact: 33,991,680 logical bytes and 34 reads.
- RK/MulRK, 202,752-byte artifact: 405,504 logical bytes and two reads.

These are code-derived in-memory work counts. The harness source counters
describe owned-source ingress and do not observe the internal fingerprint
passes, so no physical-I/O claim is made.

## Correctness gates

- CFB overlay tests pass 21/21, including exact public owned-Arc publication,
  one-pass owned direct publication, unchanged three-pass generic publication,
  composed-view preflight, mutation refusal, partial sinks, and retained
  atomic-save fences.
- OLE-common source-backed overlay tests pass 3/3 and verify identical protected
  component refusal for generic and owned ingress.
- The five XLS plan-only numeric tests pass, and the complete XLS library suite
  passes 1,015/1,015.
- Strict production Clippy with `-D warnings -D deprecated`, locked checks,
  formatting, diff, and the 64-package boundary audit pass.
- The all-target runs expose only two control-reproducible failures outside the
  patch: CFB's hostile temporary-file substitution test and the existing XLS
  writer-encryption family mismatch. OLE-common all-targets passes completely.
- Two independent frozen-tree reviewers returned SAFE. They confirmed the
  provenance boundary, generic mutation fences, stream-move inverse downgrade,
  common protection guard, bounded publication, and atomic-save behavior.

## Clean release ABBA

Control `fbb1dbbd3` and candidate `0db48c38c` were built from clean worktrees
with locked offline release dependencies. Their binary SHA-256 digests are
`9a7599cf...` and `1f890121...`. Every run was pinned to CPU 2 on the AMD EPYC
9575F host, which exposed one logical CPU to the process. The retained order
was strict `A1 control, B1 candidate, B2 candidate, A2 control`, with 20
warmups and 500 samples for each existing plan-only selector.

Positive values below mean lower candidate latency:

| Selector / phase | Statistic | A1 -> B1 | B2 -> A2 | Decision |
|---|---:|---:|---:|---|
| Number / complete workflow | p50 | 38.86% | 38.38% | accepted |
| Number / complete workflow | mean | 38.80% | 38.33% | accepted |
| Number / complete workflow | p95 | 38.40% | 38.03% | accepted |
| Number / complete workflow | p99 | 39.00% | 37.54% | accepted |
| Number / direct publication | p50 | 65.63% | 65.36% | accepted |
| Number / direct publication | mean | 65.60% | 65.32% | accepted |
| Number / direct publication | p95 | 65.50% | 65.03% | accepted |
| Number / direct publication | p99 | 65.53% | 64.44% | accepted |
| RK/MulRK / complete workflow | p50 | 37.66% | 38.44% | accepted |
| RK/MulRK / complete workflow | mean | 37.72% | 38.51% | accepted |
| RK/MulRK / complete workflow | p95 | 37.65% | 38.96% | accepted |
| RK/MulRK / complete workflow | p99 | 36.63% | 38.65% | accepted |
| RK/MulRK / direct publication | p50 | 66.52% | 66.68% | accepted |
| RK/MulRK / direct publication | mean | 66.24% | 66.61% | accepted |
| RK/MulRK / direct publication | p95 | 65.54% | 66.76% | accepted |
| RK/MulRK / direct publication | p99 | 65.61% | 66.55% | withheld |

All accepted complete-workflow and publication statistics stay within the 5%
same-implementation drift gate. RK/MulRK publication p99 is withheld because
control drift is 5.281574%. Commit latency stays effectively neutral, as
expected for a publication-only change, and receives no improvement claim.

All four reports are clean and revision-exact. Every retained timing and phase
vector has 500 samples. Corpus bytes, source and output fingerprints, semantic
reopen, Number/RK/MulRK families and values, complete CFB directory topology,
opaque streams, zero target-artifact retention, sink topology, security/no-op
gates, and partial-sink behavior remain exact. The canonical semantic
projection is identical across all legs with SHA-256 `1d12946c...`.

## Withheld scope

This is warm, in-memory, generated native-XLS evidence for forward-only direct
sequential publication. It does not optimize atomic save, generic mutable or
external sources, ordinary source-backed commits, structural/formula/string
edits, or other formats. Process RSS is 388 KiB and 832 KiB higher for the
candidate in the paired directions, so no allocation, RSS, peak/total-memory,
physical-I/O, cold-cache, producer, compression, or throughput claim is made.

## Reproduction

Build the two exact revisions from clean worktrees and run the release harness
under `taskset -c 2` with `--warmup 20 --samples 500` for:

- `xls_numeric_plan_only_number_edit_save`
- `xls_numeric_plan_only_rk_mulrk_edit_save`

Retain the strict `control A1, candidate B1, candidate B2, control A2` order.

Artifacts:

- [summary](../results/cfb-owned-numeric-publication-0172-summary.json)
- [manifest](../results/cfb-owned-numeric-publication-0172-manifest.json)
- [canonical semantic projection](../results/cfb-owned-numeric-publication-0172-semantic.json)
- raw A1/B1/B2/A2 reports and GNU Time sidecars listed in the manifest
