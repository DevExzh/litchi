# Change 0171: legacy CFB owner-validation fusion

Date: 2026-08-17

## Decision

Retain the narrow owner-validation fusion for source-backed DOC paragraph,
PPT shape-text, and XLS worksheet-visibility transactions. Each effective
operation now validates the format owner on the exact composed positional CFB
view supplied by the existing owner callback. The callback remains after CFB
reopen/range validation and before the final complete source/target
fingerprint fence.

The format APIs, patch vocabulary, no-op semantics, CFB planner, publication,
and atomic-save paths are unchanged. Exact no-op XLS visibility plans retain
their previous fallback readback. The only public surface addition is a normal
`From<OverlayError>` conversion in `litchi-ppt`; no CFB type crosses a format
facade.

## Deterministic work reduction

Before this change, each of the three format paths planned and reopened the
candidate, then called `ValidatedOverlayPlan::composed_source()` again for
semantic readback. That call performed one further complete source scan while
hashing the source and composed target together.

Each effective transaction now removes exactly:

- one complete logical scan of the source artifact;
- `ceil(artifact_bytes / 1,048,576)` source `ReadAt` calls; and
- one source/target SHA-256 pair.

This is once per DOC or PPT transaction or XLS visibility batch, not once per
changed span. Candidate CFB reopen, format-owner validation, the final complete
fingerprint fence, later sequential publication checks, and atomic-save fences
remain. Exact no-ops do not receive this reduction.

For the measured 2,135,552-byte XLS visibility corpus, the deterministic delta
is one 2,135,552-byte scan and three one-MiB logical reads per effective scalar
or 64-worksheet transaction. The harness source counters measure initial
memory-backed ingress, not those internal fingerprint reads, so this is a
code-derived logical-work count and not a physical-I/O result.

## Correctness gates

- Source-backed DOC body-text tests pass 15/15, PPT text-edit tests 26/26, and
  XLS sheet-visibility tests 12/12.
- Locked offline production checks for all three crates pass. Library Clippy
  passes with `-D warnings -D deprecated`; formatting and diff checks pass.
- The combined all-target test run completed the DOC/PPT suites and thousands
  of XLS tests. Its only failure was the unrelated, reproducible
  `xls_writer_encryption::all_profiles_round_trip_and_emit_exact_filepass_families`
  mismatch. An isolated rerun failed identically.
- Both independent reviews returned SAFE. They confirmed that source/version,
  physical expected-range, CFB reopen, owner semantic, final fingerprint,
  macro/protection/encryption, partial-output, and publication boundaries
  remain fail-closed.

## Clean release ABBA

Control `37d7e9d2f` and candidate `667a884e2` were built from clean worktrees
with locked release dependencies. Their binary SHA-256 digests are
`307aa56c...` and `313319e0...`. Every run was pinned to CPU 2 on the AMD EPYC
9575F host, which exposed one logical CPU to the process. The retained order
was strict `A1 control, B1 candidate, B2 candidate, A2 control`, with 20
warmups and 300 samples for each scalar/bounded-batch source-backed selector
and its matched eager guard.

Positive values below mean lower candidate latency:

| Selector / phase | Statistic | A1 -> B1 | B2 -> A2 | Decision |
|---|---:|---:|---:|---|
| 64-worksheet source-backed / complete workflow | p50 | 12.51% | 12.88% | accepted |
| 64-worksheet source-backed / complete workflow | mean | 12.61% | 13.36% | accepted |
| 64-worksheet source-backed / complete workflow | p95 | 12.82% | 15.38% | accepted |
| 64-worksheet source-backed / semantic staging/plan | p50 | 31.48% | 31.82% | accepted |
| 64-worksheet source-backed / semantic staging/plan | mean | 31.52% | 32.16% | accepted |
| 64-worksheet source-backed / semantic staging/plan | p95 | 31.69% | 33.16% | accepted |
| one-worksheet source-backed / semantic staging/plan | p50 | 32.37% | 31.44% | accepted |
| one-worksheet source-backed / semantic staging/plan | mean | 32.27% | 31.48% | accepted |
| one-worksheet source-backed / semantic staging/plan | p95 | 32.09% | 31.73% | accepted |

The matched 64-worksheet eager complete-workflow guard changes by only
0.23%-2.61% across the accepted p50/mean/p95 statistics. All accepted plan
statistics remain within the 5% same-implementation drift gate, and both
paired directions agree.

All four reports are clean and revision-exact. Every timing and phase vector
has 300 samples. The corpus, complete output digests, semantic projection,
source ingress, sink topology, one/64 replacement bytes, one/64 changed spans,
and reopen results are exact across all legs. The canonical semantic projection
hash is `47283bda...`.

## Withheld scope

The scalar complete-workflow result is withheld because its eager guard shifted
by a similar or larger amount. Batch p99 is withheld because the control's
same-implementation tail drift exceeds 5%. Publication alone regresses or
disagrees, so no publication-latency result is accepted. DOC and PPT did not
gain matched source-backed mutation selectors in this tranche, and therefore
receive correctness and deterministic-work claims only.

This is warm, in-memory, generated XLS evidence. It makes no allocation, RSS,
peak/total-memory, physical-I/O, cold-cache, producer, throughput, compression,
or broad legacy-Office claim.

## Reproduction

Build each exact revision from a clean worktree, then run the release harness
under `taskset -c 2` with `--warmup 20 --samples 300` and these four opt-in
selectors:

- `xls_visibility_eager_edit_save`
- `xls_visibility_source_backed_edit_save`
- `xls_visibility_eager_batch_edit_save`
- `xls_visibility_source_backed_batch_edit_save`

Retain the strict `control A1, candidate B1, candidate B2, control A2` order.

Artifacts:

- [summary](../results/cfb-owner-fusion-0171-summary.json)
- [manifest](../results/cfb-owner-fusion-0171-manifest.json)
- [canonical semantic projection](../results/cfb-owner-fusion-0171-semantic.json)
- raw A1/B1/B2/A2 reports listed in the manifest
