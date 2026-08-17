# Change 0173: CFB comment publication fusion

Date: 2026-08-17

## Decision

Retain the narrow native XLS existing-comment optimization. Effective
source-backed comment transactions now validate their semantic owner on the
exact composed CFB view inside the planner's complete fingerprint fence. The
snapshot's immutable `Arc<[u8]>` provenance is also retained through the
existing owned-source ingress, so direct sequential publication omits its two
redundant outer mutation preflights while still hashing source and target
during emission.

Exact no-ops skip both the owner callback and fallback readback. Generic
`ReadAt` sources are unchanged. Candidate composition and atomic save retain
their complete fingerprint fences. Encoding-width, fixed-length NOTE/TXO,
protection, signature, encryption, stale-source, partial-sink, topology, and
semantic readback rules are unchanged.

## Deterministic work reduction

Before this change, each effective scalar or bounded-batch transaction:

1. built and reopened the splice plan;
2. created another fingerprint-checked composed view for semantic readback;
3. performed direct publication's initial preflight, hashed source and target
   during 64-KiB emission, and performed its post-emission preflight.

The fused owned-source path removes three complete 16,995,840-byte scans,
three source/target SHA-256 pairs, and 51 one-MiB logical `ReadAt` calls per
effective transaction. It retains the 260 64-KiB emission reads and their
source/target hash pair. The reduction is per transaction, not per comment or
splice. Harness source counters describe owned ingress and therefore do not
measure these internal reads. This is not a physical-I/O claim.

## Correctness gates

- All 1,015 `litchi-xls` library tests and all 11 comment transaction tests
  pass, including 256-update batching, POI producer reopen, exact no-op,
  compressed/UTF-16 width preservation and refusal, protected/signed/encrypted
  refusal, atomic save, and short/partial sinks.
- Focused library and integration Clippy passes with
  `-D warnings -D deprecated`; formatting and diff checks pass.
- Independent review is SAFE after catching and correcting a no-op fallback
  regression before the production commit. It separately confirmed sealed
  immutable provenance, retained emission hashing, unchanged atomic-save
  fences, and exact deterministic work accounting.

## Clean release ABBA

Control `709d21717` and candidate `e204e184e` were built from clean detached
worktrees with locked release dependencies. Their binary SHA-256 digests are
`ca871946...` and `cfc37373...`. Every leg was pinned to CPU 2 on the AMD EPYC
9575F host, which exposed one logical CPU. The order was strict `A1 control,
B1 candidate, B2 candidate, A2 control`, with 20 warmups and 500 retained
samples for each source-backed selector and its matched eager guard.

Positive values below mean lower candidate latency:

| Selector / phase | Statistics | A1 -> B1 | B2 -> A2 | Decision |
|---|---|---:|---:|---|
| one-comment / complete workflow | p50, mean, p99 | 47.19%, 47.17%, 45.54% | 47.17%, 46.96%, 46.23% | accepted |
| one-comment / semantic staging and plan | p50, mean, p95, p99 | 32.42%, 32.38%, 32.29%, 30.78% | 32.37%, 32.17%, 31.08%, 31.68% | accepted |
| one-comment / direct publication | p50, mean, p95, p99 | 61.02%, 61.03%, 60.98%, 60.24% | 60.99%, 60.80%, 59.68%, 59.15% | accepted |
| 256-comment / semantic staging and plan | p50, mean, p95, p99 | 32.22%, 32.28%, 32.42%, 32.57% | 32.21%, 32.08%, 31.49%, 30.53% | accepted |

The scalar complete-workflow p95 is withheld because the second matched eager
guard direction is 5.027675%, just outside the 5% gate. The batch complete and
publication results are withheld because candidate same-implementation drift
reaches 6.673640% and 17.829569%, and the matched eager batch guard is
unstable. The stable batch semantic phase remains accepted because both source
implementations stay within 1.91% and both paired directions agree.

All raw reports are clean and revision-exact. Every timing and phase vector has
500 samples. Corpus, output hashes, semantic projection, source ingress, sink
topology, splice diagnostics, and fingerprints match across all four legs. The
canonical semantic projection SHA-256 is `0d59a6b5...`.

## Withheld scope

This is warm, in-memory, generated XLS evidence. It makes no allocation, RSS,
peak/total-memory, physical-I/O, cold-cache, independent-producer, compression,
throughput, or atomic-save latency claim. It does not broaden comment CRUD:
adding/removing comments, changing shape topology, or length/encoding-width
transitions remain explicit refusals.

## Reproduction

Build each exact revision from a clean worktree, then run the release harness
under `taskset -c 2` with `--warmup 20 --samples 500 --workers 1` and:

- `xls_comments_eager_edit_save`
- `xls_comments_source_backed_edit_save`
- `xls_comments_eager_batch_edit_save`
- `xls_comments_source_backed_batch_edit_save`

Retain the strict `control A1, candidate B1, candidate B2, control A2` order.

Artifacts:

- [summary](../results/cfb-comment-fusion-0173-summary.json)
- [manifest](../results/cfb-comment-fusion-0173-manifest.json)
- [canonical semantic projection](../results/cfb-comment-fusion-0173-semantic.json)
- raw A1/B1/B2/A2 reports listed in the manifest
