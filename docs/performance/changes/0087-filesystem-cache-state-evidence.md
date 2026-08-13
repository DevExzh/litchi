# Change 0087: filesystem cache-state evidence

Date: 2026-08-13

Harness revision: `0b3109467`

Smoke artifact revision: `0e86dfa22b247b0ce53eed774dc84b0fad5f371d`

Status: one-sample debug correctness/counter smoke; no performance claim

## Scope

The harness adds five opt-in, fresh-child filesystem cases:

- eager and source-backed OPC open;
- eager and source-backed OPC one-Part atomic save; and
- CFB same-length overlay atomic save.

Each sample runs in a fresh process and records a keyed `warm` or
`cold-requested` state. Cold-requested means the Linux cache advice request was
accepted; it does not guarantee a physically cold kernel or storage cache.
Atomic saves use same-filesystem sibling staging and rename.

## Retained smoke

`target/perf/filesystem-smoke-0096.json` is a schema-1 debug run from a dirty
worktree with zero warm-ups and one sample per state. It contains 10 result
records and five filesystem-evidence records. A compact tracked extraction is
[`filesystem-smoke-0096-summary.json`](../results/filesystem-smoke-0096-summary.json).

The retained correctness and counter facts are:

- source-backed OPC open performs 13 logical reads totaling 1,008 bytes and
  materializes zero Parts in both keyed states;
- eager OPC open materializes four Parts;
- eager and source-backed OPC atomic saves emit the same 16,783,632-byte
  output with SHA-256
  `f4bbe4de18853444cc6cd093cf561249decaa81f776afcf5de122667f5dd7009`;
- CFB atomic save reports one changed span and emits 16,913,408 bytes with
  SHA-256
  `7994759e1b2e3e520c0f0df5efb1586e34c6bc0f5744a7f4b989733cfd2830fc`;
  and
- cold-requested records contain nonzero process `read_bytes`, including
  20,480 bytes for source-backed OPC open and 16,916,480 bytes for CFB save.

## Claim boundary and next evidence gate

The elapsed values are deliberately not summarized. One debug sample from a
dirty worktree cannot support latency distributions, uncertainty, allocation,
peak-memory, throughput, warm/cold deltas, cache-temperature guarantees, or a
production-performance conclusion.

The next gate is a clean release build on a named controlled filesystem and
storage device, with repeated warm and independently verified cold samples,
CPU affinity, retained raw distributions, process I/O, allocation and peak-RSS
measurements, and balanced eager/source controls where a comparison is made.
