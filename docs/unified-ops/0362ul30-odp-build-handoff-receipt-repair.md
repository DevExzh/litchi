# 0362ul30 follow-up — ODP handoff receipt and freshness repair

Date: 2026-10-24. This source-only follow-up repairs the remaining handoff
issues found after commit `11858e697f845ade1ad0e633af6d1afcf1d99cc5`. The
118 packet is preserved; this is a fresh packet and does not claim Cargo,
native, timing, heaptrack, qualification, or acceptance work.

The review is
`/home/zhuhe/.cache/litchi-api-contract/odp-build-handoff-repair-review-11858e697f.json`,
SHA-256
`23d31045e5f59bb78cc789c4deb9414c851a82501c5717f00f1f59f1a6b1d499`.
It required a fresh plan-recheck output, generated identity-bound build and
memory receipts, and absent-or-empty target/TMP guards. Bash is declared for
the array/process-substitution memory commands.

## Fresh packet

The new packet is
`(deleted-coord-artifact)`.
The prior repaired packet at
`(deleted-coord-artifact)`
and its review artifacts remain unchanged. New packet files:

| File | Bytes | SHA-256 |
|---|---:|---|
| `input.json` | 1,749 | `92babe0b6243cdf903830b9b2a6d796604cade3a0c3b2c9724caa796391ee76d` |
| `plan.json` | 2,058 | `0e96988db9add9de7d0f91f792ef8379d995a6083dac559b70d032a1a67a589e` |
| `build-handoff.json` | 20,762 | `be60e8b150675818451a1962aea9d1dc32e3e51e74b51956ec910d0632088f2e` |
| `memory-protocol.json` | 22,912 | `c638ad86e6a7edcf1532408156684d99060020a75d04a31aa67693a1d5c6658b` |

The runner's source-only plan recheck uses the new
`qualification-run/manifest-plan-recheck` path and passed. Its retained
output is 2,093 bytes, SHA-256
`a9255a31f6a554e5e50d14f22472ae7fd44dd19af107b157233e434480d2d672`.
The earlier 118 plan receipt is preserved separately and is not overwritten.

## Repaired execution contracts

- The two build commands require each target and build TMP path to be absent
  or an empty directory before creating it. They record those successful
  freshness checks, copy/hash-check the ignored baseline lock, and generate a
  `build-receipt.json` only after both binaries, clean roots, source revisions,
  probe hashes, locks and fixture identity have been checked.
- The memory pre-run requires that generated build receipt, rechecks both
  binary hashes, source revisions and clean roots, probe hashes, fixture hash,
  CPU and exact route arguments, and records actual global-quiet and
  absent/empty memory-root/TMP status in `guard-status.json`.
- Memory commands are explicitly Bash. They create both memory legs and both
  TMP directories after the pre-run freshness check, then export per-leg
  `TMPDIR` before launching heaptrack. Captures are discovered as actual
  `.gz`/`.zst` files, hashed, and passed to
  `heaptrack_print -f <capture> -H <histogram> > <summary>`.

The +0.5 MiB historical ODP peak-heap residual remains unresolved and no
latency-only or source-only result closes it.
