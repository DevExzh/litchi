# 0487: retain consumed replay bytes across short reads

## Evidence and mechanism

0485 batches parser-consumed bytes within each filled adapter slice. The 0486
caller profiles attribute 27.38% of authored-heavy file-input whole-child sampled
period weight to stacks containing `statx`, including replay read guards and
consumed-byte sink flushing. DOCX's event encoder frequently returns a short
read before filling the adapter's existing allocation.

Retaining an already-consumed replay prefix in the same allocation should let
subsequent reads append into its unused tail. This should reduce hashing/sink
boundaries while retaining separate replay callbacks and their source checks.
No new allocation, source version cache, merged provider callback, or removed
source XML audit is proposed. Source prefix/suffix and fixed payload paths keep
their current buffer behavior. Every parser fragment retains its Work charge.

The expected benefit is largest for authored-heavy file input. Owned input and
providers already filling their windows are controls. The full 0485 matrix also
retains small, source-heavy, short-read, latency, and explicit replay-store arms
so regressions remain visible.

## Partial output contract

ADR 0005's Output section requires caller-owned non-atomic sinks to report
incomplete output and bytes written. The public OPC splice publication API and
`replay_short_sink_reports_exact_partial_output` require that count to equal
actual sink acceptance. Neither promises a flush after each provider read.

The proposed retention changes a private timing boundary: mutation or
cancellation during a later replay callback can prevent the earlier retained
prefix from reaching the sink. Already accepted output must remain counted
exactly, and no pending bytes may be published after authorization fails.
Independent read-only review reached the same conclusion: this timing change is
compatible with the documented contract. An ordinary provider I/O failure must still flush the previously consumed
prefix once while preserving its original error, unless source/cancellation or
sink failure takes precedence. Tests must exercise those distinctions directly.

## ADR obligations

| Constraint | Required implementation/evidence |
| --- | --- |
| 0001 correctness and preservation before speed | Exact candidate/source/replay digests and untouched archive oracles; revert if representative benefit is absent. |
| 0002, 0010, 0011, 0024 ownership | Private OPC adapter change only; no facade or DOCX archive dependency. |
| 0003 immutable snapshots and patches | Source-bound authorization, unchanged publication/inverse identity, reopen and patch oracles. |
| 0005 I/O, resources, and measured performance | Same bounded allocation and per-callback freshness checks; exact accepted-output accounting; matched normal/allocator measurements. |
| 0006 validation and security | Preserve source-first failure precedence, XML audit, replay EOF/hash proof, Work limits, and cancellation. |
| 0008 verification | Focused adapter tests, OPC/DOCX feature lanes, lint/docs/boundaries, and existing sanitizer fuzz target. |

All 30 ADR/README hashes match the previously read 0486 set, as recorded in
`adr-refresh.json`. This hypothesis is not a measured result or completion claim.
