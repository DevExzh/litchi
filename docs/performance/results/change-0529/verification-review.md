# 0529 verifier review

The 0529 verifier is a bounded, staged custody check for the XML attribute
probe pilot. It reports `incomplete` when planned evidence has not arrived and
does not turn a partial capture into a performance claim. It can be run during
the serial campaign with:

```text
python3 -B docs/performance/results/change-0529/verify.py --stage baseline
python3 -B docs/performance/results/change-0529/verify.py --component all --strict
```

The baseline view is deliberately independent of candidate source binding. It
replays the baseline patch from the frozen revision, requires the single
`tools/perf-baseline/src/lib.rs` harness change, validates any completed
baseline builds and children, and lists the remaining planned children as
pending. The full view additionally requires candidate replay, all receipt
artifacts, the native ABBA order, the exact analyzer replay, publication gate
arithmetic, flag reviews, final quality, disposition, cleanup, and the sealed
inventory.

The source checks distinguish the two operation-scoped allocator vectors:
`publication_allocation_metrics` supplies the allocation-call gate and
`commit_allocation_metrics` remains a separately validated diagnostic. Raw
normal reports must carry explicit unavailable samples; allocator reports must
carry balanced measured samples. The analyzer replay is byte-for-byte and an
independent gate calculation rejects a report that uses the commit vector for
the publication gate.

The publication Ir lane is intentionally left incomplete until a dedicated
raw-profile verifier binds callgrind ownership, the dump boundary, and every
shape/repeat result. A profile JSON or regex match by itself cannot satisfy
that conditional gate.

The verifier performs only read-only Git-index replays and temporary files under
the disk-backed `/home/zhuhe` filesystem. It never writes the evidence bundle
unless an explicit `--output` path is supplied, and refuses an output path
inside an existing sealed bundle. Negative probes exercise the actual path,
interval, allocation balance, publication-channel, and ABBA validators.

No ODF or iWork evidence is admitted by this verifier; the campaign remains in
the OLE2/OOXML-first priority.

## Current campaign checkpoint

The independent source, build-receipt, capture, analyzer-replay, and negative-probe checks currently pass. The analyzer records a rejected pilot: the native performance gates fail while the publication allocation-call gate passes, so the conditional profile, hardware, and eager lanes remain unmeasured by protocol. Final source binding, final quality receipts, disposition, cleanup, and the sealed inventory remain pending the root checkout refresh; this checkpoint is not a final acceptance claim.

## Root finalization checkpoint

The root extended precleanup verification to include quality, decision and real
negative probes, then verified all eight components. All 14 final quality
checks pass with 2,024 successful executions. The final auditor is byte-exact
baseline and only the public streaming tests differ from the baseline stage.
Owned paths are removed and all nine non-seal components replay after cleanup.
The final SHA256 inventory is checked separately by the all-component verifier.
