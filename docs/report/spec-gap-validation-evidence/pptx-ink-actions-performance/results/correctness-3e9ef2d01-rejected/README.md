# Rejected 3e9 correctness capture

This retention bundle preserves the complete raw receipt directory for the
fresh `3e9ef2d014a2f70bbbca18b824aa50a1567cdcc7` checkout. It is deliberately
marked **rejected-before-scenario-execution** and contains no correctness or
performance result.

The release build exited zero, but the public `--host-probe` prerequisite
exited one with this exact stderr:

```text
manifest retained generator hash differs from compiled source
```

The executable had been compiled against the stale committed corpus manifest
receipt (`8d3c6ba7ff16380e82d5192600909e24aa6432be40bac7e2480e1e6b4a4ba7c3`). The corrected generator receipt is a pending
source change and requires a new freeze, build, and host probe before any
scenario. Because this attempt's first orchestration continued after the host
probe failure, it also retained a failed matrix receipt and 42 failed lane
receipts. Those files are retained for audit only and are invalid evidence;
future orchestration must stop at the first prerequisite failure and use the
order recorded in `retention-manifest.json`.

Every regular file from the raw source `/var/tmp/pptx-ink-actions-3e9ef2d01-correctness.a2sTcr` was copied byte-for-byte under
`raw/`, including empty stdout/stderr files, Cargo metadata, source manifests,
command JSONL, host/load/environment records, build output, and all failed
lane receipts. The manifest records the source path, retained path, byte count,
and SHA-256 for each of the 124 files (4982051 bytes).

The raw `capture-provenance.json` also lists a `binary hash changed` failure. Its `raw/binary.sha256` and `raw/binary-after.sha256` receipts are byte-identical; both facts are retained and neither is promoted to scenario evidence.

The retained binary remains external at `/var/tmp/pptx-ink-actions-3e9ef2d01-target-parent.onUULP/target/release/pptx-ink-actions-performance` with SHA-256
`83319e9ff652d6b15ae118303773b105392ba7e7c8fb4951bd7317614f70ab59`; its hash and stat receipts are included in
`raw/`. Keep the source worktree, target, binary, and original raw receipt until
root completes the full raw Git byte audit. Do not delete or reuse them from
this bundle before that audit.

Pins captured in the receipts are semantic owner
`cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd`, production source baseline
`2a2ffa1cae4e6b7070082768ce84483e5d411dc8`, design blob
`597400950b1027c47cd6e4cbbedd23915bc0980e`, helper SHA-256
`bec6baafcf735d778216fb54fe6299312707e912dcaed5f89de988f7112bb58e`, and isolated harness lock SHA-256
`3c773b848134c52df596034ffbbe8688dde9924fb741dee8bc8bce0eaa772f2e`.

No `/usr/bin/time`, timed lane, allocation/RSS result, native PowerPoint
acceptance claim, or speedup claim was produced.
