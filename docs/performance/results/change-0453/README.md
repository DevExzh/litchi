# Change 0453 evidence

This bundle compares managed decoded payload sharing in PPTX cross-slide copy
plans against commit 456e19246506ffba230052df2b25b072ce4a8749 production code.
Both builds use the same final standalone timing/allocation harness. Exact source
copies and full source manifests distinguish builds carrying that same revision.

Replay after capture, sealing and cleanup:

```sh
python3 -B docs/performance/results/change-0453/verify.py --sealed --cleanup
```

The verifier also runs from a standalone copy of this directory. It checks source
and binary custody, command receipts, report oracles, exact corpus/output hashes,
derived measurements, review flags, fuzz evidence and owned cleanup. It does not
rerun Office applications or recreate machine performance. portable-probes.py
records successful standalone replay and rejected corruptions.

See protocol.md and protocol.json for the frozen matrix; measurements.md and
measurements.json for separate ordinary latency and allocator results;
regression-review.md for interpretation; source-review.md for ownership and
fallback behavior; validation-notes.md for preserved failed attempts. Logs are
retained as deterministic gzip files with raw-byte identities in compression.json.
The complete directory is covered by SHA256SUMS.

The ASan/sancov smoke executes the existing unchanged parse_opc fuzz target. It
checks the substrate only and adds no PPTX-specific fuzz target or native Office
application claim. Full non-iWork goal completion remains outside this batch.
