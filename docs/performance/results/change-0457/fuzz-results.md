# Final fuzz evidence

The README is retained byte-for-byte as the pre-capture instruction record.

Both targets completed all three stages successfully: offline lock generation,
ASAN build, and 1,000 smoke iterations with seed 457. The source snapshots,
initial corpus, stage receipts/logs, and post-run inventories are retained under
`artifacts/`. The shared `target/fuzz-asan` cache remains outside cleanup scope.

`retain-binaries.py` losslessly stores the two captured executables as
deterministic gzip files. `binary-retention.json` binds their original byte
counts/hashes, compressed identities, unchanged original inventory, and the
transformation helper. Each compressed file was streamed back through SHA-256
before its owned raw copy was removed. The 75,028,216 original bytes occupy
18,158,836 compressed bytes. The original capture inventory was not rewritten.
