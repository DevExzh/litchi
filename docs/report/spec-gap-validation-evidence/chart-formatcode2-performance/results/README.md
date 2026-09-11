# Raw receipts

The final current-source receipt contains 84 per-process JSON files and 84
`/usr/bin/time -v` files: 28 lanes, three fresh processes per lane, and 20
measured samples per process after two warm-ups. The source manifests, source
hashes, build provenance, host/toolchain data, commands, and recomputed
verification metadata are retained here. The 240 `write_to` samples include
the nonallocating sink's byte-count, checksum, and write-call work; no
comparison claim against the Vec-returning lanes follows.
