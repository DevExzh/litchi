# Clean stylesWithEffects correctness capture

All 52 lanes passed from clean source commit
`46c4568483848faa5522b8cd8f2c999e74537af7`, with production pinned to
`d1f299d00e0dd5cc5cd8ddf9811c4b1ad21d1119`. The capture contains one process,
zero warmups, and one correctness sample per lane: 29 successful operations
and 23 expected typed refusals. These are correctness receipts, not latency,
throughput, memory-scaling, or native Office acceptance evidence.

The runner's `smoke-verification.json` and the independently rerun
`root-verification.json` both report success. The source manifest covers nine
local packages, 790 source files, and 18 evidence extras. Before/after source
manifests, Cargo metadata, and binary hashes match.

- Source manifest SHA-256:
  `f10d60a07d8c874ee80c2286e64dde4f7cebbaa02b8397eebf12c25c35c4b581`.
- Cargo metadata SHA-256:
  `e992f2b1b70d1f193623269f6cc247e4be43dc40f587473a25575908eac29033`.
- Executable SHA-256:
  `75e37e666e3c026a117feaa4bd84363f0e80d576828f4c4723f0b00f6068a912`.

The existing-owner aggregate-cap control now accepts 110,151 bytes and
refuses 110,150 bytes at transaction commit, with a typed
`DocxError::Opc::ReadLimit` for `TotalPartBytes`. Its exact-fit output retains
unrelated members, as checked by member digests and reopened output. The
separate historical failure slice documents the pre-fix behavior.

All 168 original result files, including the root verification receipt, are
retained byte-for-byte. Paths inside the machine receipts identify the
original isolated source, results, and executable locations. The executable
target was removed by the runner after capture; matching before/after hashes
remain. Re-execution requires a checkout of the recorded source commit and a
fresh build. See the parent runner and verifier for the complete procedure.
