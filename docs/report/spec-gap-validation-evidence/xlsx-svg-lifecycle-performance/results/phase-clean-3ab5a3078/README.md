# XLSX SVG exploratory phase capture

Review status: independently approved for descriptive retention. The
optimization materiality gate remains indeterminate: the capture lacks an
inclusive operation clock, isolated SVG planning/proof attribution, and the
required repeated scan/census counts. No production optimization is authorized
by these observations.

The retained raw bundle was copied byte-for-byte from
`/var/tmp/litchi-xlsx-svg-phase-results-ac288-20260912T0545Z`. It contains 39
runner artifacts: nine phase JSON receipts, nine empty stderr sidecars, nine
GNU time sidecars, build and Cargo metadata, before/after source manifests,
binary digests, provenance, commands, the derived report, and verification.
No raw receipt, report, provenance file, or JSON value was rewritten during
retention.

The clean named source branch was
`perf/xlsx-svg-phase-ac288-historical`, at commit
`3ab5a3078274d0aae3d01c4bb620061d7a45778c` (historical base `8444e88ba` plus
the phase scaffold and host-probe correction commits). The approved XLSX
production source pin remained
`ac288a303264f9ea0bb4081baa44031bee5b79a7`; all 19 pinned production paths
matched its committed Git blobs, and the source worktree was clean. The
before and after source manifests are identical with SHA-256
`a8f7021a6b8beed82a576a21ec65495cd25a2eae84d8d8ee661e1899e8e9f4d9` and cover
4,875 hashed inputs (4,850 package files plus 25 declared extras). The measured
phase binary has SHA-256
`bec33231389e15593ed0c9842d5ab2a3c7e6615da208ec416c2a2f57b2e659f4`, stable
before and after capture.

The run used three fresh processes, two warmups, and 20 measured samples for
each of `multi_picture_same_drawing_16`,
`multi_picture_same_drawing_64`, and `multi_picture_same_drawing_256`: nine
receipts and 180 measured operation samples, each with six phase observations. The runner verifier accepted the
receipt schema, input identities, semantic/output gates, allocator equations,
retained-live boundaries, valid peaks, source manifests, binary identity,
sidecar exits, and derived report. The disposable Cargo target was removed
by the runner after capture.

This is an absolute exploratory decomposition of the public operation into
`open`, `stages`, `commit`, `firstsave`, `reopen_secondsave`, and complete
`validation`. The validation clock includes the harness graph, picture,
opaque-fragment, and reopen-byte checks; its observed magnitude must not be
read as a production internal-function cost. Phase values are not compared
with the sealed 69-lane acceptance baseline and provide no speedup, regression,
causal, scaling, or optimization conclusion.

The host was an AMD EPYC 9R45 with 32 cores and 129,447,068 kB reported memory
under Linux 7.0.0-1012-aws, with concrete CPU, memory, and storage probes. A
PPTX cargo/Clippy workload and a library test were active on the shared host
during capture. Therefore these observations do not claim contention-free or
isolated timing. The sealed acceptance baseline remains unchanged.

See [`../../phase-decomposition.md`](../../phase-decomposition.md) and
[`../../requirements.md`](../../requirements.md) for the exploratory contract
and measurement limits.

Independent review checked the captured commands and provenance values. It
also found that the historical phase verifier accepts some altered command
flags and compiler/allocator/phase-scope provenance on copied bundles. The
actual captured values passed independent review, but that verifier must be
hardened before it serves as the sole authority for a future capture. The
post-correction Python suite contained 41 tests.

`root-verification.json` separately records the root receipt, source-input,
production-pin, and exact-copy checks. It does not claim that the historical
verifier detects the provenance mutations described above.

Follow-up `2053dc8a3` hardens command flags, compiler/allocator/phase-scope
provenance, incremental mode, and binary-path binding. Independent replay
passed all 43 Python tests and rejected nine disposable mutation probes while
accepting this unchanged capture. Use that verifier or a reviewed successor
for future checks. This closes the tooling issue without changing the
historical source, receipts, or deferred optimization decision.
