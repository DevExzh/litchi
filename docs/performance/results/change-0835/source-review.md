# Measurement scope review

The 9,389-input source inventory matches committed 0834 final source exactly.
The new release build, qualification and formal commands all use one retained
executable. This batch changes evidence scripts and documentation only.

In `tools/perf-baseline/src/filesystem.rs`, `run_opc_eager_open` reads the entire
archive and passes a borrowed slice to `OpcPackage::from_bytes`; its local
package is destroyed before return. `run_opc_source_open` opens `FileSource`
through `CountingReadAt` and returns the package to the child, which retains it
until after timer/counter snapshots. Their lifetime scopes differ. Borrowed OPC
ingress eagerly materializes this four-Part corpus, while source open reports
zero materialized Parts. This does not describe every owned OPC ingress API.

For PPTX selected-slide lifecycle, the two routes are
`Presentation::from_bytes(fs::read(source)?)` and `Presentation::open(source)`.
Both call the selector-first `slide(PPTX_FILE_SELECTED_POSITION)` and destroy
the local presentation within the operation. The label "eager" identifies the
buffered archive route; it is not evidence that every media Part is decoded.
PPTX logical counters come from a separate untimed replay and cannot be used as
a timed syscall profile.

Verified-cold ZIP sources append proven zero EOCD-comment padding to end on a
page boundary: OPC adds 1,776 bytes; PPTX adds 1,741. The logical payloads are
unchanged. The aligned source route may perform one additional bounded 64 KiB
metadata-tail read. Therefore warm/cold observations include this documented
private layout difference as well as the observed page-cache transition.
Neither fincore nor process `read_bytes` proves a physical-media read.

The six blocks retain all observations. The original copied statistical-plan
settings were reconciled before admission; `measurement-plan-initial.json`
retains the earlier draft. The reader's first execution incorrectly expected
14 qualification children, overlooking that the paired OPC report has four.
The failure, original reader and explicit 16-child correction remain retained.
No native workload was repeated for that offline-reader correction.

ADR 0005 governs the measured I/O and preservation scope; ADR 0030 distinguishes
borrowed eager OPC ingress from retained owned-source laziness. The batch
introduces no architectural change, public API, concurrency behavior, weaker
limits, altered publication semantics, or new production permission. iWork and
unrelated worktree files remain outside the batch.
