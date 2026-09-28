# 0825 source and scope review

The baseline is reconstructed from the exact pre-0824 `opened/transaction.rs`
and `opened/xml.rs` archives. The after leg is commit
`032b0e89cdeb93b6a741c1712863b6687baeb6e5`. The complete production census
must differ only at those two paths, while Git HEAD and the ordinary-save
harness remain fixed. Neither historical timing samples nor historical
executables enter this trial. Source hashes distinguish the reconstructed
baseline even though both builds share the same recorded Git HEAD.

The 0824 edit-only result motivates this experiment but does not predict a
complete-save improvement. The real PPTX edit already avoids two duplicate
Scene reads when compaction removes only outer whitespace. A lifecycle also
opens the package, serializes it, writes a sibling, synchronizes the file,
renames it, and synchronizes the parent directory. The previously observed
durability costs make a fresh matched complete-workflow measurement necessary.

## Timed boundaries

`tools/perf-baseline/src/ordinary_save.rs::run_case_with_durability` defines
four independently prepared intervals:

- Lifecycle: `Owner::open`, `Owner::edit`, then `Owner::save_at` using the
  default durability route.
- Edit: `Owner::edit`, with the package opened outside the clock.
- Atomic publication: `Owner::save_at`, with opening and editing outside the
  clock. This includes the unchanged default full-durability atomic publisher.
- Counting publication: `Owner::write_to`, with opening, editing, sink creation,
  and sink reservation outside the clock. PPTX currently materializes an archive
  through `to_bytes` before writing it to the counting sink.

All four stop the clock before dropping the package owner. Output hashing,
readback, destination removal, and counting-sink byte classification occur
outside their clocks. The PPTX edit helper drops its returned publication
snapshot inside `Owner::edit`; this is distinct from package-owner destruction.
Allocation regions bracket the measured operation while the owner is live, so
net live bytes describe retained state at that boundary rather than a leak.

The observer adds allocator and process-counter features. Its elapsed times
are diagnostic and never pooled with native timing. Process counters include
adjacent procfs probes and are not precise production syscall or physical-device
attribution. The counting-sink byte split compares archive payloads; identical
compressed payload bytes do not independently prove that a runtime copy path
was taken. The four medians are not an additive decomposition.

## Contract and evidence

There is no new runtime or architectural change in this batch. The 0824
exact-root proof retains its ADR 0003 source-checked atomic publication and
patch behavior, ADR 0005/0032 bounded transient state, and ADR 0006 validation,
resource-limit, unknown-content, and byte-preservation contracts. Default
durability is unchanged. Root restores the after archives before the after
build and leaves those exact bytes installed at completion.

Fresh artifact admission covers the original six corpus cases and all five
policy outputs on each leg. The policies are untimed controls; all measured
path-save scenarios use the default full policy. Three real inputs enter the
timing matrix, and DOCX/XLSX are unaffected-format controls. Native timing,
allocation observations, and correctness qualification run separately. All
twelve rows and their uncertainty remain visible, including any regressions.

The claim is limited to the recorded host, three small admitted files, fixed
operations, and warm provider/filesystem state. This does not establish cold
cache, broad producers, current external Office resave, cross-platform,
remote I/O, concurrent workloads, or completion of the wider performance goal.
