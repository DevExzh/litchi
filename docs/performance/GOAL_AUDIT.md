# Non-iWork `docs/GOAL.md` audit

**Audit date:** 2026-09-04
**Audit basis:** `feat/office-format-completeness` on 2026-09-04, using
`9c6742c5212dd0e7ff2367da585abe357aae8975` as the ZIP64 control and the
integrated changes visible at audit time.  iWork is outside this audit by the
user's instruction.

**Disposition: OPEN.** The repository has substantial correctness and scoped
performance work, but the definition of done in `docs/GOAL.md` is not met. The
coverage index explicitly describes itself as representative, and the current
records do not provide a complete, independently reproducible baseline and
optimization result for the non-iWork CRUD matrix.

## Authorities and reading boundary

This audit uses the current goal, the accepted ADRs, and the current
non-iWork performance records:

- [`docs/GOAL.md`](../GOAL.md), especially the mission, measurement contract,
  CRUD matrix, deliverables, and definition of done.
- ADR 0001 (public layers and typed refusals), ADR 0002 (downward crate
  ownership), ADR 0003 (immutable snapshots and atomic edits), ADR 0005 (the
  `ReadAt`/budget/output/measurement contract), ADR 0006 (preserve-by-default
  validation and security), ADR 0008 (migration gates), ADR 0010/0011 (facade
  and OPC physical ownership), and ADR 0024 (current topology).
- [`REPORT.md`](REPORT.md), [`CRUD_COVERAGE.md`](CRUD_COVERAGE.md),
  [`crud-coverage-index-v1.json`](crud-coverage-index-v1.json), and the
  claim-registry policy.

The decisive constraints for the open ZIP work are: preserve is the default;
validation must not mutate; an owned changed OPC source may publish only
through a proven preservation plan; unsupported framing must return a typed
refusal before output; and physical ZIP ownership stays in `soapberry-zip`.
Those constraints come from ADR 0005's exact-source amendment, ADR 0006, and
ADRs 0010/0011.

## Current evidence

| Goal requirement | Evidence at audit time | Assessment |
| --- | --- | --- |
| Reproducible CRUD baseline | The coverage index maps 15 categories and 30 representative selectors. Measured rows require a validated `target/perf/container-baseline.json`; the index itself is only an identity/contract artifact. | Partial. The full required scenario matrix and metric set are not evidenced. |
| Scoped claims | The strict claim-registry check validates 6 claims. The report classifier has 167 rows: 0 `strict_claim`, 145 historical, 14 descriptive, and 8 withheld. | Claims are deliberately scoped; this is not a program completion claim. |
| Correctness and boundaries | Coordinator checks report 850 Python tests run with 20 skips and no failures and a 64-package/240-declaration boundary graph with 14 existing iWork debts. | Useful gates, but they do not establish latency, memory, I/O, or full CRUD coverage. |
| OPC/ZIP/ODF correctness | The coordinator's integrated all-feature run reports `soapberry-zip` 362 library tests, `litchi-opc` 316 library tests, ODF 287 library tests, and all three crates’ integration suites and doctests passing (1,227 total passed, two ignored). The earlier local default-feature reruns (358 ZIP and 293 OPC) were superseded by this integrated run. | Strong package/format correctness evidence; not a complete workspace performance or native-producer gate. |
| Full gate health | The prior ODF MIME/unsafe-text false positive was repaired. Scoped formatter checks now include the inherited ZIP locator/office fixes and the new ODF test fix. The late ZIP64 topology test now passes after its `Result` assertion fix. Warning-denied Clippy remains blocked by the preexisting `large_enum_variant` in `litchi-odf-common/src/package/model.rs:230` (312-byte enum variant versus 8-byte largest alternative). | The integrated package/format scope is green; the warning-denied Clippy gate remains an explicit preexisting blocker. |
| Hardware/resource profiling | `perf stat` counters, 32 affinity CPUs, 128 GiB RAM, Heaptrack, strace, fincore, and cargo-flamegraph are available. | The earlier counter blocker is gone; a pinned release baseline should now be captured. |

The passing library suites establish behavior for the exercised fixtures. They
do not establish the `docs/GOAL.md` requirements for p50/p95/p99, throughput,
allocations, peak RSS, copied/decompressed/recompressed bytes, physical I/O,
lock wait, cold/warm behavior, or Amdahl scaling. The [0404 correctness evidence](results/change-0404/validation.json) retains
commands, source hashes, and compressed logs, including the explicit ODF lint
failure. ZIP/OPC warning-denied Clippy and warning-denied documentation builds
for all three crates pass.

## ZIP64 passthrough audit

The low-level foundation is now materially further along than the last
committed control. `soapberry-zip` retains ZIP64 field origins and tail bytes;
its preservation plan preflights the complete output, patches only movable
offset fields, retains central/local metadata and ZIP64 extensible data, and
has malformed, multi-disk, truncation, descriptor, limit, offset, sink, and
no-output tests.

The integrated work changes public preservation construction to `AllowZip64`,
records whether the OPC source already uses ZIP64, and avoids the ZIP32
size/count refusal for an already-ZIP64 source. The source-backed bounds were
adjusted on the same basis. It adds successful synthetic OPC tests:

- `zip64_source_targeted_save_preserves_untouched_records_and_tail` changes
  one Part in a ZIP64 package, retains a different Part's raw local and central
  records, retains ZIP64 tail extensible bytes and the comment, reopens the
  output, and compares `to_bytes` with `write_to_stream`.
- `projected_zip64_descriptor_targeted_save_preserves_untouched_member`
  changes one Part while a different member uses ZIP64 data-descriptor fields,
  then checks raw preservation, reopen, and stream output.

Focused source-backed coverage now also includes
`zip64_one_part_overlay_preserves_raw_unknown_members_and_cold_work` and
`topology_zip64_add_remove_preserves_untouched_records_and_tail`. Both pass in the final integrated test run. The
package-writer suite also has a ZIP64 partial-sink failure case.

These tests are meaningful end-to-end package-writer correctness evidence, but
they are synthetic in-memory fixtures. They leave the following gates open:

1. Extend the source-backed selected-Part and topology matrix to signed/encrypted inputs,
   both descriptor forms, non-seek sinks, cancellation, configured metadata
   limits, and atomic filesystem finalization. Each refusal must leave the
   sink/destination untouched where the API promises that property.
2. Exercise ZIP64 add/remove/topology plans with more than the current narrow
   synthetic graph, including unknown physical members and dependency closures.
   The existing focused add/remove test is valuable correctness evidence but
   does not certify every topology operation.
3. Resolve the output-promotion boundary. The preservation writer still emits
   generated central local-offset fields as fixed 32-bit fields and rejects a
   generated ZIP64 member (`preserve.rs` around `generated_entry`). An existing
   ZIP64 tail can therefore accept only cases whose generated offsets and
   records remain representable. Either implement a validated ZIP64 generated
   record/promotion path or retain and test the typed refusal; do not infer
   general large-file support from the current tail-preservation tests.
4. Add at least one real-producer or independently generated large ZIP64 OPC
   corpus and validate all semantic Part bytes, raw member identity, archive
   layout, and reopen behavior. The current two-Part fixture does not exercise
   large offsets or large generated payloads.
5. Expand the scoped gates in
   [`changes/0404-zip64-preservation-integration.md`](changes/0404-zip64-preservation-integration.md)
   to the full non-iWork feature and native-producer matrix. The scoped final
   commands, source bindings, and pass/fail logs are now retained.

The control behavior is also important: at the control revision, a changed
owned ZIP64 source was refused rather than normalized. That refusal remains the
safe fallback for unsupported framing, suffixes, opaque topology, and
unrepresentable generated output. A successful targeted test does not authorize
a normalizing fallback for a different unsupported source.

## Prioritized remaining work

| Priority | Requirement from `docs/GOAL.md` | Next reviewable evidence |
| --- | --- | --- |
| P0 | Close the ZIP64 integration and preservation contract | 0404 binds the scoped all-feature results and fixture-generator sources; generated-offset/promotion and broader failure-atomicity coverage remain open. |
| P0 | Establish the Phase-1 baseline before selecting further optimizations | Build the release harness at the current revision and capture named non-iWork corpora with p50/p95/p99, throughput, `perf stat` counters, allocation/peak RSS, source calls/bytes/ranges, decompressed/recompressed/copied bytes, output bytes, and lock-wait fields. Run cold/warm and one-worker/explicit-bounded-worker cases. |
| P0 | Turn cache correctness into accepted observation | 0405 now retains validated direct-lock acquisition observations, cache counters, and explicit timing scope in 24 normal and 24 observed smoke rows. The release harness passes 256 tests with one ignored. Extend this descriptive evidence to representative workloads with operation-local attribution. Keep cache observations separate from latency claims until an accepted ABBA protocol exists. |
| P0 | Capture current hardware/resource evidence | Use the now-available `perf stat`, Heaptrack, strace, fincore, and cargo-flamegraph tooling in the planned 0406 release capture; bind machine, corpus, revision, and uncertainty before accepting any new claim. |
| P1 | Finish source-backed OPC CRUD adoption | Extend selective open/read/edit/save across format facades and topology changes; measure the selected-Part/compressor buffer, physical I/O, allocation/RSS, and semantic phase boundaries. The current report calls this migration incomplete. |
| P1 | Cover the high-impact CRUD categories | Add or explicitly classify conversion, stream append, structural edits, deletion/sanitization, cross-document dependency copy, merge/split, patch and inverse timing, repair/normalize, dynamic calculation, security, malformed, and real-producer scenarios. Correctness-only selectors cannot close the timing requirement. |
| P1 | Finish CFB consumer evidence | Move exact-range and overlay substrate work into DOC/XLS/PPT semantic owners; add physical-cold and high-latency range-source cases, FAT-tail behavior, and operation-local memory/resource attribution. |
| P1 | Validate output and source variants | Cover filesystem cold cache, caller-supplied range/remote sources, sequential non-seek sinks, cancellation, source-version changes, signed/encrypted/macro-enabled files, and atomic replacement across the non-iWork formats. |
| P2 | Demonstrate bounded parallelism | Use explicit execution contexts, then publish scaling curves and serial-fraction/Amdahl analysis. Lock-wait data is currently missing from the accepted report. |
| P2 | Consider layout and SIMD changes | Only after the release profiles identify a hot loop and show a material benefit with scalar fallback and differential correctness tests. |

No row in this audit should be read as a completion claim. The current
evidence supports targeted capability progress and an integrated ZIP64 change
under final review; it does not close the full non-iWork performance goal.
