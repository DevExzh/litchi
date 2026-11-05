# 0362ul89an — PPTX notes-root memo probe

This leaf adds an opt-in probe for the physical PresentationML notes-root
scan and the `SlideRootMemo` reuse path. It is harness-only: no public format
API, dependency, lockfile, default-case row, fixture binary, or production
implementation is changed.

## Probe

The selector is `pptx-notes-memo` in both `litchi-perf-baseline` and
`litchi-perf-baseline-alloc`. It requires a custody-bound 40-character source
revision and a create-new output path:

```text
litchi-perf-baseline pptx-notes-memo \
  --source-revision <baseline-or-candidate-revision> \
  --output <fresh-report.json> --samples 20 --warmup 3
```

The source code generator is `pptx-notes-memo-three-slide-v1`. Each run
creates three authored slides, each with the title `notes slide {i}`, one body
text box `slide body {i}`, and speaker notes `speaker notes {i}`. The report
records the resulting byte count and SHA-256; it does not write the fixture to
the repository.

Three rows are emitted. `cold_capture` clocks `Package::from_bytes` and
`opened_presentation`. `warm_notes_commit` opens outside the clock, changes
slide 1's notes to `rewritten notes`, and clocks the real detached commit;
all three slide roots remain unchanged and therefore remain memo-eligible.
`changed_slide_commit` opens outside the clock, changes slide 0's title to
`rewritten slide title`, and clocks the real commit; one slide root changes and
the other two remain eligible. Empty/no-op edits are rejected by the probe.

For every retained sample, serialization and an exact semantic oracle are
outside the clock. The oracle checks three slide texts, three notes texts,
the expected changed resource, commit diagnostics, and source-package
immutability for the two edit rows. The source-audited memo scope is explicit
in the report; proved-root counters remain an internal crate test and are not
fabricated as public benchmark evidence.

The report uses the existing harness `elapsed_ns` statistics object and
sample-index convention, but retains a local `pptx_notes_memo_probe_v1`
envelope. The generic `perf_abba_summary.py` has a closed catalog of the
default matrix and must not be fed this opt-in schema; the approved notes
ABBA wrapper should compare the three named rows by their `elapsed_ns.samples`
vectors and preserve the report's fidelity/allocation gates.

The allocator binary enters the shared non-nested operation-scoped allocator
region around exactly the same clocks. The normal binary emits explicit
`unavailable` allocation samples and native wall-clock timing. Allocator
elapsed samples are marked as instrumentation-only and are not latency claims.
No native timing, Callgrind, ABBA, Cargo build, or qualification run is part
of this source commit; those require the coordinator's frozen baseline and
candidate bindings and quiet-run protocol.

## Source binding and review boundary

The leaf was cut from `e80fcebe6bc2785e3b08bf415d5a545a2cca4e9a`, the current
reviewed PPTX candidate. The baseline API audit is pinned to
`76371266514feb010ef7478c26e963debcc7fa59`. The probe only calls the public
APIs present at both endpoints: `Builder`, `Package::from_bytes`,
`opened_presentation`, `Snapshot::edit`, `set_notes_text`, `set_shape_text`,
`Edit::commit`, `Commit::diagnostics`, `Package::presentation`, `Package::notes`,
and `Package::to_bytes`. No endpoint-specific compatibility alias or
production overlay is introduced.

This record is preparation evidence, not performance acceptance. The
coordinator must rebuild the normal and allocator binaries from each exact
endpoint, verify executable/source identities, and run the prescribed quiet
ABBA/Callgrind and fidelity protocol before drawing a performance conclusion.
