# Change 0429 provider and native source review

This is a source-only review of the frozen provider, range-adapter, native-image,
and shared cache-diagnostic changes. I did not run Cargo, tests, profilers,
capture commands, or CPU workloads. The review therefore records source
boundaries and evidence-contract issues; it does not certify a runtime result.

The provider lifecycle has bounded input custody and ownership phases. The
range adapter reports logical caller reads with checked counters, and its
configured zero-duration delay remains semantically configured. The native
image lifecycle uses the POI fixtures as fixed producer inputs and keeps its
image payload alive through view and slide drops. The implementation and
scopes do not authorize a cold-filesystem, physical-I/O, native cross-copy,
general leak, or comparative performance claim.

## Resolved boundaries

- `stage_files` creates a unique task directory, creates each archive with
  `create_new`, flushes and syncs it, and removes the two exact files and the
  exact directory through `StagedFiles::Drop`. Staging happens once before
  provider warmups and the file provider reopens those paths per iteration.
  The file scope explicitly describes recently written warm files and makes
  no cold-cache claim.
- Bytes and file providers pass no synthetic range cap or delay. The range
  provider requires both controls, preserves an explicit `Some(0)` delay, and
  records a configured zero-duration delay on every nonempty delegated call.
  The report records the requested and adapter configuration and the verifier
  checks that they agree.
- `PptxRangeSource` uses checked atomic additions, fails closed after metric
  overflow, validates histogram totals, and computes checked counter and
  histogram intervals. Cache points validate their managed budget gauges and
  checked event deltas; unavailable owners serialize null diagnostics rather
  than fabricated zeros.
- The provider baseline is captured before either provider is constructed.
  Source construction, sink reservations, diagnostics, checks, and drops are
  outside the open/plan/publication clocks. Native metadata performs one
  `slide.images()` call inside its metadata clock and records the `selected`
  phase afterward. Its view-owned source `Arc` is moved into the view; a
  `Weak` observes one caller owner before the caller drop and zero afterward.
  Managed Memory, Objects, and Depth gauges are checked after image/source
  release.
- The native fixture hashes and payload oracles are independently derived from
  retained ZIP/XML inputs, and the original and LibreOffice-resaved shape
  fixtures remain typed refusal controls. The source and provider scopes keep
  native image evidence separate from synthetic cross-copy evidence.

## Findings requiring follow-up

### Blocking: native first read interval is not emitted

In `pptx_native_image.rs:625-640`, `read_point` computes the first available
snapshot with `previous.map(...)`, so the `opened` phase has no `delta` and no
`counter_delta_checked` marker after the baseline unavailable phase resets the
previous snapshot. `verify-report.py:697-711` requires every first available
read point to carry a checked delta from an explicit zero baseline. This also
contradicts the native report's `checked_phase_delta` contract. Seed the first
native delta from `PptxRangeSourceSnapshot::default()` (as the provider path
does), or change the verifier and contract together.

### Blocking for the declared native descriptor oracle: selected identity is incomplete

The independent sidecar and `NativeOracle` currently retain slide/image/count,
relationship ID, media part, and payload identity. `verify_descriptor` in
`pptx_native_image.rs:671-715` does not check the producer's shape position,
shape ID, name, bounds, or exact target content type. After
`read_image`, the returned `SourceImage::descriptor()` is not compared with
the descriptor verified by `slide.images()`. The native review declares exact
descriptor identity, so either add those independently derived fields and
check both descriptors, or narrow the report gate and documentation to the
currently checked relationship/part/payload facts.

The independently observed expected fields are: POI slide shape position 1,
ID 37890, name `Picture 2`, bounds
`[1115616, 2276872, 7017380, 3262858]`, content type `image/jpeg`; POI video
shape position 0, ID 2, name `file_example_MP4_480_1_5MG_Trim`, bounds
`[3810000, 2143125, 4572000, 2571750]`, content type `image/png`.

Until these two evidence-contract items are resolved, the source review does
not certify a native lifecycle report. No runtime or performance conclusion
follows from this review.

## Resolution after native source freeze

The two native evidence-contract blockers above were resolved before capture.
`run_iteration` now seeds the first retained source snapshot from
`PptxRangeSourceSnapshot::default()` before the timed open, so the `opened`
phase carries a checked zero-base delta and remains compatible with the
independent report verifier. The source still resets read availability after
the caller adapter is dropped.

The native oracle now carries and checks shape position, shape ID, shape name,
bounds, relationship, part, and content type. The selected descriptor is
checked against those fields, and the returned image descriptor is compared
with the selected descriptor and checked again after `read_image`. The retained
POI fixture sidecar contains the same independently derived descriptor fields.
The historical findings above remain as the record of the pre-fix review; they
no longer block native lifecycle capture. Runtime results still require the
frozen driver, custody, and replay checks.
