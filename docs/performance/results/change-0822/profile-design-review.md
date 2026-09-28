# 0822 PPTX real-file edit profile design

This is a design review only. It records the smallest credible diagnostic for
the PPTX edit CPU path at base `353aa00a7d` and authorizes no source change,
workload, release build, or profiler run. The packet must keep
`performance_claim: none`: sampled cycles and instrumented timings can locate
work, but they cannot establish a phase fraction, a speedup, or a production
candidate.

## Decision

Use a small standalone packet probe that calls the public PPTX API. Keep the
production ordinary-save harness and `ordinary_save.rs` byte-for-byte
unchanged. A packet-local copy of the complete ordinary-save harness would
duplicate private corpus and owner machinery and would make code drift harder
to detect. A production feature hook would change the source being profiled,
invalidate the current source custody, and add instrumentation to the code
whose ordinary behavior is being inferred. Neither is warranted for this
bounded attribution step.

The probe should materialize an independent Cargo manifest in the packet and
depend on the current workspace `crates/litchi-pptx` by path, as the 0811 and
0814 probes did. Its Rust source is a packet input and is checked independently
of the production source. A source review must compare the probe's five public
edit calls with the production sequence before any build is admitted.

## Current production sequence and timed boundary

`tools/perf-baseline/src/ordinary_save.rs` establishes the relevant contract:

* `Owner::open` opens the PPTX with `litchi_pptx::Package::open` at lines
  699–705.
* The PPTX edit branch creates
  `Package::opened_presentation_transaction`, calls
  `set_shape_text(slide, shape, EDIT_MARKER)`, requires `true`, commits the
  transaction, requires `Commit::is_changed()`, and applies the commit at
  lines 754–777.
* The real-file target is the first admitted position. For the retained
  `shapes.pptx` input it is `(slide 0, shape 0)`, and the marker is
  `litchi-perf-0638-ordinary-save`.
* The existing `Phase::Edit` opens the owner before the clock and clocks only
  `owner.edit`, at lines 1426–1446. It keeps the owner alive while the
  allocation region closes and drops it after the timer. This is the correct
  phase boundary, but that selector does not serialize an output during the
  edit phase and does not expose a stable public wrapper symbol for sampled
  ownership.

The probe's direct helper should therefore be semantically equivalent to this
sequence, with no target discovery or pre-capture in the timed interval:

```rust
fn edit_once(package: &mut Package) -> Result<()> {
    let mut edit = package.opened_presentation_transaction()?;
    if !edit.set_shape_text(0, 0, EDIT_MARKER)? {
        return Err("the selected shape did not change".into());
    }
    let commit = edit.commit()?;
    if !commit.is_changed() {
        return Err("the edit commit reports no change".into());
    }
    package.apply_opened_presentation_commit(commit)?;
    Ok(())
}
```

The pseudocode is a review oracle, not a source patch. In an actual sample the
probe must perform these steps in this order:

1. Open the canonical `test-data/ooxml/pptx/shapes.pptx` source with
   `Package::open` and finish all source validation outside the clock. The
   driver must bind that path to the frozen 0821 source copy and hash. Do not
   call `opened_presentation` during this preparation;
   `opened_presentation_transaction` must perform the same fresh capture that
   the ordinary-save owner performs.
2. Start the monotonic timer immediately before the first call to the edit
   helper. End it immediately after `apply_opened_presentation_commit` returns.
   The interval includes transaction capture, shape lookup and rewrite,
   transaction commit, patch validation, and package publication. It excludes
   path opening, corpus preparation, target discovery, output serialization,
   readback, digesting, verification, and package-owner destruction. The
   returned `Snapshot` is ignored and dropped inside the interval, matching
   `Owner::edit`.
3. After the timer stops, serialize the edited package with `Package::to_bytes`
   and own the resulting bytes before dropping the package. Perform all
   semantic and byte checks from those owned bytes; package destruction remains
   outside the timer and need not wait for those checks.
4. Bind the `to_bytes` result to the sealed 0821 default-save output. Since the
   production source and `PackageWriter` are unchanged at this base, the
   already admitted 0821 `Package::save` bytes are the transitive publication
   witness for this edit profile. Do not add a fresh `Package::save` or fsync
   lane here: this packet measures the in-memory edit and carries no new save,
   synchronization, or filesystem durability evidence.

The direct helper should not return a value that can be discarded before the
post-clock checks. The wrapper variant should retain the result, pass a
reference to `std::hint::black_box`, and return it. This prevents a tail-call
shaped wrapper and keeps the edit result observable without moving any output
or readback work into the measured interval.

## Corpus and parity contract

The packet must use the exact real file already admitted by 0819 and 0821. It
must not regenerate a semantic fixture or substitute a nearby PPTX.

| item | required identity |
| --- | --- |
| source | canonical `test-data/ooxml/pptx/shapes.pptx` and frozen 0821 `artifacts/real-002-pptx/source.pptx`, 68,822 bytes, SHA-256 `19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571` |
| edit | `Package::opened_presentation_transaction().set_shape_text(0, 0, "litchi-perf-0638-ordinary-save")` followed by commit and apply |
| admitted output | 68,284 bytes, SHA-256 `38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf` |
| source archive | 48 members; the source and output identities are those retained by 0819/0821 artifact and ZIP-preservation admission |

The expected output should be supplied as a frozen packet input from
`docs/performance/results/change-0821/artifacts/real-002-pptx/default.pptx`.
The driver must hash both input files before the first Cargo or workload
command and must recheck them after every child. Every measured sample must
require the complete serialized byte vector to equal the expected output
vector. A per-sample SHA-256 and byte count may be recorded as a compact
receipt, but it is not a substitute for the in-process byte equality check. A
digest-only result must not silently become a semantic-only result.

The output oracle has three layers:

1. Compare the serialized bytes with the admitted output identity above. This
   retains member order, ZIP metadata, untouched compressed payloads, and
   regenerated slide bytes through the already admitted complete archive.
2. Reopen the serialized bytes with the public PPTX reader and compare the
   complete presentation text with the text obtained once, untimed, from the
   frozen admitted output. Also check that slide 0 shape 0 contains the marker
   and that the marker is not accepted merely because it appears in an
   unrelated member. The semantic check is independent evidence alongside the
   byte check; it is not a replacement for it.
3. In every qualification and native sample, compare the complete `to_bytes`
   result and its reopened semantic text to the same expected output. The
   output is the sealed 0821 default-save artifact; this transitive binding is
   sufficient for an edit CPU profile because production and its writer are
   unchanged. No destination path, rename, sync, or fresh save claim belongs
   in this packet.

The expected semantic text must be derived from the retained admitted output,
not reconstructed from a generated fixture. The probe may use
`Package::presentation().text()` and the public slide/shape view for these
untimed checks. It must not call either view before the timed edit on the
sample package, because doing so would pre-capture or cache work that belongs
to the public edit operation.

## Variant design and perturbation accounting

Build two release executables from the same frozen probe source and manifest.
The ordinary executable exposes two CLI arms; the frame-pointer executable
exposes the wrapped arm:

* `ordinary/control` (CLI mode `direct`): the ordinary release build whose
  timed call invokes the common edit helper directly;
* `ordinary/wrapped`: the same ordinary build's small `#[inline(never)]` wrapper arm.
  The wrapper retains the result, calls `black_box(&result)`, and returns it;
  this is the wrapper-control arm;
* `fp/wrapped`: the same wrapped arm rebuilt with only
  `-C force-frame-pointers=yes`, so frame-pointer call chains can be sampled
  and decoded.

The common helper should be `#[inline(always)]` only where needed to make the
direct and wrapped bodies contain the same public operation. The wrapper must
not add a second edit, pre-capture, output, or readback. Its `black_box` is
part of the wrapper instrumentation and must be named as such in the report.
The direct arm is a control for the wrapper boundary; it is not a claim about
the inlining decisions of the ordinary-save enum dispatch.

Use six counterbalanced native blocks on one pinned CPU. Each block runs all
three arms in a frozen order, with three warmups and thirty measured samples:

```text
[control, wrapped, fp]
[wrapped, fp, control]
[fp, control, wrapped]
[fp, wrapped, control]
[wrapped, control, fp]
[control, fp, wrapped]
```

This is 18 native reports and 540 measured samples. Report nearest-rank
within-process quantiles, matched-block p50 ratios, spread and tail flags, and
all raw samples. Bootstrap only the matched six block p50 ratios, using a
fresh packet seed and sorted endpoints recorded in the frozen plan. Ratios
have diagnostic meanings:

* `wrapped/control` (ordinary/wrapped divided by ordinary/control) estimates
  the timing perturbation of the non-inlined owner wrapper and its result
  barrier;
* `fp/wrapped` (fp/wrapped divided by ordinary/wrapped) estimates the added frame-pointer build
  perturbation;
* neither ratio is a source speedup, an ordinary-build phase fraction, or a
  reason to pool the wrapped values with 0819/0821 latency.

Do not add the 0819 process-counter/allocator observer feature to the timed
edit merely to obtain another column. Procfs snapshots and a counting global
allocator alter the process and are not useful for identifying the CPU leaf in
this short operation. If an allocator or procfs arm is retained for a separate
diagnostic, it must use the same start/stop boundary, be reported outside the
native timing table, and receive its own paired observer/control ratio. It
must never be subtracted from or pooled with the three native arms.

The external `perf record` observer has no credible operation-local latency
correction. Its effect should be bounded by recording the exact command,
frequency, CPU, raw-data size, lost-event count, unresolved-frame count, and
the corresponding unprofiled `fp` native timing arm. The profiled process
elapsed time is a diagnostic about the observer and whole-process work, not a
replacement for the native edit timer.

## Owner-leaf profile

Run two serial `fp` profile processes, each with the `wrapped` arm, no warmup,
and 2,000 edit samples. Use `perf record -e cycles:u -F 997 --call-graph fp`
on the same pinned CPU and retain the raw data until decoding is complete. The
profile process still opens, serializes, reopens, and verifies each sample
outside the edit clock; this is intentional for parity, but whole-process
stacks must not be called edit-only costs.

Decode against the exact `fp` executable hash before cleanup. Require an exact
demangled/mangled wrapper symbol and count:

* whole-process samples;
* samples whose stack contains the exact wrapper symbol;
* wrapper self leaves and child leaves below the wrapper;
* unresolved or unknown interior frames;
* lost-event diagnostics and any stack truncation;
* samples outside the wrapper.

Keep inclusive counts separate from self-leaf counts. Nested frame counts
overlap and cannot be added into a phase fraction. The profile is useful only
if the exact wrapper symbol is found in the binary, the owner sample count is
nonzero and decoded stacks conserve the selected-owner versus other-owner
partition. Retain unresolved and lost-event evidence rather than dropping it
to improve apparent coverage. A low owner count or a wrapper that disappears
into a tail call is a failed attribution gate, not permission to broaden the
symbol filter.

The profile owner name is the exact packet-specific symbol
`pptx_edit_profile_0822::edit_region_0822`. It must not be a short suffix that
could match an OPC or another format's function. The assembly/symbol receipt
should bind the exact symbol range to the exact `fp` binary hash, following the
qualified selector recovery used by 0814.

## Admission and stopping rules

Before any native or profile workload, run the probe's formatting, locked
all-feature check, tests, warning-denied Clippy, and rustdoc gates. Reuse the
committed production quality result only after the current source, lock graph,
normative inputs, and 0821 seal are revalidated. The probe quality result is
fresh; production quality reuse must remain labeled as reuse.

The packet should have three qualification reports (one per arm), each with
three samples, for nine qualification samples total. Qualification uses the
same timed edit and `to_bytes` parity oracle; it does not exercise `save`,
rename, or fsync. Native timing begins only after qualification passes. The
unprofiled lane is 21 reports and 549 samples (three qualification reports / 9
samples plus 18 native reports / 540 samples). Add two fp/wrapped profiler
reports of 2,000 samples each, for a final total of 23 reports and 4,549
samples, including 4,000 profiler samples. The plan must keep the 4,000
profiler samples separate from the 549 unprofiled timing samples.

Stop the packet without a profile interpretation if any of these occur:

* the source or admitted output hash differs;
* the exact public edit sequence refuses, reports no change, or changes its
  marker/target;
* any qualification or native output differs from the admitted bytes or
  semantic text;
* the direct and wrapped arms do not have the same semantic/output witness;
* the wrapper symbol is ambiguous, absent, tail-call-shaped, or has no
  qualified owner samples;
* perf loses events or leaves unresolved interiors at a level that prevents
  an honest owner-leaf table.

No source candidate, optimization, allocation saving, cross-format inference,
durability inference, or historical timing comparison follows from this
packet. If the profile passes, it identifies a bounded next source review; a
future implementation still needs its own ordinary source change, full
semantic and preservation admission, and end-to-end workflow measurement.

## Evidence lineage

The design reuses the timing-boundary discipline and exact wrapper/frame
pointer controls of the 0811 and 0814 PPTX attribution packets. It binds its
real-file and output contract to the admitted 0819 ordinary-save baseline and
the 0821 durability packet. Those earlier raw captures and outputs remain
immutable and are not pooled with 0822. The broad non-iWork performance goal
remains active; iWork is outside this packet.
