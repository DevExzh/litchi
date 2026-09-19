# `probe0693`

This directory is a standalone, retained scratch probe for the native
opened-PresentationML edit phases used by change 0693. It does not modify
production sources or provide a production performance claim. The default
binary has no counting allocator.

The probe materializes a source archive before timing. A source is either a
`.pptx` path or a deterministic `generated:<slides>x<text-boxes>` specification;
the bounded generated corpus used by the packet is `generated:12x8`. File reads,
generated-package construction, `Package::from_bytes`, target discovery, the
correctness oracle, serialization, and semantic reopen are outside the native
sample timers.

## Build

From the repository root, build the native and diagnostic binaries separately:

```sh
cargo build --release --locked \
  --manifest-path docs/performance/results/change-0693/probe/Cargo.toml \
  --target-dir "$TARGET_DIR" -j 2
cargo build --release --locked \
  --manifest-path docs/performance/results/change-0693/probe/Cargo.toml \
  --target-dir "$TARGET_DIR" -j 2 --features allocations
```

The package declares its own empty `[workspace]`, so Cargo does not treat this
probe as a workspace member. The binary name is `probe0693`.

## Modes

```text
probe0693 phases <source|generated:SxB> <samples> [warmups [one|noop|two]]
probe0693 noop <source|generated:SxB> <samples> [warmups]
probe0693 two <source|generated:SxB> <samples> [warmups]
probe0693 prefix <source|generated:SxB> <stage> <iterations>
probe0693 shape <source|generated:SxB>
probe0693 target <source|generated:SxB>
```

The original one-edit workflow remains the default. The optional final
`phases` argument selects `one`, `noop`, or `two`; the standalone `noop` and
`two` commands are equivalent explicit forms. A no-op restores the selected
shape's existing text, requires `commit.is_changed()` to be false, and checks
that the published revision and serialized output are identical. The two-edit
workflow selects admitted text shapes on two distinct slides, times both
`set_shape_text` calls in one `settext_ns` field, and validates both markers
after reopen. Its untouched-payload oracle excludes exactly those two slide
parts; every other part, relationship, content type, and non-part member stays
covered by the same preservation check as the one-edit workflow.

`phases` performs the requested warmups, then retains all requested raw samples
and prints them after the timing loop. Each sample starts from the same source
bytes and times these public operations separately:

```text
capture  opened_presentation
clone    Snapshot::edit
settext  set_shape_text
commit   Transaction::commit
apply    apply_opened_presentation_commit
total    the enclosing interval from before capture through apply
```

The total interval includes the phase timer and checking overhead between the
calls. It excludes package open, save, reopen, and the correctness oracle. No
output is written inside the measured loop. The phase table is raw TSV with
metadata lines first:

```text
sample  capture_ns  clone_ns  settext_ns  commit_ns  apply_ns  total_ns
```

The actual output is tab separated; `sample` is a zero-based row index. The
probe reports source archive byte count and SHA-256, target coordinates, sample
and warmup counts, before/after opened snapshot revision hashes, logical
semantic hashes, candidate archive size/hash, and the SHA-256 of the reopened
target marker. The semantic digest covers slide indices, slide names and text,
shape indices, shape names, and shape text through the public API. A validation
failure terminates the run instead of producing timings that lack the selected
target's semantic reopen check.
It also asserts revision change, preserved slide/shape counts, and unchanged
part inventories, content types, relationships and non-edited payload hashes
both before save and after reopen. The edited part's unknown markup and ZIP
physical metadata still require the library's existing preservation tests.

`prefix` executes complete bounded prefixes (`open`, `capture`, `transaction`,
`settext`, `commit`, or `apply`) for profiling controls and prints one completion
record after the loop. `shape` prints a bounded archive census; `target` prints
the first admitted text-shape target used by the timed edit.

## Allocation companion

Build with `--features allocations` to install a measurement-only wrapper around
the system allocator. It accepts the same command line, including:

```sh
probe0693 phases generated:12x8 100 5
```

The phase TSV appends seven fields for each of `capture`, `clone`, `settext`,
`commit`, `apply`, and `total`:

```text
<phase>_alloc_calls
<phase>_requested_bytes
<phase>_baseline_live_bytes
<phase>_peak_live_bytes
<phase>_current_live_bytes
<phase>_realloc_calls
<phase>_realloc_requested_bytes
```

Only successful `alloc`, `alloc_zeroed`, and `realloc` operations contribute to
the request counters; a successful realloc records its new requested size and
adjusts live bytes by the old/new size difference. Deallocation adjusts live
bytes. Baseline, peak, and current are absolute process live-byte values at the
phase boundary and observed phase peak, so a net release remains visible. The
allocator counters are descriptive diagnostics and are not native timing
claims. Installing the wrapper adds atomic counter operations and can perturb
timings; compare these runs only with other allocation-feature runs.
