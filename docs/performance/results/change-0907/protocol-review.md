# 0907 PPTX multi-slide capture protocol review

This review is read-only evidence for the finite current-source capture selected
by the 0906 OOXML audit. It does not run a build or workload and does not change
the PPTX runner.

## Source binding

The capture source must be the frozen `3d3650ac03be2cd8ba923a5a6e940e84b14f5eab`
source snapshot. In the current shared checkout,
`tools/perf-baseline/src/lib.rs` is SHA-256
`f712da842169399d687eb452c7042deed61682a80efdd40b2c3c178af2c036ea`, and the
same path at that commit has the same hash. A direct check is:

```sh
test "$(sha256sum tools/perf-baseline/src/lib.rs | cut -d' ' -f1)" = \
  f712da842169399d687eb452c7042deed61682a80efdd40b2c3c178af2c036ea
git diff --quiet 3d3650ac03be2cd8ba923a5a6e940e84b14f5eab -- \
  tools/perf-baseline/src/lib.rs
```

The earlier 0906 opportunity note binds the pre-help-fix harness hash
`b62775d9…`. That note is a source-audit snapshot; it must not be used as the
binary or source identity for this capture. The help/test-only change in the
frozen commit does not alter the PPTX corpus builder, dispatch, runner, or
timer, but the report and receipt must bind the complete source snapshot to
`3d3650ac03`, rather than silently reusing the earlier hash.

Build the frozen source in an isolated source/build directory and invoke the
resulting release binary directly. The retained build identity is Rust 1.95.0,
release optimization level 3, fat LTO, one codegen unit, debug information
enabled, and two Cargo build jobs. Do not build or run from the dirty shared
iWork checkout. The four reports must use one binary SHA-256 and one source
archive/lockfile receipt.

## Workload and timer

The two cases are:

```text
pptx_eager_multi_slide_batch_edit_save
pptx_source_backed_multi_slide_batch_edit_save
```

Both dispatch through `build_pptx_source_edit_corpus` and the fixed
`litchi-pptx-source-edit-media-v1` corpus. The corpus contract is 200 slides,
eight text boxes on every slide, eight deterministic incompressible 2 MiB PNG
Parts, 229 OPC Parts, and 445 physical ZIP members. The retained source is
17,017,139 bytes with logical Part bytes 17,568,429 and source SHA-256
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`.
The eight edited slide positions are exactly
`[0, 28, 57, 85, 114, 142, 171, 199]`; all eight text boxes on each selected
slide are replaced.

In `run_pptx_multi_slide_batch_edit_save`, setup, expected-output construction,
corpus checks, payload-range construction, source/sink allocation and sink
reservation are outside the clock (`lib.rs:51122-51172`). The elapsed interval
starts immediately before `publish_pptx_multi_slide_batch_edit` and ends after
that call (`lib.rs:51172-51175`). The following work is therefore inside the
reported `elapsed_ns`:

* Source-backed mode opens the source-backed editor, stages 64 shape-text
  replacements, commits, applies the patch to its source and its inverse for
  the pre-publication check, publishes the batch to the sink, then applies the
  inverse to the published snapshot and checks that reapplying the original
  patch is refused (`lib.rs:51050-51081`).
* Eager mode opens and fully materializes the OPC package, stages the same 64
  replacements, commits, publishes, applies the inverse and then its forward
  counterpart, and serializes the package in 64 KiB-or-smaller sink writes
  (`lib.rs:51084-51119`).

The source-backed and eager guard sequences are related but not identical. In
source-backed mode, forward-then-inverse succeeds before publication; after
publication, inverse succeeds and the original forward patch is expected to
fail as stale. In eager mode, the post-publication inverse followed by its
forward counterpart is expected to succeed. The source-backed post-publication
forward refusal must be reported as a passing gate, not as a failed forward
replay. Do not subtract or normalize these guards after the fact, and do not
describe this as a pure edit or publication timer.

After the clock stops, each sample is compared with the mode-specific expected
output, semantically reopened, and checked for package topology, relationships,
content types, selected/unselected Part identity, and media payloads. Digest,
source counters, sink limits and sample retention are also checked
(`lib.rs:51177-51198`). These are correctness gates and are excluded from the
latency interval. The timer does not measure physical filesystem I/O, a cold
filesystem cache, network/range latency, native Office behavior, RSS, or
allocator activity.

## Finite ABBA capture

Use four serial process legs on one fixed CPU in this order:

```text
A1 = eager multi-slide
B1 = source-backed multi-slide
B2 = source-backed multi-slide
A2 = eager multi-slide
```

Each leg should use 30 untimed warmups and 200 retained samples. Invoke the
binary directly with the parser’s actual JSON option; the harness accepts
`--json PATH`, not `--output PATH`:

```sh
taskset --cpu-list 12 "$BIN" \
  --workers 1 --warmup 30 --samples 200 \
  --case pptx_eager_multi_slide_batch_edit_save \
  --json "$OUT/pptx-eager-a.json"

taskset --cpu-list 12 "$BIN" \
  --workers 1 --warmup 30 --samples 200 \
  --case pptx_source_backed_multi_slide_batch_edit_save \
  --json "$OUT/pptx-source-a.json"

taskset --cpu-list 12 "$BIN" \
  --workers 1 --warmup 30 --samples 200 \
  --case pptx_source_backed_multi_slide_batch_edit_save \
  --json "$OUT/pptx-source-b.json"

taskset --cpu-list 12 "$BIN" \
  --workers 1 --warmup 30 --samples 200 \
  --case pptx_eager_multi_slide_batch_edit_save \
  --json "$OUT/pptx-eager-b.json"
```

Use fresh, non-existing output paths and retain each command, start/end time,
environment, CPU placement, stdout/stderr, binary hash, source archive hash,
and lockfile hashes. No other workload should run concurrently. Keep the four
legs separate in the evidence; summarize each leg before combining A and B.
The harness's sorted statistics include `elapsed_ns.sample_order`, so any
per-sample source or sink vectors must be checked against that order before
computing ratios. In particular, this runner records the PPTX `SourceSummary`
vectors in measured-iteration order and does not reorder them before returning
the `CaseResult`; do not pair a source-read or materialization vector by
position with the sorted `elapsed_ns.samples` vector without applying
`elapsed_ns.sample_order`.

## Required result checks

Every leg must pass all of these checks before its timing is admitted:

1. **Report shape and provenance.** The JSON has schema version 1, exactly one
   result, the requested case name, 200 elapsed samples, and configuration
   values `samples_per_case = 200` and `warmup_iterations_per_case = 30`. The
   binary identity is identical across all four reports. The source receipt
   binds commit `3d3650ac03`, the frozen source archive and both lockfiles.
   There must be no accidental second selector or filesystem-cache result.

2. **Corpus identity.** The result's corpus manifest is exactly the fixed
   `PPTX/OPC/ZIP`, `media-rich`, deflate corpus with generator
   `litchi-pptx-source-edit-media-v1`, 229 OPC Parts, 445 ZIP members, and the
   source byte count/hash above. The target entry is slide 100, shape 0 under
   the zero-based manifest naming used by the runner. Do not overwrite a
   changed source or output hash with the old retained value; if a fresh build
   produces a different identity, stop and re-pin the corpus explicitly.

3. **Mode-specific expected output.** The runner constructs an expected output
   independently for each mode before sampling. For every sample, require
   byte-for-byte equality with that mode's expected output, the mode's stable
   `output_sha256`, and the mode's accepted sink byte count. The retained audit
   records 17,017,145 bytes and SHA-256
   `23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9` as the
   expected identity, but that value is an audit anchor, not permission to
   assume eager and source-backed outputs are identical. Compare their expected
   bytes and hashes independently and record cross-mode equality only if it is
   observed.

4. **Semantic and OPC preservation.** Reopen every measured output and require
   200 slides in the original order; eight text shapes on every slide plus the
   media shape on each of slides 0 through 7; exact updated text on all 64
   selected shapes; and original text everywhere else. Require unchanged Part
   count, package relationships, content types and every unselected Part blob.
   Require all eight deterministic media payloads to match exactly. Require
   the source archive to remain unchanged. These are the checks implemented by
   `verify_pptx_multi_slide_edit_output` (`lib.rs:50711-50756`), and their
   successful status must be retained with the report.

5. **Patch and materialization contracts.** Require the source-backed mode to
   materialize exactly 9 Parts per sample (root plus the eight selected slide
   Parts), and eager mode exactly 229. Every measured sample must show nonzero
   ordinary source-read calls and bytes. Require the mode-specific patch guards
   described in the timer section, including the expected stale refusal in
   source-backed post-publication validation.

6. **Sink and determinism.** Require accepted bytes to equal the mode-specific
   expected output length, largest logical write no greater than 64 KiB, and
   deterministic sink summaries across measured samples. Every measured digest
   must equal that mode's expected digest. A logical `ReadAt` counter is a
   source-access count only; label it as such rather than as disk or network
   I/O.

7. **Raw ZIP preservation boundary.** The current runner verifies the logical
   OPC graph and payloads, but `verify_pptx_multi_slide_edit_output` does not
   independently assert raw ZIP member order, local/central-directory metadata,
   or comments. If the capture is described as preserving the ZIP container,
   add an untimed raw ZIP census over the source and each mode's expected output:
   bind the 445-member ordered name set and compare the raw metadata/payload of
   every member whose preservation is promised. If that census is not retained,
   limit the result wording to semantic/OPC Part preservation and the guarded
   workflow timing.

## Interpretation

Keep p50, p95, p99, max, confidence intervals, the four leg identities, source
read vectors, materialization vectors and sink vectors. Summarize A1/A2 and
B1/B2 separately before any aggregate ratio. The old PPTX timings in changes
0044, 0063 and 0077 are historical ranking evidence only and cannot be used as
the before leg. A fresh ABBA result can qualify the current guarded workflow
and identify whether the old eager/source cost separation remains; it does not
by itself establish a release speedup, producer fidelity, no-op/failed-save
atomicity, cold/range behavior, or complete PPTX CRUD coverage. Those claims
require their own selectors and gates.

The controlling audit is
`docs/performance/results/change-0906/ooxml-opportunity.md:50-83,167-213`.
The 0906 report correctly limits the next measurement to this guarded
requalification; this review supplies the corrected source binding, `--json`
invocation, exact replay semantics, mode-specific output rule, and finite
ABBA admission gates.
