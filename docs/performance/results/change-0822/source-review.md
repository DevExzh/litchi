# 0822 source and capture protocol review

This is an independent static review of the packet-local probe and its
root-owned build, capture, quality, and decode protocol at base
`353aa00a7d`. I read the files in place after the probe was handed off and
after root's Rust 2024 formatter pass. I ran no Cargo command, workload,
release build, native capture, profiler, decoder, or packet reader.

The probe is the standalone `pptx-edit-profile-0822` binary in
`probe-src/Cargo.toml`. The manifest depends on the workspace
`crates/litchi-pptx` by path and keeps the probe source packet-local. The
packet-specific owner is
`pptx_edit_profile_0822::edit_region_0822`; the plan, capture driver, and
decode driver use that qualified name. The production source and
`ordinary_save.rs` remain outside the packet source census.

## Probe boundary

The public sequence matches `Owner::edit` in
`tools/perf-baseline/src/ordinary_save.rs`:

```text
Package::open(input_path)                                  outside timer
  opened_presentation_transaction()
  set_shape_text(0, 0, "litchi-perf-0638-ordinary-save")
  commit()
  apply_opened_presentation_commit(commit)                 timed
Package::to_bytes(), drop(package), identity, reopen,      outside timer
full presentation text, slide count, target shape checks
```

The production call sequence is visible at `ordinary_save.rs:754-777`, and
the production phase opens the owner before the clock and times `owner.edit`
at `ordinary_save.rs:1426-1446`. In the probe, `execute_one` opens a fresh
`Package` at `main.rs:295-305`, starts the monotonic clock at line 306, and
dispatches either the direct helper or `#[inline(never)] edit_region_0822` at
lines 307-310. The elapsed time is sampled immediately after that helper
returns at lines 311-312. The helper's commit value is discarded at the
`apply_opened_presentation_commit` statement at lines 279-283, so its
`Snapshot` drop is inside the timed helper, as it is in the production owner.

`Package::to_bytes` is called only after the clock at lines 314-317. The
package is dropped after serialization and before identity and semantic
verification at lines 318-319. The serialized byte vector is already owned,
so this package destruction is outside the timer and does not need to wait for
the later checks. No output, hash, semantic readback, or package destruction
is in the timed edit interval.

There is no sample-package pre-capture. `opened_presentation_transaction` is
the first presentation view used by the sample helper. Every warmup and
measured iteration calls `execute_one`, which opens a fresh package from the
path. The reference is prepared once, separately, by `build_oracle` from the
frozen reference bytes using `Package::from_bytes` at lines 328-368; that
package is not the package being edited. Initial input/reference reads and
their hashes occur before the warmup loop. The warmup package is still opened
fresh and its output is verified outside the edit clock.

## Output and semantic parity

The source now has the required full output check. `Oracle` owns a copy of the
entire admitted reference byte vector at `main.rs:90-99`, and
`verify_output` compares the complete output slice with it at line 382. It also
checks the pinned output size and SHA-256, the input and reference identities,
and an independently reopened presentation's complete text, text digest,
slide count, and target shape text at lines 377-421. The reference oracle
requires shape `(0, 0)` to contain the exact marker and requires the marker in
the complete reference text at lines 335-355. Every measured sample rejects
when `all_verified` is false at lines 210-219; warmups accumulate the same
verification result at lines 194-205. The focused source test also compares
direct and wrapped output vectors against the retained 0821 reference at
lines 710-758.

The packet identities are fixed to the canonical source
`test-data/ooxml/pptx/shapes.pptx` (68,822 bytes,
`19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571`) and the
sealed 0821 default output (68,284 bytes,
`38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf`). This is
an edit-profile parity witness. It does not add a save, rename, sync, or fsync
operation, and therefore does not make a new durability claim.

## CLI and input bounds

`parse_args` requires `--input`, `--reference`, `--samples`, and `--mode` and
accepts only `direct` or `wrapped`. Samples are bounded to 1 through 10,000;
warmups are bounded to 0 through 100. Unknown arguments, missing values, and
duplicate input/reference/samples/mode/output options are rejected. The probe
rejects equal input/reference paths and rejects an existing output path unless
the output is `-` for stdout. `read_pinned_file` requires a regular file,
checks metadata length, reads through `take(MAX_FILE_BYTES + 1)`, enforces the
32 MiB bound, and checks the expected byte count and SHA-256.

One minor CLI hardening gap remains: `--warmup` is assigned directly and a
duplicate `--warmup` silently replaces the prior value, unlike the other
single-value options. The bounds are still enforced. The focused test covers
unknown, samples, and warmup bounds, but not duplicate warmup, output-path
rejection, or the bounded-read path; those are useful quality additions if the
driver's review policy requires each rejection branch to be exercised.

## Binaries and sample matrix

`build.py` builds exactly two serial release binaries from the same manifest:
the `ordinary` binary with no extra `RUSTFLAGS`, and the `fp` binary with only
`-C force-frame-pointers=yes`. The plan maps the three CLI arms as follows:

| arm | binary | mode | purpose |
| --- | --- | --- | --- |
| `control` | `ordinary` | `direct` | direct public edit helper |
| `wrapped` | `ordinary` | `wrapped` | named non-inlined wrapper control |
| `fp` | `fp` | `wrapped` | frame-pointer owner attribution |

The qualification lane is three reports and nine samples. Native capture has
six counterbalanced blocks, 18 reports, three warmups per report, and 540
measured samples. The two profiler reports each request 2,000 wrapped samples,
for 23 reports and 4,549 samples in total; the unprofiled qualification plus
native lanes remain 21 reports and 549 samples. Native processes are pinned to
CPU 12 and receive `/usr/bin/time` maximum RSS receipts. The meaningful native
comparisons are `wrapped/control` and `fp/wrapped`; they bound wrapper and
frame-pointer perturbation and do not estimate a production speedup.

## Quality and capture protocol

`quality.py` replays the committed 0821 production result through the three
named 0821 loaders without Cargo, then runs fresh packet-probe `fmt`, locked
offline `check`, locked offline tests, warning-denied Clippy, and rustdoc. The
fresh test command is release, all-features, and single-threaded; Clippy uses
`--all-targets` and `-D warnings`, and rustdoc sets `RUSTDOCFLAGS=-Dwarnings`.
The production replay and the five fresh probe gates must remain separately
labeled.

`capture.py` launches each report as
`/usr/bin/time ... taskset -c 12 <binary> --mode ... --input ... --reference
... --samples ... --warmup ... --output ...`. The perf lane launches
`perf record -e cycles:u -F 997 --call-graph fp` around the `fp`/`wrapped`
binary for two serial repeats, retains each `.data` file and report, and
records a typed unavailable result with its reason if perf is absent or
permission-denied. The input, reference, source, probe, binary, command, log,
and RSS identities are retained in receipts. `validate_capture` now requires
the report-level `all_verified` result and the admitted output identity before
each receipt is accepted; the probe's per-sample verification still supplies
the byte-vector and semantic checks underneath that aggregate.

The handoff review also caught a receipt-shape detail: the offline lane reader
requires a frozen `block` field, and the capture driver must carry that field
from the counterbalanced loop into each receipt. Root's driver pass owns that
alignment before capture artifacts are admitted. This is a receipt-schema
item, not a probe timing defect.

## Exact owner and decode evidence

`decode.py` binds static symbol evidence to the FP binary descriptor copied from
`build.json`. It runs:

```text
nm -S --defined-only <fp-binary>
nm -C -S --defined-only <fp-binary>
objdump --disassemble=<matched-mangled-symbol> --wide --line-numbers <fp-binary>
```

It requires one mangled symbol containing `edit_region_0822`, exactly one
demangled line containing
`pptx_edit_profile_0822::edit_region_0822`, and an assembly listing containing
that exact mangled symbol. The receipt records the address, size, type, symbol
range, commands, binary hash, source/probe/build descriptors, raw `nm`, and
assembly outputs. This is the right assembly binding shape: the final profile
must filter the qualified symbol, not a short `edit_region_0822` suffix.

Perf decoding uses exactly `perf script --no-inline --ns` and retains raw data,
decoded frames, logs, and deterministic gzip copies. The decoder's
`frame-owner-counts.json` substring count is a preliminary presence sanity
check. The offline analysis owns the attribution table: its retained parser
counts whole-process samples and periods, exact owner stacks qualified to the
FP binary DSO, unattributed samples, unknown interior frames, missing
descendants, inclusive child frames, call paths, and ranked leaves. It also
retains lost-event, status-line, and malformed-frame diagnostics. This keeps
the decoder small while preserving enough raw evidence for the exact owner
gate and avoids treating a substring hit as a profile conclusion.

The handoff review caught a second receipt-shape detail: compressed members
must carry `repeat` and `kind` (`raw` or `frames`) alongside the label and
identity so the offline reader can match both retained members per profile
repeat. Root's final reader/driver alignment owns this field binding.

## Pre-freeze reader/driver reconciliation

The static probe source is ready for quality admission. During the handoff,
the following reader and receipt contracts were found to be stale relative to
the formatted probe and were handed back for final packet alignment:

* `main.rs:101-114` adds `output_bytes_verified` to `Verification`, while
  `analysis.py:578-583` omits that key and requires an exact key set.
* `main.rs:41-42` includes package-owner drop and the returned Snapshot in
  `timing_scope`, while `analysis.py:608-610` requires an older shorter string.
* The source's focused test asserts six slides at `main.rs:758`, while
  `analysis.py:629-632` requires `slide_count == 7`.
* The quality adapter and offline reader now use the packet's
  `production_reuse` and nested `probe` receipt, retaining the 641/0/1
  production test witness and five fresh probe gates. This was a pre-freeze
  key-name drift and is no longer a probe-source issue.
* The report reader must bind the formatted probe's
  `output_bytes_verified` verification field, the final `timing_scope` string
  (including the Snapshot drop detail), and the six-slide semantic oracle.
  These constants were stale in the first reader draft and are being
  synchronized before final offline analysis.
* Capture receipts must retain the block selector required by the frozen
  counterbalanced order, and compression rows must retain `repeat` and
  `kind` for both raw and frame members.

These are schema alignment details for the root-owned drivers/readers. They do
not change the measured edit boundary, and no capture or profile conclusion is
drawn from this source-only review.

## Static source approval receipt

Root's Rust 2024 formatter completed before this receipt was recorded. The
formatted packet probe identities are:

| input | bytes | SHA-256 |
| --- | ---: | --- |
| `probe-src/Cargo.toml` | 574 | `de52c5103c4ff76f6e2fdf8be432566f3ca2f65f4c92f99b7ebc830d01df5708` |
| `probe-src/Cargo.lock` | 30,217 | `b709bca72e65d4f0da1f016e672cc6b940a882ed1ed759970699b51b0edf8058` |
| `probe-src/src/main.rs` | 25,423 | `bb71ca8a0cc6c25ef86d40849aaddbbbd354e3d45fbcc4578f968e74f6205087` |

These hashes approve the packet-local source, manifest, and lockfile for the
five fresh quality gates. They do not approve a production source change or
assert that a workload, profiler, or reader has run.

## Verdict

The standalone probe is the smallest credible mechanism and its actual source
has the required public edit sequence, fresh per-sample package, exact timer
boundary, full output byte equality, and independently opened semantic
reference. The ordinary/FP two-binary, three-arm matrix and 23-report/4,549
sample accounting are consistent with the frozen plan. The formatted probe is
approved for root-owned quality execution; final capture and reader receipts
remain subject to the reconciliation items above. No production source edit,
save leg, fsync evidence, or historical timing comparison is admitted.
