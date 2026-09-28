# 0811 driver and custody preflight review

This is a read-only review of the 0811 plan, custody helper, root-owned
drivers, and protocol review. It ran no Cargo command, probe, workload, perf
record, decode, or packet driver. The only independent checks here are static
file/hash comparisons and arithmetic over the declared plan.

## Source and inherited inputs

The current HEAD is `2894fbd628434bad4ef45af4f0f0b9a8d4d23468`, matching
`plan.json` and `origin.json`. The current tracked production set contains
9,196 files. Comparing every current file hash with the sealed 0810 after
manifest [`build-after/source.json`](../change-0810/build-after/source.json)
(`6f90c90e816c0fb9bd2cd03e65f28e9fa485d8157e53fd9bae4b0c7cc87b515d`) found
zero differences. The revision field advances from the 0810 after commit as
expected; the file content is identical.

The six files under [`probe-src`](probe-src) are byte-identical to the six
0810 probe files and match the hashes recorded in `inheritance.json`. The
copied root `Cargo.lock` and `rustfmt.toml` also match the live root inputs.
The probe custody helper now checks exactly six files against the sealed 0810
hash map, and every build/capture/quality/decode receipt rechecks the frozen
probe and root-input identities. `build.py` also compares the complete live
source map with the sealed 0810 after map before writing its pre-build frozen
inputs. The zero-difference comparison above independently confirms both
early-fail guards.

## Build variants and freeze boundary

The declared three variants are coherent:

* `control` uses no feature and no extra rustflags;
* `profile` uses `capture-profile`; and
* `fp` uses `capture-profile` plus
  `-C force-frame-pointers=yes`.

`build.py` rejects inherited Rust/Cargo overrides, uses one owned target,
offline locked release builds, two Cargo jobs, and disabled incremental build.
It writes `build/frozen-inputs.json` before the first Cargo invocation and
rechecks source, probe, root inputs, architecture hashes, and unrelated-file
hashes after every build. The 35 architecture inputs and three unrelated
workspace identities are bound by the current packet. `host.json` and
`toolchain.json` are included in the listed driver hashes and were captured
before the build boundary.

The measurement custody binds the plan, origin, inheritance, architecture
manifest, host/tool metadata, custody, build, capture, decode, quality, probe
quality, and root input copies. Offline analysis, root-audit, cleanup, seal,
and final validation are completion-time readers and witnesses; they do not
gate the start of measurement. The current `analysis.py` also independently
checks sealed source and probe equality after capture. The final `validate.py`
may be added by the root before the final audit/seal as stated by the
protocol.

## Measurement protocol

The native schedule is six explicit orders over `control/profile/fp`, crossed
with tiny, medium, and large shapes. The arithmetic is 6 × 3 × 3 = 54
reports, with 30 measured samples and three warmups each, for 1,620 samples.
The capture driver derives those rows from the six plan orders and records the
binary, source, probe, input, architecture, unrelated-file, log, report, and
RSS identities for each process.

The perf schedule is two large-capture runs of the `fp` binary, each with 100
operations and zero warmup. The plan binds CPU 12, `cycles:u`, 499 Hz, frame
pointer call graphs, and the exact owner
`namespace_uri_probe::capture_region_0793`; this adds 200 outputs for the
declared 56-report/1,820-sample packet total. The probe source contains the
same `#[inline(never)]` capture wrapper and the profile feature selects that
owner. `root_audit.py` is the required independent check for decoded owner
stacks, unresolved interiors, lost events, and overlapping nested counts; its
result must be available before cleanup or any interpretation.

The plan and protocol explicitly prohibit a historical timing pool. The
offline root-audit oracle imports only sealed 0810 semantic fixture/output
identities and does not import elapsed samples. No allocation or Callgrind
lane is declared, no adoption threshold is set, and no speedup or causal
phase claim is authorized.

## Quality, decode, and retention

The production quality driver reuses the six declared `litchi-pptx` gates
(format, check, test, Clippy, docs, and crate-boundary checks) against the
unchanged current source. The probe quality driver runs its three fresh gates
and requires exactly 36 passing tests, with no failures. Both lanes recheck
source, probe, root-input, architecture, and unrelated-file custody after
each command.

`decode.py` requires native completion first, verifies all three exact binary
identities while the owned target still exists, decodes both raw perf files
with `perf script --no-inline --ns`, and checks source/probe/input custody
around each decode. It writes deterministic gzip members with `mtime=0`,
retains original and compressed artifact identities plus decompressed SHA-256
values, and removes the uncompressed raw/frame files only after those checks.
`cleanup.py` removes the owned target only after the binary identities and
independent analysis witnesses are present, then rechecks the source.

## Readiness disposition

The source, six-file probe inheritance, three-build feature split, root-input
copies, native/perf schedule, exact wrapper scope, binary-live decode order,
gzip identity policy, no-historical-timing rule, and quality/test counts are
aligned with the root protocol. With the exact sealed-probe assertion now in
`custody.py` and the sealed-source assertion in `build.py`, the measurement
drivers **PASS this static review** and are ready for the root's independent
final pre-build review. Offline readers, cleanup, and final validation remain
completion tasks as specified by the protocol. No performance, semantic, or
adoption conclusion follows from this static preflight.
