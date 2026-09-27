# 0784 native sampled follow-up review

This is a read-only qualification of the retained native follow-up evidence.
It records stack ancestry and custody checks only. It does not establish a
production performance change or a historical-regression cause.

## Builds and scope

The ordinary profile binary was built from the captured probe source with the
`capture-profile` feature and no `RUSTFLAGS`. Its frozen native plan uses
`cycles:u` at 499 Hz, CPU 12, `--call-graph dwarf,16384`, two large-shape
processes, 100 measured samples, and three warmups. The raw receipt commands
are in `perf/receipts.json` and the plan is `perf-plan.json`.

The separate frame-pointer diagnostic binary was built from the same source
manifest with:

```text
RUSTFLAGS=-C force-frame-pointers=yes
```

`build-fp/receipt.json` records a successful locked release build with the
`capture-profile` feature. It binds the binary and both profile runs to the
same source manifest hash as the ordinary build. The frame-pointer flag
changes optimized code generation, so this binary is suitable for diagnostic
stack recovery only. Its samples cannot be compared with the ordinary binary
as a latency or speedup measurement.

Both native lanes used the recorded command shape:

```text
taskset -c 12 perf record -e cycles:u -F 499 --call-graph <dwarf,16384|fp> ...
```

The command receipts, exit codes, binary identities, report identities, and
plan hashes are retained for both repeats. `perf-fp/complete.json` records two
completed processes. Each frame-pointer report has 100 measured records after
three warmups; all records pass the capture schema, source identity, output
identity, semantic check, and reopen check. The source, generated output, and
readback each have 215,220 bytes and SHA-256
`9c46542b763fc4bef63dfe4336cadd2bfba2b7e7b3f18a376c3924eb5643b3e9`.

The wrapper is used only by `run_capture` in the feature-enabled probe. The
probe constructs the package before entering the capture call and performs
serialization, reopen/readback verification, report writing, and retained
owner destruction after it returns. Native `perf` still samples the whole
process, so only a stack containing the exact capture owner is admitted to
the operation-local counts. The perf stream also includes the three warmups;
the stream does not mark warmup samples separately from the 100 measured
records.

## Raw and decode custody

The ordinary DWARF records completed successfully with 3,233 and 3,158
samples. Their retained `perf/{0,1}.data.gz` files decompress to the byte
lengths and SHA-256 values recorded for the original `.data` files. The
default `perf/{0,1}.decoded.gz` streams also decode successfully, but contain
zero exact occurrences of
`litchi_pptx::package::model::Package::opened_presentation_with_limits` in
either repeat. The successful raw record therefore does not qualify capture
ancestry for the ordinary DWARF lane.

The frame-pointer records have this independently checked custody:

| repeat | raw `.data` bytes | recorded/sample blocks | exact owner blocks |
| ---: | ---: | ---: | ---: |
| 0 | 605,064 | 3,287 | 1,162 |
| 1 | 604,576 | 3,286 | 1,153 |

For both repeats, `perf-fp/decode-receipts.json` records successful
`perf script --ns` decoding of the raw file, with an empty decoder log. The
retained `perf-fp/{0,1}.decoded.gz` files are the default inline decodes. The
canonical streams used for ancestry are
`perf-fp/{0,1}.frames.gz`, produced by successful
`perf script --no-inline --ns` commands. `compression.json` binds every
retained raw and default-decoded gzip to its original byte length and
SHA-256, while `frame-receipts.json` binds the canonical frame gzip to its
original. An independent decompression and hash replay passed for raw,
default-decoded, and canonical frame artifacts.

## Exact ancestry counts

Counts below split each canonical no-inline stream at its `profile-fp` sample
headers. A helper count is the number of owner-qualified sample blocks whose
stack contains that full symbol. The counts use full demangled names from the
canonical streams; no generic `fingerprint`, `resolved`, or other short-name
matching was used.

| repeat | all process blocks | exact owner | `scan_processed_xml` | `package_fingerprint_with_memo` | `resolved` | owner blocks with `[unknown]` interior |
| ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| 0 | 3,287 | 1,162 | 1,031 | 97 | 146 | 1 |
| 1 | 3,286 | 1,153 | 1,029 | 93 | 150 | 0 |

The exact strings used for the nested checks were:

```text
litchi_pptx::package::model::Package::opened_presentation_with_limits
litchi_pptx::opened::model::capture_internal
litchi_pptx::notes::codec::scan_processed_xml
litchi_pptx::opened::model::package_fingerprint_with_memo
litchi_pptx::notes::resolved
```

The canonical frames show these helpers beneath the exact owner, followed by
the probe's `run_one` and `main` frames. Every sample block containing one of
the three reported nested helpers also contains the exact owner. The nested
counts are inclusive stack observations and are not additive. One repeat-0
owner-qualified block has an unresolved interior frame; that unresolved
ancestry prevents any phase-fraction claim under the frozen plan. The exact
owner and helper counts remain descriptive evidence of recovered ancestry.

The default `.decoded.gz` streams are retained as the normal inline decode,
but their shortened inline names are not used for these counts. The
no-inline streams are the symbol identity authority because they preserve the
full crate and module paths.

## Acceptance disposition

The frame-pointer follow-up passes raw receipt binding, successful default and
canonical decoding, full-symbol owner matching, and source/output semantic
parity. It qualifies the listed sample counts for the capture owner. The
ordinary DWARF lane remains unqualified at zero exact owner blocks, even
though its raw record and flat symbol decode succeeded.

No native phase percentage, production speedup, or before/after conclusion is
accepted. The qualified counts include warmups and whole-process sampling,
and the frame-pointer build is a different optimized code-generation variant.
