# DOCX `stylesWithEffects` bounded profile

This directory retains the approved current-source profile captured from
production source commit
`8702fd4db8723acceb7deb51bcb40ff66604bf10` with clean harness checkout
`bf21536b35380005041917e7aee03a39804b2659`. The current 52-lane correctness
smoke was completed first from frozen smoke checkout
`637082e318ae6bee2bd47ef317e4a417c1176996` and is retained in
`../smoke-current-637082e31/`; its result tree is bound to retention commit
`fa927a8a9de94891858a6bb3d44c21d5fa63697d`. The profile runner and the
independent root verifier both passed. Their `verification.json` and
`root-verification.json` files are byte-identical.

The profile has 31 lane/scale rows, three fresh process indexes, two warmups,
and twenty measured samples per process: 93 JSON receipts and 1,860 samples.
It covers native resources and deterministic 64 KiB and 1 MiB resources. The
correctness-only smoke retains the cap, malformed-input, stale-patch, signed,
absent-owner, inverse, and independence lanes; those refusal and boundary
cases were not added to this timing matrix. No 8 MiB or near-limit timing was
run. The 307 raw capture files, including eight generated fixture files, total
12,394,251 bytes. `root-retention-verification.json` records every raw-file
size and SHA-256. Three separately listed process-census snapshots are
sanitized and are not part of the raw receipt set.

Preparation generated and hashed the XML, DOCX packages, replacement
resources, limits, expected member digests, and semantic inputs before each
operation clock. Capture and projection include their named package work;
prepared mutation lanes keep setup outside the operation clock. The outer
`elapsed_ns` and phase clocks are separate from report writing. `publish_ns`
contains the public mutation/publication operation and contains the nested
`apply_ns`; `apply_ns` includes the public styles-with-effects patch and put
path. `serialize_ns` covers only package serialization after the mutation.
`reopen_ns` and `inverse_reopen_ns` are separate readback phases. Nested phase
values must not be added to their containing phase.

Allocator counters report calls, requested/direct/reallocated/deallocated
bytes, live bytes, and peak live delta for the Rust counting allocator. They
describe the measured allocator window and are distinct from RSS. The
`/usr/bin/time -v` sidecars report one process maximum RSS marker for each
fresh process invocation, including its warmups and samples. Neither measure
is a managed process-memory cap.

Some representative median observations are below. Times are milliseconds;
allocator values are bytes.

| row | elapsed | publish / apply / serialize | reopen | requested bytes | peak live delta |
| --- | ---: | ---: | ---: | ---: | ---: |
| synthetic main, 64 KiB replace | 21.513 | 13.418 / 13.118 / 0.304 | 3.548 | 24,061,974 | 799,352 |
| synthetic main, 1 MiB replace | 186.614 | 124.371 / 120.668 / 3.709 | 29.489 | 153,146,975 | 5,864,443 |
| synthetic glossary, 1 MiB replace | 189.680 | 126.136 / 122.329 / 3.706 | 30.080 | 154,794,997 | 5,865,059 |
| synthetic main, 64 KiB capture | 4.219 | capture 3.547 | — | 7,036,554 | 703,068 |
| synthetic main, 1 MiB capture | 31.361 | capture 29.515 | — | 31,295,714 | 4,637,484 |
| synthetic main, 64 KiB projection | 3.664 | capture 3.255 | — | 3,744,390 | 416,010 |
| synthetic main, 1 MiB projection | 29.864 | capture 28.915 | — | 24,710,462 | 2,841,345 |

For these synthetic main fixtures, the observed median 64 KiB-to-1 MiB ratios
were 7.43x for capture, 8.15x for projection, and 8.67x for replacement.
These are observations for these generated inputs only. They are not a
general or asymptotic scaling result, a production workload estimate, or a
before/after comparison.

The first attribution candidate is the repeated public apply path in the two
1 MiB replacement rows. Its median `apply_ns` is 120.668 ms for main and
122.329 ms for glossary, while the containing publish medians are 124.371 ms
and 126.136 ms; serialization is about 3.7 ms and reopen is about 29.5–30.1
ms. The same rows request about 153–155 MB while their measured peak live
delta is about 5.86 MB. This points to inspecting repeated traversal, copying,
and transient allocation within the apply/put path as the next narrow
investigation. The phase is composite and nested, so these receipts do not
identify a single function or establish an optimization opportunity. Any
change requires a matched source/input/toolchain capture with the same
semantic, opaque-member, inverse, allocator, and RSS gates.

The host was Ubuntu 26.04.1 on AMD EPYC 9R45 with 32 logical CPUs,
129,447,068 KiB total memory, and affinity `0-31`. One-minute load was 0.46
before and 2.76 after the capture. Physical-core count and mount identity were
not captured. The host was shared and unisolated, so there is no idle-host or
uncontended timing claim. The retained sanitized before/during/after censuses
are sampled operator observations. Under the coordinated build hold, no
competing compiler/profile process was observed in those samples; this is an
observation of the samples, not continuous proof that no such process
overlapped the run.

Source manifests and Cargo metadata matched before and after. The profile was
built with the locked offline harness, `CARGO_INCREMENTAL=0`, `LC_ALL=C`, and
the recorded default linker/toolchain. The toolchain was Rust
`rustc 1.95.0 (59807616e 2026-04-14)` and Cargo `1.95.0 (f2d3ce0bd
2026-03-21)`. The disposable binary's recorded SHA-256 is
`7aa8a918ab38a833eba602c356178ee2186866d107936152c8b59e11d2571182`;
the binary itself and Cargo target were removed only after the successful
verification sentinel. Build, source, command, fixture, host, and allocator
receipts remain here. This evidence makes no speedup, memory-cap, native Word
acceptance, or later-production claim.

## Replay layout

Root restored all 307 raw files from retention commit `23ca02f01`, reran the
frozen verifier, and obtained the exact original verification receipt. The
temporary restoration was then removed. See
[`committed-restoration-verification.json`](committed-restoration-verification.json)
for the committed-byte replay record.

The verifier requires the captured clean checkout and the recorded external
results path because `generated-fixtures.json` contains absolute generated
resource and package paths. The retained repository result directory cannot
be replayed directly. Restore the exact external path
`/var/tmp/litchi-docx-styles-effects-profile-results-bf21536b3-20260912-a` and
copy only the paths listed in `root-retention-verification.json`, including
`generated-fixtures/`, without rewriting their bytes or embedded absolute
paths. The target path may remain absent after successful cleanup.

The durable source bundle is `captured-source.bundle` (SHA-256
`199e7bef6f3ac67d8e2a2de6b25dc20987a897d8ca0b0448f6f21087119702c1`). It
contains `bf21536b35380005041917e7aee03a39804b2659` and requires
`8702fd4db8723acceb7deb51bcb40ff66604bf10`:

```sh
git bundle verify /path/to/profile-clean-bf21536b3/captured-source.bundle
git fetch /path/to/profile-clean-bf21536b3/captured-source.bundle \
  HEAD:refs/heads/docx-styles-effects-profile-capture-bf21536b3
git worktree add --detach \
  /var/tmp/litchi-docx-styles-effects-profile-capture-bf21536b3-a \
  bf21536b35380005041917e7aee03a39804b2659
```

After restoring the raw files at the recorded external path, run the captured
checkout's `verify_profile.py` with that checkout as `--root`, its evidence
directory as `--evidence`, the restored external directory as `--results`, and
the restored before/after manifest and metadata paths. Write replay output to
a new file in the same external results directory. The verifier's output
containment rule rejects a direct replay against this retained repository
directory, and the original raw receipts must remain unchanged.
