# 0814 — current-production native attribution

This packet records a fresh evidence-only diagnostic for the current PPTX
production source at `5eb8629254fbf88009b988734af59eb90f8c6271`. The tracked
production manifest has 9,196 files. Its file hashes match the retained 0813
after source; the newer commit revision is recorded as the authoritative 0814
source identity. There is no candidate, source allowlist, allocation lane,
Callgrind lane, adoption threshold, or speedup claim. iWork remains out of
scope.

The six-file public probe is copied byte-for-byte from 0813. It retains the
`litchi.pptx.public-workflow-probe-0806.v1` schema, the repaired
`valid-4attr` fixture, 36 tests, the
`litchi-perf-0780-static-mce-capabilities` marker, and the exact owner
`namespace_uri_probe::capture_region_0793`. Root `Cargo.lock` and
`rustfmt.toml`, all 35 architecture inputs, the three unrelated workspace
files, and the driver identities are frozen before the first build.

The six-gate production quality result is reused from the sealed 0813 after
result because all 9,196 tracked production file hashes are equal. That
witness records 1,241 passed tests, zero failed tests, three ignored tests,
and 85 suites. The 0814 probe quality lane remains fresh and runs formatting,
the 36-test release suite, and release Clippy with warnings denied.

The root agent builds three release binaries serially in
`/home/zhuhe/code/litchi-target-0814`: `control` has no features, `profile`
uses `capture-profile`, and `fp` uses the same feature with
`RUSTFLAGS=-C force-frame-pointers=yes`. Builds are offline, locked, release
builds with two Cargo jobs and no incremental compilation. Each executable is
copied under an immutable variant name before the next build. Root-owned
`assembly.py` records exact scanner and inspector symbols for all three
binaries while those binaries remain available.

Native timing covers `capture` on `tiny`, `medium`, and `large` with six
three-variant blocks in this exact order:

```text
control/profile/fp
profile/fp/control
fp/control/profile
fp/profile/control
profile/control/fp
control/fp/profile
```

Each process uses CPU 12, three warmups, and thirty measured samples. The
matrix has 54 reports and 1,620 measured samples. `/usr/bin/time` records
maximum RSS in KiB for every report. After native capture, two large-capture
perf runs use the `fp` binary, 100 samples, no warmup, CPU 12, `cycles:u` at
499 Hz, and `--call-graph fp`; `--no-buildid-cache` prevents an owned-target
run from writing a build-id cache elsewhere. These runs add two reports and
200 samples, for a total of 56 reports and 1,820 samples.

The sampled lane is decoded while all exact executables remain in the owned
target with `perf script --no-inline --ns`. Raw perf data and decoded frames
are retained as deterministic gzip members with uncompressed and compressed
identities. Whole-process sample counts, exact-owner counts, unresolved
frames, and lost-event diagnostics remain descriptive attribution evidence.
Nested stack counts overlap; no phase fraction, latency, RSS, allocation,
causal, or universal workload claim is authorized.

The root-only sequence is:

```text
python3 -B docs/performance/results/change-0814/quality.py
python3 -B docs/performance/results/change-0814/build.py
python3 -B docs/performance/results/change-0814/probe_quality.py
python3 -B docs/performance/results/change-0814/capture.py native
python3 -B docs/performance/results/change-0814/capture.py perf
python3 -B docs/performance/results/change-0814/decode.py
python3 -B docs/performance/results/change-0814/assembly.py
```

Independent readers run only after capture, decode, and assembly are
terminal. The owned target and three copied binaries are removed only after
the readers and cleanup witnesses pass. This packet does not authorize a
production change or adoption decision.

The bootstrap seed reserved for independent numerical readers is `814814`
with 10,000 median resamples and sorted zero-based endpoints `[250, 9749]`.
Any later report must retain all tail/RSS diagnostics and unresolved/lost
sample diagnostics and must scope every statement to this source, probe,
machine, build, and metric.
