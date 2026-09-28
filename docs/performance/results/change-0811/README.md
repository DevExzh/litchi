# 0811 — current-production native attribution

This packet records an evidence-only diagnostic for the current production
source at `2894fbd628434bad4ef45af4f0f0b9a8d4d23468`. Its tracked production
manifest has 9,196 files and its file hashes match the sealed 0810 after
manifest. There is no candidate, source allowlist, allocation lane, Callgrind
lane, adoption threshold, or speedup claim. iWork remains out of scope.

The six-file public probe is byte-for-byte copied from 0810, including the
`litchi.pptx.public-workflow-probe-0806.v1` schema, repaired `valid-4attr`
fixture, 36 tests, and exact owner
`namespace_uri_probe::capture_region_0793`. Root `Cargo.lock` and
`rustfmt.toml`, all 35 architecture inputs, the three unrelated workspace
files, and every driver identity are frozen before the first build.

The root agent builds three release binaries serially in
`/home/zhuhe/code/litchi-target-0811`: `control` has no features, `profile`
uses `capture-profile`, and `fp` uses the same feature with
`RUSTFLAGS=-C force-frame-pointers=yes`. All builds are offline, locked, and
use two Cargo jobs; each completed executable is copied under an immutable
variant name before the next build.

Fresh correctness consists of the 36-test probe fmt/test/Clippy lane in its
owned `probe-quality` subtarget. The six production gates are reused from the
sealed 0810 after result because all 9,196 production file hashes are equal;
`quality.py` verifies that witness without executing Cargo a second time.

Native capture covers `capture` on `tiny`, `medium`, and `large` with six
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
maximum RSS in KiB for each report. After native capture, two large-capture
perf runs use the `fp` binary, 100 samples, no warmup, CPU 12,
`cycles:u` at 499 Hz, and `--call-graph fp`; `--no-buildid-cache` prevents an
owned-target run from writing a build-id cache elsewhere. These runs add two
reports and 200 samples, for a total of 56 reports and 1,820 samples.

The sampled lane is decoded while all exact binaries still exist with
`perf script --no-inline --ns`. Both raw perf data and decoded frames are
retained as deterministic gzip members (`mtime=0`) with uncompressed and
compressed identities. Whole-process sample counts, exact-owner counts,
unresolved frames, and lost-event diagnostics remain descriptive attribution
evidence. Nested stack counts overlap; no phase fraction, latency, RSS,
allocation, causal, or universal workload claim is authorized.

The root-only sequence is:

```text
python3 -B docs/performance/results/change-0811/quality.py
python3 -B docs/performance/results/change-0811/build.py
python3 -B docs/performance/results/change-0811/probe_quality.py
python3 -B docs/performance/results/change-0811/capture.py native
python3 -B docs/performance/results/change-0811/capture.py perf
python3 -B docs/performance/results/change-0811/decode.py
```

Independent readers may replay only after all capture and decode processes
are terminal. The owned target is removed only after those readers and cleanup
witnesses pass. The 0811 batch does not by itself authorize production change
or adoption.

Completed results are in [the report](../../0811-pptx-current-native-attribution.md),
[analysis.json](analysis.json), [root-audit.json](root-audit.json), and
[results-review.md](results-review.md). Run `python3 -B validate.py --final`
from this packet directory to replay checks after owned-target cleanup.
