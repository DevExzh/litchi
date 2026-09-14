# change-0569 evidence packet: the detect-then-open saving, priced

Change record:
[`docs/performance/0569-ooxml-detect-then-open-priced.md`](../../0569-ooxml-detect-then-open-priced.md).
Disposition: attribution only. No production change, `performance_claim: none`.

## Contents

| Path | What it is |
| --- | --- |
| `shapes.txt` | What every facade opener returns for every mismatched input, the evidence that no facade error names the detected format. |
| `bench-run1.txt`, `bench-run2.txt` | Two independent paired A/B runs, 2,000 and 3,000 iterations on different cores, agreeing within 0.6%. |
| `cascade.txt` | The try-each-facade workaround, measured, showing it is worse for a format tried late. |
| `traces/` | `strace` captures of the one-call and two-call patterns for each format, segmented per `openat` of the fixture. |
| `probe/` | The scratch probe's sources, so the measurement can be rebuilt. |

## Replay

Rebuild the probe against the workspace and run its three modes:

```sh
cargo build --release --manifest-path <probe>/Cargo.toml
<probe target>/release/shapes
<probe target>/release/bench 2000
<probe target>/release/cascade 2000
```

The probe takes path dependencies on the workspace crates and must be built with
the pinned 1.95.0 toolchain, which means building from a directory the
repository's `rust-toolchain.toml` governs.

## Why the traces are segmented per `openat`

Segmenting a trace by process and descriptor is unsafe on this host: the dynamic
loader uses descriptor 3 before `main`, and the package file later reuses it.
That is the error change 0567 corrected in change 0561. These captures segment
at each `openat` of the fixture instead, which cannot be contaminated by loader
reuse.

## What is not here

The probe's 3.0 GiB of build output was removed after capture. Warm cache, one
host, one fixture per format, one feature set, one cascade ordering. No
cold-cache, latency-bearing-source, allocation or throughput result is claimed.
Host load average was 21 to 25 from concurrent builds, so absolute medians may
run a few percent high; the paired deltas are drift-cancelled by construction
because the two patterns are interleaved in one loop.
