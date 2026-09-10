# Managed DOCX source read-ahead (0493)

This bundle measures exact reads against an explicit 4 KiB production read-ahead
policy in the same managed OPC/DOCX lifecycle. It follows the benchmark-only
request-amplification experiment in change 0492. See [methods.md](methods.md)
for the operation boundary, provider model, accounting, and limitations.

## Reproduction and verification

The build source is base revision
`e8ee19b063285cb0bb73dea819f1be1ac5e82133` plus `build-source.patch`.
Apply the patch in an isolated checkout of that revision; it includes the
preexisting Keynote compilation-input delta solely to reproduce the measured
source. That unrelated working file is excluded from this batch's source commit.
`source-reproduction.json` binds the patch and complete source manifest.

The two `build-*.json` records bind the actual commands, environment, source
manifests, retained executable hashes, and successful build receipts. The build
environment uses Rust 1.98.1, four Cargo jobs, release debug information, frame
pointers, unwind tables, and disabled incremental compilation. The benchmark
package has its own workspace; use its explicit manifest path.

```sh
python3 -B docs/performance/results/change-0493/measure.py plan
python3 -B docs/performance/results/change-0493/measure.py verify --pilot --attempt pilot1
python3 -B docs/performance/results/change-0493/measure.py verify --attempt formal1
python3 -B docs/performance/results/change-0493/verify_bundle.py verify
```

The verification commands recompute report oracles and analysis from retained
raw observations. They do not execute a benchmark. Existing evidence filenames
are immutable; fresh measurements require distinct attempts and authenticated
build records. After building and retaining both roles, the original capture
sequence is:

```sh
python3 -B docs/performance/results/change-0493/measure.py freeze
python3 -B docs/performance/results/change-0493/measure.py capture-all --pilot --attempt pilot1
python3 -B docs/performance/results/change-0493/measure.py analyze --pilot --attempt pilot1
python3 -B docs/performance/results/change-0493/measure.py capture-all --attempt formal1
python3 -B docs/performance/results/change-0493/measure.py analyze --attempt formal1
```

Capture takes the shared advisory CPU lock itself. Do not wrap capture in
`gate.py`, which uses that same lock. The 16 formal children contain 480 measured
samples; the eight pilot children contain 24. Warmups are validated separately
and excluded from those counts. Normal and allocator binaries each run exact
and read-ahead policies under both transport service models, with reversed order
in the second formal repeat. Whole-child RSS is one observation per child.

`final-gates.json` selects final-source validation. Earlier failed or changing
source attempts remain visible under `validation/` and are explained in
[development.md](development.md). Independent reviews cover
[publication](publication-review.md), [measurement](measurement-review.md), and
[bundle verification](bundle-review.md), and [formal results](results-review.md).

Cleanup is restricted to the batch's exact Cargo target and temporary root.
It authenticates the two retained executables and active-process state before
removal; the release executables remain available for reproduction. The bundle
seal binds the final evidence inventory and invokes canonical capture and
cleanup verification.

This batch does not establish native-producer, cold-filesystem, borrowed-source,
multicore-scaling, or comprehensive CRUD completion. The full goal remains open.
