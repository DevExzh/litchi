# 0557 driver notes

`run.py` is a serial custody wrapper for the XLSX provenance merge campaign. It
does not implement the benchmark or the candidate. The Rust binary constructs
the deterministic corpus, executes the requested case, writes the measured
JSON, and writes its corpus catalog; the wrapper supplies one fresh process for
one case and one shape.

The intended sequence is:

```text
python3 -B docs/performance/results/change-0557/run.py freeze --stage baseline
python3 -B docs/performance/results/change-0557/run.py build-normal --stage baseline
python3 -B docs/performance/results/change-0557/run.py build-alloc --stage baseline
python3 -B docs/performance/results/change-0557/run.py noise --stage baseline --execution-stage baseline
```

The baseline noise command must finish before candidate output is observed. It
runs the eight primary case/shape rows twice with CPU 2, 20 warmups, and 1,000
samples. Freeze the resulting `N` and the floor formula in the coordinator's
machine-readable analysis before candidate capture. The noise report may show
provisional per-row floors for diagnostics; the admission floor is recomputed
from the matched native baseline p50 for each paired row. `N` is a
percentage-point value:

```text
N = 100 * abs(noise_r2_p50_ns - noise_r1_p50_ns) / noise_r1_p50_ns
floor = max(1.0, 3 * N, 100 * 50000 / matched_baseline_p50_ns)
```

The last term is also percentage points: 50,000 is a nanosecond absolute
floor, and the multiplication by 100 converts its ratio to percent. If `N`
exceeds 5 percentage points, retain both repeats and stop the candidate gate;
do not rerun the pilot to obtain a more favorable threshold. `run.py` does not
derive `N` or authorize candidate capture, so the coordinator must invoke
`analyze.py --noise-only`, retain its result, and enforce this stop before
launching any candidate child.

After the candidate source is prepared, freeze and build it in its own stage:

```text
python3 -B docs/performance/results/change-0557/run.py freeze --stage candidate
python3 -B docs/performance/results/change-0557/run.py build-normal --stage candidate
python3 -B docs/performance/results/change-0557/run.py build-alloc --stage candidate
```

The normal and allocator lanes each use two ABBA repeats over all 32 rows. A
repeat-2 matrix is traversed in reverse order to reduce order bias. The four
legs are baseline-R1, candidate-R1, candidate-R2, and baseline-R2. Native
elapsed p50/mean primary and control gates use only the normal lane. Allocator
elapsed vectors remain diagnostics; allocator per-child RSS and measured
allocation phase vectors have their own gates. The final baseline leg is invoked
with `--stage baseline --execution-stage candidate`; it
uses the retained baseline binary and baseline output folder while checking
the candidate execution manifest.

```text
python3 -B docs/performance/results/change-0557/run.py native --stage baseline --execution-stage baseline --repeat 1
python3 -B docs/performance/results/change-0557/run.py native --stage candidate --execution-stage candidate --repeat 1
python3 -B docs/performance/results/change-0557/run.py native --stage candidate --execution-stage candidate --repeat 2
python3 -B docs/performance/results/change-0557/run.py native --stage baseline --execution-stage candidate --repeat 2
```

Use the same four invocations with `alloc` for the allocator lane. Native
children use 20 warmups and 1,000 samples. Allocator children use three
warmups and 30 samples. `/usr/bin/time` writes a separate RSS and process-time
sidecar for every child. The case and shape are always singular, so the sidecar
has one process scope.

Every stage is frozen with a full source manifest. The first freeze creates an
exclusive `workspace-lock.json` binding and `workspace-Cargo.lock` copy when
those evidence files are absent; later stages only validate them. A child
checks both the output-stage and execution-stage manifests and both lock
identities before launch and after termination. Binary-backed children hash the
retained binary before and after launch. Receipts bind these identities, the
exact argv, selected environment, plan hash, driver hash, host observation,
exit status, and every child artifact.

Every child receives `TMPDIR=<owned-target>/tmp` and
`CARGO_TARGET_DIR=<owned-target>`; quality commands therefore keep Cargo
outputs inside the campaign target even when their argv has no explicit target
directory.

Evidence paths are exclusive. A pre-existing receipt, host record,
stdout/stderr file, benchmark JSON, catalog, or `/usr/bin/time` sidecar aborts
before launching a child. `run_child` writes the receipt before reporting a
non-zero exit or a failed post-run source guard, so source mutation and other
failed attempts remain reviewable. A source guard failure is never converted
to a successful or partial result.

After both manifests are frozen and before the first candidate build, run the
pure asymmetric-guard check:

```text
python3 -B docs/performance/results/change-0557/run.py guard-test \
  --stage baseline --execution-stage candidate
```

It accepts the retained baseline output manifest without comparing it to live
candidate source, checks the candidate execution manifest against live source,
and verifies that a synthetic execution digest mutation is refused. The test
does not modify repository or evidence files.

The optional `check` action lets a quality runner reuse the same source and lock
guards without a benchmark binary:

```text
python3 -B docs/performance/results/change-0557/run.py check \
  --stage candidate --execution-stage candidate --name xlsx-tests -- \
  cargo test --locked -p litchi-xlsx --all-features
```

For rustdoc checks, pass the command-specific environment explicitly, for
example `--env RUSTDOCFLAGS=-D warnings` before the `--` separator.

The driver owns only `run.py`, `plan.json`, and this note during preparation.
Generated stage evidence and the owned target are campaign outputs; the root
coordinator decides their retention and cleanup after pre-cleanup verification.
