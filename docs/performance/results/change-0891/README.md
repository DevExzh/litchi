# 0891 x14ac paired candidate packet

This packet measures the reviewed XLSX `x14ac` candidate against the retained
0885 release binaries. The candidate source is sealed before any build and is
restricted to `crates/litchi-xlsx/src/raw/worksheet/x14ac.rs` plus the new
`crates/litchi-xlsx/src/raw/worksheet/x14ac/reuse_tests.rs`. The production
candidate, patch, source review, quality receipt, and methodology review are
independent inputs; this packet does not turn a source review into a
performance claim.

The builder reconstructs the retained 0885 source archive, overlays the two
sealed candidate files, checks the complete 10,765-file baseline inventory plus
the new file, and builds native and allocator release binaries with Rust 1.95,
offline locked dependencies, fat LTO, one codegen unit, and two Cargo jobs.
Build roots are marked and remain outside the committed packet. The retained
0885 native and allocator executables are the before side.

After the release build, create one content-identical executable copy for each
role and mode. Each copy is reused by its scheduled fresh children. This
controls pathname and inode history; it does not establish a causal RSS result,
so RSS is retained as a review signal.

The capture has 72 sequential fresh children pinned to CPU 12 with a warm
filesystem cache: eight native AB/BA blocks over the four reviewed cases,
three warmups and 50 samples per native child, followed by four allocator
AB/BA primary-case blocks with three warmups and 15 samples. Whole-child pairs
are reduced with bootstrap seed 891, 10,000 draws, and indices 249/9749.
Samples are never pooled. The primary adoption gate is native primary-case p50
median ratio at most `0.98` and bootstrap upper endpoint below `1.0`.
Controls, p95/p99, p50 spread, allocator counters, and RSS remain visible to
review; no allocator or RSS result is an automatic adoption decision.

After the source owner has sealed `candidate-v3.patch` and
`candidate-freeze-v3.json` (retaining the superseded v1 and v2 freezes), run the candidate release build and make the four
managed copies:

```sh
python3 -B docs/performance/results/change-0891/build.py
python3 -B docs/performance/results/change-0891/prepare_exec.py \
  --before-native docs/performance/results/change-0885/binaries/native \
  --before-allocator docs/performance/results/change-0885/binaries/allocator \
  --after-native docs/performance/results/change-0891/binaries/native \
  --after-allocator docs/performance/results/change-0891/binaries/allocator
```

After the source review and methodology review are present, run the supplied
quality runner and aggregate its passing receipt. Both helpers are create-only;
the runner stops at the first failed command:

```sh
python3 -B docs/performance/results/change-0891/quality_runner.py \
  --run --attempt quality-v3 --freeze candidate-freeze-v3.json
python3 -B docs/performance/results/change-0891/quality.py \
  --attempt quality-v3
```

After the passing `quality.json` is present, freeze the measurement plan. This
validates the exact protocol and binary/copy identities without starting a
workload:

```sh
python3 -B docs/performance/results/change-0891/freeze.py
```

Only then run the capture and reduction:

```sh
python3 -B docs/performance/results/change-0891/freeze.py --verify
python3 -B docs/performance/results/change-0891/experiment.py \
  --plan docs/performance/results/change-0891/plan.json \
  --output-root docs/performance/results/change-0891/capture
python3 -B docs/performance/results/change-0891/reduce.py \
  --plan docs/performance/results/change-0891/plan.json \
  --experiment-root docs/performance/results/change-0891/capture \
  --output docs/performance/results/change-0891/reduction.json
python3 -B docs/performance/results/change-0891/verify_packet.py \
  --plan docs/performance/results/change-0891/plan.json \
  --experiment-root docs/performance/results/change-0891/capture \
  --reduction docs/performance/results/change-0891/reduction.json \
  --output docs/performance/results/change-0891/verification.json
```

The release builder and managed-copy preparation are deliberately separate
create-only commands. The source owner must first seal `candidate-v3.patch` and
`candidate-freeze-v3.json`; then run `build.py`, followed by `prepare_exec.py`
with the retained 0885 binaries and the two candidate binaries. A fresh
checkout does not contain ignored release binaries, the source archive, or
managed copies. `verify_packet.py` can replay the plan and reduction after
source/build roots and managed copies have been removed: it reconstructs the
candidate in a temporary directory from the retained archive and patch and
does not build or launch a workload.

All packet JSON and receipt writers are create-only. Preserve failed attempts
under a new attempt directory; never overwrite a sealed receipt or rerun a
failed child in place.
