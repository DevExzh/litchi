# 0457 final source-tail candidate epoch

This directory retains the final `odp_source_tail_append_lifecycle` epoch
after the authored-fragment capacity/lease correction. Its separate build,
source manifest, executable bindings, fresh output fixtures, and twelve lanes
(360 samples) passed verification. It reuses the frozen candidate oracle and
deterministic corpus reference. The original `candidate/` bundle remains
historical evidence.

The lane remains specialized source-backed publication-plan evidence. Its
retained-result contract differs from the owned existing-append control, so
the later comparison script reports numeric endpoints and explicit deltas
only. It does not authorize an ordinary Commit/Patch or general CRUD speedup
claim. Normal lanes keep allocator metrics unavailable; allocator lanes retain
allocation volume and regional-peak vectors. All source/sink/semantic/oracle
gates are delegated to the copied independent report oracle.

The directory was prepared without copied build receipts, source manifests,
binary/output bindings, fixtures, runs, or reports. Those final artifacts were
created once, in the order below, after the final build succeeded.
Use `-B`; the binders and capture driver refuse to overwrite retained evidence.

1. Build the final source epoch using the root-owned build procedure and
   retain its successful receipt at
   `docs/performance/results/change-0457/checks/candidate-final-build.json`.
   The receipt must bind an unchanged source snapshot and the final revision.

2. Copy the two final release executables to an isolated path, then bind the
   receipt, source snapshot, and executable identities into this directory:

   ```sh
   mkdir -p /tmp/litchi-goal-0457/candidate-final
   cp tools/perf-baseline/target/release/litchi-perf-baseline \
     /tmp/litchi-goal-0457/candidate-final/litchi-perf-baseline
   cp tools/perf-baseline/target/release/litchi-perf-baseline-alloc \
     /tmp/litchi-goal-0457/candidate-final/litchi-perf-baseline-alloc
   python3 -B docs/performance/results/change-0457/candidate-final/bind.py \
     --build-receipt docs/performance/results/change-0457/checks/candidate-final-build.json \
     --repo-root /home/zhuhe/code/litchi \
     --normal-source tools/perf-baseline/target/release/litchi-perf-baseline \
     --allocator-source tools/perf-baseline/target/release/litchi-perf-baseline-alloc \
     --normal-copy /tmp/litchi-goal-0457/candidate-final/litchi-perf-baseline \
     --allocator-copy /tmp/litchi-goal-0457/candidate-final/litchi-perf-baseline-alloc
   ```

3. Export a fresh source/output fixture set with the final build's retained
   fixture example. The output directory must be new. The generated source
   files are checked against `corpus-bindings.json`; the candidate files are
   independently checked by `native/verify-output.py` before they are copied
   into this bundle:

   ```sh
   tools/perf-baseline/target/release/examples/odp_source_tail_append_fixtures \
     /tmp/litchi-goal-0457/candidate-final-probe
   python3 -B docs/performance/results/change-0457/candidate-final/bind-output.py \
     --repo-root /home/zhuhe/code/litchi \
     --native-oracle docs/performance/results/change-0457/native/verify-output.py \
     --protocol docs/performance/results/change-0457/candidate-final/protocol.json \
     --fixture tiny /tmp/litchi-goal-0457/candidate-final-probe/tiny-source.odp \
       /tmp/litchi-goal-0457/candidate-final-probe/tiny-candidate.odp \
     --fixture medium /tmp/litchi-goal-0457/candidate-final-probe/medium-source.odp \
       /tmp/litchi-goal-0457/candidate-final-probe/medium-candidate.odp \
     --fixture large /tmp/litchi-goal-0457/candidate-final-probe/large-source.odp \
       /tmp/litchi-goal-0457/candidate-final-probe/large-candidate.odp
   ```

   Preserve an existing probe directory and select a new path for any retry.
   Compare the newly authenticated archive and XML identities with the
   historical results; let the independent oracle establish this epoch's
   output instead of copying old fixtures or bindings.

4. Capture the two serialized six-lane phases. R1 is normal tiny/medium/large
   followed by allocator tiny/medium/large; R2 is the reverse order. Each lane
   retains three warmups and thirty samples on CPU 2 with one worker:

   ```sh
   python3 -B docs/performance/results/change-0457/candidate-final/capture.py \
     --phase R1 --attempt final \
     --repo-root /home/zhuhe/code/litchi \
     --protocol docs/performance/results/change-0457/candidate-final/protocol.json \
     --custody-driver docs/performance/results/change-0457/check.py
   python3 -B docs/performance/results/change-0457/candidate-final/capture.py \
     --phase R2 --attempt final \
     --repo-root /home/zhuhe/code/litchi \
     --protocol docs/performance/results/change-0457/candidate-final/protocol.json \
     --custody-driver docs/performance/results/change-0457/check.py
   ```

5. Verify and derive the final epoch before any control/candidate endpoint
   comparison:

   ```sh
   python3 -B docs/performance/results/change-0457/candidate-final/verify.py \
     --attempt final --precleanup --repo-root /home/zhuhe/code/litchi
   python3 -B docs/performance/results/change-0457/candidate-final/derive.py \
     --attempt final --precleanup --repo-root /home/zhuhe/code/litchi \
     --write --output docs/performance/results/change-0457/candidate-final/summary.json
   python3 -B docs/performance/results/change-0457/comparison.py \
     --control docs/performance/results/change-0457/control/measurements.json \
     --candidate docs/performance/results/change-0457/candidate-final/summary.json
   ```

The source manifest and build receipt identify this final source epoch; the
phase receipts additionally bind the final protocol, capture driver, output
binding, executable copies, source custody, allocator environment, reports,
and oracle logs. Preserve any failed attempt directory and its failure receipt.
