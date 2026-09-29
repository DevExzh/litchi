# 0840 — fresh CFB emission order experiment

The experiment tests moving ministream emission after the large streams that
already precede it in the allocated physical layout. The candidate preserves
all successful output bytes and tests whether avoiding Cursor gap initialization
improves fresh public DOC/CFB writer operations.

## Evidence map

- `origin.json`, `host.json`, `environment.json`, `source-before.json`: exact
  base, toolchain/host, normative and unrelated input hashes, source inventory.
- `plan.json`, `capture-design.json`, `preflight-amendment.json`: frozen cases,
  reducers, guards, equal-length argument paths, and separate allocator preflight.
- `candidate.patch`, `source-before`, `source-after`, `candidate-source.json`:
  source change and five focused tests. The initial candidate identity and
  formatting-only adjustment are retained separately.
- `quality-before.json`, `quality-after.json`: affected-owner gates. Candidate
  passing command labels use `after-v2`; the first formatting failure remains.
- `probe`, `probe-quality-*`, `build-*`, `freeze-*`: standalone native and
  allocator companions, fresh qualification, exact inputs and binary hashes.
  The final baseline probe command labels use `before-v2`; initial helper
  visibility/unused-function Clippy errors and draft source remain retained.
- `allocator-provenance.json`, `lock-alignment.json`: byte-identical copied
  observer modules and dependency versions/checksums aligned with the root lock.
- `runs`: exclusive process folders with argv, source binary identity, output
  reports, GNU-time counters, terminal receipts, and artifact hashes. Only
  qualification and untimed seek observers retain complete CFB artifacts.
- `readers.py`, `qualify.py`, `reader-tests.json`: independent input generation,
  CFB parsing, source/artifact equality, cursor event replay, allocation
  conservation, and mutation/statistical checks. Earlier schema/order mistakes
  and successful replacement invocations remain in `commands` and `drafts`.
- `mechanism.json`: untimed before/after gap and seek counts; byte-identical
  artifacts bind the physical-order observation to the public writer outputs.
- `admission.json`, `analysis-repair.json`, `analyze-v2.py`, `analysis.json`, `paired.md`, `paired.csv`: immutable formal
  inputs, all individual metrics/intervals, and the frozen adoption decision.
- `closure.json`, `cleanup.json`, `seal.json`: command outcomes, source/custody
  checks, marker-owned temporary-root removal, and the exact owned-file seal.

Native samples time fresh writer construction, prepared-input registration and
`write_to`. Input generation, verification and output destruction are excluded.
Allocator regions cover the same operation; the native binary has no global
allocation observer. Seek traces and preflight observations are not pooled with
native timing. GNU-time RSS covers the whole process, including untimed work.

Replay from the recorded checkout after temporary build roots are removed:

```sh
python3 -B docs/performance/results/change-0840/analyze-v2.py
python3 -B docs/performance/results/change-0840/close.py verify
```

Raw receipts retain original absolute command paths. A fresh experiment needs
new exclusive run/build directories and fresh identities; capture deliberately
refuses to replace an existing measurement. Relative Cargo manifests, exact
locks, deterministic generators and complete source hashes retain build inputs.

This is a fresh in-memory seekable-writer experiment. It does not measure the
source-layout reuse route, filesystem durability, sequential writer, concurrent
writers, native Office interoperability, or physical cold-cache behavior. No
compression policy, layout rule, durable wire or unrelated work is changed.
