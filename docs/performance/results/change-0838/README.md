# 0838 — current-source DOCX inverse diagnostic

[Report and interpretation](../../0838-docx-durable-inverse-physical-diagnostic.md).
Production is unchanged from `4435bae2c9`; the rejected 0837 compression candidate
is not present. Three fresh processes execute six cases each, retaining 63 ZIPs
and 18 durable patch wires. Durable inverse is byte-exact for the ordinary
borrowed-regenerated control, but changes the main document compressed payload
from 135 to 141 bytes for both staged-source variants. Logical restoration,
retained-source exact restoration, and stale-source rejection pass throughout.

## Reproduce retained analysis

From the repository root:

```sh
python3 -B docs/performance/results/change-0838/audit.py
```

The independent stdlib reader verifies ZIP member hashes, canonical content
hashes, XML paragraph order, unchanged local spans, exact retained inverse,
durable exact/non-exact classifications, patch-wire hashes/magic/version, and
cross-process determinism against `analysis.json`. Rust separately asserts
canonical wire re-encoding and stale refusal with an empty sink.

For a fresh native run, build `probe/Cargo.toml` offline and locked with an
isolated target directory, then pass `fixtures/fresh-source.zip` and a new output
directory to `source-inverse-diagnostic`. `driver.py`, `quality-v2.py`, and
`capture.py` document the recorded environment and sequence; they deliberately
refuse to overwrite existing receipts and are not rerun-in-place helpers.

## Evidence map

- `origin.json`, `plan.json`, `fixture-origin.json`, `host.json`: frozen scope,
  input provenance, normative/unrelated hashes and host facts.
- `quality-reuse.json`: exact production source binding to prior final gates.
- `quality-source.json`, `drafts/main-initial.rs`, initial check receipts:
  initial probe, including the unused-import Clippy failure.
- `quality-v2-source.json`, `quality.json`, `probe-v2-*` receipts: final fresh
  formatting, compilation, strict Clippy and rustdoc gates.
- `lock-alignment.json`, `probe/Cargo.lock`: resolved workspace-aligned versions.
- `freeze-diagnostic.json`, `build-diagnostic.json`, `admission.json`: source,
  executable hash and capture/reader/input bindings.
- `runs/run-00` through `run-02`, `capture.json`, `analysis.json`: complete
  diagnostic artifacts, process IDs, raw and logical facts and independent audit.
- `commands`: terminal argv, environment, log hashes, exit codes and time bounds.
  `probe-clippy` is the only nonzero command; final checks pass.
- `close.py`, `closure.json`, `cleanup.json`, `seal.json`: evidence validation,
  marker-checked temporary-root removal and owned-file hash manifest.

The retained executable is removed with the isolated build root after capture;
its hash and successful build receipt remain. This packet has no performance
samples or speedup claim. Existing exact-artifact tests remain unchanged.
