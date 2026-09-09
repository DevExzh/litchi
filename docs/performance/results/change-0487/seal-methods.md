# 0487 evidence seal

`verify_seal.py` has a verify-only default.  It creates `seal.json` only with
an explicit `seal` command, and both the seal and the optional
`seal-verification.json` receipt are created exclusively.  The receipt is the
only documented non-sealed file excluded from the inventory; `seal.json` is
also excluded because it contains that inventory.  Every other regular file
under this directory is recorded with its byte length and SHA-256.  Symlinks,
special files, and Python bytecode (`__pycache__`, `.pyc`, or `.pyo`) fail
closed.

Run the seal after all retained artifacts exist, selecting only terminal final
gates.  Historical failed receipts may remain in the directory without being
selected:

```text
python3 -B verify_seal.py seal \
  --gate validation/format-final2.json \
  --gate validation/tests-all-features-final2.json \
  --gate validation/tests-no-default-final2.json \
  --gate validation/clippy-final2.json \
  --gate validation/rustdoc-final2.json \
  --gate validation/bench-tests-serial1.json \
  --gate validation/python-final2.json \
  --gate validation/boundaries-final2.json \
  --comparison comparison-summary.json \
  --profile profiles/profiles1/result.json \
  --fuzz fuzz-stream/candidate1/smoke.json
```

The exact final gate list is part of `seal.json`; focused test gates may be
selected with repeated `--gate` options alongside the terminal build and
quality gates.  A selected gate must use
the retained gate schema, exit with code zero, and have equal before/after
source snapshots (`source_unchanged: true`, called stable by the verifier).
Failed development attempts are not silently promoted into that list.

The comparison check imports `compare.py` and calls its retained validators
directly.  The 0487 lane binds its `before` phase to the retained 0485
baseline and its `after` phase to the 0487 build; the 0485 evidence itself
remains the earlier comparison against the sealed 0484 original.  It checks
the protocol, build and binary bindings, capture gate receipts, every formal
process receipt/report/resource, content identity, and the retained summary
rows.  It also requires every selected final gate to bind the after-build
source snapshot.  It requires the declared 18-arm matrix of 144 formal
processes and 4,320 measured samples.  It never calls `capture`, `analyze`,
or writes a comparison output.

The profile check requires the terminal `pass`, the fixed 12-label inventory
(`before-perf-*`, `after-perf-*`, and `after-strace-*` for the owned and file
arms of both profile workloads), a zero process return code, and every
artifact listed by each receipt to have the recorded hash.  The fuzz check
requires the prepared seed hashes, exactly 27 positive cases whose embedded
receipt equals `positive.json`, two `-runs=10000` runs with the 64 KiB cap,
and exact hashes for every retained `post-run/run-N` tree.

Verification is the default and does not write anything:

```text
python3 -B verify_seal.py
```

To retain the excluded verification receipt, add `--receipt`.  It is created
exclusively and must be removed before a new receipt is requested.
