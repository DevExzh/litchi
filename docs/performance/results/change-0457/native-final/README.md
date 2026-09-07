# Change 0457 native-final ODP fixture oracle

This is a fresh native capture bundle. It contains only the copied runner,
read-only ten-fixture inventory, independent ZIP/XML oracle, verifier, binder,
preflight helper, and retained upstream notices. It intentionally contains no
old binding or run records. The fixture archives remain external and untouched.

The final runner writes live producer outputs under
`/tmp/litchi-goal-0457/native-final/` and retains source/output gzip archives
under `native-final/runs/`. Its parent custody gate is the standard
`../checks/native-final.json` receipt. The final `../SHA256SUMS` seals the
complete change bundle, including that gate.

Run after binding the fresh candidate binary and completed build receipt:

```text
python3 -B docs/performance/results/change-0457/native-final/bind.py \
  --build-receipt docs/performance/results/change-0457/checks/candidate-final-build.json \
  --binary /path/to/odp_native_tail_append_probe \
  --retained-binary /tmp/litchi-goal-0457/candidate-final/odp_native_tail_append_probe
python3 -B docs/performance/results/change-0457/check.py --tag native-final -- \
  python3 -B docs/performance/results/change-0457/native-final/run.py
python3 -B docs/performance/results/change-0457/native-final/preflight.py
```

`verify.py --precleanup` checks the retained live binary, original fixtures,
exact runner paths, and output hashes. `verify.py --portable` copies the
complete change bundle and replays retained archives through the copied
independent oracle without the original binary or fixture checkout.
