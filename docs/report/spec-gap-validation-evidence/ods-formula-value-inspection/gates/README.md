# Value-inspection gate receipts

This directory retains the seven locked/offline integration gates for the
sixteen-function value-inspection profile. `Cargo.lock` is the retained gate
copy (`58b4be6cf88d7f7c5c2b16bd069a589e261e2a68e45a808a5cf3f12e1340a3e3`).
The workspace lock is a separate ambient input and remains at the repository
root.

The final sequence is:

1. create a clean checkout at the baseline commit in `baseline.json`;
2. run `stage.py CHECKOUT`, which copies the source closure declared by the
   performance profile and writes the candidate hash map;
3. run `run.py CHECKOUT TARGET`, which records the seven command logs and
   before/after source manifests; and
4. run `verify.py` and the evidence-root `verify.py`.

`stage.py` and `run.py` refuse to operate on the shared working tree. The
verifiers derive the focused test totals, changed source set, and performance
case set from retained manifests. They do not contain provisional totals or a
fixed candidate case count. Until the source freeze and gate run exist,
`verify.py --allow-pending` reports a pending receipt; a normal verification
fails closed when those receipts are absent. A present but failing receipt is
always an error, including in pending mode.

The root verifier also checks the independent Python oracle, the retained
LibreOffice fixture and its typed XML rows, and performance captures when a
complete performance results tree is supplied. Native differences are kept as
observations against the contract; they are never converted into agreement by
the gate scripts.
