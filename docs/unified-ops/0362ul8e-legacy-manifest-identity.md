# 0362ul8e — legacy native manifest identity repair

Date: 2026-10-24.  This is a source and validation-tool repair only.  It
does not claim a Cargo build, native execution, timing result, allocation
capture, or acceptance.

## Defect and binding

The generic runner already records the exact manifest supplied with
`--manifest`, including its absolute path and SHA-256.  The legacy validator
discarded that identity and substituted
`tools/native_qualification/manifests/{lane}.json`.  Consequently a build
receipt produced from an external final driver (for example,
`legacy-ole-properties-final.json`, `legacy-xls-properties-final.json`, or
`legacy-rtf-final.json`) was rejected even when its source, preparation receipt,
workload, and manifest hash were correct.

The repair binds one manifest identity through all three stages:

* `SafeRunner._context` exposes `{manifest_path}`.  Each of the three legacy
  manifests passes that rendered absolute path as the final validator argument.
  The workload guard requires the explicit token, output-bound validator cwd,
  and the active validator script/argv pair; it does not waive external
  manifests or relax the reviewed routes.
* `legacy_native_validator.py` accepts the manifest path explicitly, rejects
  symlinked/non-absolute paths, checks the manifest file's schema and lane,
  compares the build receipt's path and SHA-256 to that exact file, and uses
  that manifest's dirty-path allow-list when validating the preparation
  receipt.  The summary now reports the same exact manifest path.
* `prepare_legacy_native_inputs.py` rejects symlinked/non-absolute manifest
  paths before resolving and hashing the manifest.  Its existing receipt
  remains the source of the preparer/manifest/source identity binding.

The reviewed RTF, OLE-properties, and XLS-properties workloads, source fences,
consumer adapters, locks, fixtures, caps, and route checks are unchanged.
The old external packets and their false-readiness/history records are
retained; they were not rewritten in this source repair.  After this leaf is
integrated, each final packet must be freshly prepared with the repaired
manifest templates and current final source pin before any build slot.

## Build-free validation

The source-only focused suite passed with 46 tests:

```text
PYTHONDONTWRITEBYTECODE=1 python3 -m unittest \
  tools.test_legacy_native_qualification \
  tools.test_native_qualification_runner \
  tools.test_no_3rdparty_references
Ran 46 tests
OK
```

The new tests load an external final-manifest copy, render its explicit path,
reject a wrong validator template, and exercise the validator's external
manifest path/hash/lane checks before receipt provenance.  `py_compile` and
`git diff --check` also passed.  No Cargo/native helper was invoked.

The repaired source hashes at preparation time are:

```text
tools/native_qualification_runner.py: 0864679291b3a2aaac93c6ae324da5e3a9c57506328efcbf3eac39fbb64d3013
tools/native_qualification/legacy_native_validator.py: 9f2d2f7a4eac2eb51b8918b3e8e8d0670cec6d5c20ad25b30bf4010f80593dc7
tools/native_qualification/prepare_legacy_native_inputs.py: 3aad0d288aa653643d2efa2782b8ac5d96ad82fbdc22a07d57cb9fa1b47871bf
```

The leaf is based on integration tip `472b6705475b073e7fe08f6e11ee17fb1e29f838`.
The commit produced from this worktree is the source of truth for the final
hashes after review; no push or merge is performed here.
