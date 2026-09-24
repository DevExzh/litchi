# Candidate integration gates

Run these gates after the candidate sources have been synchronized to the
isolated workspace and its build target has been released by the current owner:

```sh
python3 docs/report/spec-gap-validation-evidence/ods-formula-array-reference-evaluation/gates/run.py \
  --workspace /home/zhuhe/code/litchi-array-worktree
```

The runner writes logs and `results.json` beside itself. It requires matching
canonical and isolated source hashes before starting Cargo. It records both
source trees again after the checks, together with the compiler version,
commands, statuses, elapsed times, runner hash, workspace HEAD, and selected
build environment. An unchanged HEAD alone is insufficient because candidate
source is copied into the detached baseline workspace before validation.

The source manifest covers the ODS crate, its five local dependency crates,
their source/test/example files, Cargo manifests and lockfile, build scripts,
and workspace Cargo configuration. Files in the external `test-data` tree
are outside this manifest; it is not a complete archive of every test input.
The six-crate list must be updated if the candidate adds a local dependency.

Checks are all-feature/all-target ODS tests, warning-denied Clippy, warning-denied
public documentation, doctests, and formatting. Boundary-policy checks and
performance captures are separate gates. The target and compiler temporary
directory are on disk under `/home/zhuhe/code/litchi-array-*`, not tmpfs.

Adding or smoke-checking this runner is not evidence that the candidate passes.
Only completed logs and a matching `results.json` establish a gate result.
