# Safe W4b/W5 native qualification runner

`tools/native_qualification_runner.py` replaces the historical W4b and W5
shell launchers for final execution.  It does not contain a source commit or
binary identity.  The caller supplies the final baseline/candidate worktrees
through the manifest environment, so a returned source repair cannot silently
reuse an old receipt.

The two manifests preserve the recorded workloads and order:

- `manifests/w4b-xlsb.json`: `xlsb_w4b_probe`, `sheet_dispatch_bench`, then
  `xlsb_crud` for each leg, 31 rounds, and
  `before after after before`.
- `manifests/legacy-w5.json`: `legacy_w5_source_probe`, then
  `legacy_w5_probe` for each leg, 31 rounds, `LITCHI_PROBE_BATCH_ROUNDS=7`,
  and the same ABBA order.

Use a new external output directory for each attempt.  The target directories
must also be external to both source worktrees.  The runner refuses an output
directory containing an existing phase journal/receipt and never deletes an
existing target, binary, log, or measurement output.

Example W4b setup (replace every path with the final frozen paths):

```sh
export NATIVE_QUAL_OUTPUT=/external/native-final/w4b-20261023-run1
export W4B_BEFORE_SOURCE=/external/sources/w4b-before
export W4B_AFTER_SOURCE=/external/sources/w4b-after
export W4B_TARGET_ROOT=/external/native-final/targets/w4b-run1
export W4B_XLSB_FIXTURE=/external/fixtures/testVarious.xlsb
export LITCHI_COORD_BIN=/home/zhuhe/code/litchi-worktrees/.coord/bin
export W4B_ANALYZER=/external/harnesses/w4b-xlsb/perf/analyze.py

python3 tools/native_qualification_runner.py \
  --manifest tools/native_qualification/manifests/w4b-xlsb.json \
  --phase build
python3 tools/native_qualification_runner.py \
  --manifest tools/native_qualification/manifests/w4b-xlsb.json \
  --phase abba --cpu "$CPU"
```

Use the analogous `W5_BEFORE_SOURCE`, `W5_AFTER_SOURCE`, and
`W5_TARGET_ROOT` variables for `legacy-w5.json`, and set `W5_ANALYZER` to the
record-0359a `analyze.py`.  The W5 probes synthesize their inputs, so that
manifest has no fixture path.

Build commands are executed directly, never through a shell.  The runner
injects `--offline` and `--locked` only when missing and only before Cargo's
`--` separator, sets `CARGO_TARGET_DIR` to the manifest's explicit target,
rejects Cargo target/config/manifest escapes, forces one Cargo build job and
disabled incremental compilation, and records the effective argv and
environment policy.  Each build target must be a distinct fresh external
directory (absent or empty); no stale target is deleted.  The runner uses
fixed decimal bounds: a 19,200,000,000-byte preflight ceiling, a
20,000,000,000-byte hard target ceiling, and a 20,000,000,000-byte child
address-space cap.  These values cannot be weakened in the manifest.

A build receipt records the observed Git HEAD/tree/status, lock hashes,
probe/fixture hashes, resolved source/target/output paths, target sizes, and
every binary SHA-256.  Before ABBA, the current manifest/runner hashes,
receipt phase/lane/limits, complete event-journal hash/count/order, and all
resolved identities are revalidated.  ABBA refuses to start if any identity or
binary hash differs from the build receipt.

Each child receives a separate stdout/stderr log and a journal event.  Nonzero
exit, empty required output, missing declared output, source mutation, target
growth beyond 20,000,000,000 bytes, post-exit descendants, an unstable target,
or an unrelated Cargo/rustc/rustdoc/rustfmt/clippy or profiler process under
any `/home/zhuhe/code/litchi*` worktree aborts the run.
The process guard follows the launched child subtree, so its own compiler
descendants are allowed while another Litchi worktree's build is rejected.
It identifies the actual executable/comm/cwd; a Python command containing the
word `cargo` is not treated as a compiler.  Failed events and logs remain on
disk and the failed receipt is written before the command returns nonzero.

The receipts are orchestration evidence only.  They do not assert correctness
or performance.  The manifests route probe/sheet/source/ops stdout to the
historical `probe-*`, `sheet-*`, `source-*`, and `ops-*` filenames, retain CRUD
JSON under its historical name, and run the real lane `analyze.py` as a
fail-closed validator producing `summary.json`.  A nonempty stdout stream
alone is never accepted.
