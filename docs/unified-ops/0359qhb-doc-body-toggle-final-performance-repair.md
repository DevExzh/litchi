# 0359qhb — final DOC body-toggle qualification repair

- **Status:** source-only protocol repair; no Cargo, native, allocation, or
  profiler command was run.
- **Date:** 2026-10-24.
- **Supersedes:** the executable acceptance details in 0359qha. The accepted
  472b source, binary, correctness, CHPX, and historical diagnostic identities
  remain unchanged.

## Review findings closed

The 0359qha handoff had four acceptance holes. This packet closes each one in
the tracked runner and its executable protocol test:

1. `verify_source` is called in a standalone status-checked statement, and
   every check returns failure explicitly. A bad early source hash cannot be
   hidden by Bash's `||`-list errexit behavior. The candidate must be clean;
   the baseline must have exactly the two reviewed overlay paths, both as
   unstaged modifications:

   ```text
   crates/litchi-doc/src/body_text.rs
   crates/litchi-doc/src/tracked_revision/package.rs
   ```

2. The selected DOC routes pass `batch_rounds` directly to the probe's
   `time_rounds`. The runner sets both the process argument and
   `LITCHI_PROBE_BATCH_ROUNDS` to 31 and validates every JSON row's actual
   `samples == 31`. It no longer treats `/usr/bin/time` `%e` (which includes
   fixture/setup work and rounds to centiseconds) as an operation threshold.

3. Operation timing is checked from the probe receipt. Every row must contain
   finite `min_us` and `median_us`; `doc/batch1k_leaf_write` must have both at
   least 250,000 us. The single-toggle `doc/leaf_write` route is still checked
   for exact samples and output but has no artificial 250-ms floor.

4. Allocation now fails if Callgrind or Massif output is absent, empty, or
   cannot be annotated (`callgrind_annotate`/`ms_print`). A `check-profile`
   phase exercises incomplete and complete artifact directories without
   running a profiler.

The native route remains two complete A-B-B-A blocks per route, with fresh
per-run and per-leg output/TMP roots, pinned CPU, a 20-GiB no-swap systemd
scope, timeout, and before/during/after host quiet checks. No performance
claim is made by this source-only repair.

## Executable evidence

The protocol test is:

```text
docs/unified-ops/results/0359qha/test_protocol.sh
SHA-256 8d09537ac9307bac4bc1ed4da22361660c25a977b79748c0eb29a595d4dfb10f
```

The repaired runner is SHA-256
`1f3eeacca4baf44cf5d4ebc8976c1853bf33e1c1b24217b07ac65b2f7c69b194`; the
receipt validator is SHA-256
`a3a8f696c0ad02c5707d50aed3f17c4365b52dee88cb4d78b98d80dc7f7dd80f`.

It runs `bash -n`, the Python receipt validator, an early bad-hash check,
exact profile artifact/annotation failure and success checks, missing-receipt,
249,999-us subthreshold, and seven-sample count rejection cases. The runner,
test, and validator hashes above bind this exact source-only packet.

The source-bound identity and output hashes are the exact final-472b values
already recorded in 0359qha; this record intentionally adds no large capture,
binary, or profiler artifact.
