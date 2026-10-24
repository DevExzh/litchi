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

2. The historical reviewed schedule is explicitly 31 outer rounds with
   `LITCHI_PROBE_BATCH_ROUNDS=7`; the final symmetric schedule is 31/31. The
   selected DOC routes pass `batch_rounds` directly to the probe's
   `time_rounds`, so the runner sets both the process argument and
   `LITCHI_PROBE_BATCH_ROUNDS` to 31 and validates every JSON row's actual
   `samples == 31`. It no longer treats `/usr/bin/time` `%e` (which includes
   fixture/setup work and rounds to centiseconds) as an operation threshold.

3. The 250,000-us minimum is aggregate timed work per leg. The leaf route uses
   16 fixed 31-sample invocations and the batch route uses one fixed
   31-sample invocation, identically for candidate and baseline. Every row
   must contain finite, non-negative `min_us` and `median_us`, with
   `min_us <= median_us`. For each exact 31-sample receipt, the validator
   certifies the conservative sorted-sample lower bound
   `16*median_us + 15*min_us`, then applies the floor to the sum across the
   fixed invocation count. This is a certified lower bound, not the actual
   sample sum. A legitimate 1,000-us leaf operation passes its
   16-invocation calibration, while a genuinely short aggregate rejects; the
   counterexample where `median_us*31` passes but this bound fails is tested.
   No individual 1–2 ms leaf operation is required to meet 250 ms.

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
SHA-256 ced8ddcb6f3eed942d893773ee3703e7532be3c9a6606c0d3bb5d2b8191887dd
```

The repaired runner is SHA-256
`2d854437d19e359c327b42b89727f581338b659f4c2ba1589131253893f43a61`; the
receipt validator is SHA-256
`5a22f9503831e03b6766143cc07a00271ef2c439225f0fdf6fe7bbbcd541e750`.

It runs `bash -n`, the Python receipt validator, an early bad-hash check,
exact profile artifact/annotation failure and success checks, missing-receipt,
fast-leaf aggregate success, genuinely short aggregate rejection, and
seven-sample count rejection cases. The runner, test, and validator hashes
above bind this exact source-only packet.

The source-bound identity and output hashes are the exact final-472b values
already recorded in 0359qha; this record intentionally adds no large capture,
binary, or profiler artifact.
