# 0359qha — DOC body-toggle final calibrated performance protocol

- **Status:** executable handoff corrected; no Cargo, native, profiler, or
  allocation command was run in this turn.
- **Date:** 2026-10-24.
- **Supersedes:** the execution instructions in 0359qh only. The provenance,
  accepted source hashes, correctness output, CHPX result, and old diagnostic
  hashes in 0359qh remain unchanged.

## Why this follow-up exists

0359qh described only one short native block with binary argument `1` and
`LITCHI_PROBE_BATCH_ROUNDS=3`. That is not the approved W5 qualification
protocol. 0359a §7 requires 31 measured samples per scenario and two complete
A-B-B-A blocks. The reviewed W5 runner and calibration references are retained
at:

~~~text
(deleted-coord-artifact)/abba.sh
SHA-256 0dcde7efdea349e95822abf9a54063b2b1e5e73e0fd1148a359702d524c4989b
(deleted-coord-artifact)/analyze.py
SHA-256 14bf15e41202cd610c8bcecb2d239d9ceef3256b354c9144b8b6aa1a808f4d21
(deleted-coord-artifact)/summary.json
SHA-256 6419d8c05d8494abfc922883e62b34897ccd67a78d42f8e178d79a743587be90
~~~

The historical W5 schedule is recorded explicitly as 31 outer rounds and
`LITCHI_PROBE_BATCH_ROUNDS=7`. The source-bound DOC routes call
`time_rounds(..., batch_rounds, ...)` directly, so the process argument does
not control their sample count. The final symmetric schedule therefore sets
both `ROUNDS=31` and `LITCHI_PROBE_BATCH_ROUNDS=31`, then validates every JSON
receipt's actual `samples` field as exactly 31; setting the environment to 7
would silently produce only seven samples. The runner uses a pinned CPU, a
20-GiB scope, and A-B-B-A twice.

The 250,000-us minimum is aggregate timed work per native leg, not a
per-operation floor. The leaf route uses 16 fixed 31-sample probe
invocations, while the 1,000-toggle batch route uses one fixed 31-sample
invocation; both endpoints use the same route schedule. For each exact
31-sample receipt, the validator certifies the conservative lower bound
`16*median_us + 15*min_us` (the sorted median is sample 15), then sums those
bounds across invocations and requires at least 250,000 us. This is a
certified lower bound, not an actual sample sum. Process wall time is retained
for diagnosis only because it includes fixture/setup work.

## Correct route meaning

The original queued residual is the *single* leaf body toggle on the 1,000-
paragraph fixture (`doc/leaf_write`, paragraph 500). It is not a 1,000-toggle
queue. `doc/batch1k_leaf_write` is a separate required CHPX coalescing and
fidelity route: it stages 1,000 body toggles and publishes once. Both routes
must be qualified, but the result must name these concerns separately.

## Accepted final-472b identities

The protocol reuses, without rebuilding, the source-bound binaries from
0359qg/0359qh:

| item | candidate | baseline |
|---|---|---|
| source commit | `472b6705475b073e7fe08f6e11ee17fb1e29f838` | same |
| source tree | `ebc1cee05f124c0c63977aff07a23288c19e7d00` | same |
| Cargo.lock SHA-256 | `4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` | same |
| body/revision SHA-256 | `db01cfce38c13c6c806382ffede990fda2b3bbe4dd53f59c4c5cb9feba9777d3` / `029721d34e7ad658e0064afd258cdb2b5a40e330110255b157aafb83f0d7b12a` | `6dc26da8130659c212dd29f26499cdaba7c4849268583b44dfcd9413a36887a3` / `b7155dc7c9aae6e318845b5073da4c0f275889125d38eae61df091aa2a2166df` |
| probe source SHA-256 | `b808e4dcba33dcafd4dff7dd22dc250e4170f2613839044fa8097263f5020965` | same |
| binary SHA-256 | `ba7042fd413c26764986b0ed9dcd79beca85e91231f0beb2f850550152e14877` | `2d62dacae12b49214a086c109b506ca8eb69a33cf5e770627794efecdbcba810` |

The full correctness receipt remains outside the checkout at
`(deleted-coord-artifact)/actual-receipt.txt`,
SHA-256 `a0dfb1d1b66c16249939a702d26565466a3ca8b86d11e9bc87461bf5df414b3a`.
The expected JSON output hashes remain the 0359qh values; batch raw-byte
inequality and the passing CHPX oracle remain explicit.

## Executable packet

The reviewed runner is tracked beside this record:

~~~text
docs/unified-ops/results/0359qha/run.sh
SHA-256 2d854437d19e359c327b42b89727f581338b659f4c2ba1589131253893f43a61
~~~

It performs no build. It refuses a pre-existing `RUN`, verifies every source,
lock, probe, and binary hash, creates fresh per-leg/route/invocation TMP
paths, validates every JSON scenario/sample/output digest and aggregate
schedule, and writes a manifest outside the checkout. It has two phases:

~~~sh
RUN_ID=replace-with-a-new-id
CPU="${CPU:?set quiet pinned CPU}" PHASE=allocation RUN=/external/fresh/doc-toggle-472b-alloc-"$RUN_ID" \
  docs/unified-ops/results/0359qha/run.sh
CPU="${CPU:?set quiet pinned CPU}" PHASE=native RUN=/external/fresh/doc-toggle-472b-native-"$RUN_ID" \
  docs/unified-ops/results/0359qha/run.sh
~~~

The allocation phase uses one route-local profile sample (`PROFILE_ROUNDS=1`,
`PROFILE_BATCH_ROUNDS=1`) for each leg and route, under a pinned CPU, 20-GiB
scope with swap disabled, and a 900-second profiler timeout. It retains
Callgrind/Massif output and annotations; this is allocation evidence, not the
31-sample native result.

The native phase first runs candidate and baseline calibration legs for each
route with the final `31/31` settings and its fixed invocation count. The
leaf calibration demonstrates that a fast operation can meet the aggregate
floor through repeated bounded probe invocations; a genuinely short
aggregate rejects. Calibration validates output, reported operation timing,
aggregate count, and host quietness but does not use process wall time as the
threshold. It then runs, for each route and for `block=1` and `block=2`,
exactly:

~~~text
baseline, candidate, candidate, baseline
~~~

Every native leg is wrapped in a 900-second timeout and a 20-GiB,
no-swap systemd scope, pinned with `taskset`. The runner records actual host
snapshots before, during, and after each leg. A load above the fixed 1.0
threshold or any `cargo`, `rustc`, `ninja`, or `make` process marks the packet
rejected; it does not get called timing noise.

The static executable-protocol test is:

~~~text
docs/unified-ops/results/0359qha/test_protocol.sh
SHA-256 ced8ddcb6f3eed942d893773ee3703e7532be3c9a6606c0d3bb5d2b8191887dd
~~~

It runs `bash -n` and checks the fixed 31/31 sample contract, two ABBA
blocks, order, per-route TMP, fresh-root guard, output validation,
CPU/cgroup, timeout, profile artifact presence/annotation, early source-hash
failure, certified aggregate timing, and before/during/after quiet checks. It
passed in this turn:

~~~text
0359qha protocol static checks passed
~~~

## Acceptance boundary and remaining gap

No native or profiler result is claimed here. The previous 3cad/9e1cd
Callgrind/Massif and loaded-host ABBA artifacts remain historical diagnostics;
they cannot qualify final-472b. Final acceptance still requires one fresh
external packet from the executable runner with all source/binary/lock/output
hashes, calibration legs, two complete ABBA blocks per route, profile hashes,
and quiet/load rejection evidence.
