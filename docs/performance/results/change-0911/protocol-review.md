# 0911 PPTX owning-route opened-snapshot Callgrind protocol review

This is a finite, read-only review of the proposed 0911 attribution capture.
It does not build the harness, run the benchmark, collect a profile, or change
production code. The capture is intended to isolate the existing owning route's
opened-presentation snapshot phase from the later edit, apply, publication, and
semantic checks.

## Frozen binary and route schedule

Use the retained 0910 executable at
`docs/performance/results/change-0910/binaries/native`, SHA-256
`f4ad7ee158061cc80354c3c8674c9fe6a8100ed76a0c8f743653efcb06bb2e8d`, with no
rebuild and no source overlay. Its recorded size is 9,301,456 bytes. Every
child receipt must bind that exact path, size, hash, command, environment,
Valgrind/Callgrind identity, and CPU affinity. The 0910 source, build, and
lockfile receipts remain the custody authority; the dirty shared checkout is
not a substitute for them.

Run four fresh processes serially on CPU 12 in ABBA order:

| leg | route | role |
| --- | --- | --- |
| owned-a | `owned` | first owning snapshot profile |
| materialized-a | `materialized` | first materialized snapshot profile |
| materialized-b | `materialized` | second materialized snapshot profile |
| owned-b | `owned` | second owning snapshot profile |

Each child must use `warmup=0`, `samples=1`, a new artifact directory, and a
new Callgrind output path. The benchmark command shape is the retained
`pptx_owned_benchmark` example:

```text
taskset --cpu-list 12 valgrind --tool=callgrind --collect-atstart=no \
  --toggle-collect=litchi_pptx::package::model::Package::opened_presentation_transaction \
  --zero-before=litchi_pptx::package::model::Package::opened_presentation_transaction \
  --dump-after=litchi_pptx::package::model::Package::opened_presentation_transaction \
  --callgrind-out-file=<fresh-callgrind-path> \
  docs/performance/results/change-0910/binaries/native \
  <route> 0 1 <fresh-artifact-directory>
```

Do not add a warmup, repeat a child after inspecting its profile, substitute
`timed_iteration` as the owner, or use a different executable. The two
ABBA pairs provide an order check only; four one-sample profiles do not
establish instruction variance or a confidence interval.

## Why the selected owner is admissible

`nm -C --defined-only` on the retained executable reports exactly one defined
text symbol named
`litchi_pptx::package::model::Package::opened_presentation_transaction`.
The same executable also retains the caller
`litchi_perf_baseline::pptx_owned_benchmark::timed_iteration`. The owner is the
method timed as `snapshot_ns` at
`tools/perf-baseline/src/pptx_owned_benchmark.rs:294`, between the route
ingress and the 64 shape-text replacement loop. The later
`apply_opened_presentation_commit`, `Package::to_bytes`, and bounded sink write
are after the owner returns and therefore outside the selected collection
interval.

The benchmark's one measured operation calls the owner once. With zero
warmups and one sample, `run` reaches one `timed_iteration` call, and that
function contains one `package.opened_presentation_transaction()` call. The
source corpus builder at `lib.rs:17052` uses `Package::from_bytes` and OPC
inspection; it does not open a transaction. The output oracle at
`lib.rs:50745` also uses `Package::from_bytes` and OPC inspection. It runs
after `timed_iteration` returns. The selected example is an opt-in entry point
that does not execute the general harness matrix or the unrelated control
fixture. This static audit must be repeated against the frozen source receipt
before accepting the raw profiles.

The owner is preferable to the whole `timed_iteration` helper for this
question. The helper includes clone, ingress, edit/commit, apply, publication,
and sink work, while the selected method is the existing source-level boundary
for the opened snapshot. The owner filter does not alter the workload; it only
selects the instruction interval around the one dynamic snapshot call.

## Collection and contamination gates

`--collect-atstart=no` must be present. `--zero-before` and `--dump-after`
must name the exact same owner. These settings are necessary but are not
themselves proof of scope. A raw reader must inspect every dump and every
positive owner edge. The accepted profile must satisfy all of the following:

1. The Callgrind header has the expected `Ir` event, the command names the
   frozen binary and the exact owner options, and the profile's binary path
   and child command match its receipt.
2. There is exactly one positive owner collection interval. Its owner record
   is the exact demangled method name, its incoming caller is
   `litchi_perf_baseline::pptx_owned_benchmark::timed_iteration`, and its
   dynamic call count is one. The raw summary and totals for that interval
   are positive and equal.
3. There are no positive owner edges from corpus construction, output
   verification, report writing, or any unrelated caller. An extra positive
   owner edge is not discarded as setup; it makes the profile ambiguous and
   stops admission until the cause is explained from the raw graph.
4. The positive dump is attributable to the selected owner interval, and a
   later process-termination dump, if emitted, has zero summary and totals.
   A nonzero termination dump or a second positive owner dump is a failure.
5. The profile reader accounts for all positive function costs. The owner
   self cost plus its immediate direct-child inclusive costs must reconstruct
   the owner edge without double-counting nested descendants. Preserve the
   complete raw graph; do not use `callgrind_annotate` text as the authority.

The normal JSON report must independently pass its one-result route, `samples=1`,
`warmup=0`, frozen corpus, source identity, output identity, sink bound,
semantic output, and source-immutability gates. Those checks are outside the
Callgrind owner interval and cannot be used to identify the interval. The
report's semantic verifier must remain outside the collected method; do not
move it into the owner, call it from a wrapper, or subtract its instructions
after collection.

## Attribution boundary and debug-info limitation

The retained release binary was built with debug information disabled. Raw
Callgrind records must therefore be admitted by exact function names and
caller/edge structure; source line numbers are unavailable or unreliable for
this capture. In particular, do not claim that an instruction row came from
the source line for `snapshot_ns`, `apply_opened_presentation_commit`, or any
other phase merely because that line exists in the source. The snapshot-versus-
apply boundary is established by the selected owner returning before the
apply call, not by a line-number attribution.

Within the owner, retain self Ir and immediate direct-child inclusive Ir as a
disjoint partition when the raw graph exposes them. Named children such as
package open/catalog construction and presentation snapshot construction may
be shown as diagnostic rows only when their exact function records and edge
relationships are present. Inclusive descendants must never be added to their
parent's row. A missing or inlined child is reported as unavailable; a nearby
function name is not a substitute.

The selected owner measures the implementation of
`opened_presentation_transaction` and its direct call tree. It does not
measure the caller's preceding input clone or ingress, the later shape-text
staging and commit, `apply_opened_presentation_commit`, serialization, sink
writes, semantic reopen, raw preservation checks, digesting, or JSON/artifact
writes. The owning and materialized profiles are descriptive instruction
counts for the corresponding snapshot method on this one generated in-memory
fixture. They are not milliseconds, native cycles, allocations, RSS, cache
misses, physical I/O, or a causal proof of the 0910 wall-clock ratio.

## Route and corpus gates

Both routes must use the exact 0910 media-rich corpus: 200 slides, eight text
boxes on each slide, eight deterministic 2 MiB media parts, 229 logical OPC
parts, 445 ZIP members, source length 17,017,139 bytes, and source SHA-256
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`. Each
child's one measured output must be 17,017,145 bytes with the stable 0910
output SHA-256
`23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9`, and
must pass the existing semantic, sink, digest, and source-immutability gates.

The two routes have different ingress implementations, so their snapshot
instruction totals are descriptive route observations. The profile must not
be combined with 0907/0908 counts or converted into an elapsed-time ratio.
The source-backed route's materialization occurs before the selected owner;
the owning route's preservation provenance also exists before the owner. Those
route differences are intentionally outside this narrow snapshot attribution.

## Stop conditions and reporting

Stop admission and retain all receipts and raw dumps if the binary hash,
owner name, caller, dynamic call count, route schedule, affinity, corpus, or
collection options differ; if the exact owner is absent or duplicated; if an
extra positive owner edge appears; if source/line information is incorrectly
needed to classify the interval; if the termination dump is nonzero; or if
the JSON semantic report fails. Never repair a failed profile by choosing the
largest dump, subtracting the semantic oracle, or selecting a dump by order.

If all gates pass, report one measured owner `Ir` row per route and retain
both ABBA observations. A second sample per route may be described as a
replication, but no confidence interval or native speedup claim is warranted.
The result may identify direct function-level contributors inside the snapshot
owner, with the debug-info limitation stated above. It does not authorize a
production change; the next step still requires correctness-first evidence
and broader producer/input coverage.
