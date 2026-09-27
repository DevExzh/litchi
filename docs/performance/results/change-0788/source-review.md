# 0788 RSS phase probe source review

This is an independent read-only review of the current archived probe source.
It does not modify production sources or `tools/perf-execution`, and it makes
no memory, RSS, allocator, or performance claim. The reviewed files are
`probe-src/main.rs`, its two manifests, the copied lockfile, and the probe
README.

## Disposition

The probe is suitable for the frozen diagnostic lane. Its source diff against
the native `tools/perf-execution/src/main.rs` is limited to the opt-in phase
protocol, phase plumbing around the Parts route, explicit report destruction,
and two probe tests. The corpus generator, payload bytes and order, source
versions, limits, budget construction, `measure` closure, verification, source
observer, resource snapshots, report schema, and normal CLI fields remain
unchanged. The `run_sample` compatibility wrapper is test-only under `cfg(test)`
and therefore does not introduce an unused production item.

The copied `Cargo.lock` has SHA-256
`93a1e7b103f7d4055945faaeda3f3b0d2e225778dffc7447a3ef640f2327e232`, matching
the frozen native lockfile. The materialized manifest retains the native
dependencies and release settings, adds only the independent binary name
`cached-part-memory`, and the template retains `@SRC@` for replay. The packet's
quality receipt records successful format, check, test, Clippy, and doc gates
for this exact probe source; the repository-boundary gate is a separate packet
check and is not inferred here.

## Lifecycle and marker ordering

The global sequence is stable:

1. The enabled handshake creates its standard stdin/stdout handles, caches the
   child PID, and emits `startup` before the output existence check or corpus
   construction.
2. `corpus_ready` follows deterministic corpus construction.
3. Warmups use `usize::MAX` as their internal sample value and emit no sample
   marker. `warmup_done` follows all warmups.
4. Measured samples are collected with the original report allocation and
   lifecycle. `samples_done` follows the final sample.
5. `report_written` follows the existing create-new report write. The report is
   then explicitly dropped, and `report_dropped` follows that drop before
   `run` returns.

For an observed first or last measured Parts sample, the markers have these
boundaries:

| marker | source location and boundary | retained state at the marker |
| --- | --- | --- |
| `package_ready` | immediately after `SourceBackedPackage` construction and before URI-vector allocation | package, source, context, root budget |
| `after_preload` | primed: after verified preload drop and observer reset; fresh: after observer reset at the point immediately before the operation setup snapshot | package, URI vector, source, context, root budget |
| `after_operation` | after the original `measure` call and resource snapshot, before `verify_parts` | returned batch plus package and setup state |
| `after_batch_drop` | after source-metric snapshot and explicit `drop(batch)` | package and setup state |
| `after_package_drop` | after explicit `drop(package)` and `drop(context)`, before the original after-drop resource snapshot | source Arc, URI vector, root budget, and local handshake state |

All five markers are outside the `measure` closure and therefore cannot enter
the reported wall or process-CPU interval. The fresh `after_preload` marker has
the same protocol name as the primed marker; the command's `--state` identifies
which meaning applies. The explicit `drop(context)` before
`after_package_drop` matches the packet's frozen “package and context drops”
scope. The source Arc, URI vector, and root budget intentionally remain alive
through that observation, as documented by the packet plan and README.

`should_mark_sample` excludes the warmup sentinel, marks sample zero and the
last sample, and marks a one-sample run once rather than twice. Thus the
expected marker counts are 11 for the one-sample protocol and 16 for the
30-sample protocol: six global markers plus five sample markers per observed
endpoint.

## Protocol and calibration review

When and only when `LITCHI_RSS_PHASES=1`, `from_env` accepts the Parts route,
constructs `BufWriter<Stdout>` and `BufReader<Stdin>`, and records
`std::process::id()`. Other routes fail closed before emitting a startup marker.
With the variable unset, no stdin/stdout handles are constructed and no
diagnostic record or ACK read occurs. The normal route logic still uses the
same source, operation, verification, and report code. The disabled binary is
therefore the correct diagnostic-off control for handshake perturbation, while
its timing and RSS must remain separate from the frozen native lane.

Each enabled record is exactly:

```text
RSS0788\tphase\tsample\tpid\n
```

The writer flushes before reading exactly two bytes. Only `+\n` is accepted;
EOF, a short read, a wrong ACK byte, or a wrong trailing byte returns an error.
The phase tests exercise a valid record, a wrong ACK, and a truncated ACK. No
`/proc` parsing, allocator counter, unsafe code, SIGSTOP, or runtime
instrumentation is present in the child. The PID is available for the parent
to verify against `/proc/PID/exe` and its parent relationship before replying.

The parent must retain the diagnostic limitation: waiting for an ACK can fault
or allocate pages, and the marker/ACK itself can change point-in-time RSS. It
must snapshot only after receiving a marker and before sending its ACK, retain
the raw snapshot files, and never pool enabled or disabled diagnostic runs with
native timing or RSS. The current packet plan freezes the exact `+\n` protocol,
phase names, PID field, and 11/16 snapshot counts. At this review point the
capture driver and live `/proc` receipts are not retained, so external
executable/PID/PPid validation and end-to-end ACK sequencing remain capture
responsibilities rather than claims made by this source review.

## Source and report invariants

The Parts sample continues to reset source metrics after setup/preload, starts
the original `measure` boundary at the same public operation, verifies bytes
after the timer, snapshots source metrics before destruction, and checks worker
and I/O release after the original context teardown. The report still uses
`litchi.execution-baseline.v1`; no diagnostic phase data is serialized into the
native report. `source-metrics` remains a separate feature-enabled observer
build, while the native tool remains feature-off and byte-identical.

The probe's extra protocol state is local to the diagnostic executable. A
failure while waiting for an ACK can occur after a report has been created or
after `report_written`; the parent must treat the nonzero child exit as a
failed diagnostic process and retain the output for custody inspection rather
than interpreting it as a successful sample. This does not alter the existing
create-new report behavior or the source/output verification contract.

No source-level blocker was found for the frozen 0788 capture. Final
acceptance still depends on the parent capture receipts proving the expected
phase namespace, PID/executable/PPid identity, snapshot custody, terminal
success, and the packet's independent replay checks.
