# Change 0343: OOM-safe non-iWork validation gate

Date: 2026-08-31

Status: validation protocol; no measurement or performance claim

## Scope

This record defines a bounded validation gate for the non-iWork crates.  It is
an execution and evidence policy, not a change to the crate graph or to the
production API.  iWork packages remain outside this gate because their
`prost-build`/`protoc` path and feature closure have separate host
requirements and should be scheduled by their own bounded gate.

The unit of scheduling is a validation root: one package or package/feature
selection, one command class, and one private target directory.  Roots may be
listed in a matrix, but they are never executed concurrently.  A root that
fails, is cancelled, or is ineligible produces a record and does not unblock a
later root until its cleanup decision has been made.

One mode invocation writes one mode-level JSON report.  It is not one report
file per root: the report's ordered `phases` array contains the root commands,
facade commands, and any cleanup phases executed by that invocation.

## Hypothesis and mechanism

The working hypothesis is that host OOM is primarily caused by the aggregate
working set of overlapping Cargo roots, retained incremental state, debug
information, and simultaneous test execution.  A broad workspace command can
therefore exhaust memory even when each individual package is buildable.  This
gate tests that hypothesis operationally by making the maximum scheduled
parallelism one and by removing the largest avoidable retained-build inputs.

Every root is run with the following resource contract:

```text
CARGO_BUILD_JOBS=1
CARGO_INCREMENTAL=0
CARGO_PROFILE_DEV_DEBUG=0
CARGO_PROFILE_TEST_DEBUG=0
```

The gate also passes `--test-threads=1` to every test harness.  Thus there is
one Cargo rustc job, one test thread, and no overlapping validation root.  A
mode invocation uses its explicit or default isolated `CARGO_TARGET_DIR`, and
its roots reuse that directory only in the serialized order defined by the
runner.  A scheduler must not replace these constraints with a larger `-j`, a
second concurrent invocation, or a parallel test runner.

Mode-specific lint strictness is explicit.  Clippy carries exactly
`--no-deps -- -D warnings`; documentation and documentation tests set
`RUSTDOCFLAGS=-D warnings`; the deprecated-API mode sets
`RUSTFLAGS=-D deprecated`.  These required values are enforced at the phase
boundary and are recorded only through the bounded environment allow-list;
ambient flags are not a substitute for the gate's exact lint policy.

Metadata, dependency-tree, and host-`rustc` probes stream stdout and stderr
through finite byte caps.  The caps apply independently to each captured
stream, and a cap hit fails closed with a bounded capture-limit diagnostic
rather than materializing unbounded output.  The same rule applies to retained
phase diagnostics and report errors.  Target-directory telemetry also has
finite traversal caps for visited directories, entries, and accumulated file
bytes; exceeding a cap yields an incomplete/capped scan status and never a
false complete footprint.  Scans use `lstat` and do not follow symlinks.

The child is placed in a reapable process group.  Ctrl-C requests termination
of that child group with SIGTERM, waits for a bounded grace interval, escalates
to SIGKILL when necessary, and reaps the child before the mode reports
`interrupted`.  This is best-effort process cleanup, not a host-wide process
kill.

These controls bound known sources of amplification.  They do not guarantee
that a compiler, linker, proc macro, test, or host kernel cannot allocate more
memory than is available.  A killed process is a bounded-gate failure, not
evidence that the package is invalid.

## Telemetry schema and claim boundary

The runner emits one JSON object for the complete mode invocation.  The
mode-level shape is:

```text
{
  "version": 1,
  "mode": "check|clippy|doc|lib-tests|doc-tests|deprecated",
  "outcome": "running|passed|failed|interrupted",
  "phases": [
    {
      "mode": "<same mode>",
      "index": 1,
      "scope": "bulk-test/<package>|bulk-clean/<package>|facade...",
      "status": "passed|failed|error|interrupted",
      "returncode": "<optional integer>",
      "elapsed_ns": "<integer>",
      "target_before": "<bounded footprint>",
      "target_after": "<bounded footprint>",
      "child_rss": "<bounded RSS observation>",
      "env_keys": "<allow-listed names>",
      "cargo_env": "<allow-listed bounded values>"
    }
  ]
}
```

The top-level `version`, `mode`, and `outcome` describe the invocation, not an
individual root.  `phases` is ordered by actual execution and includes the
cleanup phase when one is run.  The runner may include additional top-level
clock, host, target, environment, cleanup, feature-unification, and limitation
objects, but root facts belong inside the nested phase records.  Command lines,
environment values, diagnostics, and errors are bounded; secrets and
unbounded command output are not serialized.

`child_rss` is a sampled sum of readable descendant `VmRSS` values with a
20-ms sampling interval and a high-water value when available.  Its status is
explicit: `available` means samples were obtained, `partial` means some
samples or descendants were unreadable while other samples were obtained, and
`unavailable`/`not_collected` means no usable sample was available or RSS was
not collected.  Even `available` is not a complete OS peak-RSS measurement:
short-lived processes, kernel memory, filesystem cache, compiler-server memory,
and unrelated processes are outside the stated scope.

The target scan is a logical regular-file inventory with bounded traversal; it
is evidence about the selected target path, not a full disk-usage or kernel
memory measurement.  The report records scheduling and cleanup evidence, not
compiler correctness, reproducibility, storage performance, or causal proof
that a particular allocation caused an OOM.

## Package and target cleanup policy

`lib-tests` emits a cleanup phase immediately after each successful bulk test
root.  That phase runs `cargo clean --package <package>` for the root that just
passed, reducing the test-binary accumulation before the next root.  The
currently failing root is retained: execution stops before its cleanup phase,
so its target artifacts remain available for diagnosis.  A successful facade
test root is handled by the runner's corresponding facade cleanup phase.

Other modes retain their shared per-invocation target artifacts by policy;
their roots are serialized for memory control, but successful roots are not
cleaned between phases.  This permits later phases to reuse the target and
keeps the report's before/after footprint meaningful.  There is no claim of a
bounded failure bundle: failure retention is the target/report policy above,
not a promise to package, cap, or preserve a separate diagnostic archive.

The explicit `CARGO_TARGET_DIR` and `--record-file` paths are trusted caller
inputs.  The gate does not recursively delete an arbitrary explicit path.
After Ctrl-C, SIGTERM, SIGKILL escalation, or a process that exits outside the
reapable group, target artifacts and same-directory report temporary files may
remain.  Such residual state is reported as unknown/retained where it can be
observed and must be cleaned by the operator using a path they control.

## Validation status

- Focused runner validation: `python3 -m unittest tools.test_non_iwork_gate` ->
  **51 tests passed in 0.328s** after adding the post-reap RSS regression.
- The mode-level gate execution and its JSON report remain unrun in this
  documentation batch; host/toolchain identity, streaming-cap outcomes, RSS
  status, and cleanup outcomes are therefore still placeholders.
- No Cargo build, `rustc` invocation, broad Cargo matrix, or performance
  benchmark was run for final validation.
- One intermediate failed unit-test run exposed stale subprocess mocks and
  serially invoked non-building `cargo metadata/tree`; those processes exited.
  The mocks were corrected to the capped boundary, and the final unit tests
  use fakes for command execution.
- The broad Cargo matrix (`cargo check --workspace --all-features --lib
  --tests`, the corresponding workspace test suites, and related crate-wide
  gates) was deliberately not run in this batch.  This change only records the
  gate policy, and launching that matrix here would both violate the
  documentation-only scope and reintroduce the concurrent/high-memory
  conditions this protocol is intended to control.

## Performance claims

`performance_claim: none`

No throughput, latency, wall-clock, allocation, RSS-capacity, compilation-time,
or OOM-freedom claim is made.  Single-job execution, disabled incremental
state, disabled debug information, serialized roots, one test thread, bounded
captures, and cleanup phases are resource-control and diagnosability measures;
they are not benchmark results.

## Feature-unification limitation

`--all-features` is applied per selected package root, while the facade's
default, safe-feature, and combined-safe-feature closures are separate roots.
Cargo may still unify features across dependencies inside one selected root.
Serializing roots limits process overlap but does not prove the feature
unification that a single aggregate workspace invocation would produce.
Aggregate workspace feature behavior remains outside this gate's claim.

## Follow-up

Populate the placeholders only from a run that enforces the complete contract,
writes one mode-level report with ordered nested phases, and records cleanup
outcomes.  Any future change to the matrix, retention policy, telemetry fields,
capture/traversal caps, or iWork boundary should update the schema/version and
document why the memory and evidence boundaries remain valid.
