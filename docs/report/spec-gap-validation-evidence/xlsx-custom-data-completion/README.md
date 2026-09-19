# XLSX Custom Data caller-limit completion

This batch finishes caller-selected resource profiles for the existing Custom
Data package owner and recognized connection bindings. It preserves inert
payload storage, source XML, atomic edits and reversible source-checked patches.
See the [contract](contract.md) for precise byte-accounting and temporary-buffer
semantics, and the [independent review](review.md) for issue dispositions.

The completion includes namespace declaration/prefix admission, bounded
connection/query-table XML processing, source and output limit separation for
shrinking XML edits, and connection-owner relationship provenance and output
accounting. The shared XML codec exposes independent source/input and replacement
output ceilings; existing codec entry points retain their original profile.

## Validation

The isolated gates passed 1,860 XLSX, 619 OPC and 302 shared OOXML tests
(2,781 total). One pre-existing external ZIP64 corpus test remains ignored.
Strict Clippy, rustdoc, changed-file formatting, crate boundaries and diff checks
passed. Sources remained unchanged throughout the run. Whole-crate formatting
reports only the unchanged baseline `drawing_svg_read.rs` differences; the runner
proves the file matches the base commit and records that exception rather than
claiming every command passed. See [gate results](gates/results.json) and
[verification](gates/verification.json).

The public Custom Data suite has 28 cases, including source-preserving no-ops,
stale/namespace/signature refusals, exact resource boundaries, namespace prefix
and declaration budgets, query-table raw/MCE admission, shrinking replacements,
moved payload storage, relationship presence and inverse restoration. Two new
shared-codec cases separately test source/input versus output limits and empty
root expansion. Historical captures under `xlsx-custom-data-limits-final` are
not validation of this batch.

Independent review passed. The [matched performance report](performance/results/matched-report.md)
retains 450 accepted samples across five workloads and three authored sizes.
The added admission checks cost 13.6–25.1% in median runtime and 4.9–10.6% in
requested allocation bytes; peak live allocation changes by less than 0.4%.
These measurements include allocator-observer overhead and shared-host noise.
An earlier 225-sample diagnostic capture exposed an excessive relationship
reservation, which was fixed before the accepted capture and final gates.
[Root verification](root-verification.json) independently validates retained
hashes, sample invariants, source identity and reported median comparisons.

The selected-source checkout is based on `1d39eea516248c758b8ea429c879a4a3245e4fbb`.
It excludes the unrelated pending OPC synthetic relationship map and drawing-test
edits. The [gate runner](gates/run.py) records the exact source before/after,
lockfile, toolchain and command logs. [verify.py](verify.py) independently checks
retained evidence and selected workspace sources.

The [performance plan](performance/PLAN.md) compares the same authored inputs and
five complete lifecycle workloads against the committed baseline. These are
bounded conformance fixtures; they do not establish native Office acceptance.

## Scope

No Custom Data payload or connection is executed. Native producer
interoperability remains unverified. Limits describe the modeled OPC aggregate
and individual bounded transformations; they are not a physical ZIP-size limit
or an operation-wide allocator budget. Existing unsupported graph/MCE shapes and
changed signed packages remain refusals.
