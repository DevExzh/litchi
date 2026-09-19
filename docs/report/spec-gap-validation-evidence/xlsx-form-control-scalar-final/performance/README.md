# XLSX form-control scalar lifecycle performance evidence

This directory is the reproducibility bundle for the final XLSX
`formControlPr` scalar lifecycle. It is intentionally separate from the
historical [`xlsx-form-control-owner-performance`](../xlsx-form-control-owner-performance/)
owner-read smoke. The historical harness measures owner projection and does
not cover scalar edits, paired VML publication, inverse replay, or save/reopen.

The bundle is prepared but remains **pending final-source freeze**. The
authoritative capture must be run only after the root agent records the exact
source tree, compiler/toolchain, dependency lock, and fixture hashes. Until
then, no result under this directory is a performance claim.

`harness/run.sh` builds the small process harness in a temporary Cargo target,
uses a caller-supplied source root and fixture root, and removes the target,
staging project, and generated packages on exit. It emits one raw JSON sample
per lane and repetition. Each sample contains elapsed time, allocator event
counts and requested bytes, live-byte delta and peak live bytes, RSS before
and after, high-water RSS before and after, and counted `ReadAt` requests.
The output also records source and harness SHA-256 manifests, compiler
identity, and the release-binary hash. After collection, the runner derives
`stats.json` and `report.md` directly from the retained raw samples; the
percentile method and those report hashes are recorded in the receipt index.

The harness has no production dependencies and does not copy native fixtures
into this directory. It measures both the ordinary eager workbook facade and
the source-backed scalar editor where the API permits a matched workload:

* owner read after an already-open package;
* exact semantic no-op commit and publication;
* one changed scalar commit and publication;
* forward and inverse source patch application;
* changed save/publication followed by workbook reopen and typed readback.

The fixture selector derives a field/value pair from the first retained
control, with explicit field overrides for fixtures whose first authored
scalar is not the scalar proven by its VML mirror. The no-op uses the authored
scalar and the changed case toggles a bounded Boolean or checked value. The
selected fixture and derived case are printed in the correctness record and
rechecked on every sample. The workloads preserve the paired `ctrlProp`/VML
closure; the harness does not claim a one-part edit.

The final matrix excludes `button-form-control.xlsx` because its authored
`LockText` has no VML `LockText` occurrence and the bounded source editor
refuses insertion/removal of a missing mirror. That fixture remains useful for
owner-read checks, but cannot provide a paired scalar lifecycle measurement
under the current facade contract.

See [`performance-plan.md`](performance-plan.md) for the measurement contract
and the final capture checklist. Result files are deliberately absent until
the source freeze is supplied.
