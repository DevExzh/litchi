# 0362ha — corrected unified File qualification harness

- Status: protocol and source candidate; the coordinator review below still
  requires harness repair before native qualification.

The historical `file_w8_probe` and its receipts remain frozen. This harness uses
`file_w8_qualified_probe`, which pairs the facade with the corresponding leaf
package for every format family (DOCX, ODT, DOC, RTF, XLSX, XLSB, XLS, ODS,
PPTX, PPT and ODP).

The probe exposes a route list from the selected source before constructing
fixtures. The historical open/save rows remain present, and the corrected
probe adds the actual edit families from the original W8 workload matrix:
open/get, save, repeated save, set/edit/commit, batch commit, import commit,
merge commit, refresh commit, and their separate
publication routes where the family admits them. A route is listed only when
both the native leaf and the facade have the same semantic operation; an
unsupported operation fails route discovery instead of silently becoming an
open/save substitute.

The read and publication rows are `open_get`, `save`, and `repeated_save`.
Mutation rows are `set_commit`, `set_publication`, `batch_commit`,
`batch_publication`, `merge_commit`, and `merge_publication`; workbook families
also expose `import_commit` and `import_publication`, and DOCX exposes the
native field `refresh_commit` and `refresh_publication` pair. Commit rows stage
the actual leaf edit (or the corresponding import/merge/refresh edit) and time
the native commit. The leaf import routes decode the same CSV and apply its
cells through the native worksheet edit, because leaf packages do not have a
format-neutral import verb. Publication rows reuse the committed owner and
time only the output operation. Repeated save performs two publications per
logical call so its exact bytes are qualified independently of first
publication.

The probe calibrates a common logical-call count for each route across both
endpoints and both protected revisions. Each measured process performs one
real leaf, facade, facade, leaf inner block; the runner repeats independent
before, after, after, before outer blocks at least five times. It emits raw
elapsed nanoseconds, logical calls, operation counts and totals; analysis
normalizes totals by operation counts rather than comparing independently
calibrated wall times. Target calibration is 500 ms with headroom; a sample
below 150 ms or a calibration that cannot reach the target within its explicit
ceiling fails closed. Mutation/commit and publication are separate routes. A
validation call runs outside the timed span and re-publishes the untouched
owners after staging, checks semantic edited state, exact no-op and
repeated-output bytes, changed output identity, and the observed output
digests.

The example also accepts `--mode validate --route FAMILY/ROUTE`. This executes
the same leaf/facade staging and validation path without `Instant` sampling or
calibration, so a compile/correctness check can run before any native
qualification window.

The guarded runner performs fresh ABBA legs for every route, records binary,
source commit, lock and fixture identities from a separate endpoint manifest
for each revision, verifies every manifest artifact hash, rejects compiler,
profiler, test, peculiar-build and overlapping qualification processes using
full `/proc` command lines and executable identities before, after, and every
100 ms while a child runs, preserves every raw receipt and host snapshot, and
propagates any child failure. It refuses to overwrite an existing output
folder. Run it only with protected release binaries and endpoint manifests:

```text
bash docs/unified-ops/results/0362h/abba.sh BEFORE AFTER OUTPUT [ROUNDS] [CPU] \
  --before-manifest BEFORE-ENDPOINT.json --after-manifest AFTER-ENDPOINT.json
python3 -B docs/unified-ops/results/0362h/analyze.py OUTPUT
```

Each endpoint manifest has schema `litchi_file_endpoint.v1` and contains a
`source_commit`, a `dependency_projection_sha256`, and `binary`, `source`,
`lock`, `probe`, and `fixture` objects. Each object has a real `path` and
`sha256`; the runner hashes the files again. The lock, dependency projection,
probe, and fixture identities must match between endpoints, while source
commits and binaries are allowed to differ. This prevents a shared environment
variable or reused target directory from making two endpoints appear paired.

The runner does not build binaries or claim acceptance. The coordinator records
and reviews the resulting paired, per-call ratios, raw totals, semantic output
receipts, and any allocation evidence after a quiet native window. The
superseded `a271e1413f` publication-only receipt remains historical and is not
used for qualification.

The intended W8 source is the Part A API from `b1ae817dc4`: `Commit` and
`BatchCommit` keep their result fields private and expose the triad through
`file()`, `patch()`, and `diagnostics()` accessors. Their consuming extraction
method is `into_file()`, which moves the already-created immutable result
without adding a clone to the measured setup. The probe uses that method for
the unified mutation routes. The earlier `cd64450c7f` rootcheck used `.file`
against the coordinator's later BIFF/Graph-integrated `ddf2ac8ae2` source,
whose fields were public. That receipt is retained only as a superseded
API-shape compile receipt; no W8 qualification is claimed from it.

```text
CARGO_BUILD_JOBS=1 cargo check --offline --locked -p litchi-ops \
  --example file_w8_qualified_probe \
  --features docx,xlsx,xlsb,pptx,doc,xls,ppt,rtf,odt,ods,odp
```

The command above is the command for the intended W8 source and must be run
after the probe and the Part A facade are built together. A prior untimed
rootcheck ran the equivalent route preflight against the coordinator's later
`ddf2ac8ae2` source, where result fields were public; it passed for
`docx/set_publication`, `xlsx/import_publication`, `pptx/merge_publication`,
`ods/batch_publication`, and `odp/open_get`. Those receipts are retained
outside the checkout at `.coord/w8-file-correctness/` as superseded API-shape
evidence only. They do not validate this `b1ae817dc4` candidate, and no W8
performance qualification is claimed from them.

## Coordinator integration candidate

The coordinator assembled this candidate on `b34cb1b104` from Part A
`b1ae817dc4` plus its DOC refusal-kind dependency `216c1ad422`, the detection API and its consumers from `84a64e8285`, and the
corrected qualification probe/protocol from `3ce77d193a`. The detection patch
excludes `transaction.rs`'s independent content-replay work. Earlier worker
PPT/CFB/OPC/ODF/PPTX/XLS/RTF Part B implementation changes are not included.
The already accepted PPT repair, character-style policies and Pages registry
error fidelity remain in the candidate base. The historical benchmark table
is retained with no acceptance budget or passing verdict.

This integration is a source candidate. Compilation and untimed route checks
must run against its private-field Commit API before any native qualification.
Final matched binaries will also require the W4/W4b integration. No native
performance acceptance is claimed by assembling or committing this candidate.

The first assembled check exposed the separately committed DOC refusal-kind
dependency; the coordinator added it to the same source candidate as its File
consumers. Review also repaired two detector-policy defects: every plain and
compressed RTF entry point now applies its source-byte ceiling, and the
OpenOptions bridge no longer applies a DOC-only ceiling to PPT detection.
Boundary tests cover equal/one-under RTF input limits across byte, reader,
path, owned and prepared detection, restored reader cursors, and unrelated
CFB owner policies. Parser/decompression budgets remain with the owner open.

The assembled source passes all-target/all-feature Clippy with warnings denied
for ops, facade, CLI, Python and DOC; the no-format ops build and the RTF-only
input-boundary regression also pass. Matrix regeneration and its eight unit
tests pass. [The candidate receipt](results/0362ha/candidate-validation.json)
records changed source hashes, successful checks and all earlier failed
attempts. The full current-source ops library test was still running when this
candidate was committed; this receipt does not claim its result or the final
workspace gate.

## Remaining harness review

The coordinator returned three implementation defects in the assembled probe.
Some native read legs inspect a whole document (or only count sheet names)
while the facade reads one logical target. The DOCX refresh fixture contains
no fields and therefore exercises only an empty refresh. Source preservation
compares one semantic marker rather than the owner's exact published bytes
before and after staging. These must be corrected and validated on the actual
Part A API before this protocol can establish performance acceptance. The
historical and superseded receipts remain attached to their source versions.
