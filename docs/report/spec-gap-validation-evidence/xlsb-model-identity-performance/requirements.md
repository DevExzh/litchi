# XLSB Data Model table-identity performance requirements

This file is the bounded design contract for the source-backed XLSB table-name
identity and dependency-closure rename profile. It is a measurement scaffold,
not a result. No timing, allocation, RSS, native-producer, or speedup claim is
admissible until the host rename API, the complete XLDM 140 fixture builder, and
the correctness gates are frozen.

The profile follows ADR 0003 for immutable snapshots, selector-based edits,
source-checked reversible patches, atomic failure, and exact no-ops. It follows
ADR 0005 for borrowed positional input, finite hierarchical limits, fallible
reservations, process-isolated measurements, and the distinction between
requested allocator bytes, peak live allocator bytes, and process RSS. It does
not add a native XLSB Data Model claim: the local native inventory found no
inspectable workbook containing a complete Data Model, so this profile uses
deterministic synthetic complete-XLDM-140 sources only.

## Scope and API binding

The neutral XLDM owner currently exposes a source-borrowed
`Xldm140Closure<'storage, 'source>` and the typed operations
`rename_table_name` and `rename_table_name_with_relationships`. The XLSB host
operation being implemented may expose a different selector or transaction
method. The adapter must bind to the final public host API after source freeze;
it must not call private codec functions or construct a private patch merely to
make a lane measurable.

The host profile must identify these phase boundaries explicitly:

| Phase | Included work | Excluded work |
| --- | --- | --- |
| `open` | package admission, workbook/model owner discovery, bounded source capture, and the closure/projection required by the public read API | fixture construction, source hashing, result comparison |
| `stage_noop` | selector resolution and staging of a name equal to the source name | package generation, save, reopen, report encoding |
| `stage_rename` | selector resolution, identity/closure preflight, and candidate staging exposed by the public edit API | save and post-run assertions |
| `commit` | public commit validation, candidate construction, closure readback, and source-bound patch creation | ZIP/save and independent reopen unless the API necessarily includes them |
| `save` | serialization of an already committed workbook to a caller-owned sink | edit staging, fixture generation, reopen, hash calculation |
| `reopen` | opening the saved bytes and reading the resulting identity projection | the preceding save and fixture construction |
| `inverse` | applying the public source-checked inverse to the reopened forward result and validating the restored state | forward fixture construction and report generation |

If a public method combines phases, the receipt records the combined scope and
marks unavailable sub-phases as unavailable. It must not infer a stage timing
by subtracting unlike end-to-end samples. Fixture construction, expected output
derivation, source hashes, semantic comparisons, and receipt serialization stay
outside the measured interval unless a lane explicitly names them.

Every profile receipt records `recipe.id`, `recipe.version`, the source fixture
path, the selected scale case, endpoint layout, name profile, and the complete
caller-limit object used for admission. It also records `source_bytes`,
`staged_bytes`, `candidate_bytes`, and `output_bytes` as separate values. A
value is `null` when the public API does not expose that boundary; the harness
must not infer candidate or staged bytes from an unrelated serialized package
size. The phase object always contains the seven named fields
`open_ns`, `stage_ns`, `commit_ns`, `save_ns`, `reopen_ns`, `inverse_ns`, and
`validation_ns`. Each timed field covers only its named operation, and
validation, semantic comparison, hashing, and receipt encoding run after the
timed operation unless explicitly called out in the lane.

The correctness receipt also records the separate `OlapProofLimits::DEFAULT`
object (`max_items`, `max_string_bytes`, `max_source_bytes`, and
`max_work=4_000_000`). These graph-proof quotas are distinct from the XLSB
`Limits::DEFAULT` caller limits and must not be folded into the latter.

The semantic receipt contains the source and observed table identities,
relationship IDs, relationship metadata paths, endpoint fields, and
time-grouping source/calculated IDs. A successful changed lane must report
separate equality gates for table IDs, table names, relationship IDs,
relationship endpoints, relationship metadata paths, and time-grouping IDs.
The relationship ID is the generated `RelId` identity from the admitted XLDM
metadata member; an outer relationship ordinal is not sufficient evidence.
The complete identity vectors are compared after save/reopen and inverse, not
only their counts.

The correctness verifier independently recomputes the expected renamed vector
from the source vector and selected name profile. Producer-supplied equality
booleans are checked against that recomputation rather than accepted as proof.
Each successful correctness result carries complete source and forward
member-hash manifests for OPC parts, relationship XML, content types, and
XLDM inner members. Only the admitted workbook/model and declared identity
metadata paths may differ; native, generated, opaque, and unrelated members
must remain byte-hash identical.
For the table-name lane, the mutable inner-path set is derived independently
from the selected `T1` table metadata, relationships whose containing or
primary endpoint is `T1`/`Table1`, and the structural `BackupLog`; all other
table and relationship metadata remain preservation-gated.
The verifier also regenerates the relationship pair recipe for the selected
or distributed endpoint layout, including generated `RelN` identities and the
closure's containing-dimension projection order. A source vector that changes
a dimension ID, table metadata path, relationship index key, relationship
endpoint/path, or time-group identity coherently in every reported vector is
still rejected unless it matches that independent recipe. The receipt's
mutable-path list must equal the derived closure exactly; it cannot be widened
by adding an unrelated member such as `Model.1.db.xml` or a `T2` table file.

## Deterministic complete-XLDM-140 corpus

The adapter must generate one deterministic source recipe and report its byte
length, FNV-1a-64 value, and SHA-256 in every receipt. The source recipe must
contain a complete, internally consistent version-140 closure:

- the section 2.2 directory/file-group and object identity records;
- section 2.5 table metadata with unique qualified TableIDs, XML table names,
  columns, and relationship metadata;
- section 2.6 Dimension/Attribute objects for every projected table and column;
- section 2.3 native data and section 2.4 generated data for every admitted
  column and relationship index; and
- the OLAP and generated position/identity mappings required by
  `prove_xldm140_closure` and the host's writable closure proof.

The generator must use the existing bounded XLDM codec/test fixture machinery
where possible. It must not manufacture a byte sequence that merely passes the
outer storage parser while omitting native, generated, or OLAP ownership. The
receipt records the recipe version and all scale parameters, rather than
retaining an unbound generated binary as if it were a native fixture.

The primary scaling matrix is a bounded Cartesian matrix:

| family | table count `T` | relationship count `R` | purpose |
| --- | ---: | ---: | --- |
| `tiny` | 1 | 0 | no relationship rewrite and exact no-op controls |
| `small` | 4 | 0, 3, 8 | table-only and sparse/dense closure controls |
| `medium` | 16 | 0, 15, 32, 64 | separate table and relationship growth |
| `large` | 64 | 0, 63, 128, 256 | bounded stress shape; no unbounded corpus claim |

The checked-in corpus catalog names these twelve cases and the two endpoint
layouts (`selected_table` and `distributed`). The correctness runner emits all
`12 × 2 × 5 = 120` points with complete source/candidate/readback/inverse
semantic gates and no timing or allocator sampling. Under the public host's
default OLAP graph-work ceiling, the ten `T=64,R=256` points are expected
typed `limit_exceeded` refusals after generation and package admission; they
must remain visible in the receipt rather than being omitted or counted as
successful renames. This correctness run does not constitute a performance
measurement or a native XLSB acceptance result. Performance sampling may begin
only after the source/semantic gates below have been reviewed on a clean
committed source pin.

The scaled fixture retains the optional `Year` grouping through the medium
family, where time-group identity checks are exercised, and uses a compact
Key-only table shape for `T=64` points. That fixture choice keeps the bounded
graph proof focused on the requested table/relationship growth; it does not
claim that large native models omit time groupings.

Every relationship has a validated containing table, primary table and column,
foreign column, generated relationship identity, and closed relationship-index
member. At least one matrix family must place all `R` endpoints on the selected
table so relationship replacement work is observable; a paired distributed
family must spread endpoints over other tables to separate endpoint scanning
from output growth. Table IDs, paths, dimension IDs, Attribute IDs, native
values, and generated identities remain stable across a rename.

Each scale is exercised with these name profiles:

1. `same_length_ascii`: a same-size replacement suitable for fixed-allocation
   controls;
2. `shorter` and `longer`: variable-size replacements that force the
   closure-aware writer path;
3. `escaped_xml`: a valid XML string containing characters requiring XML text
   escaping, with the exact escaped output length precharged; and
4. `unicode`: bounded non-ASCII UTF-8 names, proving byte-based rather than
   character-count limits.

The selected table is resolved by its stable TableID or final public semantic
selector. The harness must never pass a generated filename, physical member
offset, relationship ID, or storage path as a public identity selector.

## Required lane matrix

The sealed profile must contain at least the following lanes. The runner may
add scales, but may not remove a lane or silently turn a refusal into success.

### Read and baseline lanes

- `open_complete_{tiny,small,medium,large}`: open the synthetic XLSB package,
  discover the model, and obtain the complete typed identity/closure required
  by the public read path.
- `projection_only_{T,R}`: inspect the neutral XLDM closure with the host
  package layer excluded, when the public owner permits that separation. This
  is a decomposition control, not an end-to-end host result.
- `noop_commit_{T,R}` and `noop_save_{T,R}`: perform an exact semantic no-op,
  requiring the source-backed patch/state to remain borrowed or shared where
  the API exposes that property and the saved logical members to remain byte
  identical.

The baseline for a changed lane is the same source, limits, package shape,
open state, and validation scope with no rename intent. A baseline must not be
an empty or model-free workbook. Ratios are reported only for matched work;
the profile may report absolute observations for lanes whose APIs cannot be
matched.

### Rename and lifecycle lanes

- `stage_rename_{name_profile}_{T,R}`: stage one table XML-name rename with
  relationship closure enabled. Record whether the selected table has zero,
  sparse, or dense affected endpoints.
- `commit_rename_{name_profile}_{T,R}`: commit the staged operation and require
  a source-checked candidate patch.
- `save_reopen_rename_{name_profile}_{T,R}`: save the committed result to a
  sequential caller sink, reopen it, and read the identity projection.
- `inverse_reopen_{name_profile}_{T,R}`: apply the inverse to the reopened
  forward result, save/reopen when the host API permits it, and prove exact
  restoration of all retained package members and source metadata.
- `repeat_determinism_{name_profile}_{T,R}`: run the same source and edit twice
  in fresh processes and compare candidate bytes, member ordering, IDs,
  relationship XML, and output hashes.

The final candidate must change only the admitted table XML name and proven
relationship endpoint/name values. The following must remain unchanged and be
checked outside the timer: TableID, generated paths, Dimension/Attribute IDs,
native and generated values, relationship IDs/index identities, unrelated
table metadata, workbook records, connections, content types, relationships,
unknown XLDM members, and opaque model payload bytes.

### Refusal and preservation lanes

- `reject_unknown_member_{T,R}`: add a bounded unknown XLDM member to an
  otherwise complete source. A structural rename must refuse before candidate
  output allocation when the closure cannot prove preservation.
- `reject_incomplete_closure_{T,R}`: remove or corrupt one required
  native/generated/OLAP owner or mapping. The operation must return a typed
  refusal and leave the source untouched.
  A graph-work limit refusal records
  `source_proof_status=unavailable_graph_work_limit`; it must not emit a
  partial semantic vector or claim that the full closure was independently
  established.
- `reject_ambiguous_identity_{T,R}`: duplicate or ambiguously bind a table
  XML name, relationship endpoint, or closure owner. No order-based choice is
  allowed.
- `retain_opaque_payload_{T,R}`: retain an unrelated opaque XLDM member and a
  large inert XLSB model-part payload. A no-op and any accepted rename must
  preserve exact bytes; an unsupported structural shape must refuse rather than
  drop the payload.
- `reject_stale_source_{T,R}`: change one lexical source byte, relationship
  member, content-type token, or opaque payload byte after staging. Apply must
  fail closed without mutating the target.
- `reject_signed_change`: changed publication on a signed package must retain
  the existing explicit signature-policy refusal; signed exact no-ops remain
  no-ops.

Expected refusals are successful profile outcomes only when the exact public
error class is recorded, the source bytes and source allocation remain usable,
and no partial package or XLDM candidate is published. Unexpected acceptance,
an unrelated error, or an altered source is a hard verifier failure.

## Correctness and preservation gates

Every successful changed sample has untimed gates for all of the following:

- complete source and candidate closure proof succeeds;
- old and new semantic names are exactly the requested values, with XML
  entities decoded for comparison and original lexical bytes preserved where
  they are outside the admitted rewrite;
- the selected TableID, all column identities, dimensions, attributes,
  relationship IDs, generated paths, native/generated values, and OLAP object
  identities are unchanged;
- every proven relationship endpoint that names the selected table changes,
  while unrelated endpoints and unknown members remain byte-identical;
- `Xldm140Patch::before` is the exact source slice, a no-op uses the borrowed
  source variant, and a changed patch owns only the candidate bytes required
  by the public contract;
- forward output is deterministic across repetitions and is valid after
  package save/reopen;
- inverse application restores the exact original decompressed XLSB member
  bytes, including workbook records, model payload, content types,
  relationships, unknown XLDM bytes, and lexical metadata; and
- stale, limit, malformed, signed, and unsupported inputs are atomic and leave
  the original package reopenable.

The verifier must compare complete member manifests, not only the projected
table names. A semantic digest is supplementary evidence and cannot replace
exact source/member comparisons.

## Limits and allocation contract

The adapter binds every available caller limit by name in the receipt. At
minimum this includes the current XLDM/XLSB fields for raw records, tables,
relationships, rewrite bytes, retained part bytes, graph parts, graph
relationships, metadata bytes, connection bytes, and connection count. Any
new identity/closure limit introduced by the host coder is added to the same
manifest rather than hidden in a fixture constant.

For each positive limit that the lane exercises, the fixture must include:

- an exact-boundary case that succeeds when the source and resulting candidate
  fit exactly;
- a one-byte-under case that returns the typed limit error before the relevant
  candidate buffer or retained value is constructed; and
- an over-boundary case when a distinct source or output resource is relevant.

The verifier records observed and maximum values and checks checked arithmetic
for source bytes, candidate bytes, metadata, closure members, relationship
occurrences, rewrite work, and output bytes. A final-output check after a large
temporary allocation is insufficient evidence for a preallocation limit.
Fixture generation and the expected-size calculation are outside the measured
library allocation boundary, but the library must perform its own exact
preflight.

Each raw sample records, where observable:

```text
elapsed_ns,
direct_allocated_bytes, realloc_new_bytes, realloc_old_bytes,
deallocated_bytes, live_before, live_after, peak_live_delta,
alloc_calls, dealloc_calls, alloc_failed, alloc_underflow,
source_bytes, candidate_bytes, bytes_copied, staged_bytes,
tables, relationships, affected_relationships, closure_members,
source_reads, source_read_bytes, output_bytes, output_write_calls,
phase_ns: {open, stage, commit, save, reopen, inverse},
hardware_counters: {cycles, instructions, branches, branch_misses,
                    l1d_misses, llc_misses, page_faults}
```

`source_reads` and read bytes are unavailable rather than zero for an
in-memory XLSB API. Requested allocation bytes count direct allocations plus
successful reallocations' new sizes; old realloc sizes remain separate. Peak
live allocation is the incremental high-water above the timed interval's
live-before baseline. `/usr/bin/time -v` RSS is process-wide and is never
described as per-operation memory. The allocator equation and failed/underflow
flags must pass for every retained sample.

The no-op and expected-refusal lanes must additionally record whether the
source/opaque storage allocation is shared or borrowed when the public API
exposes an observable pointer/variant. A source hash alone does not prove
allocation sharing. Changed lanes must distinguish retained source bytes from
new candidate bytes; a full source clone hidden inside setup is not counted as
zero-copy.

## Runner and provenance

The future runner must use at least three fresh processes, at least two warm-up
iterations, and twenty retained samples per lane for sealed evidence. It must
record CPU model, logical/physical core counts, memory, OS/kernel, Rust and
Cargo versions, target, compiler flags, allocator, Cargo.lock hash, fixture
recipe, source hashes, binary hash, and relevant environment variables before
and after the run. Nonempty `RUSTFLAGS`, bootstrap overrides, source drift,
missing receipts, allocator failures, or failed correctness gates invalidate the
run.

The runner owns only its private temporary target and generated fixture area.
It must not delete the repository target, native corpus, another agent's
worktree, or historical evidence. It must retain raw JSON, commands, source
manifest, build provenance, and verifier output; a future measurement batch
must clean only its own temporary artifacts.

Before a sealed run, every local Cargo source, manifest, build script, fixture,
and harness input must be tracked at the pinned Git commit and byte equal to
that commit's `git cat-file` blob. This check must reject unstaged, staged,
untracked, and `assume-unchanged` edits. The source manifest records the exact
Cargo metadata digest and commit; the build receipt records the profile binary
digest before and after the run, and verification must bind both digests back
to those receipts. A provisional smoke result from a dirty or uncommitted
checkout cannot be promoted to sealed evidence.

The final report must provide p50/p95/p99 and an uncertainty summary for each
matched lane, plus source/output bytes and allocation/peak-live distributions.
Hardware counters, perf profiles, RSS, and syscall observations are supporting
evidence and must retain their whole-process or stage scope. No result may
claim native Excel compatibility, native Data Model producer coverage, a
general XLDM speedup, O(1) behavior, linear scaling, or an end-to-end memory
bound without a separately matched experiment proving that claim.

## Implementation hotspots to inspect after the source freeze

The first profiling pass should attribute work across the existing seams rather
than assume that the rename itself dominates:

- `litchi_xldm::identity`: identity projection, closure proof, relationship
  endpoint lookup, exact output-length preflight, owned replacement payloads,
  candidate closure reparse, and OLAP proof;
- `litchi_xldm::codec`: source directory scan, variable-size payload rewrite,
  checked reservations, and outer storage reconstruction;
- `litchi_xlsb::data_model::{Snapshot, Transaction, Patch}`: source capture,
  typed staged metadata, opaque model-part sharing, package clone/materialize,
  readback, stale guards, and inverse publication;
- `litchi_xlsb::data_model::package`: graph/relationship/content-type
  preflight and retained source-token capture; and
- XLSB package save/reopen and workbook-record patching, including any
  connection-owner validation performed for every stage.

The profile should report counts and bytes for these boundaries where the
public instrumentation can observe them. It must not label a linear or
quadratic hotspot from source inspection alone; such a claim requires the
table/relationship scaling matrix and a retained profile or counter receipt.

This scaffold deliberately leaves measurements, native evidence, and any
optimization decision open until the final host API and source/test freeze.
