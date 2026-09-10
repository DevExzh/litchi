# 0495 managed versus unmanaged DOCX edit harness plan

Status: draft design for the new, unsealed 0495 measurement.  This document
does not authorize a build or a capture.  It is intentionally separate from
the retained 0494 reports and helpers.

## Decision

Add a new benchmark entry point/module for 0495 after the DOCX source-backed
edit implementation lands.  Keep the 0494 entry point and its reports as the
unmanaged historical regression control.  The new entry point measures the
same one-paragraph edit through two constructors of the same public API:

* `unmanaged-api`: a source-backed `Package` without an
  `ExecutionContext`, followed by `edit_document()`,
  `replace_paragraph_text(Position::new(0), replacement)`, `commit()`, and
  `publish_document_commit_to_stream`.
* `managed-api`: the same lifecycle and arguments, with a finite explicit
  `ExecutionContext` and the managed constructor.

The exact managed constructor and any ownership type used for the edit
snapshot must be read from the landed implementation before the harness is
compiled.  The calls above describe the required public lifecycle, not a
permission to add a second private transaction path.  A typed refusal from the
current managed boundary is a capability result (`unsupported` or a negative
probe), never a successful managed benchmark row.

The report and protocol names should be new (`docx_edit_provider_managed_v1`
and `docx-edit-provider-managed-protocol-v1`).  This prevents a validator from
mistaking a managed row for a 0494 `docx_edit_provider_v1` row.  The protocol
binds the exact 0494 corpus and records the retained 0494 binaries as a
historical unmanaged control.  It does not make a before/after optimization
claim across binaries, source revisions, or the two API modes.

## Why this boundary is needed

The current source-backed package has a typed refusal in
`main_document_snapshot` for managed main-document edits.  The existing test
`managed_main_document_edits_refuse_with_typed_boundary_error` proves that
boundary; it is not a performance baseline.  The coder's production change
must replace or extend that test with a successful managed ownership contract
before 0495 formal runs are admitted.

The source-backed transaction owns `Snapshot` values and the package already
charges managed publication through the output budget.  The missing evidence
is whether the managed edit snapshot, commit, and any cache/catalog owner are
charged and released when their last owner is dropped.  The harness therefore
records per-sample budget gauges before, live, and after the package/document/
commit drops.  It must not infer release from cumulative input, output, or work
counters.

The 0494 retained normal and allocator binaries remain useful for an
unmanaged regression check.  They were built from an earlier source state and
cannot be used to claim that managed code is faster or that a current change
improved the old path.  A 0495 report must show build, source-manifest, corpus,
protocol, and API-mode identities before any comparison is made.

## Corpus and provider matrix

Every 0495 row uses the exact 0494 corpus identity:

* version `0188-media-v1`, generator `litchi-docx-source-edit-media-v1`;
* 16,793,036 archive bytes, SHA-256
  `a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4`;
* 20 archive members, 200 paragraphs, and eight 2 MiB PNG-signature media
  members;
* expected original text 10,000 bytes, SHA-256
  `ad4fe690f0ef2281ad8e64a78d1f4d64e7c8625d672b3ac2f7e9fbc28a82f4af`.

Use the six 0494 provider arms so that transport behavior remains comparable:

| arm | construction and fixed controls |
| --- | --- |
| `owned` | `OwnedSource`, no synthetic range delay |
| `instrumented` | owned source plus the logical/physical counting wrapper |
| `file-warm` | the same archive staged once and reopened through `FileSource`; warm-cache observation |
| `short` | maximum range 4,096 bytes |
| `delayed` | maximum range 65,536 bytes, 1,000 microsecond delay, 100 MiB/s minimum-service pacing |
| `range-zero` | the delayed transport with zero fixed delay, otherwise the delayed controls |

Source construction, file staging, wrapper creation, trace-capacity
reservation, finite limits, and context construction happen once per sample
outside the operation clock.  A fresh source adapter is still required for
each measured iteration.  No row may silently use ambient filesystem or
network state.  The file arm is explicitly warm; it is not a cold-I/O claim.

The formal matrix is provider × API mode × binary role.  Run both normal and
allocator binaries with the same source and protocol configuration.  The
allocator binary surrounds the same operation region with the global-system-
allocator observer; it is an allocation observation, not a second latency
claim.

## Minimum Rust harness patch

Do not edit the sealed 0494 profiler, report, or recovery helpers.  The
smallest maintainable patch is a new
`tools/perf-baseline/src/docx_edit_provider_0495.rs` (the final module name may
follow the coder's routing convention) plus a new binary selector.  It may
reuse private-free corpus and provider helpers only if doing so does not change
the 0494 source or its hashes.  Copying the small setup scaffolding is safer
than adding a version branch to the old report schema.

The new module needs only these changes:

1. Parse one `edit_api` mode and the existing provider/range/sample arguments.
   Reject an unknown mode, missing managed limits, or mixed mode values in one
   process.  A process measures one provider and one API mode so its report
   identity is unambiguous.
2. Construct the same corpus and provider matrix as 0494, then add a named
   finite-context factory.  Reuse the managed read-ahead policy's finite
   values initially: 64 MiB memory/input/output, 1,000,000 objects, depth
   1,024, 2 GiB work, one worker/task, 64 MiB in-flight bytes, and the existing
   bounded cache limits.  If the landed edit implementation needs a different
   finite value, record that change in the protocol and use it for both managed
   controls; never substitute an unlimited sentinel.
3. Split only the operation call in `run_sample`:

   ```text
   provider/context setup (outside)
   start operation clock and allocator region
     Package open
     edit_document
     replace_paragraph_text(Position::new(0), replacement)
     commit
     publish_document_commit_to_stream(&mut sink, &commit)
     commit/package/document/managed payload drops
     timed commit diagnostics and source/target identity checks
   stop clock
   budget-after-drop, source/range/physical snapshots, output and oracles (outside)
   ```

   The unmanaged and managed functions must share the same sink, replacement,
   output ownership, and validation scaffolding.  Do not implement the
   managed row by calling the old `publish_docx_source_edit` helper.  That
   helper is retained as the 0494 historical control only.
4. Add report records for `edit_api`, constructor/context identity, all budget
   snapshots, cache diagnostics, and release checks.  Keep the existing 0494
   physical/range/zero-length counters and source-version fence.
5. Add the allocator twin and the new selector/usage text.  Add focused parser
   and schema tests before any pilot capture.  No production API workaround or
   fallback to the old refusal belongs in the harness.

The driver should fail closed if a managed call returns the old typed refusal,
if a managed row has no `budget_managed=true`, or if an ownership gauge does
not return to its per-sample baseline.  A separate `managed-refusal-probe`
may retain the typed error and payload-read count for capability documentation,
but it must not be mixed into formal timing statistics.

## Timed and untimed work

The operation clock has one definition for both modes:

> Fresh source-backed DOCX package open, one paragraph edit staging and commit,
> sequential publication, commit/package/document drops, and timed commit
> diagnostics plus source/candidate XML identity checks.

The following are setup and must be outside the clock: deterministic corpus
and expected-output construction, expected replacement text, source adapter
and file handle construction, trace-vector reservation, sink reservation,
finite read/cache limits, `Budget`/cancellation/`ExecutionContext`, and all
preflight patch-oracle preparation.  Build and file staging are outside the
sample entirely.  The managed context must be constructed before starting the
clock, while the package's production cache/edit ownership is constructed by
the measured package open.

The following are outside the clock after the output has been retained:

* exact output byte and SHA-256 comparison;
* semantic reopen, paragraph replacement verification, and all eight media
  payload comparisons;
* source-version and physical/range counter snapshots;
* forward replay, inverse restoration, stale-target rejection, and foreign-
  source rejection;
* final budget/cache release checks and report serialization.

The output sink remains owned after the clock.  Commit diagnostics and source
and candidate XML identity comparisons stay inside the clock, matching the
0494 scope.  The managed `Budget` and `ExecutionContext` must outlive the
package and commit until the after-drop snapshot is taken, then be dropped
outside the clock.  This is necessary to observe persistent owners without
charging context teardown as document work.

## Managed budget contract

Record `SourceCacheDiagnostics` and the core budget at three points for every
successful managed row:

* `before`: after setup and before package open;
* `live`: after publication and before package/document/commit drops;
* `after_drop`: after all package, document, edit, commit, and managed payload
  owners have been explicitly dropped.

At minimum each snapshot contains:

* `budget_managed`, reservation-failure count, memory/input/output/work/object
  and depth used/limit values;
* cache hits, cold loads, successful/failed loads, retained bytes and
  in-flight bytes;
* `budget_cache_reserved_bytes`, `budget_catalog_reserved_objects`, and
  `budget_cache_reserved_objects`;
* source physical returned bytes and the managed input/output deltas.

The formal success invariants are:

1. `budget_managed` is true and no reservation failure occurs.
2. `memory_after_drop == memory_before` and
   `objects_after_drop == objects_before`, using the per-sample baseline
   rather than assuming either baseline is zero.
3. Cache/catalog reservation gauges that belong to the package return to their
   `before` values.  This is the persistent-ownership release evidence.
4. Input, output, and work are reported as cumulative deltas.  They are not
   required to return to baseline.  Where the implementation charges at the
   source/sink boundary, validate input against the authenticated physical
   returned-byte count and output against the sink's accepted-byte count; if a
   lower layer deliberately charges a different boundary, retain both values
   and state the relation in the row instead of fabricating equality.
5. Any nonzero live or after-drop retained owner is a row failure with an
   explicit resource and value.  Do not hide it by dropping the `Budget` before
   reading diagnostics.

The unmanaged row records `managed=false` and its budget fields as unavailable
with the reason `no ExecutionContext`; it does not pretend that unmanaged
memory is budgeted.  Cumulative values and gauges from the old 0494 helper are
not retrofitted into this schema.

## Correctness and failure semantics

Both API modes must prove the same output contract before a row is eligible for
timing statistics:

* source version is unchanged during the operation;
* one materialized main-document source and one changed commit with one
  operation are observed;
* commit source and target XML identities match the preflight source and
  candidate;
* the output is byte-exact where the preservation contract requires it,
  semantically reopens, contains the requested paragraph text, and preserves
  all media members;
* the patch replays forward, its inverse restores the original, stale targets
  are rejected, and a foreign source is rejected.

Patch and rejection checks run outside the operation clock and are correctness
oracles, not successful baseline operations.  A managed boundary refusal,
cancellation, budget exhaustion, malformed source, source conflict, sink
failure, or release invariant failure receives an explicit status such as
`unsupported`, `cancelled`, `limit`, `source_conflict`, `sink_failed`, or
`ownership_leak`; it is excluded from latency summaries and retained with the
error type.  In particular, the old managed refusal cannot be labeled a
successful managed baseline.

## Protocol and report shape

The top-level protocol should contain:

```json
{
  "schema": "docx-edit-provider-managed-protocol-v1",
  "version": 1,
  "change": 495,
  "case": "docx_opened_document_managed_vs_unmanaged_one_paragraph_edit_save",
  "claim_authorized": false,
  "comparison_policy": "descriptive same-build API-path evidence; no optimization claim",
  "corpus": { "version": "0188-media-v1", "archive_sha256": "..." },
  "edit_apis": ["unmanaged-api", "managed-api"],
  "providers": ["owned", "instrumented", "file-warm", "short", "delayed", "range-zero"],
  "budget_profile": { "managed": true, "finite": true },
  "historical_0494_control": { "managed": false, "comparison_only": true }
}
```

The real receipt must expand the abbreviated corpus and historical-control
objects with exact hashes, source manifest, build receipts, driver/helper
hashes, CPU lock, environment, command argv, and expected process/sample
counts.  Per-report identity must include `edit_api`, provider configuration,
constructor name, source revision, binary SHA-256, and protocol SHA-256.  A
validator must reject a row whose protocol or source binding differs from the
top-level receipt.

For the initial formal matrix use three warmups and 30 measured samples per
provider/API/role/repeat, with repeats 1 and 2.  That is 48 gated child
processes (`2 roles × 6 providers × 2 API modes × 2 repeats`).  A pilot should
use one warmup and three samples per cell and must pass all correctness and
release checks before formal expansion.  If the driver chooses a different
repeat schedule, the receipt must state the resulting expected process count
before the gate starts.

The timing, setup, allocation, provider, physical, range, zero-length, and
media scope strings should retain the 0494 meanings, with the managed context
and budget ownership additions stated above.  Do not call synthetic range
counts filesystem or network I/O, and do not call allocator observations
throughput results.

## Gate and build sequence

After the coder lands managed support and the relevant source tree is frozen:

1. Run focused production tests for managed edit, publication, cancellation,
   source conflict, and ownership release. Update the old refusal test to test
   the successful contract; retain a separate typed-refusal test only for an
   unsupported document shape.
2. Add the 0495 harness and parser/schema tests. Run the new binary's `--help`
   and a no-workload argument-validation smoke test.
3. Capture normal and allocator build receipts with the same source manifest,
   locked dependencies, frame pointers/unwind tables, CPU lock, and temporary
   directory policy used by the gated measurement. The build must authenticate
   the current source, not the 0494 retained binary.
4. Run a pilot through the 0495 gate for every provider and both API modes.
   Retain stdout/stderr, report, raw counters, output hashes, and the process
   group terminal/hash. Stop if any managed row is unsupported or leaks a
   budget owner; do not silently fall back to unmanaged.
5. Run the formal normal and allocator matrix only after the pilot is fully
   validated. Validate reports through the canonical measurement validator,
   then publish the protocol, build, source, and report custody receipts.

Canonical command shape once the new selector exists (the gate supplies the
CPU lock and authenticated environment; the exact binary path is a build
receipt field):

```text
python3 -B docs/performance/results/change-0495/gate.py pilot-r1 \
  -- <retained-0495-binary> docx-edit-provider-managed \
  --provider owned --edit-api managed --samples 3 --warmup 1 \
  --source-revision <frozen-40-hex-revision> \
  --output <pilot-dir>/owned-managed.json

python3 -B docs/performance/results/change-0495/gate.py formal-r1 \
  -- <retained-0495-binary> docx-edit-provider-managed \
  --provider <provider> --edit-api <unmanaged-api|managed-api> \
  --samples 30 --warmup 3 --source-revision <frozen-40-hex-revision> \
  --output <formal-dir>/<role>-<provider>-<api>-r<repeat>.json
```

The gate wrapper must run each canonical child under the CPU lock and preserve
the whole-child terminal.  The command shape is intentionally not executable
until the coder's routing and argument names are frozen; this plan does not
run Cargo, a benchmark, or a capture.

## Acceptance checklist for coding

The implementation is ready for a pilot when all of these are true:

* the managed production test succeeds through `edit_document`, one
  `replace_paragraph_text`, `commit`, and publication;
* managed and unmanaged rows use the same public lifecycle and corpus;
* setup and timed scopes are represented in the report and no context/cache
  construction is hidden in the operation clock;
* every managed row has before/live/after-drop budget and cache diagnostics;
* persistent memory/object/cache reservations return to their sample baseline;
* cumulative input/output/work values are labeled as cumulative and are not
  used as release evidence;
* exact, semantic, media, source-version, and patch oracles pass;
* the old refusal is either absent from the formal route or is recorded only
  as a non-success capability result;
* report, binary, source, protocol, gate, and driver hashes are authenticated;
* the 0494 retained reports remain byte-for-byte unchanged.

This plan recommends implementing one new 0495 module and one new protocol,
then measuring only after the managed ownership contract is observable.  It
keeps the useful 0494 unmanaged regression while making the managed resource
accounting and persistent-release question independently reviewable.
