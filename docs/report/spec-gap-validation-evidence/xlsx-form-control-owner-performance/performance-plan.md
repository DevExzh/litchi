# XLSX form-control owner performance smoke plan

Status: pre-correction candidate evidence, pending owner review. The frozen
snapshot used here predates a later owner-reported canonical-SML
entity-escaped-URI correction; this artifact must not be treated as final
approval until the corrected source gate is supplied and replayed.
Owner review also currently holds source-v2 on duplicate VML attributes,
diagnostic projection-byte undercharge, and name-limit `Resource` error
labeling; this smoke does not certify those corrections.

This run is a bounded allocation/correctness smoke for the new public
worksheet form-control owner APIs. It is not a runtime baseline, optimization
claim, or before/after comparison. The source under test is the frozen root
gate snapshot recorded in `.codex-tmp/form-control-owner-root-gate.json`:

- base commit: `047d14953b4b2e674abb7c0783367225258aa287`;
- source checkout: `/var/tmp/litchi-form-owner-root-s450o_y9/worktree`;
- isolated Cargo target: a separate temporary target, never the root-gate
  target;
- changed public-owner files: the five paths and SHA-256 values in the gate.

The replay path is reproducible from the retained artifacts: run
[`restore-source.sh`](restore-source.sh) against a repository containing the
base commit, then pass its bounded output tree to
[`harness/run.sh`](harness/run.sh). The restore applies
[`source-delta.patch`](source-delta.patch), installs
[`source-Cargo.lock`](source-Cargo.lock), and checks the resulting files
against the gate, the full [`source-manifest.sha256`](source-manifest.sha256),
and [`source-bundle.tar`](source-bundle.tar). The harness runs from a
temporary workspace carrying the pinned Cargo/toolchain/lint configuration,
and verifies the full source manifest, build configuration, harness manifest,
locks, compiler identity, and relevant environment before and after build and
collection. By default it writes a new output directory under `runs/` and
leaves earlier receipts untouched.

The harness uses only the retained XLSX form-control fixture corpus already in
the source checkout. It does not copy native fixtures into the evidence
directory. The harness is relocatable: provide
`LITCHI_FORM_OWNER_SOURCE_ROOT`, `LITCHI_FORM_OWNER_FIXTURE_ROOT`, and
optionally `LITCHI_FORM_OWNER_OUTPUT_DIR` when replaying it. The output
directory must be empty when explicitly supplied.

## Contract

The harness records seven raw repetitions for each operation and fixture. Each
receipt includes elapsed nanoseconds, process RSS sampled before/after the
operation where Linux procfs is available, and raw global-allocator counters:
allocation/deallocation/reallocation calls and requested sizes, cumulative
requested event bytes, phase-end live-byte delta from the pre-phase global
live-byte baseline, and the maximum live-byte increase above that baseline
during the phase. Allocator counters are process-global observations, not
physical heap usage and not a language-level allocation proof. Live bytes are
maintained across phase boundaries by successful alloc/dealloc/realloc deltas;
RSS remains a separate process snapshot.

The measured public operations are split at the package-open/owner-projection
boundary so package work is not hidden inside a projection number:

1. `eager_package_open` and `source_package_open`: the public
   `Workbook::from_bytes` and `SourceBackedWorkbook::from_read_at` open calls.
   The source lane uses a caller-owned counted `ReadAt` adapter and reports
   logical read calls, requested bytes, returned bytes, and maximum request.
2. `eager_owner_projection` and `source_owner_projection`: the public
   worksheet `form_controls()` call after the corresponding package open.
3. `eager_single_query` and `source_single_query`: one position selector on a
   retained collection.
4. `eager_many_query` and `source_many_query`: all position selectors and
   public iteration on a retained collection.
5. `eager_cheap_clone` and `source_cheap_clone`: collection/view clones with
   no source mutation.

The `button-form-control.xlsx` fixture is the one-control case. The retained
two-control fixtures provide bounded many-control cases. A temporary generated
copy of the button fixture adds one deterministic, incompressible 1 MiB
unrelated `xl/opaque/unrelated.bin` ZIP member. It is hashed and used during
the run, then removed; no large native member is retained or copied into this
directory. This lane makes package-open and selected-owner source reads
visible separately without comparing different semantics as a speedup claim.

Correctness probes run before timing and fail the harness if any check fails:

- eager and source-backed collections agree on count, position, authored
  names, typed object tokens, shape identity, and source properties bytes;
- source-backed collections retain a read set and source version, while eager
  collections do not expose a source read set;
- repeated position/name selectors resolve the same control, while the
  duplicate-name fixture remains an explicit ambiguity;
- cloned collections and views remain equal and retain exact source bytes;
- lowering `OwnerLimits::max_controls` below the fixture count returns a
  typed error for both eager and source-backed public entry points;
- retained property-member hashes match the fixture archive bytes.
- the generated unrelated-member package has the same selected control and
  source-byte assertions as its native input, while its package size and
  source read counters are recorded independently.

This smoke intentionally measures public API behavior and bounded ownership
retention. It does not measure save/rewrite, native Office behavior, a cold OS
page-cache state, concurrency scaling, decompression throughput in isolation,
or a baseline implementation. Such claims require a separately approved
corpus and protocol.
