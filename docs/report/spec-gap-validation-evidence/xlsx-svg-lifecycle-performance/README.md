# XLSX ordinary worksheet SVG lifecycle profile scaffold

This directory is a bounded profiling scaffold for the ordinary worksheet
`SpreadsheetDrawing` SVG lifecycle described by
[`xlsx-svg-lifecycle-design.md`](../xlsx-svg-lifecycle-design.md). It is
wired to the current source-backed owner and public attach/detach signatures,
but remains freeze-gated for acceptance. It contains no final or acceptance
measurements or performance claims. A separately labeled
`results/exploratory-before/` bundle retains one bounded current-working-tree
receipt per same-drawing 16/64/256 attach cost lane for optimization
comparison. A later `results/exploratory-detach-before/` bundle uses six
shared and distinct multi-detach lanes. Neither bundle must be read as frozen
evidence.

The default profile boundary is one worksheet drawing relationship containing
direct `xdr:pic` owners. The `multisheet_attach_detach` lane adds a second
ordinary worksheet and drawing so public worksheet selection is exercised
across independent owners. The profile covers an existing PNG compatibility fallback, an
embedded `asvg:svgBlip` owner, source-backed capture, snapshot cloning,
source splice and graph-planning work, candidate publication, and package
reopen. The matrix exercises `twoCellAnchor`, `oneCellAnchor`, and
`absoluteAnchor` pictures. A picture is selected by the ordinary semantic
worksheet / drawing / picture selector from the current settled public API; the harness must
not require callers to provide relationship IDs, package paths, or XML
offsets.

The following remain outside this batch: chartsheets, chart user-shape
drawings, grouped-picture ownership, `mc:AlternateContent` ownership,
linked-resource fetching, SVG rendering or conversion, DOCX, and picture
creation from an empty drawing. Unknown extensions are retained or refused as
specified by the design; they are not silently interpreted as SVG owners.

`harness/Cargo.toml` now links the settled `litchi-xlsx` and `litchi-opc`
interfaces. `harness/main.rs` and `harness/adapter.rs` call
`Workbook::edit`, semantic worksheet selection, `PictureSelector`, borrowed
`SvgInput`, `SourceDrawing`, and `Commit`. The shell runner still requires an
explicit API-wiring and freeze gate, so compiling the adapter does not collect
measurements while production semantic checks are still being reviewed.

The deterministic synthetic corpus is recorded in
[`corpus-manifest.json`](corpus-manifest.json). It describes small and large
opaque payloads, 256/1,024-picture inventory cases, shared and distinct SVG
targets, inherited namespace pressure, three anchor forms, strict-host and
incoming-edge lifecycle cases, captured-owner clone cases, the multisheet
attach/detach package, and refusal cases.
Each raw receipt records both the legacy FNV-1a identity and a SHA-256 digest
of the exact bounded input identity; the verifier requires those identities to
remain stable across fresh processes.
The native producer fixture is used for read/capture validation only:

[`fixtures/tdf169496_hidden_graphic.xlsx`](fixtures/tdf169496_hidden_graphic.xlsx)

Its archive hash is the design hash
`0b647da300a085f39914fdfae961463ae9e54ffe772b2e0eb9860a841ab93f72`.
The unchanged LibreOffice corpus file is retained in Git so a clean checkout
does not depend on the ignored `3rdparty` tree. Its original path and retained
license notices are documented in [the fixture provenance](fixtures/README.md).
The adapter does not mutate that fixture. The
`multi_picture_same_drawing_{16,64,256}` lanes stage several lifecycle changes
on one drawing in one transaction. They retain the current composed-planner
cost shape, including its per-intent source rescans, without making a scaling
claim. `source_copy_bytes` and `staged_bytes` remain explicitly unavailable
in receipts because the public API does not expose those stage counters;
allocator totals are not used as a substitute.

The 69-lane acceptance matrix also includes
`inverse_attach_detach_{anchor}_{size}` for exact in-memory inverse restoration
and replay refusal across all three anchors and both payload sizes, plus the
composite `mixed_caps_rejection` lane. The latter exercises the retained part,
aggregate bytes, relationship count/XML bytes/XML events, and content-type
mapping ceilings against one shared detach+attach transaction and records each
requested ceiling in the receipt; it is an atomic refusal gate rather than a
per-cap performance comparison.

The acceptance matrix also contains a strict-host attach lane, a strict-host
final-owner detach lane, and an incoming-edge final-owner lane. The strict
lanes check worksheet and drawing relationship dialects after publication; the
incoming-edge fixture is built with a separate opaque package relationship so
the retained SVG leaf cannot be removed solely because the selected drawing
owner disappeared; it also asserts that the selected drawing-to-SVG
relationship itself was removed. Shared-final fixtures are assembled directly in the corpus
generator and never call production detach to prepare their input.

`clone_captured_owner_{small,large}` scans and validates one source drawing
before warm-up and measurement, then times only a clone of the already
captured `PictureSource`. It is kept separate from the workbook open/save
clone lanes so its receipt scope does not imply a workbook snapshot or capture
parse cost. The namespace refusal fixture derives and records the first
refused active-binding boundary by probing the admitted host policy. Its
receipt records `namespace_generated_bindings` for pressure-fragment
declarations, `namespace_active_bindings` including the seven fixed drawing
root bindings, and `namespace_active_limit`; the first refusal is
`active_limit + 1` and the preceding generated count is accepted.
Opaque extension and descendant checks compare exact bytes and occurrence
counts after lifecycle operations. Its allocator snapshot intentionally keeps
the cloned owner live until after the snapshot; byte/owner/reference validation
and clone cleanup happen after the timed region.

The `same_picture_attach_detach_{two_cell,one_cell,absolute}` lanes invoke
public attach and detach in separate commits for the same semantic picture
and require source drawing/content-type restoration. The
`multisheet_attach_detach` lane creates two ordinary worksheets with
independent drawings, attaches Sheet1/two-cell and Sheet2/one-cell owners in
one transaction, then detaches both after reopen. It checks independent SVG
parts, anchor geometry, and graph cleanup through public worksheet selection.

The six exploratory detach lanes,
`multi_picture_same_drawing_detach_{shared,distinct}_{16,64,256}`, use the
same bounded picture counts and exercise the fallback path with shared and
distinct SVG targets. They live outside the 69-lane acceptance list.

The exploratory `inventory_shared_root_namespace_32` lane inventories 32
distinct SVG owners under 128 large inherited root bindings. Its source-shape
checks require raw SVG source and namespace context projections for all 32
owners. Its source-shape receipt is paired with the retained scope probe at
[`../xlsx-svg-source/retained-scope-probe.rs`](../xlsx-svg-source/retained-scope-probe.rs),
which records retained owner source bytes only. Neither artifact is a timing,
peak-memory, or final performance result.

The separate [`phase-decomposition.md`](phase-decomposition.md) harness is an
exploratory diagnostic for the three same-drawing attach fixtures. Its
`xlsx-svg-lifecycle-phase-profile` binary and `run_phase_profile.sh` runner
split each public operation into `open`, `stages`, `commit`, `firstsave`,
`reopen_secondsave`, and complete `validation` clocks. Each phase records
phase-local requested allocation and the live bytes retained across the
boundary; phase values are never summed or subtracted across unlike live
sets. This harness is outside the 69-lane acceptance matrix and has no
baseline comparison or performance conclusion.

When the owner and adapter are frozen, the intended command is:

```sh
XLSX_SVG_PROFILE_SOURCE_PIN=ac288a303264f9ea0bb4081baa44031bee5b79a7 \
XLSX_SVG_PROFILE_RESULTS=/var/tmp/litchi-xlsx-svg-lifecycle-results-$$ \
PROFILE_FROZEN=1 \
XLSX_SVG_PROFILE_API_WIRED=1 \
CARGO_TARGET_DIR=/var/tmp/litchi-xlsx-svg-lifecycle-target-$$ \
bash docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle-performance/run_profile.sh
```

The runner rejects unset freeze or wiring gates before target setup, and it
rejects a missing/mismatched committed source pin after the initial metadata
probe but before build or measurement. The
`profile_pins.py`, `committed_inputs.py`, and `source_manifest.py` checks read
Git blobs directly and reject untracked or dirty production inputs; a working
tree hash alone is not sealed evidence. It also rejects `RUSTFLAGS`, bootstrap,
and similar suppression variables; runs each lane in three fresh processes with twenty measured
samples after two warm-up samples; records `/usr/bin/time -v` RSS; and removes
only a newly created isolated Cargo target. The runner never removes the main
Cargo target. Set `ALLOW_EXISTING_TARGET=1` only when an explicitly isolated
target has been inspected and is intentionally reused.
The source manifest includes the tracked root `rust-toolchain.toml`,
`.cargo/config.toml`, and the harness `Cargo.lock` as build inputs; a sealed
run therefore requires those files to remain committed and unchanged. The
isolated harness lock owns dependency resolution for this profile.
`XLSX_SVG_PROFILE_RESULTS` must name a nonexistent external directory; the
runner creates it and writes `report.md`, raw receipts, source manifests,
provenance, and `verification.json` there. Historical evidence under this
directory is never deleted, and failed-run diagnostics remain in the fresh
external output for review.

Each process retains a stderr file and an exit-status receipt. Verification
requires empty stderr, exit status zero, exact process identities, consistent
input sizes and digests, and identical executable digests before and after
measurement. The executable itself is disposable; retained digest receipts
record its measured identity rather than promising later executable rehashing.
Run this command from the directory to exercise the adversarial receipt,
runner-output, committed-source snapshot, and derived-summary checks without
compiling or measuring production code:

```sh
python3 -B -m unittest test_verify test_runner test_source_snapshot test_uncertainty
```

After `verify.py` accepts a retained run, `uncertainty.py` can produce a
separate descriptive view without rewriting raw receipts:

```sh
python3 -B uncertainty.py \
  --results /var/tmp/litchi-xlsx-svg-lifecycle-results \
  --output /var/tmp/litchi-xlsx-svg-lifecycle-uncertainty.md \
  --json-output /var/tmp/litchi-xlsx-svg-lifecycle-uncertainty.json
```

It requires exactly three fresh process receipts per lane, at least two
warmups and twenty measured samples, and matching semantic, allocator, and
typed-refusal gates. It reports each process median and min/max sample range,
the range of those process medians, and the `/usr/bin/time -v` RSS values.
The n=3 ranges are descriptive; the derived files do not provide confidence
intervals or speedup, regression, causal, or scaling claims.
Fresh `run_profile.sh` runs capture this tool in both source manifests and
generate `uncertainty.md` and `uncertainty.json` after verification succeeds.

Refusal lanes emit a typed expected-refusal outcome only after the public API
returns a matching limit/owner error and the source bytes remain unchanged.
Unexpected acceptance, source mutation, setup failure, or an unrelated API
error remains a failed lane and cannot be relabeled with the lane's requested
refusal class.

The adapter must preserve the receipt contract in
[`requirements.md`](requirements.md). It must report the timed scope for
capture, clone, source splice, graph planning, commit validation, publication,
and reopen where the public API permits those boundaries. End-to-end lanes
include the named closure and reopen checks; fixture construction and
post-run assertions stay outside isolated operation timers unless the lane
explicitly includes them. `report.md` remains a no-measurement placeholder
until a frozen adapter has produced receipts.
