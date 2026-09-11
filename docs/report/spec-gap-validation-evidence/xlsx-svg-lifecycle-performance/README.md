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

The profile boundary is one worksheet drawing relationship containing direct
`xdr:pic` owners. It covers an existing PNG compatibility fallback, an
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
incoming-edge lifecycle cases, captured-owner clone cases, and refusal cases.
Each raw receipt records both the legacy FNV-1a identity and a SHA-256 digest
of the exact bounded input identity; the verifier requires those identities to
remain stable across fresh processes.
The native producer fixture is used for read/capture validation only:

`3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx`

Its archive hash is the design hash
`0b647da300a085f39914fdfae961463ae9e54ffe772b2e0eb9860a841ab93f72`.
The adapter does not mutate that fixture. The
`multi_picture_same_drawing_{16,64,256}` lanes stage several lifecycle changes
on one drawing in one transaction. They retain the current composed-planner
cost shape, including its per-intent source rescans, without making a scaling
claim. `source_copy_bytes` and `staged_bytes` remain explicitly unavailable
in receipts because the public API does not expose those stage counters;
allocator totals are not used as a substitute.

The 65-lane acceptance matrix also includes
`inverse_attach_detach_{anchor}_{size}` for exact in-memory inverse restoration
and replay refusal across all three anchors and both payload sizes, plus the
composite `mixed_caps_rejection` lane. The latter exercises the retained part,
aggregate bytes, relationship count/XML bytes/XML events, and content-type
mapping ceilings against one shared detach+attach transaction; it is an
atomic refusal gate rather than a per-cap performance comparison.

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

The six exploratory detach lanes,
`multi_picture_same_drawing_detach_{shared,distinct}_{16,64,256}`, use the
same bounded picture counts and exercise the fallback path with shared and
distinct SVG targets. They live outside the 65-lane acceptance list.

The exploratory `inventory_shared_root_namespace_32` lane inventories 32
distinct SVG owners under 128 large inherited root bindings. Its source-shape
checks require raw SVG source and namespace context projections for all 32
owners. Its source-shape receipt is paired with the retained scope probe at
[`../xlsx-svg-source/retained-scope-probe.rs`](../xlsx-svg-source/retained-scope-probe.rs),
which records retained owner source bytes only. Neither artifact is a timing,
peak-memory, or final performance result.

When the owner and adapter are frozen, the intended command is:

```sh
PROFILE_FROZEN=1 \
XLSX_SVG_PROFILE_API_WIRED=1 \
CARGO_TARGET_DIR=/var/tmp/litchi-xlsx-svg-lifecycle-profile-target \
bash docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle-performance/run_profile.sh
```

The runner rejects unset freeze or wiring gates before it creates a profile
target. It also rejects `RUSTFLAGS`, bootstrap, and similar suppression
variables; runs each lane in three fresh processes with twenty measured
samples after two warm-up samples; records `/usr/bin/time -v` RSS; and removes
only a newly created isolated Cargo target. The runner never removes the main
Cargo target. Set `ALLOW_EXISTING_TARGET=1` only when an explicitly isolated
target has been inspected and is intentionally reused.

Each process retains a stderr file and an exit-status receipt. Verification
requires empty stderr, exit status zero, exact process identities, consistent
input sizes and digests, and identical executable digests before and after
measurement. The executable itself is disposable; retained digest receipts
record its measured identity rather than promising later executable rehashing.
Run `python3 -B -m unittest test_verify` from this directory to exercise the
adversarial receipt checks without compiling or measuring production code.

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
