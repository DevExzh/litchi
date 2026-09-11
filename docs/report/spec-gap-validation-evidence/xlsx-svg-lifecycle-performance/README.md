# XLSX ordinary worksheet SVG lifecycle profile scaffold

This directory is a bounded profiling scaffold for the ordinary worksheet
`SpreadsheetDrawing` SVG lifecycle described by
[`xlsx-svg-lifecycle-design.md`](../xlsx-svg-lifecycle-design.md). It is
intentionally unwired: the XLSX source-backed owner and its public attach /
detach API are not frozen yet, so this directory contains no measurements and
makes no performance claim.

The profile boundary is one worksheet drawing relationship containing direct
`xdr:pic` owners. It covers an existing PNG compatibility fallback, an
embedded `asvg:svgBlip` owner, source-backed capture, snapshot cloning,
source splice and graph-planning work, candidate publication, and package
reopen. The matrix exercises `twoCellAnchor`, `oneCellAnchor`, and
`absoluteAnchor` pictures. A picture is selected by the ordinary semantic
worksheet / drawing / picture selector once that API exists; the harness must
not require callers to provide relationship IDs, package paths, or XML
offsets.

The following remain outside this batch: chartsheets, chart user-shape
drawings, grouped-picture ownership, `mc:AlternateContent` ownership,
linked-resource fetching, SVG rendering or conversion, DOCX, and picture
creation from an empty drawing. Unknown extensions are retained or refused as
specified by the design; they are not silently interpreted as SVG owners.

No XLSX production dependency is present in `harness/Cargo.toml` while the
public API is provisional. `harness/main.rs` has the lane vocabulary, a
process-local counting allocator, and the receipt support used by the future
adapter, but it exits with an explicit `api-wiring-pending` refusal. This is
deliberate: copying a guessed API into a benchmark would create evidence for
an operation that does not exist.

The recipe-only synthetic corpus is recorded in
[`corpus-manifest.json`](corpus-manifest.json). It describes small and large
opaque payloads, 256/1,024-picture inventory cases, shared and distinct SVG
targets, inherited namespace pressure, three anchor forms, and refusal cases.
The native producer fixture is used for read/capture validation only:

`3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx`

Its archive hash is the design hash
`0b647da300a085f39914fdfae961463ae9e54ffe772b2e0eb9860a841ab93f72`.
Synthetic package generation and semantic assertions belong in the eventual
API adapter, not in this pre-freeze scaffold.

When the owner and adapter are frozen, the intended command is:

```sh
PROFILE_FROZEN=1 \
XLSX_SVG_PROFILE_API_WIRED=1 \
CARGO_TARGET_DIR=/var/tmp/litchi-xlsx-svg-lifecycle-profile-target \
sh docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle-performance/run_profile.sh
```

The runner rejects unset freeze or wiring gates before it creates a profile
target. It also rejects `RUSTFLAGS`, bootstrap, and similar suppression
variables; runs each lane in three fresh processes with twenty measured
samples after two warm-up samples; records `/usr/bin/time -v` RSS; and removes
only a newly created isolated Cargo target. The runner never removes the main
Cargo target. Set `ALLOW_EXISTING_TARGET=1` only when an explicitly isolated
target has been inspected and is intentionally reused.

The future adapter must preserve the receipt contract in
[`requirements.md`](requirements.md). It must report the timed scope for
capture, clone, source splice, graph planning, commit validation, publication,
and reopen where the public API permits those boundaries. End-to-end lanes
include the named closure and reopen checks; fixture construction and
post-run assertions stay outside isolated operation timers unless the lane
explicitly includes them. `report.md` remains a no-measurement placeholder
until a frozen adapter has produced receipts.
