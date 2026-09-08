# 0467 source review: XLSX cell-attribute scan

This is the source and semantics review for the 0467 candidate. The production
change is committed as
`87733cf3b86c5500ee7aee4cf6a4cb0f3cabf6a7` (`perf(xlsx): scan eager cell
attributes once`), on top of the measured control revision
`33243bc2e6bd63ab056eb45493b0717818501f4b`. The control binary has been rebuilt
in a clean role worktree, while the candidate clean build, formal A/B captures
and remaining repository validation are owned by the coordinator. This file
therefore records the mechanism and review obligations; it does not authorize
a performance claim.

The review was made against the accepted ADR set listed by
`docs/adr/README.md`. In particular, the change remains inside the XLSX raw
codec owner (ADR 0002 and ADR 0024), keeps validation and preservation
fail-closed (ADR 0006), and is evaluated under the measured-evidence contract
in ADR 0005. It does not change a public API, package ownership, snapshot/edit
publication, source limits, or the validated-store handoff.

## Baseline path and measured relevance

The 0466 source review identified `Parser::start_cell` in
`crates/litchi-xlsx/src/raw/worksheet/codec.rs` as a direct hot path. For each
`<c>` element, the old implementation called the shared
`unqualified_attribute_value` helper five times, for `r`, `s`, `cm`, `vm`, and
`t`. Each call created a fresh checked quick-xml attribute iterator, walked the
complete attribute list, ignored qualified names, checked malformed and
duplicate syntax, and decoded the selected value into an owned string.

The 0466 frame-pointer profile contained 15,483 samples and 135,131,742,466
weighted event periods. Inclusive Parser ancestors accounted for
35,501,949,212 periods within the exact commit marker, and the shared
`unqualified_attribute_value` context accounted for 11,458,403,980 periods
across the whole-process denominator. The Heaptrack export reported
56,349,806 allocation calls, with repeated quick-xml `RawVec` growth under
checked attribute iteration identified as a concrete allocation lead. These
are whole-process, inclusive observations: they include warmups, expected
output construction, reopen/verification and teardown where applicable. They
do not prove an operation-local allocation delta or speedup.

The ordinary dense corpus remains two 256-by-256 worksheets with 131,072
stored cells and 1,311 staged updates. Its source archive is
`5dd3ad701eb686f6d2d14e9f177a4e9433445728b57b484d53f663b2f87a7714`.
The timed region is ordinary `Edit::commit` followed by
`Workbook::write_to`; source open, staging, expected-output construction,
reopen, complete verification and teardown are outside that timer. The
candidate targets one parser operation inside this larger path and should be
judged only by the frozen 0467 matched-role protocol and its correctness
gates.

## Candidate mechanism

The candidate adds a private `CellAttributeView<'a>` and
`scan_cell_attributes` in the XLSX worksheet codec.

1. `scan_cell_attributes` creates one `element.attributes().with_checks(true)`
   iterator and consumes the complete list once.
2. It ignores qualified names by requiring an unqualified key before matching
   the five SpreadsheetML cell fields.
3. It decodes `r` as soon as that attribute is encountered and owns the
   resulting `String`, preserving the old coordinate-decoding point.
4. It retains `s`, `cm`, `vm`, and `t` as the raw borrowed
   `quick_xml::events::attributes::Attribute` values in five fixed `Option`
   fields. No attribute vector or per-cell map is introduced.
5. `parse_cell_u32` decodes and parses the three numeric fields at the old
   style, cell-metadata, and value-metadata points. The cell type is decoded at
   the old `PendingCell` construction point.

The expected mechanism is removal of four repeated attribute-list scans and
their per-iterator duplicate-check bookkeeping for each cell. The candidate
still performs one complete checked scan and still decodes every modeled value
that the old path decoded. It does not disable duplicate checking, retain a
document-sized index, change XML ownership, or bypass numeric bounds.

## Semantic review

The important error-order property is preserved by the shape of the helper.
The old first `r` lookup completed its checked iterator before `parse_a1` ran.
The candidate also completes the checked scan before parsing the coordinate.
Thus malformed attribute syntax and duplicate names anywhere in the start tag
remain observable before an invalid coordinate. The candidate decodes `r`
during the scan, so an invalid entity in `r` still fails before a later
duplicate or value-decoding error, as it did in the old first lookup. The
remaining four values stay raw until after coordinate validation, preserving
the existing style, cell-metadata, value-metadata, and cell-type order.

The following behavior is intentionally retained:

- `with_checks(true)` rejects malformed and duplicate attributes, including
  duplicates of otherwise ignored names; the iterator error is converted to
  the existing shared `XmlError` boundary.
- A qualified `x:r`, `x:s`, `x:cm`, `x:vm`, or `x:t` is not substituted for its
  unqualified SpreadsheetML field. Qualified and unknown attributes remain
  outside the typed cell state while their raw XML remains owned by the
  surrounding preservation path.
- Entity and XML 1.0 attribute normalization uses the same
  `decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)` operation.
- `s`, `cm`, and `vm` retain their existing integer parsing and Office bounds;
  the row's `last_column` is still updated before those later checks, matching
  the existing parser control flow for a rejected cell.
- `t` remains an optional decoded string in `PendingCell`; formula and other
  worksheet element parsers are untouched.

The candidate tests add coverage for entity-encoded core fields, duplicate
core and ignored attributes, qualified extensions, coordinate/style/error
precedence, metadata bounds, and a shared-string read. The retained
`cargo test --locked -p litchi-xlsx --all-features` receipt passes all 1,242
tests across 59 targets: 935 library tests plus the integration and doctest
targets in `validation/xlsx-tests.log`.

The other completed scoped gates are XLSX formatting, the all-feature
workspace check, XLSX `-D warnings` Clippy with `--lib --no-deps`, rustdoc with
warnings denied, and the crate-boundary checker. Workspace-wide formatting is
not green because the unchanged excluded iWork file
`crates/litchi-keynote/src/document.rs:592` is the only retained difference;
`validation/format-scope.json` records its base identity and exclusion. Both
the original separate-path role builds and the corrected same-shared-tree
control/candidate builds completed clean. All four fixed 500-sample captures
completed and their primary mean, p50, p95 and p99 statistics pass the unchanged
ABBA qualification policy. DOC guards and process RSS remain within the
five-percent review threshold; the short full-matrix flags remain disclosed.

`source-bindings.json` binds 6,992 present Rust/TOML/lock source files for
each role. It separately hashes the two compile-time included fixtures and
explicitly records 40 omitted historical evidence/probe files, including the
unused root `Cargo.lock`; the performance harness uses
`tools/perf-baseline/Cargo.lock`. The candidate binding lists only the two
XLSX production files as changed from control.

## Review risks and required gates

The fixed view is small, but the borrowed `Attribute` values must remain tied
to the `BytesStart` event until they are decoded. The all-feature test,
workspace-check and rustdoc gates provide compile-time coverage for this
borrow boundary, and both clean role builds cover the performance binary. The
fixed ABBA comparison is complete; its raw reports and exact recomputation
are retained in `fixed-qualification.json`.

The duplicate and malformed-input cases should be checked at the typed error
boundary, not only with `is_err()`. The candidate deliberately maps iterator
errors through `XmlError::Malformed`; this matches the existing quick-xml
iterator error path, but exact message and variant parity should be confirmed
for duplicate `r`, `s`, `cm`, `vm`, `t`, unknown, and qualified names. Entity
decode errors must remain distinct from duplicate errors when `r` is the first
modeled field encountered.

Validation should include the existing worksheet parser suite, the new focused
tests, warning-denied lint and formatting, preservation/round-trip checks, and
the repository's applicable boundary and verification gates. The normal
release candidate must then be built in its own clean worktree with the same
flags and captured in the frozen A1/B1/B2/A2 order. The full default guard and
Heaptrack lanes are supplementary; instrumented elapsed times must not be
compared with normal lanes.

## Limits of this review

This change addresses one repeated parser operation. It does not remove the
four complete worksheet Store parses expected in the ordinary dense commit,
alter publication validation, improve the writer, change cache retention, or
make the selected-cell path source-backed. It introduces no parallelism,
unsafe code, archive dependency, or public semantic surface.

No 0467 speed, throughput, allocation, RSS, tail-latency, or Amdahl claim is
supported until the clean matched-role captures and their individual rows,
same-role drift, and correctness receipts are available. The 0466 profile is
the hypothesis evidence for this candidate, not an after measurement. Any
retained result must remain scoped to the named XLSX case, dense corpus,
binary/build identity, CPU affinity, worker count, sample protocol and metric.
