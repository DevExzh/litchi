# 0362ul43 — Require retained limits for the XLSX drawing parser

- Parent source: `30aaafbf98e097aeec8cf6fe3df6caaae62a37d6`.
- Audit: `(deleted-coord-artifact)/raw-record-limit-audit-3431ba566c.md`, SHA-256
  `fe35d28ff10d953179aa2768c35abaf80e6bd178d475825bcbef07cb05ea67fd`.
- Diagnostic source tips: `91045f088e` (API and owner migration),
  `01682c04eb` (glossary argument identity), `4d30f3e9f8` and `f1d98c3bde`
  (remaining integration consumers), all preserved under the named
  `refs/diagnostics/uo2-xlsx-raw-limits-0362ul43/` refs.
- Final source and record delivery is one atomic commit based on the parent.
- Status: source repair, bounded validation, and static validation complete.

## Contract repair

`litchi_xlsx::drawing::parse` is now the single public drawing codec entry
point and requires `&litchi_opc::SourceLimits`.  The former public
limit-less wrapper and the private `parse_with_limits` name are gone.  The
worksheet package path passes its retained `owner.package.read_limits()` and
the pivot-chart path passes its retained `package.read_limits()`, so a pivot
load cannot silently fall back to default limits.  The parser continues to
intersect the caller's part, XML-event, and XML-depth ceilings with its
drawing-specific bounds before typed allocations.

All current consumers were migrated in this lane: the worksheet and
pivot-chart owners, the drawing unit tests, the drawing content-part,
SVG-lifecycle, and anchor-geometry integration tests.  The existing
exact-budget tests now exercise the public canonical entry point: an exact
no-MCE part budget is accepted, an MCE-expanded part is refused before typed
parsing, and an insufficient depth budget is refused.  No drawing capability
or default alias was added.

A focused adjacent audit found no second gap in the drawing source scanners:
`WorksheetSourceScan::scan` requires `WorksheetSourceLimits`,
`SourceDrawing::scan` requires `ScanLimits`, and their public output helpers
take explicit byte bounds.  This record does not claim closure for unrelated
XLSX raw codec families.

The `xlsx-minimal` and `xlsx-full` glossary profiles register the required
signature `litchi_xlsx::drawing::parse(xml, caller)` with the PPT owner's
`required_signatures` schema and the nominal `litchi_xlsx::SourceLimits` type
(the checker unwraps the required reference before comparing nominal identity).

## Validation

The eight changed Rust files were formatted with
`rustfmt --edition 2024`; `git diff --check` passed.  The glossary JSON was
parsed successfully, and a source audit confirmed that no
`parse_with_limits` drawing consumer remains and that every call supplies a
caller limit.

The final bounded receipt is external at
`(deleted-coord-artifact)/receipt.json`, SHA-256
`f1659bf765dcfcfa272b54b71b52ef042566901ce3a8494fc34a60dab4ea9a63`.  It
binds clean source `f1d98c3bde4e89efb4315842238534c6b1695c05` and the copied
workspace lock SHA-256
`4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1` from
`full-gate-9781cf6d87/source/Cargo.lock`.  The accepted f5 runner is
`uo2-gate-runner-self-attribution/docs/unified-ops/results/0362ul29/tooling/run_safe_gate.py`,
SHA-256 `1904177040fc7c523582433028d9bba56efc30fe0d9d96beb554a2fab343691e`.

With `LITCHI_CAPPED` and `LITCHI_BLESS` unset, one build job, one test
thread, incremental/debug information disabled, offline locked dependencies,
`LITCHI_MEM_MAX=20G`, and the external target
`(deleted-coord-artifact)`, the receipt records:

- 20 `drawing::` library tests passed (2,101-byte log, SHA-256
  `bb82c4c88b940c8f66258088b8d6a27d2b7fe507926ab6a5012446f258d4db9c`).
- 12 `drawing_content_part` integration tests passed (1,348-byte log, SHA-256
  `52ea7b0e51c132ab28f60c80a3cf44fcbbd0fc1b0ac668c0e0579141aacaf5b1`).
- `litchi-xlsx` all-target, all-feature clippy passed with `-D warnings`
  (177-byte log, SHA-256
  `909ff6d6535ad815dde69555f71ae26abbe4a8a3a0cbed3ce39215b20079a828`).

The maximum sampled target was 810,835,755 bytes, below both the 19 GB live
guard and 20 GB hard guard.  Earlier sparse-materialization and
compiler-discovered-consumer failures remain as immutable `receipt-pre-*`
diagnostics under the same external state directory; they are not used as the
final result and no large evidence artifact is committed.  Before the atomic
history rewrite, the tracked source diff over `crates/` and `tools/` from the
parent through the tested tip had SHA-256
`9867a2b9d374151972546a8b8cf86ea8792ca5001a1e52a1d998cea9a2d622f9`; the
final commit is checked against that same digest.
