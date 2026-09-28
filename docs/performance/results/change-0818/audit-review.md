# 0818 independent artifact-auditor review

This review covers the new `artifact_audit.py` only. The 0817 auditor and its
failed admission report remain immutable inputs. The new reader uses only the
Python standard library, ZIP reading, and XML parsing; it imports no project
crate and does not run Cargo, a binary, a workload, or a subprocess.

## Manifest accounting

The artifact exporter has six cases and five byte-identical policy outputs per
case. The auditor keeps the physical and logical identities separate:

* `archive_member_count`, `archive_bytes`, and `archive_sha256` are checked
  against the ZIP inventory for every case.
* Real-file `entry_count`, `target_entry`, target payload identity, and
  `uncompressed_payload_bytes` are checked against the decoded physical
  members, matching `real_file_manifest` in
  `tools/perf-baseline/src/ordinary_save.rs:1003`.
* Generated `entry_count`, `entry_bytes`, `target_entry`, target payload
  identity, and `uncompressed_payload_bytes` are recomputed from bounded
  semantic projections. They are never compared with the number or total
  decoded size of ZIP members.

The generated projections are fixed to the medium constructors used by this
artifact export. DOCX checks the 200 deterministic paragraph strings and sums
their UTF-8 payloads, matching
`tools/perf-baseline/src/lib.rs:16553-16580`. PPTX checks 12 slides with eight
deterministic shape strings per slide and sums their UTF-8 payloads, matching
`tools/perf-baseline/src/lib.rs:21906-21938`. XLSX checks the four 48-by-48
numeric sheets, each expected A1-style cell value, the 93-update metadata, and
the eight exact `litchi-cell-crud-00.png` through `07.png` 512 KiB media
payloads. Its logical byte total is recomputed as `cell_count * 4 + 8 * 512
KiB`, matching `tools/perf-baseline/src/lib.rs:22710-22810`.

The generated target hashes are hashes of the logical target payload (`0`, the
first paragraph string, or the first shape string), while real-file target
hashes are hashes of the named main XML member. The distinction is intentional
and is now checked independently.

## XLSX edit closure

An admitted XLSX edit may change the target worksheet cell and the owner's
direct workbook `calcPr` node. The auditor removes only direct spreadsheet
`calcPr` children from deep copies of source and output workbook roots and
requires the remaining XML to have exactly equal canonical form. It separately
requires:

* equal workbook root tags and complete root attribute dictionaries;
* equal workbook namespace-declaration census;
* no more than one source direct `calcPr` and exactly one output direct
  `calcPr`;
* source `calcPr` markup has no unsupported attribute, child, or non-empty
  text that could be silently discarded;
* output `calcPr` attributes exactly
  `calcId=0`, `fullCalcOnLoad=true`, `calcCompleted=false`,
  `calcOnSave=true`, and `forceFullCalc=true`, with no unknown attributes;
* no output `calcPr` child elements or non-empty text.

Thus an unknown workbook attribute, child, namespace declaration, relationship,
worksheet structure, formula, or extension outside the direct `calcPr` span
still fails. The allowance does not cover a workbook part wholesale. The
source-backed closure documentation identifies the owner span as worksheet
dimension/sheetData plus workbook `calcPr` at
`crates/litchi-xlsx/src/cell_values/validation.rs:5-20`, and the invalidation
helper is at `crates/litchi-xlsx/src/raw/recalc.rs:1-12`.

## DOCX preservation

The DOCX allowance remains exactly `word/document.xml` for an admitted
append. Every other changed member must retain decoded bytes, including
`word/_rels/document.xml.rels`; the ordered relationship edge list is still
compared directly. Canonical or graph equality cannot admit a relationship
child reorder. The target-part comparison removes only the appended marker
paragraph and compares the remainder canonically, while the package-level
member check rejects any untouched decoded-byte change.

## Preflight against the immutable 0817 artifacts

The new auditor was run with `python3 -B` against
`docs/performance/results/change-0817/artifacts` and wrote a distinct temporary
report. The result was deliberately failed admission with exactly three errors,
all on `real-000-docx`:

```text
real-000-docx: OPC relationship graph changed outside the edit closure
real-000-docx: decompressed member changed outside target: word/_rels/document.xml.rels
real-000-docx: canonical XML changed outside target: word/_rels/document.xml.rels
```

The other five cases were green: generated DOCX, generated XLSX, generated
PPTX, real XLSX, and real PPTX. The durable witness is retained under
`change-0818/preflight/`: `artifact-audit.json` is the report, `audit.log` is
the exact stdout/stderr capture, and `receipt.json` binds both to the auditor,
the frozen 0817 auditor, and the immutable 0817 manifest. The new script hash
for that run was
`f3fbb57f7a8e10aa1b45a18874d49003cd5b033c62548265b134f8338b049026`.
The preflight report hash was
`5d61b31c43d3985395fab3ebc707b8fb6ce30bb28e2150997d3603a926509c49`; it had
schema `litchi.performance.0818.artifact-audit.v1`, six cases, and 30 outputs.
The retained log hash is
`f156e38af25990a3c058f5c9e9687303b99061a2f2c17f4ff773f55bdd94b379`.

`--check` retains the existing replay contract: it recomputes the report and
requires exact JSON equality, then returns success even when the retained
report has `ok: false`. Replaying this failed preflight report returned zero;
replaying the old 0817 report correctly returned a mismatch because the
corrected auditor produces a different report schema and accounting.

This preflight establishes the intended reader boundary. It does not admit the
real DOCX package, and it does not imply any timing result.
