# Admission-0 artifact audit diagnosis

This is a read-only diagnosis of the terminal failed witness at
`admission-0/artifact-audit.json`. It does not change that report, the frozen
`artifact_audit.py`, the export manifest, or any timing admission decision.

The retained receipt binds the run to auditor SHA-256
`f8df2023231e0f61ea67db712310654657c778d35263fd5c148c369b6afddc3c`,
manifest SHA-256
`8293d809be070c2eee6cd8773551103f36c1b2e33363b80395ff90b9c3a0e402`, and
report SHA-256
`01da3f5b2768ad5ad2a74b4b3b3204f6f72bfadb02e05830e51705dc2d84aa82`.
The report is `ok: false` with 22 errors.

## Error classification

| case | audit errors | diagnosis |
| --- | ---: | --- |
| `generated-docx-medium` | 5 | Generated logical-payload fields were read as ZIP inventory and main-part identity. Reader/schema mismatch. |
| `generated-xlsx-medium` | 7 | The same 5 logical-payload mismatches, plus the expected `calcPr` invalidation in `xl/workbook.xml` was outside the frozen allowed set. Reader closure mismatch. |
| `generated-pptx-medium` | 5 | Generated logical-payload fields were read as ZIP inventory and main-part identity. Reader/schema mismatch. |
| `real-000-docx` | 3 | `word/_rels/document.xml.rels` has one relationship-order normalization; the edge set and every edge value are unchanged. Graph identity is unchanged; decoded-byte and ordering preservation remain unmet under the frozen 0817 contract. |
| `real-001-xlsx` | 2 | `xl/workbook.xml` changes only the exact dirty-calculation `calcPr` node required by a cell edit. The frozen closure omitted this owned workbook span. |
| `real-002-pptx` | 0 | Passed the frozen audit. |

No error indicates an invalid ZIP, an unparseable XML member, an invalid
content-type graph, an invalid relationship target, a missing or extra member,
or a non-target cell/shape/document semantic change. These observations separate reader closure/accounting mistakes from the
DOCX lexical/order change. They do not establish full preservation under
0817 or the ordering requirement in `docs/GOAL.md`.

## Generated corpus fields are logical workload identity

The generated manifest fields are intentionally semantic workload counters.
They do not describe a ZIP member inventory. The constructors make that
distinction explicit:

* `build_semantic_docx_corpus` sets `entry_count` to the paragraph count,
  `target_entry` to `paragraph:0`, and `uncompressed_payload_bytes` to the sum
  of paragraph text bytes ([`tools/perf-baseline/src/lib.rs`](../../../../tools/perf-baseline/src/lib.rs:16553)).
* `build_xlsx_cell_crud_corpus` sets `entry_count` to the cell count,
  `entry_bytes` to the logical `i32` cell width, `target_entry` to the A1 cell,
  and `uncompressed_payload_bytes` to logical cells plus generated media
  ([`tools/perf-baseline/src/lib.rs`](../../../../tools/perf-baseline/src/lib.rs:22710)).
* `build_semantic_pptx_corpus` sets `entry_count` to slide-shape count,
  `target_entry` to `slide:0/shape:0`, and `uncompressed_payload_bytes` to
  logical shape-text bytes ([`tools/perf-baseline/src/lib.rs`](../../../../tools/perf-baseline/src/lib.rs:21906)).

The source ZIP measurements and declared logical values are:

| case | ZIP members / decompressed bytes | declared `entry_count` / `uncompressed_payload_bytes` | declared target / bytes |
| --- | ---: | ---: | --- |
| generated DOCX | 12 / 51,896 | 200 / 10,000 | `paragraph:0` / 50 |
| generated XLSX | 17 / 4,466,367 | 9,216 / 4,231,168 | `Sheet1!A1` / 1 |
| generated PPTX | 61 / 141,726 | 96 / 4,992 | `slide:0/shape:0` / 52 |

The generated target hashes likewise hash the logical text or cell value, not
`word/document.xml`, `xl/workbook.xml`, or `ppt/presentation.xml`. The source
archive identity and `archive_member_count` are correct. The five generated
manifest errors per case are thus auditor assumptions that apply to the
real-file manifest, not evidence of exporter output loss.

The real-file constructor uses the physical member count, decompressed member
sum, and main-part bytes for these fields ([`ordinary_save.rs`](../../../../tools/perf-baseline/src/ordinary_save.rs:1003)); those fields agree with the three staged real inputs:

| case | input path | source bytes / SHA-256 | members / decompressed bytes |
| --- | --- | ---: | ---: |
| `real-000-docx` | `test-data/ooxml/docx/documentProperties.docx` | 23,503 / `1cff7a0a94dfce307a70032d21070d26ae34b9fdf742cf70fa66d4a2078ec9d5` | 12 / 44,910 |
| `real-001-xlsx` | `test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx` | 8,435 / `d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4` | 10 / 15,304 |
| `real-002-pptx` | `test-data/ooxml/pptx/shapes.pptx` | 68,822 / `19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571` | 48 / 154,250 |

## Real DOCX relationship member

All five real DOCX outputs are the same 23,279-byte archive with SHA-256
`9b36ea76d4bed6dca0d1b567923673d071ca3c2def46a8d0581e7f2d58031e1c`.
The only changed members are the intended `word/document.xml` append and
`word/_rels/document.xml.rels`.

The exact relationship sequences are:

```text
source: rId4/fontTable.xml, rId1/styles.xml, rId2/settings.xml,
        rId3/webSettings.xml, rId5/theme/theme1.xml
output: rId1/styles.xml, rId2/settings.xml, rId3/webSettings.xml,
        rId4/fontTable.xml, rId5/theme/theme1.xml
```

Every tuple `(Id, Type, Target, TargetMode)` is present exactly once in both
files. The edge sets are equal (5 source edges, 5 output edges); only the
child order moves `rId4`. There is no added, removed, retargeted, or
retyped relationship, and all five internal targets remain present.

The raw XML difference is also lexical:

```xml
source: <?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships ...><Relationship Id="rId4" .../><Relationship Id="rId1" .../>...</Relationships>
output: <?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships ...><Relationship Id="rId1" .../><Relationship Id="rId2" .../><Relationship Id="rId3" .../><Relationship Id="rId4" .../>...</Relationships>
```

The source/output member sizes are 817/815 bytes and compressed payload sizes
are 249/241 bytes. ZIP metadata changes from the producer's
`flag_bits=6, create_system=45, create_version=45` to the writer's
`flag_bits=8, create_system=20, create_version=20`; the namespace census is
identical (`http://schemas.openxmlformats.org/package/2006/relationships`).
Canonical XML differs only in the `Relationship` child sequence. The
relationship graph itself remains valid.

An earlier graph-based oracle accepted this serializer normalization. The prior
independent oracle models relationships as a map keyed by `Id` and explicitly
allows an adjacent `.rels` lexical rewrite when that graph is equal
([`docs/performance/results/change-0778/oracle.py`](../change-0778/oracle.py:1084)).
The frozen 0817 reader instead stores an ordered edge list and compares it
directly ([`artifact_audit.py`](artifact_audit.py:376),
[`artifact_audit.py`](artifact_audit.py:913)). That ordered comparison turns a
same-edge-set normalization into the three DOCX errors. 0817 already requires untouched decoded bytes and order to remain unchanged,
and `docs/GOAL.md` explicitly requires ordering preservation. The earlier
oracle does not override those requirements. An order-independent edge check
alone cannot admit this case; the ordering behavior needs explicit resolution
or repair before fresh admission. This is not an unexplained graph change.

## XLSX workbook `calcPr`

The XLSX ordinary edit is `Workbook::edit().sheet(...).set("A1", marker)` and
commit ([`ordinary_save.rs`](../../../../tools/perf-baseline/src/ordinary_save.rs:726)).
The source-backed value editor's documented closure includes the owner's
workbook `calcPr` span ([`crates/litchi-xlsx/src/cell_values/validation.rs`](../../../../crates/litchi-xlsx/src/cell_values/validation.rs:5)).
The recalculation helper is specifically named “Workbook `calcPr`
invalidation after semantic cell edits” and sets the dirty flags
([`crates/litchi-xlsx/src/raw/recalc.rs`](../../../../crates/litchi-xlsx/src/raw/recalc.rs:1)).

### Real XLSX

All five real XLSX outputs are the same 8,521-byte archive with SHA-256
`0f6902152f1b8c40c47086206023887e87f54ef57f1da417bc91d8494a4d3e68`.
The target worksheet cell is `Munka1!A1`; the only changed members are
`xl/workbook.xml` and the target `xl/worksheets/sheet1.xml`.

The workbook's exact semantic delta is:

```xml
source: <calcPr calcId="152511"/>
output: <calcPr calcId="0" fullCalcOnLoad="true" calcCompleted="false" calcOnSave="true" forceFullCalc="true"/>
```

All workbook children and attributes outside that direct `calcPr` are equal.
The source/output workbook sizes are 1,275/1,351 bytes and compressed payload
sizes are 618/639 bytes. The namespace census is identical. The output flags
are the exact dirty-calculation contract recorded by the earlier independent
oracle ([`change-0778/oracle.py`](../change-0778/oracle.py:683)).

This is an expected edit closure, not a preservation defect. The frozen 0817
reader's XLSX allowed-member set contains the target worksheet and optional
`xl/sharedStrings.xml` but omits the owned workbook part
([`artifact_audit.py`](artifact_audit.py:719)); it therefore reports the
workbook's exact `calcPr` rewrite as an unexplained outside-target change.
The corrected closure must compare workbook structure after removing the
direct `calcPr`, then require exactly the dirty flags above. It must continue
to reject any other workbook structure, relationship, content-type, or
namespace change.

### Generated XLSX

All five generated XLSX outputs are the same 4,226,568-byte archive with
SHA-256
`20335c4480405e051be2f4fea0e2a5aa8c5912b729cb1df980112ea7c5056687`.
The target is `Sheet1!A1`; the changed members are
`xl/worksheets/sheet1.xml` and `xl/workbook.xml`.

The source workbook has no direct `calcPr`; the output appends exactly:

```xml
source: <sheets>...</sheets></workbook>
output: <sheets>...</sheets><calcPr calcId="0" fullCalcOnLoad="true" calcCompleted="false" calcOnSave="true" forceFullCalc="true"/></workbook>
```

No sheet name, sheet ID, relationship ID, workbook attribute, or other
workbook child changes. Source/output workbook sizes are 424/527 bytes and
compressed payload sizes are 226/279 bytes; the namespace census is identical.
The same direct `calcPr` invalidation is expected when the generated workbook
starts without one. The two generated XLSX workbook errors are therefore the
same closure-reader false positive as the real XLSX errors, independently of
the five generated logical-manifest errors.

## Output verification strength and admission consequence

For every case, all five policies (`default`, `full`, `file-only`, `no-sync`,
`stream`) reproduce one reference archive byte-for-byte. The exporter manifest
reports source unchanged and successful reopen for all 30 outputs. The audit
independently parsed all 30 archives, found exact source/output member sets,
and found valid content-type and relationship graphs for every source and
output. Outside the listed changed members, decoded bytes, compressed payloads,
ZIP metadata, and XML namespace censuses are equal in the report. The DOCX,
XLSX, and PPTX target projections pass; the real PPTX case is fully green.

That verification establishes valid packages, equal relationship edge sets,
and the expected XLSX edit closure. It does not establish DOCX lexical/order
preservation or make the frozen report green. Correct the generated-manifest
and XLSX closure interpretation, resolve the DOCX ordering requirement, and
perform fresh complete admission before considering timing.

The read-only replay helper is [`admission-diagnosis.py`](admission-diagnosis.py).
It imports no project code and runs no exporter, Cargo command, workload, or
subprocess:

```text
python3 -B docs/performance/results/change-0817/admission-diagnosis.py \
  docs/performance/results/change-0817/artifacts
```
