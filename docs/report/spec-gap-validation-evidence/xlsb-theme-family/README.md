# XLSB applied-theme family metadata

**Policy follow-up:** [removal closure](../theme-family-removal-closure/README.md)
supersedes the original empty-wrapper retention behavior. The gate, review,
caller, and performance receipts below describe their recorded `9facb6057`
revision; they must not be treated as measurements of the changed removal.

This evidence directory covers optional family metadata inside an XLSB-owned
complete Theme part. The prior `drawingml-theme-family` evidence covers the
standalone fragment grammar; `xlsb-theme` covers the base Theme and whole-part
package lifecycle. See [requirements.md](requirements.md) for this batch's
ownership and preservation requirements.

## Architecture

| Contract | Application to this batch |
|---|---|
| ADR 0001/0004: typed public API | Family values use validated names and GUIDs; optional metadata participates in the existing Theme edit workflow. |
| ADR 0002/0024: vocabulary ownership | DrawingML locates and edits family XML; XLSB owns the workbook relationship and publication. |
| ADR 0003: source-checked reversible edits | Theme commits retain before/after state, exact no-op behavior, stale-source checks, and inverse patches. |
| ADR 0005: bounded I/O and memory | Family discovery inspects the accessed Theme only; source-backed views retain managed part data. Profile claims are scoped to recorded lanes. |
| ADR 0006: preservation and security | Unrelated XML and extension containers survive edits; signatures follow existing explicit publication policy. |
| ADR 0008: verification | Independent tests, review, native fixtures, offline XSD checks, downstream caller, and source-bound gates support the stated surface. |
| ADR 0011: physical package ownership | Existing OPC publication and save paths handle physical package changes. |

## Reproduction

From the repository root, run:

```sh
python3 docs/report/spec-gap-validation-evidence/xlsb-theme-family/run_gates.py
python3 docs/report/spec-gap-validation-evidence/xlsb-theme-family/run_caller.py
python3 docs/report/spec-gap-validation-evidence/xlsb-theme-family/validate_schema.py test-data/ooxml/xlsb/date.xlsb
PERF=1 docs/report/spec-gap-validation-evidence/xlsb-theme-family/performance/run_profile.sh
python3 docs/report/spec-gap-validation-evidence/xlsb-theme-family/verify_root.py --performance
```

`validate_schema.py` needs Python `lxml` and reads only vendored schemas. It
validates the complete Theme with ECMA XSD and recognized family subtrees with
the Microsoft schema, mapping Microsoft's imported GUID name to the equivalent
ECMA shared type. Strict-family nested DrawingML imports are mapped to their
Strict ECMA namespace. This is explicit schema adaptation, not evidence of
native application acceptance.

The caller uses only public APIs and the native `date.xlsb` fixture. It records
updated, removed, added, and newly authored family XML after actual XLSB save/reopen, and checks
that inverse patches restore exact source XML. Generated ZIPs remain in memory.

The final root gate passes **1,127 tests across 37 targets**, with **12 existing
skips** (one optional corpus case and eleven writer doctests). Strict Clippy,
rustdoc, formatting, topology checking, and **957 policy tests** pass. The
source-bound verifier checks **866 input files**, the downstream caller, and
**eight schema reports**. See [validation scope](validation.md),
[gate receipt](gates/receipt.json), and [root verification](root-verification.json).

The final [profile](performance/report.md) verifies **1,620 samples across 18
lanes** against **4,619 source paths**. The [root assessment](performance-assessment.md)
explains absolute costs and measurement limits.

Both independent reviewers approved the final source; see [review](review.md)
and its source hashes. The broader audit remains open.
