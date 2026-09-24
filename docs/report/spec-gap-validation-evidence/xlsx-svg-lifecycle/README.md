# XLSX worksheet SVG lifecycle evidence preparation

The implementation requirements are in
[`../xlsx-svg-lifecycle-design.md`](../xlsx-svg-lifecycle-design.md).
This directory prepares independent validation tools; it is not evidence that
the XLSX lifecycle is implemented or approved.

`validate_schema.py` validates complete Transitional worksheet drawing XML
against the vendored ECMA schema and validates each recognized direct SVG owner
against the vendored MS-ODRAWXML schema. It recognizes direct pictures in all
three worksheet anchor forms, checks duplicate admitted SVG owners, and leaves
unknown extension URIs opaque. It does not validate package relationships or
media reachability, process MCE branches, render SVG, or prove native application
acceptance. Strict dialect validation remains separate work pending the exact
extension schema/relationship namespace policy.

Run the validator on generated XML files:

```sh
python3 docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle/validate_schema.py drawing.xml
```

The source-bound `schema-tool-check.json` records a native drawing and synthetic
one-cell/absolute-anchor variants accepted by the validator, plus malformed
marker, element-only text, and duplicate SVG-owner controls that are rejected.
The synthetic variants are validator tests, not library-generated output.
Reproduce them without creating temporary fixtures or Python caches:

```sh
PYTHONDONTWRITEBYTECODE=1 python3 docs/report/spec-gap-validation-evidence/xlsx-svg-lifecycle/check_schema_tool.py
```

Final lifecycle evidence must additionally cover public API read/attach/detach,
all anchor geometries, native shared-owner cleanup, exact source/inverse/stale
behavior, configured limits, package save/reopen, independent review, strict
crate gates, and source-bound profiling after production freeze.
