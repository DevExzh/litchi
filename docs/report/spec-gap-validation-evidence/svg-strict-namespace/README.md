# Strict SVG namespace evidence

This directory records a bounded downstream proof for the namespace boundary
between Strict OOXML core drawing markup and the Microsoft SVG extension.

The central result is visible in `outputs/strict-attached.xml`: the slide root,
`a:blip`, and core `r` binding use Strict URIs, while the newly authored
`asvg:svgBlip r:embed` locally binds `r` to the Transitional
`officeDocument/2006/relationships` namespace. The attached relationship part
uses the Strict physical image relationship type. The complete Strict slide
passes the vendored ECMA Part 1 schema, and the direct SVG child passes the
unmodified MS-ODRAWXML §5.24 schema. `outputs/direct-strict-invalid.xml` is the
negative control: the same child with a Strict `r:embed` fails §5.24 because
that schema imports Transitional `a:AG_Blob` and Transitional relationship
attributes.

`outputs/xlsx-mixed-strict.xml` provides the equivalent native-derived
SpreadsheetDrawing schema control. It is a synthetic Strict conversion of the
retained Transitional drawing fixture: its XDR and DrawingML core are Strict,
its SVG child relationship attributes remain Transitional for the MS child
schema, and both schema layers pass. The input fixture and member hashes are
recorded in the schema result.

The retained package harness is in `harness/main.rs`. It uses the existing
PPTX lifecycle API on the retained synthetic source fixture and does not add a
private production or test source path. The source fixture is retained in the
earlier lifecycle evidence directory and is read-only input here.
`raw-results/` contains the harness log, schema result, and tool provenance.
`gates/` contains the compile-first receipt and source manifest;
`probe-receipt.json` binds those inputs to every generated output.

Run from the repository root:

```text
python3 docs/report/spec-gap-validation-evidence/svg-strict-namespace/run_checks.py
python3 docs/report/spec-gap-validation-evidence/svg-strict-namespace/run_probe.py
python3 docs/report/spec-gap-validation-evidence/svg-strict-namespace/verify.py
```

The scope is limited to schema and downstream byte/lifecycle proof. It does
not establish native application acceptance, rendering, or final feature
approval.
