# 0362dvf — DrawingML diagram raw limits

- Status: source and consumer migration plus focused capped qualification
  complete; ready for coordinator merge.
- Base: `2f7db7ee7fa01893e9099217fc4243687d196f63`.

## Policy

The diagram module now has one explicit `diagram::Limits` policy. Its default
retains the prior 16 MiB part, 200,000 XML-node, 128-depth, 100,000-point and
connection, and 1 MiB text ceilings. The policy also carries output and
SmartArt-node budgets; zero and over-hard-cap values return the shared typed
`Error::InvalidLimit` before parsing or output allocation. Getters use the
`*_value()` spelling.

The bounded raw APIs are:

* `DiagramDataModel::{parse,to_xml,write_xml}`;
* `DiagramDefinition::{parse_layout,parse_quick_style,parse_colors}`;
* `SmartArt::parse_data_xml` and all five `generate_smartart_*` producers; and
* `diagram::edit::{Outline::read,set_text,drawing_text,set_drawing_reference}`.

There are no limit-less twins or compatibility aliases. Writers preflight the
selected output budget before reserving their sink or replacement buffer, and
the source-preserving editors retain bounded XML-node/depth, point/connection,
and text scans, including nested text-body and drawing events. Typed
`LimitKind` values carry the observed value and active maximum for XML bytes,
nodes/depth, data points/connections, semantic nodes/depth/text, and drawing
bytes. SmartArt validation bounds its work stack before output generation and
generators reserve from the checked output estimate.

## Consumers

DOCX and PPTX inventories, opened SmartArt edits, the DOCX writer, operations
SmartArt leaves, examples, tests, and the unified SmartArt performance harness
pass `diagram::Limits` explicitly (default policies are selected at callers
that do not retain a narrower package policy). The authoring recognition path
also carries the selected policy through `Spec::{from_parts,recognize}`.

No fixture or performance artifact is added by this source migration. The
focused qualification used one build job, offline locked dependencies, the
repository memory-capped cargo shim, and an external target below 20 GB. The
DrawingML all-feature library/integration run passed 240 unit tests and all
integration suites. All-feature library checks passed for `litchi-docx`,
`litchi-pptx`, and `litchi-ops`; the four changed `litchi` SmartArt examples,
the DOCX residue benchmark example, and the standalone unified SmartArt
performance harness also compiled. `rustfmt --check` and `git diff --check`
passed. No fixture, binary, or performance artifact was added; matched ABBA
measurement remains coordinator work.
