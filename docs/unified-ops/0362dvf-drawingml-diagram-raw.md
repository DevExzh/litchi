# 0362dvf — DrawingML diagram raw limits

- Status: follow-up grammar, fidelity and authoring-limit repair source and
  historical focused capped qualification are complete; ABBA measurement
  remains coordinator work. The root-lock-less qualification recorded here is
  diagnostic only; final 52ed acceptance is recorded in `0362dvfb`.
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
* `diagram::edit::{Outline::read,set_text,drawing_text,set_drawing_reference}`;
* `diagram::authoring::{Spec::{parse_outline,outline,data_xml,layout_xml,
  drawing_xml},quick_style_xml,colors_xml}`; and
* `litchi_pptx::shape::diagram::graphic_frame`.

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

## Coordinator repair pass

The follow-up pass closes the five review findings from the first packet:

* `edit::read_drawing` counts every `dsp:sp` globally (including sibling and
  empty shapes), rather than using the open-element stack as a shape budget;
* both raw readers track `root_seen`, `root_closed`, depth and stacks, reject
  duplicate or trailing roots and non-whitespace trailing markup, and refuse
  EOF while a root remains open;
* `Spec::recognize` checks both direct inputs against the selected input policy
  and regenerates through the bounded `data_xml` and `layout_xml` paths;
* authoring `Spec::{data_xml,layout_xml,drawing_xml}` and
  `{quick_style_xml,colors_xml}` now require `&Limits`, preflight their
  scaled output estimates, and retain final typed output checks; and
* PPTX `shape::diagram::graphic_frame` now uses the canonical `&Limits`
  policy, bounded relationship/output sizing, and migrated tests/examples.

## Follow-up review repair

The second review pass adds behavior coverage and closes the remaining
grammar/fidelity findings:

* `edit::read_drawing` rejects nested empty `dsp:drawing` roots and counts
  sibling shapes against the global node budget;
* `DiagramDefinition` rejects nested or duplicate roots, unterminated roots,
  and non-whitespace markup after a closed root;
* `SmartArt::parse_data_xml` requires the DrawingML data-model root, rejects
  wrong roots, plaintext and empty roots, and rebuilds the complete parent
  graph for nested points;
* the generic SmartArt data and cached-drawing generators emit presentation
  points, parent connections and cached shapes for every nested node, with a
  nested round-trip regression proving that no descendants or references are
  lost;
* `Spec::parse_outline` and `Spec::outline` enforce caller input/output
  policies before their resource-scaled buffers are reserved; and
* live DOCX/PPTX/ops consumers and the standalone SmartArt performance
  harness use the new authoring signatures, with custom graphic-frame and
  parser-boundary regressions.

The historical focused validation for this repair passed with one build job,
offline locked dependencies, the memory-capped cargo shim, and an external
target below 20 GB: all 245 `litchi-drawingml` unit tests and integration tests
passed;
all-feature library checks passed for `litchi-docx`, `litchi-pptx`, and
`litchi-ops`; and the standalone unified SmartArt harness passed its locked
check after its tracked lockfile was refreshed for the current workspace
dependencies. The earlier 8f49 qualification is not used as evidence for
these source changes. This run has no root-lock pre/post receipt or persisted
logs, so it is not 52ed acceptance evidence. The superseding `9786` source
correction passed 252 tests and the affected consumer/harness checks under the
verified 52ed root lock; see `0362dvfb`. No new fixtures, binaries, or
measurement artifacts are included.

No fixture or performance artifact is added by this source migration. The
historical focused qualification used one build job, offline locked
dependencies, the repository memory-capped cargo shim, and an external target
below 20 GB. The DrawingML all-feature library/integration run passed 245 unit
tests and all integration suites. All-feature library checks passed for
`litchi-docx`,
`litchi-pptx`, and `litchi-ops`; the four changed `litchi` SmartArt examples,
the DOCX residue benchmark example, and the standalone unified SmartArt
performance harness also compiled. `rustfmt --check` and `git diff --check`
passed. No fixture, binary, or performance artifact was added; matched ABBA
measurement remains coordinator work.
