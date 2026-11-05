# 0362ul108a — Preserve empty Web Extensions conformance intent

The first 0362ul108 change correctly ignored unused relationship namespace
declarations, but its no-relationship fallback lost the requested profile for
an empty graph.  Consequently an empty Transitional task-pane graph could be
accepted as a Strict signed no-op, even though the profile is part of the
edit intent.

The planner now aggregates relationship-profile evidence over the complete
task-pane and add-in graph.  Resolved relationship-valued attributes (`r:id`,
`r:embed`, and `r:link`) are authoritative whenever any graph part contains
one: the requested profile must be present and the opposite profile absent.
Only when the whole graph has no such attribute does the planner use the
already-retained root namespace scope from each parsed XML document as the
profile fallback.  Thus an empty root declaring Transitional remains an exact
Transitional signed no-op and refuses a Strict request, while an unused or
vendor-only declaration cannot change a graph with actual relationship
tokens.  The root scope is captured during the existing bounded parse; no
second source scan or eager materialization was added.

The existing mixed-profile and signed-empty-task-pane regressions remain the
focused coverage for this boundary.  This is a source-only repair.  No Cargo,
build, rustdoc, runtime, native, metadata, or performance command was run;
the coordinator must rerun the affected common and DOCX web-extension gates.
