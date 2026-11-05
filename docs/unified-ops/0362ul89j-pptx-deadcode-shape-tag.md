# 0362ul89j — PPTX shape/tag dead-code return

Date: 2026-10-25. This is a source-only dead-code repair; no compiler,
runtime, performance, or qualification result is claimed.

The repair is based on frozen source
`237d6801a641b85b4477c432a26ac53e531987ed`. The actual all-features compiler
diagnostic log is retained outside Git at
`(deleted-coord-artifact)/01-pptx-all-features-check-attempt-actual-f5-01.log`
(49,586 bytes, SHA-256
`e3004cf8f79ec32a2dd2b082f4204dc1efe5a673dab7df97dcb63f9203aa06ea`). Its
diagnostics index is 16,760 bytes with SHA-256
`775ff0196bcbcf8ec70f1bb66ee969419181664f262944dd3775f308c09f235e`.

The index contains 25 owned diagnostics. Production routes retain the
selected-limit variants throughout: classification and Designer package
operations keep their `*_with_limits` paths; slide shape adapters keep
`*_shapes_with_limits`; master-layout publication keeps the shape-limited
preflight/source functions; and tag package/shape edits keep the explicit
limit routes. Zero-consumer private default twins were removed. Default
helpers used only by unit/test seams are now `cfg(test)`: master-layout
package/placeholder helpers, tag package/shape helpers, and Designer-property
load/remove helpers. The public family facades are unchanged.

Static checks completed: `git diff --check`, source-level consumer/disposition
audit against the indexed corpus, no conflict markers, and no upstream fixture
paths or compiled-binary references in changed files. Cargo, rustdoc, and
runtime tests were not run as required by the coordinator.
