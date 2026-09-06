# 0436 source review and validation scope

Independent source review found no release blocker in the private span encoder.
It checked UTF-8 boundaries, the 256-byte cap, per-scalar admission checks,
position-sensitive whitespace, content-before-paragraph refusal ordering and
hierarchical Work rollback. The production fragment buffer shares the paragraph
limit and cannot partially accept a span because of a tighter capacity.

The first focused test compile failed on eight redundant std::io paths forbidden
by the crate's unused-qualifications lint. Those paths were shortened, and the
focused suite passed. The deterministic ordinary-span cancellation assertion
was then tightened to opening-tag bytes plus exactly 256 bytes. The full final
ODT release suite passed 1,002 tests with none ignored, and scoped all-feature,
all-target Clippy passed with the inherited large_enum_variant allowance.

The tests independently retain scalar write chunk boundaries and compare every
threshold through 257-byte ASCII, multibyte and mixed input. They compare exact
local-limit tuples, Work resource/observed/limit/scope, accepted output prefixes,
and parent/child usage. Content tests use a nonzero prior prefix. Combined
paragraph/content and content/Work ties preserve content-first attribution.

Cancellation checks are cooperative: an already charged span may finish copying
up to 256 bytes before the next check. The deterministic scratch writers verify
pre-cancel, cancellation during opening-tag short writes and after one complete
ordinary span. They do not assert scalar-equivalent asynchronous progress or
arbitrary scratch-writer partial-error Work equality. Existing public integration
tests exercise actual ZIP sink failures and cancellation. No new public API,
shared budget semantics, dependencies, unsafe code or concurrency was introduced.

This batch reuses the unchanged harness and common package layer. Their complete
0435 gates remain prior evidence; the current batch reruns all ODT tests and
rebuilds the standalone harness for the changed provider. Native fixture tests
are functional coverage, not a fresh native Office application run. There is no
new native interoperability, cold/range-source, parallel scaling or copy-count
claim from this private serialization change.
