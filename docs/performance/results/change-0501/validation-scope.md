# 0501 validation scope

The changed code is a private digest input reduction, with no new parser,
unsafe code, concurrency, dependency, or public API. Exact plan payload and
candidate readback checks remain mandatory. The two new private regression
tests independently mutate equal-length image and chart bytes and cover both
plan matching/publication refusal before output and candidate reread refusal.

The full PPTX all-target gate includes the existing
`source_backed_cross_copy_adversarial` and `source_backed_cross_copy` integration
suites, covering graph closure, source changes, typed resource refusal,
cancellation, raw package preservation, and publication behavior. The repository
has no dedicated PPTX fuzz target; the unchanged generic OPC parser fuzz target
is not represented as a new test of this private plan behavior. No fuzz-run or
new unsafe/concurrency validation claim is made for this batch.

The first new test compile found an unused import. The next execution found
that authored slide titles do not assign distinct common-slide names; the
fixture now sets those names explicitly. Clippy then identified three test-only
style issues, which were fixed. Every failed gate log and receipt is retained.
The final all-target suite executes the final test bytes after these fixes.

The staged raw Cargo logs retain terminal blank lines and raw perf exports
retain trailing alignment spaces. An unrestricted `git diff --cached --check`
therefore reports raw-output whitespace. Those hash-bound artifacts are not
normalized; the source and authored documentation whitespace check excludes
the raw evidence directory.
