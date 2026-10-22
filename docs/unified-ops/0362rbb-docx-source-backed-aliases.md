# 0362rbb — DOCX source-backed canonical module paths

This follow-up removes redundant root-level names from the source-backed DOCX
API. Story text records are now reached through `source_backed::story_text`,
one-paragraph records through `source_backed::tail_append`, and stream records
through `source_backed::tail_append_stream`. The matching tests, SVG lifecycle
consumer, example, source-read-policy consumer, and stream fuzz target use
those canonical modules. The glossary profiles no longer register the removed
root aliases.

The pure `Package::apply_tail_append_patch` forwarder was removed with the
alias family. `Package::apply_tail_append_stream_patch` remains the durable
stream replay entry point. Its resolver, source and artifact proofs, exact
inverse authorization, cancellation, bounded replay, and full-artifact tests
are unchanged; this cleanup does not remove durable stream capability.

## Validation

The affected integration tests passed: 24 SVG lifecycle, 19 glossary-story,
25 secondary-story, 14 story, 40 tail-append, 34 durable
tail-append-stream, and one source-read-policy test (157 tests total). The
final log is outside
the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/w8-docx-story-text/alias-tests-v5.log`,
14,036 bytes, SHA-256
`f430556c3de7df59cfc8da86227fecc138788ba15e693808223936c3a1e41f61`.

All-target/all-feature `litchi-docx` Clippy passed with `-D warnings`; the log
is at
`/home/zhuhe/code/litchi-worktrees/.coord/w8-docx-story-text/alias-clippy-v2.log`,
171 bytes, SHA-256
`d3499a29a90e7d18648ad1c066d2ba8a71d8d3a4c3ec9305c6a618e5851fa5aa`.
`cargo fmt --all -- --check` and `git diff --check` also passed. No binary or
large evidence file is committed.
