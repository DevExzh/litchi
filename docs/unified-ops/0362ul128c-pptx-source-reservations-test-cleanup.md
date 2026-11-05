# 0362ul128c — PPTX source reservation test ownership cleanup

Date: 2026-10-25

The 128b real-consumer regression now respects the existing edit ownership:
`commit_checked()` consumes the edit, so its invalid post-commit `drop(edit)`
was removed. The nearby source tests were audited for the same moved-value
pattern; no production code or API behavior changed.

Validation is source-only: rustfmt and `git diff --check`; no Cargo, build,
rustdoc, metadata, or runtime command was run.
