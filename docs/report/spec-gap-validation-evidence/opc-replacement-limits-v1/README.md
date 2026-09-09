# Explicit OPC replacement limits

Content-type and relationship source-token replacement can now use caller-selected read limits for both the current source and replacement. Existing entry points delegate to default limits. Changed signed sources require an explicit signature policy, stale snapshots refuse before mutation, and exact source bindings remain checked. This provides a prerequisite for owners that read and later restore XML under their own bounded policies.

Rust 1.95.0 validation passed 544 unit/integration tests and five doctests, with one existing integration test ignored, strict all-target/all-feature Clippy, changed-file formatting, and whitespace checks. New tests demonstrate over-limit replacement refusal is atomic and exact-limit restoration succeeds. The dedicated temporary directory avoided unrelated filesystem quota failures encountered in an earlier agent run.

The receipt binds the two reviewed source files and compressed root validation logs. This batch makes no performance or native-producer compatibility claim.
