# Validation history

The initial `harness-tests` command ran the complete harness library in the
debug profile. It reached the existing
`media_rich_odt_scalar_and_batch_resource_replacements_are_matched` test,
which constructs a 17,060,208-byte media package and exercises repeated scalar
publication of 64 resource replacements. After more than ten CPU minutes in
that case, root explicitly terminated the owned test child (PID 2458266) to
move the full performance-harness suite to the release profile. The command
returned 101 with unchanged source custody. Its partial log and failed receipt
are retained; it is not a passing suite or an inferred test failure.

This was an intentional execution-profile correction after inspecting the
workload, not a restart caused by a polling timeout. The same live session was
polled until its terminal result. Focused streaming/oracle debug checks and the
complete release library suite are the replacement validation gates.

Independent review also found the initial XLSX expected-cell oracle did not
reject cells outside A:D or extra package members. The capture will use the
strengthened all-stored-cell and exact-member oracle and explicit mutation
regressions. These additions affect untimed correctness checks only.

The first focused debug compile found a missing parent-module qualification
for `ArchiveReader` in the new mutation test. The corrected test uses
`super::ArchiveReader` and rebuilds the same six members with an injected E1
cell, then requires the coordinate-specific refusal. Its extra-member case
requires the member-set-specific refusal. The failed compile remains retained.

The initial complete standalone formatting check reported one batch-owned
`lib.rs` wrapping change and preexisting formatting in five other files.
The owned change was formatted with Rust 1.98.1. `formatting-debt.json` binds
the other files and proves they were unchanged from the starting revision.

The first `derive` invocation imported `verify_report`, while the retained file
is named `verify-report.py`. The corrected driver uses an explicit file-based
module import. `derive-v2` passed without changing any report or capture.

The staged-file whitespace check caught one extra trailing blank line in
`next-work.md`; it was removed before the final seal and commit.
