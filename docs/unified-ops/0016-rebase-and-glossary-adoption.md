# 0016 — Rebase and API glossary adoption

Date: 2026-10-08. Status: kickoff complete; W0 enforcement and API waves pending.

## Owner direction and scope

The owner asked to set up and achieve `docs/UNIFIED_OPS_GOAL.md`, refactoring
API shapes to the glossary/default-policy ADR in the base checkout, beginning
with a rebase. The persistent goal is active. This record starts that work;
it does not declare the overall goal or a new milestone complete.

The existing accepted ADR 0033 is the unified-operations design (acceptance
record 0006). The base's untracked glossary draft reused that number. Preserve
the established identity and historical citations: adopt the glossary as
[ADR 0034](../adr/0034-unified-api-glossary-and-default-policies.md), with a
pointer at the requested draft path. The owner's explicit instruction to
refactor according to it authorizes this branch's adoption. ADR 0004 now
points to the newer shape/default contract; the glossary is tracked.

Source provenance, read without changing the base checkout:

| Base file | Bytes | SHA-256 |
|---|---:|---|
| `docs/adr/0033-unified-api-glossary-and-default-policies.md` | 20612 | `581009b25841cbcc60afde0c4b1c9b948df5bd5faa77fffcc506aef1a597283b` |
| `docs/API_GLOSSARY.md` | 14905 | `372a81700a3fc0f359f9fd350c589c86b27bd2888c806815b378601abe925ec7` |

The imported ADR changes its number/status and adds adoption context; its
substantive design is retained. The glossary adds explicit adoption status.
Its inherited conformance labels must be verified during each wave, not used
as evidence of completion.

## Rebase and conflict resolution

- Old branch tip: `dbe66611f1`.
- New base: `32ad64e320` (`feat/office-format-completeness` at rebase start).
- Rebased branch tip before this record: `61630e5635`.
- Replayed the existing branch commits with `git rebase
  feat/office-format-completeness`. Historical merge commits were flattened;
  historical record hashes remain evidence references.
- Reviewed recorded resolutions for the incrementally assembled ops scaffold,
  signature policy, and dependency inventory.
- Combined the base's `OwnedSourceFingerprintMemo` with the branch's cheap
  `SourceBackedPackage::clone`: retain and share the optional memo instead of
  discarding it or copying cached state. Updated the exhaustive clone test and
  added a regression: fingerprint through the original, drop it, and reuse the
  same initialized memo through its clone.
- Removed obsolete family-detection helpers with M0a S9 and retained the base's
  `TransferDelayPolicy` instrumentation import while deleting `Workbook`.
- Verified the base is an ancestor and that no old hash from the 2026-10-05
  rewrite map is reachable from the rebased tip.

## Rollout reconciliation

M0a and M0b already landed. The glossary draft's instruction to unify leaves
before M0a cannot apply retroactively. Each leaf wave must now migrate the
existing `litchi-ops`, facade, Python, benchmark, test and example consumers
in the same buildable change. Existing M1/M2 branches require reconciliation
before integration; their unmerged work is not claimed here.

W0 next: rustdoc-JSON public-name checker, explicit package/wave and
transaction-family registrations, focused tests, CI wiring. Start with no
banned names or registered families, as the ADR specifies. Do not globally
reject pending waves or legitimate domain IDs such as DOCX `glossary::Id`.
Then migrate DOCX, PPTX and XLSX in order, preserving the no-op, signed-write,
encryption, snapshot and resource-budget contracts. Continue remaining leaves
and the goal's format tiers and milestones; glossary work does not reduce the
all-format CRUD coverage requirement.

The goal prohibits edits to independently owned iWork crates. W1 is therefore
an owner-coordination dependency, recorded in `IWORK_API_REQUESTS.md`, not an
excuse to modify those crates. Its public API is consumed at M6. Two read-only
review agents inspected ADR constraints and adoption conflicts. The requested
`opus` implementer model is not available in this session; no implementation
was delegated to a substitute model.

## Verification

- `cargo check -p litchi-ops --all-features`: passed.
- `cargo test -p litchi-ops --all-features --lib`: 388 passed.
- `cargo check --manifest-path tools/perf-baseline/Cargo.toml --tests`: passed,
  including the facade dependency and the retained harness instrumentation.
- `cargo test -p litchi-opc --lib source_backed::shared_clone_tests`: 6 passed.
- `cargo test -p litchi-opc --lib owned_source_fingerprint_memo`: 2 passed.
- `cargo test -p litchi-opc --test source_backed_shared_clones`: 10 passed.
- `cargo clippy -p litchi-opc --lib --no-deps -- -D warnings`: passed.
- `python3 tools/check_crate_boundaries.py`: passed (66 packages, 267
  internal dependency declarations, 11 explicit debt items).
- Focused Rust formatting and `git diff --check`: passed.

The first OPC test compile caught the exhaustive test destructuring missing
the newly retained memo; the test was updated before the successful runs.
An initial combined workspace/harness check command was rejected because the
harness is a separate workspace; the manifest-path command above passed.

Full milestone gates (all feature combinations, warning-denied rustdoc,
workspace tests, fuzzing, pytest, OfficeCLI parity and paired measurements)
remain required. No throughput, memory, parity or full-milestone claim follows
from these checks. OfficeCLI behavior is unchanged by this kickoff.

## Cleanup

No new worktrees were created. Existing `uo/*` worktrees predate this session
and contain independent work; they are not disposable session artifacts.
Session-generated build targets and the generated root lockfile are removed
after verification. No base-checkout files were edited and no push was made.
