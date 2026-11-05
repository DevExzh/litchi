# 0017 — W0 API glossary gate

Date: 2026-10-08. Status: W0 complete; leaf migration waves pending.

## Change

Implements ADR 0034's governance gate in `tools/check_api_glossary.py`, with
explicit feature profiles in `tools/api_glossary.json`. Each profile builds
rustdoc JSON and checks it before another feature profile can overwrite the
output. `--json --profile` supports inspecting an existing document, with
caller responsibility for build provenance. Temporary build output lives
outside the checkout on the same disk and is removed on completion; an
explicit `--target-dir` retains the cache.

Names are exact and scoped to API paths. The public graph includes local
module/glob re-exports, renamed exports, inherent methods through type aliases,
public fields, variants, trait items and hidden public APIs. Public signature
references expose nominal and dependency type names at the exporting API.
Private siblings, documentation, argument names and generated blanket methods
are not name targets. Missing rule scopes fail instead of silently disabling
checks. Rules contain replacement guidance for diagnostics.

Explicitly registered public transaction modules must export six types:
`Edit`, `Patch`, `Commit`, `Diagnostics`, `Error`, `LimitKind`. Functions or
nested/private types with these names do not satisfy the requirement. This
is a vocabulary/shape-existence check, not proof of signatures or behavior.

The initial `core-minimal` profile runs against the current tree with empty
banned-name and family lists, as W0 specifies. The framework accepts other
workspace packages without a hardcoded crate list. No unregistered format or
feature surface is certified. Register profiles/rules with each completed
rename and update the glossary status in that same commit.

CI runs the checker and its real-rustdoc tests in the Rust-toolchain job,
before unrelated formatting gates. Glossary-only changes now trigger this
workflow. `docs/API_GLOSSARY.md` documents registration, scope, commands,
exit codes and remaining limitations.

## Verification and review findings

- `python3 -m unittest tools.test_check_api_glossary`: 19 tests passed.
  Tests compile a temporary Rust library with the installed rustdoc instead
  of relying solely on hand-written JSON shapes. Negative cases exercise
  public aliases and methods, scoped selectors vs domain IDs, fields, trait
  members, hidden items, incomplete families, malformed graphs and policies,
  missing scopes, unresolved dependency globs, and failed builds. A failed
  build cannot fall back to an existing JSON file.
- `python3 tools/check_api_glossary.py`: passed on the current `litchi-core`
  tree; reports 1 profile, 0 bans, 0 families explicitly.
- A review probe built DOCX rustdoc and applied a temporary `edit_document`
  ban: it diagnosed four paths, including root `Package::edit_document` and
  `source_backed::Package::edit_document`. The probe did not enable hidden
  items and is only a negative scanner check, not W2 conformance evidence.
- `git diff --check`: passed.

The reviewer found and the implementation corrected type-alias method scope,
dependency signature attribution, nonexistent scope handling, and malformed
known-local path/index mismatches. It also found that real rustdoc omits some
opaque/generated signature IDs from both index and paths. Their use-site
spelling is checked, but unavailable members cannot be inspected. This is
separate from a known-local path missing its item or a missing exported child,
which refuses the input. No missing opaque ID is assumed to be a local type.

A Pages probe confirmed that external public globs may lack their target
module data in rustdoc JSON. Such a profile fails closed, not as a false pass.
Dependency export-data support is required before activating such profiles.
Public family registrations name public modules; private implementation
module names are not export scopes merely because some types are re-exported.

No Rust library code or API changed in W0. Rust library build/test, Clippy,
rustdoc, fuzz, Python and benchmark milestone gates remain necessary for leaf
waves and the unified-read/write milestones; this tooling change does not
claim those gates or any new OfficeCLI parity. No performance claim is made.

## Next work and cleanup

Begin W2 DOCX owner/projection and transaction migration, with consumers and
verification in the same changes. Coordinate W1 via the existing iWork API
requests rather than editing independently owned crates. M1/M2 work already
on other branches still needs reconciliation before integration.

No worktrees were created. Rustdoc probe JSON and generated build output,
root lockfile and new Python cache files are removed after verification.
No large artifact, compiled executable or dependency snapshot is committed.
