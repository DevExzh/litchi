# 0362ut — ODG joined plans stage through the primary edit

This slice starts from the W8 ODF runtime line and is implemented by source
commit `253fae9d4513cae98fb754e48c6b20c5f4af648a`.

## API and ordering

ODG no longer publishes a `JoinedEdits` through owner methods. The old
`Package::apply_joined*` routes are removed. A caller creates the bounded plan
with `Package::joined_edits(limits)`, prepares disjoint durable sub-edits, and
stages that plan with `Edit::join(joined)`. The ordinary `Edit::commit()` then
returns the existing ODG commit triad: the new `Package`, its reversible
`Patch`, and content-free diagnostics. Replay remains `Package::apply(&Patch)`.

`Edit::join` is admission-only. It checks content-matching lineage, validates
all durable payloads, checks every operation's source precondition against one
fingerprint, rejects package replacement mixed with semantic effects, and
checks conflicts before moving the plan. A rejected plan is returned by the
recoverable `JoinError`, so the edit remains unchanged and the plan can be
retried or routed elsewhere. One pending plan is retained rather than merging
independently bounded plans: its original `CompositionLimits`, unique IDs and
stable identifier order remain authoritative, and a second plan is rejected
without being consumed.

The commit order is explicit: ordinary staged operations run first, followed by
the joined sub-edits in core's deterministic identifier order and each durable
payload's operation order. Effect-key conflicts are refused before admission,
including package/resource/structure exclusivity, page-name uniqueness, and
presentation style dependencies. Distinct page transitions whose source
styles differ remain composable; shared-style transitions and page-style
retargets are refused. The joined route stages semantic operations into the
same edit and performs one rebuild/readback. Package replacement reopens the
replacement once and publishes one direct patch; it never loops through the
standalone per-operation `apply_durable_patch` publication route.

All ODG joined consumers and transition/semantic/capability tests were migrated
in the source commit. Structural and semantic joined commits cover exact
replay, inverse replay, no intermediate publication, policies, stale lineage,
payload/source mismatch, deterministic order, limits and recoverable atomic
refusal.

The source commit also registers direct `litchi_odg::Package` glossary bans for
`apply_joined`, `apply_joined_with_policies`, and
`apply_joined_with_security_policy`; the old bans were not left scoped only to
the read-only `Drawing` projection.

## Validation

The source commit above was validated with the memory-capped cargo shim, one
build job, disabled debug information, `CARGO_INCREMENTAL=0`, and
`TMPDIR=/tmp/claude-1001/odg-joined`. The isolated target was
`(deleted-coord-artifact)`, measured at
413 MiB after the run. These external receipts are intentionally not tracked:

| Command | Receipt | Size | SHA-256 |
|---|---|---:|---|
| `cargo check --offline --locked --all-targets --all-features -p litchi-odg` | `(deleted-coord-artifact)/check-final.log` | 162 bytes | `fae313f565339b7b03e998bd01cdb4767541cdc4d23e63c4190aa48ee1dcecef` |
| strict `cargo clippy --offline --locked --all-targets --all-features -p litchi-odg -- -D warnings` | `(deleted-coord-artifact)/clippy-final.log` | 162 bytes | `ead0e57dd22a0aaa12e2e4065904af2b38851392f96181ca20d6b5c11e711396` |
| focused `cargo test --offline --locked --all-features -p litchi-odg --test package_semantics --test transition_review --test capability_transactions` | `(deleted-coord-artifact)/tests-final.log` | 4368 bytes | `c52fa40ef7b8f3e1108afe96383bd393467e66ba88723f44fb84b2f42b2bebb8` |

The focused run passed 26 package-semantic tests, 11 transition tests and 8
capability tests. Rustfmt, `git diff --check`, the API glossary JSON parse and
the repository artifact guards passed. No benchmark, fixture, binary or large
repository evidence artifact was added.
