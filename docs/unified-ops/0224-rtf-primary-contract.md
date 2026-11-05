# 0224 — RTF primary transaction contract

Primary RTF commits expose `package`, `patch`, `diagnostics`, `into_package`
and the consuming `(Package, Patch)` pair `into_parts`. The old `snapshot` and
`into_snapshot` names are removed. Primary replay is `Package::apply(&patch)`;
the patch implementation's apply helper is crate-private. `Patch::is_noop`
compares the exact artifact, and commit changed diagnostics derive from the
same authoritative bytes rather than a caller-supplied flag.

Previously, in-memory replay accepted equal source bytes and returned the
patch's cached target owner, bypassing a recipient's stricter ingress profile.
Changed replay now reparses the target under the recipient's retained options
when policies differ. Exact no-op replay returns the recipient snapshot itself,
preserving its owner identity and options. An identical policy may reuse the
already admitted immutable target. Source conflicts remain typed refusals.

This wave concerns the primary `edit` family. Specialized tail-append and
redaction families retain their separate current contracts and are not claimed
canonical here.

## Verification

The two recipient-policy regressions were run with the previous replay body
restored temporarily: both fail (lost snapshot identity and accepted oversized
target). The corrected body passes all three canonical primary regressions.
The candidate was restored before full validation.

Consumer commits: `d2f4830bd9` (picture publication), `5258aa1460` (tests),
`9eb2c42516` (harness and docs). Final validation:

- `cargo test -p litchi-rtf`: 1,430 tests/doctests.
- `cargo test -p litchi-ops --all-features`: 1,610 tests/doctests.
- All-target warning-denied Clippy for both crates.
- Native-resave RTF and benchmark library compile checks.
- Crate boundaries: 66 packages; OfficeCLI delta and its 15 Python tests.
- Glossary checker unit suite: 21 tests.

Warning-denied rustdoc passes for both crates. The RTF rustdoc JSON profile
passes 41 banned-name checks, including primary commit names and projection
lifecycle restrictions. Finite semantic profiles remain outside this claim.

Cargo used external copied locks (root validation used cached cc 1.6.0,
consumer checks kept their lock contents). Repository lockfiles are unchanged.
Temporary 0224 worktree, branch, build target, copied locks and logs are removed
after validation. No compiled artifacts or large evidence are committed.

## Remaining work

RTF semantic profiles and focused transaction vocabulary, minimal creation and
Builder, native atomic save, and unified CRUD remain required. This change does
not establish complete RTF coverage or runtime performance acceptance. The full
all-format goal, Python/CLI work and broader qualification remain active.
