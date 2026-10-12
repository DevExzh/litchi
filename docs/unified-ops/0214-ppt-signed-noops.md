# 0214 — Signed PPT primary transactions

ADR 0034 requires exact no-op edits to preserve signed artifacts, while changed
writes refuse with `Error::Signed`. The PPT primary transaction previously
rejected signed input during initial edit admission, before determining whether
an operation changed anything.

The canonical primary error now exposes `Signed`, with content-free diagnostics.
The operations facade retains that typed cause as `Source::PptEdit` and maps it
to `DocumentProtected`. Limits, patch conflicts, missing targets and native
closure refusals retain their respective error categories.

Native snapshots inspect signature stream/storage names through their existing
CFB index with an aggregate directory-entry bound. Signed snapshots may enter
read-only planning, including nested signature storage. Same-position moves,
same-value text/anchor/visibility/advance/transition/media operations and empty
text batches preserve the complete source bytes. Changed requests refuse before
calling the mutation encoder; ordinary and profiled commits also enforce the
policy at the final publication boundary. Lower-level protected-container
mutation guards remain intact.

Read-only persisted-record and shape inspection retain the selected source
limits and leniency. Durable semantic replay uses the same signed/no-op split;
exact canonical no-op patch replay and inverse preserve the original owner.
Source freshness fences and failure-atomic staging remain in force. No signature
is removed or regenerated implicitly. Tests use opaque signature markers to
exercise preservation policy; they do not assert cryptographic verification.

## Verification

The native all-feature suite passes 1,399 tests/doctests (11 existing ignored
cases), including seven new canonical signed-policy regressions and a native
durable replay regression. A subsequent repaired-chain regression brings the
final focused canonical suite to eight passing tests, under both all features
and no default features. The operations facade suite passes 1,559
tests/doctests; its final focused typed-classification regression also passes.

All-target, all-feature Clippy passes with warnings denied. Warning-denied
rustdoc, PPT-only facade compilation, the PPT-minimal glossary profile (35
banned names and one transaction family), crate boundaries (66 packages, 268
dependencies, 11 existing debt entries), and 15 parity-tool tests plus generated
parity checking pass. The PPT-only check exposed and fixed a DOC-only gate on
the shared patch-composition error classifier.

Validation uses an external target directory and a copied external lockfile
with the cached dependency resolution; both repository lockfiles retain their
previous hashes. No performance improvement is claimed. Temporary validation
artifacts are removed after recording the results.

## Remaining scope

Unified PPT navigation and mutation dispatch, broader native signature verbs,
performance qualification and the full all-format/Python/CLI goal remain open.
