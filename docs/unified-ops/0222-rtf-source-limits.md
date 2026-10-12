# 0222 — RTF source limits and open policy

RTF ingress now uses checked `SourceLimits` and `OpenOptions`. The old
`ParseLimits`, `ParseOptions` and `read::{Limits, Options}` names are removed.
The builder has fallible consuming setters and a fallible `build`; dimensions
are input bytes, events, binary bytes, aggregate binary bytes, decompressed
bytes, opaque nodes, opaque node bytes and aggregate opaque bytes. Input defaults
to 512 MiB under ADR 0034; other existing finite defaults are preserved. Zero
means explicit denial. Every dimension rejects values above `isize::MAX` with
`Error::InvalidLimit { kind, value, maximum }`.

Ingress violations use `Error::Limit { kind, observed, maximum }` across the
lexer, parser, binary/opaque retention, file reads and compressed expansion.
The selected profile and explicit leniency survive immutable document edits
and candidate reopening. Strict default behavior and recorded repairs remain
unchanged. Existing specialized `*_with_limits` entry points accept the new
checked profile; the extensible envelope is available through `*_with_options`.

The unified opener now forwards `OpenOptions::rtf_options` into native parsing
instead of ignoring per-format policy. Native limits apply in addition to the
facade's physical detection ceiling. Native invalid profiles classify as
`invalid_argument`; exceeded profiles classify as `limit_exceeded`.

## Verification

Native changes are committed as `b3b2fd7428`; native consumer migration as
`807cca87f4`. Validation passes:

- `cargo test -p litchi-rtf`: 1,423 tests/doctests.
- `cargo test -p litchi-ops --all-features`: 1,610 tests/doctests.
- RTF-only facade profile: five policy regressions, covering byte/path ingress,
  generic ceiling precedence, both compressed transports, explicit repair,
  zero opaque-node denial and invalid-profile classification.
- All-target warning-denied Clippy and warning-denied rustdoc for both crates.
- Rustdoc JSON glossary profile: 17 banned names; checker unit suite: 21 tests.
- Crate boundaries: 66 packages; OfficeCLI delta check and 15 Python tests.

Native tests cover all eight checked dimensions, owned/raw/file/validation
entry points and retained edit policies. The facade's prepared-package opener
is OOXML-only, so no RTF prepared-route coverage is claimed. No runtime
performance claim is made for this policy/API migration.

Cargo used an external copied lock with cached cc 1.6.0; the repository lock
is unchanged. The base remains `32ad64e320` and is already an ancestor.
Owned 0222 worktrees, branches, copied lock, build target and logs are removed
after validation; no compiled artifacts or large evidence are committed.

## Remaining work

This is the first tier-3 migration after the tier-1 and tier-2 formats acquired
working read/write loops. It does not complete RTF: the canonical package and
projection split, primary transaction vocabulary, Builder and unified CRUD
adapter remain required. Separate semantic/fixed structural `LimitExceeded`
errors and unchecked standalone `transport::DecompressionLimits` are not migrated here.
The latter is an explicit remaining codec policy gap, not covered by the
checked document-ingress claim. Broader format coverage, Python/CLI work and
performance/corpus acceptance remain open under the active full goal.
