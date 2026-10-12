# 0225 — Native RTF Builder and unified creation

`litchi_rtf::Builder` owns a fresh plain-text body, a checked source profile and
an explicit character set. `build()` returns the immutable `Package`;
`Package::new()` uses the same empty-document path. Encoding delegates to the
native writer, preserving its escaping, Unicode and paragraph semantics.
Builders have no byte, stream or filesystem publication methods.

Text admission precedes its copy. Encoded output is bounded before allocation,
with geometric buffer growth capped by the selected source-byte ceiling. The
complete artifact is strictly reopened under the same profile, including
non-byte dimensions. Failures retain typed native errors. The builder Debug
implementation reports text length without disclosing the authored body.

Unified `create(Rtf)` directly adopts the native owner. Template, Strict and
initial-metadata options return explicit capability refusals and appear in the
capability gaps. Creation is reported independently of the pending CRUD and
save adapters; existing read behavior is retained.

## Verification

Native regressions cover exact text escaping, charset declaration, pre-copy
admission, encoded output bounds, amortized bounded growth, zero-event refusal,
immutable editing, deterministic empty creation and the builder publication
boundary. Facade regressions cover creation, option/capability agreement and
typed error classification. The capability invariant was extended to account
for the new RTF creation gaps.

Final qualification:

- Native suite: 1,441 tests/doctests; all eight builder tests passed again after
  the final I/O error-construction lint correction.
- Facade all-features suite: 1,613 tests/doctests.
- RTF-only creation and opening-policy tests: eight passed.
- All-target Clippy and rustdoc for both crates: warnings denied, passed.
- RTF glossary JSON: 44 banned-name checks passed.
- Crate boundaries: 66 packages; OfficeCLI capability delta unchanged.
- Glossary and OfficeCLI checker unit suites: 36 tests passed.

Native builder commits were integrated as `ab4cf11131` and `597848a912`;
this record accompanies the final integration and qualification commit.
Cargo used an external copied lock with cached cc 1.6.0. Repository lockfiles
remain unchanged. The owned 0225 worktree and branch were removed; external
build artifacts, copied lock, wrapper and logs are removed after qualification.
No compiled artifacts or large evidence are committed.

## Remaining work

Atomic publication currently lives in other format families. RTF native save
requires a format-neutral publisher with the same pre-publication safety and
committed-failure contract; it must not introduce a plain write masquerading as
atomic save. Native and unified filesystem publication remain pending here.

Semantic profiles, focused transaction vocabulary, unified RTF CRUD and broader
authoring remain required. Other tier-3 formats, Python/CLI completion and full
performance/corpus acceptance remain part of the active goal. This wave makes
no runtime performance acceptance claim.
