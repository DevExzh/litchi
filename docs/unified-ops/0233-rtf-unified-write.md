# 0233 — Unified RTF paragraph text writes

RTF now supports immutable paragraph `set`, scoped query fanout, explicit edits,
atomic/best-effort batches and process-local history. The adapter resolves native
paragraph positions against one snapshot and submits the selection as one native
paragraph batch. It retains origin and native open policies through edits and
publication. The property registry declares `paragraph.text` writable;
capabilities retain explicit structural-verb and native source-closure gaps.

Target expansion, native operation count, aggregate replacement/action staging,
and retained snapshot weights are bounded. Incoming replacement strings are
copied only after aggregate admission. Native edit failures retain a typed
`Source::RtfEdit`, including protection, source limits and unsupported formatting
closures. Native patch history weights now include private change descriptors
and owned semantic payloads as well as both source artifacts, fixing undercounts
when long source text is replaced by a short value.

Single-step changed detection uses the native patch's no-op result. Multi-step
changed detection compares complete serialized artifacts and propagates failures;
flattened text equality cannot discard formatting or structural changes. Exact
no-ops retain native snapshot identity, and replay rejects foreign owners even
when their input bytes match.

Qualification exposed a native paragraph-batch preservation defect: rebuilding
flattened body text converted untouched inline line breaks to paragraph breaks.
For batches over ordinary bodies containing inline breaks, the native editor now
validates paragraph source ranges, splices selected replacements into retained
bytes and reparses the candidate with the original policy. Empty-paragraph
insertions delimit the preceding control word, including ASCII text, leading
spaces and numeric signs. Candidate text and opaque-preservation checks still
run before publication. Other native closure restrictions remain explicit
refusals; this does not claim general rich paragraph rewriting.

## Verification

Thirteen unified integration tests cover native batch byte parity, scoped
selection and zero matches, Unicode, native structural boundaries, empty
paragraphs, typed budget/protection/source-policy failures, immutable history,
exact no-ops, atomic rollback, best-effort batches, registry capabilities and
save/reopen with origin preservation. Native regressions cover inline-break
preservation, ASCII insertion/fanout and semantic history weights.

The all-features facade suite passes 1,639 tests/doctests; the native RTF suite
passes 1,465. The RTF-only facade suite passes 504 tests/doctests. Warning-denied
all-target Clippy passes for both crates with all features and for the RTF-only
facade. Warning-denied rustdoc, no-feature compilation, scoped formatting and
diff checks pass. Crate boundaries pass for 66 packages; 36 Python checker tests
pass; the generated OfficeCLI parity delta is unchanged. These Python checks
cover repository tooling, not a new Python binding milestone.

The dispatch, adapter and native preservation changes were integrated as
`5d97fdd413`, `7e00825646`, `b00c6c3e52` and `0b4cb24442`; final qualification also
fixes native range typing/composition initialization and uses typed no-op proof
for single-step changed detection. The owned agent worktree and branch were
removed. Validation used external copied lockfiles and build output; all three
repository lockfile hashes remain unchanged. Owned 0233 scratch/build artifacts
are removed after qualification. No binary or large evidence is committed.

## Remaining work

RTF add/remove/move/swap, richer nodes/properties, transaction coalescing and
paired performance qualification remain open. Multi-step content comparison
currently materializes endpoint bytes, and generic repeated sets retain separate
native commits. No runtime performance acceptance is claimed. Native source
vocabulary/publication policies and the other full-goal requirements remain
active. This is a paragraph-write milestone, not completion of the goal.
