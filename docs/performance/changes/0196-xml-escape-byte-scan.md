# Change 0196: XML escape byte-scan replaces Aho-Corasick automata

Date: 2026-08-18

## Decision

Replace the two Aho-Corasick automata behind
`litchi_core::xml::{escape_xml, unescape_xml}` with plain left-to-right byte
scans, and drop the `aho-corasick` dependency from `litchi-core` (its only
consumer in the workspace; the crate remains in the lockfile as a transitive
dependency of `regex`).

Both rewrites are exactly equivalent to the automaton semantics:

- `escape_xml` patterns are five distinct single ASCII bytes (`& < > " '`).
  No two patterns can match at the same position, no pattern byte can appear
  inside a multi-byte UTF-8 sequence, and replacement text is never
  rescanned, so a byte loop with a five-way `match` produces byte-identical
  output for every input.
- `unescape_xml` patterns (`&amp; &lt; &gt; &quot; &apos;`) share only the
  leading `&`, and no pattern is a prefix of another, so at any position at
  most one pattern can match: leftmost-longest degenerates to "match the
  single applicable entity at each `&`, otherwise copy the `&`", which is
  what the scan does. Output is built from input slices only, so replaced
  text is never rescanned (`&amp;lt;` still decodes to the literal `&lt;`).

The signatures, escaping table, and error behavior (infallible) are
unchanged; every caller in every format crate sees byte-identical output.
Profiling of the ODS source-backed commit path attributed 3.47% of
commit-phase samples to the Aho-Corasick DFA inside `escape_xml`, reached
through `write_rows_bounded`; the DFA machinery (state-ID special-case
checks, dense transitions) is disproportionate for five one-byte needles on
short cell texts.

## Mechanism and invariants

Both functions keep a no-match fast path that returns `s.to_string()`
without allocating an output buffer piecemeal. Slicing stays on char
boundaries because every skipped or matched byte is ASCII. New unit tests
pin: every special byte, doubled ampersands, embedded multi-byte UTF-8,
literal-entity input to the escaper (`&lt;` becomes `&amp;lt;`), incomplete
and unknown entities, adjacent entities, the no-rescan rule
(`&amp;amp;` → `&amp;`), and escape/unescape round trips. The existing
doc examples continue to pass unchanged.

## Matched release timing

Two frozen release binaries differ only in `litchi-core/src/xml/escape.rs`
and the litchi-core manifest; both contain changes 0193-0195 as baseline and
the identical 341-case selector matrix. Control SHA-256
`4df50556e468396d1f1ff6f4a89d5b17459976f349bf240a70e52be7fe428bc3`,
candidate SHA-256
`db708afa17eddc7dab9911429c51d6d0cd676550f8c33f4458893f2ea1201cff`.
Fresh CPU-2-pinned processes ran
`A1 control, B1 candidate, B2 candidate, A2 control`, 30 warmups and 500
retained samples per leg over the three existing ODS source-backed edit
selectors. The predeclared p50/mean/p95/p99 drift ceilings are
5%/5%/10%/15%; a statistic is accepted only when both paired directions are
lower and both drifts pass its ceiling. Every leg reports all embedded
verification flags true.

### Repeated-edit selector (`ods_source_backed_repeated_edit`)

All sixteen statistics (total/stage/commit/publication x p50/mean/p95/p99)
are withheld as neutral: every paired-direction pair straddles zero within
4.2% and no drift ceiling failed except where noted in the summary. The
per-commit escaping volume here (four one-row splices) is too small for the
mechanism to clear noise.

### One-edit guardrail (`ods_source_backed_one_edit_save`)

All eight lifecycle/commit statistics are withheld as neutral (directions
straddle zero within 3.0%; the commit p99 pair at +8.39%/-8.70% straddles
with healthy drifts). No regression trigger fired.

### One-percent guardrail (`ods_source_backed_one_percent_edit_save`)

The escape-heavier workload (21 written cells per commit) accepts:

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Control drift | Candidate drift | Decision |
|---|---:|---:|---:|---:|---|
| lifecycle p50 | 0.67% | 1.95% | 2.18% | 0.86% | accept |
| lifecycle mean | 0.44% | 3.02% | 3.32% | 0.64% | accept |
| commit p50 | 1.53% | 3.48% | 2.11% | 0.10% | accept |
| commit mean | 1.73% | 4.87% | 3.06% | -0.23% | accept |
| commit p99 | 8.62% | 20.15% | 1.06% | -11.70% | accept |

Lifecycle p95/p99 and commit p95 are withheld on control
same-implementation drift (13.57%-20.48% against 10%/15% ceilings), not on
direction. No withheld statistic shows a regression pattern.

The claim is scoped to the measured ODS selectors; escape-heavy paths in
other format families share the mechanism but are not re-measured here.

## Verification

```text
cargo test --locked -p litchi-core --all-targets
cargo test --locked -p litchi-core --doc
cargo test --locked -p <each of the 14 escape_xml consumer crates> --all-features --lib --tests
cargo clippy --locked -p litchi-core --lib --tests -- -D warnings
RUSTDOCFLAGS='-D warnings' cargo doc --locked -p litchi-core --no-deps
cargo fmt --all -- --check
python3 tools/check_crate_boundaries.py
```

The full-workspace gate could not complete in this environment (the
all-features debug build exceeds the available disk); the consumer-scoped
gate above covers every crate whose code calls the changed functions, and
the facade compile fix was verified with
`cargo check --locked -p litchi --all-features --lib`.

litchi-core passes 172/172 lib and integration tests plus 44 doctests,
including the new boundary/rescan tests. A differential harness compared
both functions against the original Aho-Corasick automata over 2,000,000
seeded randomized inputs drawn from an entity-heavy alphabet: byte-identical
output on every case. All fourteen workspace crates that call
`escape_xml`/`unescape_xml` pass their `--all-features --lib --tests` suites
(1204 tests) with two pre-existing, unrelated exceptions that reproduce
identically with this change reverted:

- `litchi-docx` `source_backed_paragraph_copy::
  partial_sinks_complete_and_write_zero_fails_without_false_progress`
  expects an `OpcError::ZipError` variant where current HEAD code surfaces
  `OpcError::IoError(WriteZero)`; the write-zero refusal itself still fires.
  Proven pre-existing by rerunning with this change reverse-applied.
- The full-workspace `--all-features` build of the `litchi` facade failed at
  HEAD before this change: `Slide::Keynote` gained a `name` field without
  updating the constructor in `crates/litchi/src/presentation/prs.rs`. This
  change's working tree carries the one-line fix
  (`name: slide.name().map(str::to_owned)`), which is independent of the
  escape mechanism and required for any all-features facade build.

Scoped strict Clippy, rustdoc, formatting, and crate-boundary checks pass.
`cargo sort` is not installed in this environment; the manifest edit removes
one alphabetically-sorted line and cannot disturb ordering.

Artifacts:

- repeated-edit: [summary](../results/ods-repeated-edit-0196-summary.json),
  [manifest](../results/ods-repeated-edit-0196-manifest.json)
- one-edit guardrail: [summary](../results/ods-one-edit-save-0196-summary.json),
  [manifest](../results/ods-one-edit-save-0196-manifest.json)
- one-percent guardrail:
  [summary](../results/ods-one-percent-edit-save-0196-summary.json),
  [manifest](../results/ods-one-percent-edit-save-0196-manifest.json)
- compressed A1/B1/B2/A2 raw reports listed in each manifest
