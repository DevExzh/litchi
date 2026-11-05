# 0361c — Complete the iWork substrate error declarations

- Date: 2026-10-21.
- Base: `4e7dadaf1b`.
- Contract: 0351 §2.4, final paragraph; ADR 0034 §6.
- Status: error declaration slice validated; integrated acceptance remains pending.

The follow-up to 0361b found four more public `Error` enums that permitted
exhaustive downstream matching: the physical IWA substrate, archive ingress,
text appearance and legacy IWA adapter. Each now has `#[non_exhaustive]`.
Their existing variants, payloads, messages and wire representations are
unchanged.

Every affected caller migrates with the declarations. The format adapters
keep their explicit resource, allocation, I/O, source-change and refusal
mappings. Terminal groups of invalid-source or invalid-format variants now
also cover future variants. Detection retains unknown physical errors in its
existing typed wrapper; its archive fallback uses the fixed message
`iWork archive ingress rejected the source`. The soundtrack work-budget
adapter still treats the two decompression limits specially and delegates
every other error to its existing mapper.

An independent source review found no changed mapping for a current variant
and no new authored-content exposure. This change affects downstream
exhaustiveness requirements and future fallback paths; it changes no
currently reachable codec, serialization or allocation algorithm. No native
performance claim is made.

## Validation

Workspace all-target/all-feature compilation passes. Two intermediate attempts
identified remaining exhaustive consumers; those callers are migrated here.
All-target/all-feature Clippy with warnings denied passes for the eight
affected crates. The physical, archive, text and detection suites pass all
413 tests across 15 targets, with no ignored tests. Workspace formatting,
fresh warning-denied rustdoc JSON for the four declaration crates, the `iwa`
glossary profile and both independently locked tool manifests pass.
`native-resave` still reports its four existing unused-import/dead-code
warnings.

The fresh [inventory](results/0361c/inventory.tsv) has no `exhaustive_error`
rows. It retains 45 `with_limits` rows in the physical and archive layers;
these remain required migrations to configured owners or explicit raw limits.
The remaining `EntryEdit` row is a naming-heuristic false positive: this
`Copy` raw descriptor contains only a borrowed member name and replacement
bytes, with a constructor and getters. It owns no package, stages no mutable
state and has no commit lifecycle. The bounded physical reassembler consumes
these descriptors, like the raw descriptors distinguished in 0351 and 0361b.
The row remains in the evidence for review.

The [validation receipt](results/0361c/validation.json) records commands,
counts, log sizes and hashes. Full logs and rustdoc JSON stay outside the
committed evidence. The coordinator's full-gate `toolunit` command now also
includes `tools.test_check_api_glossary`, covering the staged-output checks
and error-inventory regression introduced in 0362e and 0361b.

The pending W4/W4b/W5/W8 merges, other raw/shared API work, performance
repairs and the final integrated gate remain required. Nothing is pushed.
