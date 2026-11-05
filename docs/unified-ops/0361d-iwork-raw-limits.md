# 0361d — Required limits on canonical IWA raw operations

- Date: 2026-10-21.
- Base: `e63a1fe3534bcaaae2ce367339f7f8275877baef`.
- Contract: 0351 §2; follow-up to 0361c.
- Status: implementation and repaired native comparison accepted; integrated acceptance pending.

The physical IWA records and archive catalogs expose one canonical method for
parsing, validation, mutation and serialization. Each operation requires its
limits argument. The default wrappers and `*_with_limits` spellings are
removed, including the policy-plus-limits reference-inspection spelling.
`litchi_iwa_core::Limits` is the sole physical archive profile; the public
`ArchiveLimits` and `ResourceLimits` aliases are removed from the substrate
and its re-exports. Every affected production, test, example and fuzz consumer
migrates in this commit. Existing explicit profiles remain explicit; callers
of former default wrappers select their profile at the call site.

Catalog byte-copy and sink output also take required limits, rejecting an
oversized result before allocating or writing. Canonical object-framing
validation takes limits and checks source, object, header and message bounds
during its existing traversal. It retains the canonical-prefix, contiguous
range and metadata/payload agreement checks.

Snappy decoding, decoded-length inspection, compression and output-bound
calculation require `SnappyLimits`. Compression honors decoded bytes, chunk
size, frame count, compressed chunk size and the actual compressed stream
size. The bound calculation uses checked multiplication and a remainder,
without a loop over frames. A compressed stream ceiling may be smaller than
one maximum-size frame: empty output remains legal, and compressible input
is judged by its actual encoded size. Default framing and encoded bytes are
preserved.

The physical package profile carries its entry ceiling into Snappy's encoded
stream ceiling. Derived aggregate decoder profiles retain that ceiling.
Compressed-input refusals keep the archive's `EntryBytes` classification;
decoded aggregate refusals retain their existing component/aggregate rules.
The legacy package's archive replacement and insertion use its configured
profile before publication, preserving the original owner on failure.

The first full test pass found an exact-budget caption creation regression:
a conservative 1,382-byte output estimate was compared with an 856-byte
ceiling even though the actual encoded result fitted. The affected explicit
physical preflights and encoders now use the same profile. Output reservations
and encoder frame scratch remain separate: capped output bounds cannot stand
in for the temporary raw encoder allocation. The value-axis planner also
uses the configured chunk size and counts its emitted frames.

The glossary adds physical-core and archive profiles, bans the removed names,
and checks immutable catalog/source projections. The crate-boundary guards
and their regression tests use the canonical raw methods.

## Validation and measurement

The final seven-crate pass ran 220 test targets through 138 bounded commands:
5,088 passed, none failed and seven existing tests were ignored. The first
pass's caption failure remains in the evidence; its exact-budget creation and
patch replay now pass. Final workspace all-target/all-feature compilation,
seven-crate Clippy with warnings denied, formatting and three fresh
warning-denied rustdoc builds pass. All three glossary profiles pass.
After the private Snappy preflight/reservation repair, both substrate
crates pass all 313 unit, integration and documentation tests across eleven
targets, plus Clippy with warnings denied. No public signature or consumer
changed after the full seven-crate pass. Exact command and artifact receipts
are in [validation.json](results/0361d/validation.json).
The direct-output encoder repair then passes 314 tests across those same
targets and warnings-denied Clippy. Its first compile attempt exposed a
duplicate test import; that failed receipt is retained alongside the repair.
The 1,170 tooling tests and crate-boundary audit also pass. The separate fuzz
workspace exposed additional arity migrations and compiles after those
callers were updated; both independently locked tool manifests compile.

The fresh [inventory](results/0361d/inventory.tsv) removes all 45 physical/raw
`with_limits` rows. Its sole remaining row is the previously reviewed
`EntryEdit` descriptor: a borrowed member name and replacement bytes, with no
owner, staged mutation or commit lifecycle. It is a raw reassembly input,
not an edit family. The manual constructor and owner-policy work below
remains required despite being outside this heuristic's categories.

The native comparison uses the exact base above and an isolated snapshot of
the candidate sources, identical locks and a shared probe with only API-arity
adapters. It covers small, large and incompressible Snappy data, the output
bound, canonical framing, exact catalog copies and catalog sink writes.
Allocation measurements use separate instrumented binaries. The final
comparison and its longer large-copy qualification pass the coordinator
review in [0361e](0361e-iwork-raw-performance-repair.md); all earlier rejected
attempts remain in the evidence.

The [first native receipt](results/0361d/native-v1.json) preserves all five
paired blocks for each of twelve routes. Every measured leg exceeds 431 ms;
encoded bytes and allocation counts agree. It is rejected: canonical framing
costs 1.073× base, and small catalog copy/write costs 1.599×/1.588×. The bound
calculation is 0.175× base, but that improvement does not waive the regressions.

The repair removes repeated profile validation from catalog output: the
profile's private fields and validated constructors already establish its
invariants, while output size is still admitted before allocation or writing.
Snappy decoding retains its complete allocation-free framing/budget preflight
and raw payload validation, then uses the admitted ranges without repeating
every bound check. It initializes each frame immediately before decoding.
Canonical framing uses the final varint group to detect overlong prefixes;
its admitted message count and `u32` lengths prove that the payload sum fits
`u64`. The output-bound calculation keeps constant-divisor arithmetic for the
ordinary 64-KiB chunk path. Independent source review found no removed source,
payload, canonicality or resource check. Further paired measurements exposed
a small-compression regression and prompted a direct-output encoder repair.
The final twelve-route comparison preserves encoded bytes and reduces encoder
allocations: small compression is 0.951× base and incompressible compression
0.682×. The longer large-copy qualification has a 1.002× median with mixed
blocks. The full comparison, allocation counts and review limits are recorded
in [0361e](0361e-iwork-raw-performance-repair.md).

## Required remaining work

This record does not close the whole raw/shared API audit. The raw metadata
constructors `FieldPath::new`/`From<Vec<_>>`, `FieldInfo::new`,
`ArchiveInfo::new` and `MessageInfo::new` still need explicit limits and
fallible allocation with every consumer migrated together. Their public raw
fields remain subject to publication-time validation; identifier-zero
construction must retain its existing later validation classification.

The remaining package writers that explicitly select default Snappy profiles
need an owner-policy propagation audit, including separate encoder scratch
accounting. Detection, legacy package/document wrappers and text-wire
`*_with_limits` methods remain required migrations. These are outstanding
engineering work, not accepted exceptions. W4/W4b/W5/W8 integration, the
other raw/shared migrations and the final integrated gate also remain
required. Nothing is pushed.
