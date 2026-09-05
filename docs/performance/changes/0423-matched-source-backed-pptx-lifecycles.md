# 0423: matched source-backed PPTX lifecycle baselines

The existing source-backed cross-copy benchmark started after opening and cache
warming. It could not describe the same ingress-to-publication interval as the
owned lifecycle. Two new opt-in selectors now measure public source-backed
opening, planning and publication for the same plain and eight-image slide-copy
inputs used by the owned baseline. The old phase-only selector and owned
runner boundaries remain unchanged.

The implementation is `6ca9962c7818e173538a40ed45f3a3c32cc1aa6a`. It changes
benchmark and coverage code only. No production optimization, dependency or
unsafe implementation is added. The [design and ADR review](../results/change-0423/design.md)
records the input/output contract, ownership and measurement boundaries.

Independent output checks bind each image embed to its relationship, content
type and exact payload; all other inserted slide XML bytes must remain exact.
The fixture-specific XML comparison normalizes only generated double-quoted
`r:embed` values. Exact added member/part sets, presentation relationship
cardinality, deterministic output, source stability, refusal behavior and raw
untouched ZIP preservation remain hard gates. Valid-archive mutation tests
exercise payload, relationship target, geometry and relationship count failures.

The [frozen protocol](../results/change-0423/protocol.json) uses 16 fresh,
serialized processes on CPU 2: two roles, two corpora, two repeats, with separate
100-sample normal and 30-sample allocator lanes. Each lane/corpus uses owned R1,
source R1, source R2, owned R2. Inputs, selectors and bounded sink ceilings match;
output serialization may differ by role. Input clones and sink reservation are
outside the interval. Source adapters count logical reads inside it.

The source publisher consumes its editor during publication; other caller
objects remain live at exit. The owned API retains different artifacts.
Allocator V3 records absolute process callback-order live bytes, including
entry live bytes; whole-process RSS also includes setup and correctness checks.
These measurements do not establish cross-role memory reductions or post-drop
retention. The [resource review](../results/change-0423/resource-review.md)
describes these limits and finite source-cache behavior.

The [replay bundle](../results/change-0423/README.md) retains exact source,
binary, configuration, corpus and output identities, individual samples,
verification journals, failed checks and passing reruns. Portable replay uses
four unchanged hash-pinned shared validators and needs no original build or
binary. This batch is a measurement enabler; no cross-role speedup,
prior-version improvement, physical-I/O or zero-copy claim is authorized.

Validation includes 64 applicable Rust tests across retained focused runs,
35 CRUD-index tests, all nine registered strict claim replays, crate boundaries,
and warning-denied rustdoc. Strict Clippy still reports existing debt; the
27-warning diagnostic library pass is not warning-free validation. The index
now maps 32 selectors across 15 categories, with 10 measured and 22 correctness
only. Defaults remain unchanged.

The [result table](../results/change-0423/result-table.md) retains all 1,040
observations through the underlying reports. Normal source-backed p50 values
were 2.165 / 2.165 ms for plain and 272.762 / 261.779 ms for media. Owned
media repeats drifted +7.083% in p50 and +7.034% in mean, exceeding the 5%
limits; those statistics are descriptive. The other three corpus/API pairs
passed all declared within-role limits. No cross-role delta is calculated.
All 16 report replays passed; eight R1 mutation suites rejected 88 probes,
and portable replay reproduced these checks using the pinned tools.

Explicit plan/publication/drop snapshots and stack-attributed allocation
profiles are the next measurement work. Native producers, cold/range sources,
near-limit behavior, bounded scaling and broader CRUD coverage remain open.
The full non-iWork goal remains incomplete.
