# Independent source and test review

The design/source reviewer found no correctness blocker in the concrete patch.
Checkpoint capture follows the existing successful worksheet-start seek;
publication still follows complete scanning and final fences. Each replay uses
an independent local hint. Duplicate order, SST/formula/XF decoding and error
precedence remain unchanged.

Weak parsed-index identity is valid under the current private constructor
invariant: each SharedOleFile creates a distinct parsed-index allocation.
Foreign and expired checkpoints fall back to cold hints. The weak identity
prevents ABA and does not retain source bytes or parsed metadata. A future
shared-index constructor would need an explicit per-reader identity boundary.

The 192-byte index overhead conservatively includes the new inline checkpoint;
cache-table fixed overhead remains 128. Slot capacity, managed reservations,
concurrent candidate weights, pinning/eviction and release remain accounted.
The checkpoint adds no separate allocation and is discarded on abandonment.

A separate tester added ten CFB tests for lifetime, foreign/expired identity,
FAT/MiniFAT, backward/stream mismatches, corrupt chains, source changes and
concurrent local hints. Another tester added two XLS tests for alternating
first/last/missing/random coordinates and concurrent cloned handles. All final
owner/facade checks pass: 1,870 combined CFB/XLS tests, two existing ignored,
and 61 facade tests, plus formatting/check/Clippy/rustdoc.

The reviewer also approved the final direct pointer-identity comparison:
ParsedOleIndex is sized and nonzero-sized, the pointers are never dereferenced,
and the retained Weak keeps the control block allocated even after value drop.
This removes the initial temporary Arc downgrade and its two weak-count atomic
operations. Initial measurements and the longer tiny-query controls are archived;
no cold-query or open-time variation is attributed to those removed atomics.

Final independent performance review recommends retention with an explicit
tradeoff: the 200-sample controls reproduce about 25% lower large-sheet warm
latency and 19–20% lower Plan1 latency, supported by reduced repeat-loop
instructions. Simple owned queries regress 8–16%, the missing open-plus-three
workflow regresses 5.12–5.65%, and refusal routes regress 5–11%. Tight controls
mean these costs must not be dismissed as noise. No first/open/refusal gain is
attributed to checkpoint reuse. Plan1 file N=10 RSS is a +5.92% review trigger;
N=1,010 is neutral, so no RSS improvement or checkpoint-causal claim follows.

Retain only the scoped repeated selected-query optimization for meaningful
worksheet prefixes. Tiny/no-benefit admission and refusal-path costs remain
priority follow-ups. The goal's 5% threshold is a review trigger; there is no
aggregate performance claim or coverage promotion. Measured allocator request
and retention grow 32 bytes per successful build, distinct from the conservative
64-byte increase in logical cache accounting.

A separate final evidence reviewer cross-checked the six timing-table rows,
long controls, complete regression disclosures, profile shares, instruction/RSS
numbers, all 88 allocation groups and quality/corpus/sample counts against the
retained artifacts. No material factual blocker was found. Review was read-only;
no additional builds were run.
