# 0799 — two-attribute fast-path preflight

0798 measured full consumption for every observed OPC iterator instance in its
generated PPTX operation regions. One- and two-attribute tags together comprised
99.777% of large capture instances. This motivates a new prototype that avoids
first-attribute replay on both classes. The rejected 0794 and 0797 prototypes
remain rejected; no historical timing is pooled.

This packet archives a candidate for all five canonical checked-attribute helper
copies and tests it outside production. The first two items use an unchecked
quick-xml iterator, with an explicit second-name duplicate check before scanning
the value. A possible third item returns to the checked iterator after replaying
the successful prefix. The existing 32-name/ordered-map bound remains required.
Exact malformed-key, equals, quote, duplicate, clone and terminal semantics are
mandatory. The source archive and review describe the final implementation.

Root alone runs isolated helper tests, builds, native captures, and Callgrind,
serially. Source agents may work concurrently; heavy offline replay waits until
every capture is terminal. One standalone binary contains both exact helper
legs. Thirty-nine frozen inputs extend 0797 with the two-to-three transition:
duplicate-after-two, long quoted/unterminated duplicate-after-two, and the three
syntax tails after two successful attributes. The probe checks exact first-error
results and positions against quick-xml, plus clones and repeated terminal calls.

Six alternating native block pairs use 30 samples, three warmups, and 4,096
repetitions. Construction includes opaque materialization and destruction;
consumption includes construction, iteration, checksum work, and destruction.
All inputs are hot and repeatedly reused. Two Callgrind repeats use one sample,
one iteration, no warmup, and the four exact non-inlined owners. Guest counters
and simulated branch events are diagnostic and never native cycle fractions.

The final plan freezes advancement criteria before compilation or measurement.
Advancement can authorize only fresh public-workflow, resource and cross-format
trials; production adoption is outside this packet. All regressions remain
visible, including cases outside any advancement predicate. The measured 0798
counts cannot be multiplied by micro ratios to predict a workflow speedup.

Source restoration is trivial because production is never edited. The packet
retains exact source and architecture hashes, fixture oracles, source-bound build
receipts, every raw sample/counter dump, independent audits, and cleanup witnesses.
Only the owned target is removed. Unrelated files and worktrees remain intact.
No producer, CRUD, cold/range, concurrency, or format coverage is promoted.
