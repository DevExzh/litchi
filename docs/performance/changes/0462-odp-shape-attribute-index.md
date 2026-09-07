# 0462 — reject ODP shape-attribute index on the practical gate

0462 evaluates a private fixed seventeen-key first-occurrence index for ODP
shape parsing. Generic `ElementAttrs` callers retain their existing layout;
only `ShapeAttrs` carries the index. Typed hits still decode lazily and the
source review preserves exact namespace identity, first occurrence, iterator
progress, malformed and decode error reachability, eager style fallback and
drawing-attribute harvest order.

The frozen A1/B1/B2/A2 matrix retains 24 reports and 720 samples. Normal p50
candidate-minus-baseline deltas are -2.3623% / -3.0049% for tiny,
-3.4649% / -3.1917% for medium, and -2.6602% / -2.6279% for large in R1/R2.
The predeclared 3% medium/large gate fails because both large rows remain below
threshold, although every normal p50 bootstrap interval is below zero. All
allocation metrics are exactly equal and no adverse >5% elapsed or RSS flag is
present. The fixed index adds 280 bytes of per-element state (`ElementAttrs`
144 bytes, `ShapeAttrs` 424 bytes), with no measured allocation reduction; the
R1 tiny allocator interval also crosses zero. The candidate is therefore
rejected without a retained speedup or memory benefit claim.

Supplementary public-API phase clocks, separate from the keep-decision matrix,
exclude setup, warmups and checks. Their R1/R2 p50 deltas are transaction
-1.2255% / -3.8972%, snapshot opening -3.9477% / -5.6825%, commit -3.6631% /
-4.1715%, add +0.1066% / +0.0797%, and publication -0.1857% / -0.2066%.
Whole-process counters include setup, warmups and checks: instructions
-3.3494%, cycles -3.0144%, branch misses +1.2247%, cache misses +1.2504%,
branches -2.7522%, context switches -2.4540% and page faults +2.1945%.
These supplementary values are diagnostic and do not establish causal API
attribution.

Manual assembly review finds 17 baseline `ElementAttrs::get` static call sites in
`shape_builder` with a 1,400-byte (`0x578`) frame and 17 candidate
`ShapeAttrs::get_known` call sites with a 1,688-byte (`0x698`) frame. The generic
`ElementAttrs::get` frame remains 328 bytes in both builds. Indexed known-key
hits bypass the cached linear loop, while a defensive fallback remains when a
cached-key predicate does not validate. The automatic
`known_cached_scan.eliminated` field is a narrow direct-call heuristic and is
not assembly evidence for eliminating the scan.

The candidate records 379 ODP tests and warning-denied owner Clippy as passing;
the initial compilation failure remains preserved in its historical receipt.
All 387 harness tests pass (one ignored). Both Rust files are restored
byte-exact to baseline revision `dbd2f8ece`; final Clippy/docs/format/boundaries,
portable replay, tamper rejection and owned temporary cleanup pass. Four staged
executables totaling 233,043,088 bytes were removed. The experiment adds no selector or corpus
coverage: the registry remains 439 selectors / 36 defaults, 0460's accepted
optimization remains retained, the full non-iWork goal remains open, and iWork
is excluded. See the [comparison summary](../results/change-0462/summary.json),
[phase summary](../results/change-0462/phase-summary.json), [source review](../results/change-0462/source-review.md),
and [assembly review](../results/change-0462/assembly-review.md).
