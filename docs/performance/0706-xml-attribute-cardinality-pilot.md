# 0706 — XML attribute-cardinality proof pilot

`performance_claim: none`; the candidate is rejected and production is restored.

The 0706 candidate reuses the successful lexical attribute-layout traversal
already required by `xml-minifier`. That traversal returns a private saturated
cardinality (`None`, `One`, or `Many`) to attribute inspection. Only a proven
zero- or one-attribute start tag disables duplicate-key bookkeeping; `Many`
keeps the original checked iterator. The change adds no probe, replay pass,
retained collection, dependency, public API, or unsafe code. Lexical failures
still precede attribute inspection, and the existing decoder, limits,
`xml:space` inheritance, root/depth/token/event checks, and guarded stream
boundaries remain authoritative.

The candidate was reviewed independently before measurement. The review found
no blocking correctness issue: the lexical proof uses the same four XML
whitespace bytes and boundary as quick-xml, a successful zero/one scan cannot
contain a second attribute, and both slice and guarded-stream callers receive
the proof from their existing layout scan. Declaration checks discard it.
The source-compatible XLSX path still retains both original and replacement
audits; existing XLSX layout facts and insertion-only splice proofs do not
prove the generic OPC XML contract.

## Decision and gate results

The experiment uses revision `cfd2aa4f3fb051f3914503fce597ef70593a26dd`,
CPU affinity 12, separate release native and allocator children, and the
frozen order `baseline-noise-1`, `baseline-noise-2`, `baseline-A1`,
`candidate-B1`, `candidate-B2`, `baseline-A2`. The primary workload is the
matched source-backed XLSX one-percent cell edit/save route on medium and
dense-sparse shapes. Each primary leg retains 200 samples after 20 warmups;
guard legs retain 30 samples after 10 warmups, and the producer control retains
200 samples after 20 warmups. The native packet has 60 timing children and the
separate allocation packet has eight children.

Admission required every primary row to improve total p50 and total mean by at
least 2% and publication p50 by at least 5%. The allocator diagnostic required
at least an 8% publication allocation-call reduction. The native gate fails in
seven of eight paired rows, so the pilot is rejected even though the allocator
gate passes.

| Pair | Shape / repeat | Total p50 reduction | Total mean reduction | Publication p50 reduction | Row |
| --- | --- | ---: | ---: | ---: | --- |
| A1 → B1 | dense-sparse / 1 | 1.487% | 1.595% | 2.312% | fail |
| A1 → B1 | dense-sparse / 2 | 3.755% | 3.741% | 6.607% | pass |
| A1 → B1 | medium / 1 | 1.754% | 1.780% | 3.152% | fail |
| A1 → B1 | medium / 2 | 1.919% | 1.954% | 2.686% | fail |
| A2 → B2 | dense-sparse / 1 | −0.467% | −0.009% | −2.028% | fail |
| A2 → B2 | dense-sparse / 2 | 1.716% | 1.711% | 2.406% | fail |
| A2 → B2 | medium / 1 | 1.712% | 1.677% | 2.611% | fail |
| A2 → B2 | medium / 2 | 2.559% | 2.558% | 3.001% | fail |

Thus B1 has three of four rows below the total threshold, and B2 has no row
that meets all three gates. The one passing row does not support a claim.
The same threshold magnitude was intentionally used as in rejected 0529, but
the mechanisms differ: 0706 reuses an existing lexical proof, while 0529
added an unchecked two-attribute probe and then replayed the checked iterator.
The 0529 probe/replay design is not revived by this result.

The packet reports 157 native absolute-change flags over 5% (124 paired
comparisons and 33 repeat-drift entries). Of the paired flags, 53 are adverse
and 71 favorable. Repeat two is higher in 18 drift entries and lower in 15,
with changes ranging from −29.86% to +11.32%. Adverse examples include a
dense-sparse primary open p50 increase of 11.57% and managed dense-sparse
reopen p50 increase of 10.30% in A2/B2; reopen remains outside the primary timer. Baseline A/A diagnostics have a maximum absolute shift of 45.03%
across the recorded noise and baseline-repeat metrics; this is a disclosed
diagnostic with no fixed pass threshold. The allocator comparison has 16
over-5% flags, all favorable publication allocation/deallocation count or byte
reductions. Operation-region peak-live-byte changes are below 5% and differ by
three bytes in the primary publication rows. These flags describe the paired
and repeat diagnostics; they are not additional retained performance claims.

The allocator comparison is useful attribution, but cannot override the native
gate. Publication allocation calls fall from `19,197` to `365` on medium
(`−98.10%`) and from `36,573` to `365` on dense-sparse (`−99.00%`). Publication
requested bytes fall by 29.21% and 40.32% respectively. Commit allocation
calls remain `42,772` and `80,835`; operation-region peak live bytes differ by
only three bytes in these captures. No RSS was captured for this pilot.

The follow-up auditor microguards are separate evidence, not primary speed
claims. The corrected analysis covers 60 paired rows across zero, one, two,
many-64, `xml:space`, and source-noncompact families, with five warmups and
100 samples of 100 auditor calls per row. It retains 45 metric-level review
flags over 5% after correcting the second-pair labels, plus 14 baseline A/A
drift flags over 5%. Representative adverse
batch results include mixed-chunk stream/authored one-attribute p50 at −6.52%,
stream/authored two-attribute p50 at −10.70%, and stream/authored
source-noncompact p50 at −14.31%. These are 100-call batch timings; they are
not per-call claims. The raw oracle rows and the archived pre-correction
analysis remain in the packet.

The publication-instruction profile and the conditional `publication Ir ≥ 3%`
gate were not required after the primary pilot failed. No broad consumer or
profile claim follows from the diagnostic reductions.

## Correctness and validation

The public differential oracle exercises 4,148 deterministic cases and 29,036
calls across compact, authored, source, stream, limit, malformed, and chunked
policies. Baseline and candidate outcomes have identical full JSON results and
result digests, with zero panics. The exact oracle recapture corrected an
earlier whitespace normalization that could hide Debug-detail changes; the
normalized preflight is
archived, exact Debug strings were recaptured, and both builds remain equal.
An earlier oracle probe correction only discarded ignored `black_box` results
under `deny(warnings)`; it did not change semantics or production code. The
test preflight also corrected an aggregate fixture count and the expected
unterminated-attribute offsets/details against unchanged baseline behavior.

The restored production checkout passes the focused `xml-minifier`
all-features locked run with 63 tests passed and one existing ignored test.
The candidate checkout passed 66 tests and one ignored test because it included
the three private unit tests added with the candidate. Eight independent public
integration test functions in
[`crates/xml-minifier/tests/attribute_cardinality.rs`](../../crates/xml-minifier/tests/attribute_cardinality.rs)
are retained with the restored production source. They cover zero/one/two/many
cardinality, duplicate refusal, malformed attributes, `xml:space`, lexical
policy, BOM/encoding offsets, and inclusive resource boundaries across slice
and chunked stream routes.

The final restored checkout also passes `cargo fmt --all --check`, the recorded
`xml-minifier` all-features/all-targets check, and the recorded
`xml-minifier` Clippy command with `-D warnings`. All six evidence gates pass:
crate boundaries, strict performance claims,
structural claims, report classification, CRUD coverage, and the non-iWork
audit. The packet retains their commands, source bindings, and logs.

## Scope and limits

All raw native timing, allocator, oracle, microguard, build, source-census,
review, and correction artifacts are retained under
[`results/change-0706/`](results/change-0706/). The native timing evidence uses
synthetic in-memory XLSX source shapes and separate children; it does not
measure RSS, physical cold storage, native Office producers, cross-platform
behavior, or broad consumer quality. The primary total includes the measured
open/planning/commit/publication region defined by the harness, including the
returned `MultiSnapshot` drop, and excludes sink setup, remaining handle
destruction, reopen, and oracle work. The allocator values are operation-region
diagnostics and are not process RSS or summed phase peaks.

Production retains no 0706 auditor implementation. The independent test file
is retained as public coverage.
The broader non-iWork performance goal remains active; iWork is outside this
batch.

The packet's [README](results/change-0706/README.md) maps every raw and derived
artifact, while [`analysis.json`](results/change-0706/analysis.json) records
the machine-checkable rejection and paired gate calculations.
