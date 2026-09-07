# Next work after the 0462 ODP attribute experiment

The full non-iWork performance goal remains open. The 0462 shape-index candidate
is rejected: both large normal rows miss the frozen 3% gate, with unchanged
allocation metrics and 280 bytes of extra shape state. The next optimization
should move beyond small `ElementAttrs` matcher changes.

## Selected ODP bottleneck; global ranking remains open

Within the recently profiled ordinary existing-ODP append workload, the large
commit phase is the largest measured phase. The 0458 phase matrix reports a
152.056/152.334 ms large normal lifecycle p50 in R1/R2. Commit consumes
45.64%/45.55% of the phase envelope and 118,096,476 of 211,442,207 lifecycle
allocated bytes (55.85%). Snapshot opening and transaction construction are
also substantial at about 27% and 26%; append is about 1.55% and sequential
publication about 0.03%. These are the accepted owned
`odp_existing_append_lifecycle` phase boundaries, not a source-tail result.

This is not a global end-to-end ranking. The 0452/0453 media-rich simulated
range PPTX lifecycle is roughly 1,767--2,573 ms depending on the capture
variant, but it has a different provider, result contract and coverage status.
Cross-format ranking therefore remains unresolved until comparable source,
output, semantic and resource scopes are captured.

The 0459 frame-pointer/DWARF diagnostic gives a useful internal ranking but is
not an ordinary-build timing attribution. Within commit, sampled periods put
`Snapshot::from_owned_package` at 59.36%/60.01%,
`Parser::parse_slides_with_styles` at 57.23%/59.71%, and
`MutablePresentation::to_bytes_bounded` at 27.41%/29.98%. These inclusive
shares overlap and include warmups. They identify candidate work families,
not removable work or a causal speedup. The 0458 source audit also records
candidate package reopen, semantic readback, compact/provenance validation,
and patch retention as separate materialization boundaries.

0461 does not justify continuing with a local matcher as the primary target:
its normal p50 improvements are only 2.060%/2.155% in R1 medium/large, below
the 3% gate, and every allocator metric is unchanged. 0460's fused staging
optimization remains the accepted baseline; its transaction benefit should not
be conflated with a commit-path result.

## Concrete next measurement

The 0459 diagnostic already narrows the ODP commit work to semantic readback
and serialization: `Snapshot::from_owned_package`/slide parsing has the
largest sampled share, with `to_bytes_bounded` next. A further broad
attribution-only turn would not answer the next design choice. Run one
decision-oriented diagnostic build or harness-only probe for the existing
public lifecycle, keeping the production timer and `Snapshot`/`Commit`/`Patch`
contract unchanged. Split `Transaction::commit` into the following bounded
stages and retain non-overlapping ownership notes:

1. `MutablePresentation::to_bytes_bounded`, including
   `generate_content_xml` and writer/compression work.
2. Candidate `OwnedPackage::from_bytes` and physical archive indexing.
3. `Snapshot::from_owned_package` and
   `Parser::parse_slides_with_styles` semantic readback.
4. `validate_compact_xml_parts`, including the source/provenance splice attempt
   and any compact-audit fallback.
5. Media/RDF/chart checks, `Patch` construction, and source/after retention.

The relevant owners are `crates/litchi-odp/src/authoring/edit.rs` (commit at
roughly lines 2164-2395), `crates/litchi-odp/src/authoring/mutable.rs`
(`generate_content_xml` and `to_bytes_bounded`),
`crates/litchi-odp/src/codec/parser`, and
`crates/litchi-odf-common/src/core/writer.rs`/`xml_splice.rs`. The existing
phase driver is `tools/perf-baseline/src/odp_append_attribution.rs`; its
public phase clocks and lifecycle oracle should remain the enclosing reference.

Use the established 64/4,096/8,192-slide corpus, normal and allocator lanes,
R1/R2 ordering, three warmups and 30 retained samples. A diagnostic profile
may use the 0459 frame-pointer/DWARF setup, but the final candidate needs a
profile bound to its actual binary epoch. Record per-stage allocation and
working-set observations without summing phase peaks, and record the counts and
bytes for candidate slide readback, source `content.xml` loading, splice
success/fallback, and compact-audit passes. The decision probe should answer
whether a proof-reuse candidate below removes a whole source/candidate scan or
only moves the same work. Preserve exact source bytes, output bytes, semantic
readback, untouched-member, no-op and reversible-patch checks. If the proof
probe is neutral, the existing sampled ranking supports examining readback
reuse; if it shows a material validated duplicate scan, pursue proof reuse
before redesigning snapshot ownership.

The concrete proof-reuse candidate to inspect concurrently is a private,
source-identity-bound `content.xml` audit token from
`MutablePresentation::generate_content_xml`/the writer to
`validate_compact_xml_parts`: carry the source version and member identity,
exact source/candidate spans or hashes, generated-fragment audit result, and
the splice classification. The final validator may consume it only when every
binding still matches; otherwise it must execute the current splice/fallback
path. This could avoid repeating the source load and authored-fragment audit
already performed during generation. It must not skip candidate reopen,
well-formedness, compactness, malformed-input refusal or source-freshness
checks. First measure token hit/miss and source/candidate bytes before any
production change.

Any optimization must retain atomic commit, complete semantic readback unless
an equivalent proof is accepted, compactness/provenance refusal behavior,
resource limits, and exact ordinary Patch semantics. A lazy source-backed
replacement for ordinary `Snapshot`/`Patch` remains an API/ADR decision, not a
local commit optimization.

## Broader completion gap and follow-on

The coverage index still has 15 categories, 33 representative mappings, 10
measured mappings and 23 correctness-only mappings. The generated ODP append
matrix is a formal supplementary baseline but remains correctness-only in the
checked-catalog index. Native application roundtrips, distinct-package pairs,
cold/physical I/O, bounded-worker scaling, broader semantic CRUD, and source
backed editor adoption remain open.

After the ordinary commit attribution, the highest broader-ROI follow-on is to
finish one source-backed OPC semantic lifecycle rather than another isolated
parser microbenchmark. The prepared substrate is in
`crates/litchi-opc` and the PPTX consumer in
`crates/litchi-pptx/src/presentation/source_cross_copy.rs`. Changes 0452/0453
show why this is worth measuring: the media-rich simulated-range lifecycle
falls from 2572.729/2572.933 ms to 1766.853/1774.087 ms in R1/R2 after
retained capture, and shared decoded ownership removes 16,777,408 planning
bytes. Those are provider-specific measurements and memory evidence, not a
general PPTX speedup: the representative cross-copy rows remain
correctness-only, generated/native breadth is incomplete, and the simulated
range service is not physical network evidence.

That follow-on needs an explicit public-editor result contract, checked
source/destination identities, dependency-closure and untouched-member oracles,
allocator/source/sink counters, and separate native/different-package and
cold/range matrices. Do not promote the row or claim full-goal completion from
the existing substrate or simulated-range results.
