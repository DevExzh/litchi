# Change 0159: PPTX source-backed cross-copy evidence

Date: 2026-08-16

Status: correctness/counter evidence only. No eager/source speedup, cache,
allocation/RSS, physical-I/O, media-rich, real-producer, or release-ABBA
claim is accepted.

## Scope

The standalone performance harness adds one opt-in selector:

- `pptx_source_backed_cross_copy_plain`

It uses the exact deterministic plain three-slide source and two-slide
destination bytes already owned by `pptx_cross_copy_plain`. It copies source
slide 3 into destination slide 2 at zero-based insertion position 1 through
the public `SourceBackedPresentationEditor` APIs:

- `plan_cross_slide_copy`; and
- `publish_cross_slide_copy_to_stream`.

No media-rich source-backed control or synthetic eager control is added. The
selector is independent evidence over matched input bytes, not an eager versus
source-backed benchmark.

## Timing and evidence boundary

The reported `elapsed_ns` is the sum of two separately timed public calls: the
source-backed plan call and the public publication call. Publication includes
the API's required preparation rerun, source/destination revision checks,
topology construction, and sequential ZIP/OPC publication. There is no
synthetic commit phase and no reopen phase in `elapsed_ns`; `plan_ns` and
`publication_ns` are emitted separately.

Corpus setup, source/destination opening, semantic setup, output reopen,
semantic verification, topology and raw ZIP preservation gates, typed stale
source/destination refusal checks, foreign-editor refusal, and output hashing
are outside timing. Raw preservation checks compare untouched destination
members' local and central records (with only the relocatable local offset
normalized), payloads, physical local order, central order, and archive
comment. The expected additions are one OPC slide Part and its relationship
member; the summary names the measured added OPC Part count, archive-member
count, and slide payload bytes explicitly.

Measured samples reset both instrumented sources after setup and report source
and destination logical `ReadAt` calls and bytes separately. These counters are
source-adapter evidence, not physical I/O, filesystem, decompression, cache,
allocation, or memory-volume evidence. Cache counters are intentionally not
serialized because the consuming public source-backed publisher does not
expose a post-publication destination cache snapshot.

## Reproduction

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 1 --samples 5 \
  --case pptx_source_backed_cross_copy_plain \
  --json target/perf/pptx-source-backed-cross-slide-copy.json
```

The focused harness test
`pptx_source_backed_cross_copy_plain_emits_counter_only_evidence` verifies
selector parsing, deterministic matched corpus bytes, plan/publication phase
vectors, source/destination logical counters, output digest, raw/topology
gates, and typed refusal gates. The selector increases the selectable matrix
from 301 to 302 while leaving the default 36 cases / 198 records unchanged.

## Remaining gaps

This tranche does not measure or accept a performance comparison with eager
PPTX, source-backed media closure, arbitrary dependency graphs, physical or
cold I/O, memory or allocation behavior, filesystem save, real-producer
documents, or a release ABBA result. Those require separately matched
controls and evidence.
