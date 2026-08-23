# Change 0265: PPTX slide-boundary publication harness

## Status

Landed in `0b62b26b4033a6d6f358a1178c57fc901215c1f2`. This change adds two
opt-in PPTX whole-slide publication selectors for boundary CRUD correctness
and phase evidence. It does not establish a latency, allocation, RSS, or
physical-I/O claim.

## Scope

The selectors `pptx_slide_remove_boundary_save` and
`pptx_slide_move_boundary_save` use one deterministic four-slide PPTX corpus:

- four dependency-free plain slides;
- 45 ZIP members;
- 32,396 source archive bytes;
- source archive SHA-256
  `685a1805ad291e8f9852d3ccd584320f20847bd0ac8fdf29857f96efe1109477`.

The removal selector covers the first, middle, and last positions of the
four-slide presentation (`0`, `1`, and `3`). A one-slide package is separately
required to refuse removal of its final slide. The move selector covers both
boundary directions (`0 -> 3` and `3 -> 0`), and moving a slide from a position
to itself is an exact no-op. The two selectors are opt-in `Case` entries: the
selectable matrix is now 385 names, while the default remains 36 cases and 198
records.

## Production path and timing

Removal uses the production opened-presentation `Snapshot` and its typed slide
removal plan; move uses the production `Snapshot`/`Transaction` flow and its
commit. The measured representative operation reports separate plan, commit,
sequential OPC-publication, and semantic-reopen phase vectors. Corpus and sink
preparation, expected outputs, all independent package/ZIP oracles, durable
patch checks, refusal probes, and source-immutability checks are outside the
phase clocks.

## Correctness and evidence boundaries

Each output is semantically reopened and checked for the expected slide order
and count. Untouched members retain both their raw ZIP local records and their
central-directory records after local-header offsets are normalized. Removal
allows only its presentation, relationship, content-type, and selected-slide
records to change; move requires strict identity for every untouched member,
including `[Content_Types].xml`.

The corpus is built twice and must be byte-identical. The harness also checks
exact no-op behavior, source immutability, dependency, unknown-member,
markup-compatibility, signed-package, and limits refusals, plus stale and
foreign serialized-patch refusal. Partial and zero-output sinks must fail with
the required accepted-byte behavior. Representative removal and move patches
are serialized, decoded, applied, and inverted; forward and inverse results
must reopen to the expected output and source bytes. These are bounded
production-contract and publication-locality gates, not claims about arbitrary
PPTX producers or unsupported relationship topologies.

## Reproduction

```sh
cargo run --manifest-path tools/perf-baseline/Cargo.toml --release -- \
  --case pptx_slide_remove_boundary_save,pptx_slide_move_boundary_save \
  --warmup 3 --samples 30 \
  --json target/perf/pptx-slide-boundaries.json
```

The selectors provide correctness and phase evidence only. No latency,
allocation, RSS, decompression, throughput, or physical-I/O claim follows.
