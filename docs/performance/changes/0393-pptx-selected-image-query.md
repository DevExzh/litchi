# Change 0393: PPTX selected-image query without full descriptor retention

Date: 2026-09-03

Status: implemented, independently reviewed and tested, with a narrowly scoped
release claim

`performance_claim: scoped`

`claim_authorized: true`

## Scope and mechanism

`SourceSlide::image` and `SourceSlide::read_image` formerly called
`SourceSlide::images`, reserved a descriptor vector for the complete scene,
constructed every picture descriptor, and then retained only one. They now
use a private selected-query mode that retains only the requested descriptor
and the final picture count. `SourceSlide::images` keeps its original
full-inventory behavior and allocation policy.

The selected mode is deliberately not an early exit. It still parses every
picture, validates every picture relationship, and resolves every target in
scene order, including pictures after the selected one. It then performs the
existing final execution and source-version fences before reporting an
out-of-bounds position. `read_image` continues with the existing external-
target refusal, internal Part and relationship validation, payload read, and
final freshness checks only after that metadata pass. MCE refusal, malformed
grammar, missing or mistyped targets, cancellation, freshness, and error
precedence therefore remain fail closed.

Test-only counters prove that a three-picture `image` or `read_image` query
uses one selected query, performs zero descriptor-vector reservations, and
resolves all three targets. The matching `images` query uses one full query
and retains its descriptor reservation. Additional tests cover a malformed or
missing later target when position zero is selected, exact final length for an
out-of-range request, malformed-before-out-of-bounds precedence, MCE refusal,
and source-change-before-out-of-bounds precedence.

No public API, package ownership boundary, dependency edge, executor, cache,
archive representation, unsafe code, or parallel behavior changed.

## Validation

Validation explicitly selected Rust and Cargo 1.98.1 because the repository's
pinned 1.95 installation lacks Cargo. All commands used `--locked`, one Cargo
job, and an isolated target.

- `cargo check -p litchi-pptx`: passed.
- New selected-query tests: `3/3` passed.
- Picture-related tests: `20/20` passed.
- Source-version, cancellation, MCE, outbound-target, and filesystem selected-
  image tests: `11/11` passed.
- Full library gate with the one documented exclusion: `534` passed, `1`
  skipped.
- The excluded
  `opened::tests::stale_and_unsupported_raw_xml_fail_before_publication`
  failure reproduced at exact baseline `4ba030c2a6c…` with
  `Invalid("presentation slide-ID list contains a non-slide child")`.
- Clippy reported only four pre-existing lints in untouched files and no lint
  in `presentation/source.rs`.
- `git diff --check` and the single-file format check passed.

Independent source review found no material correctness or ownership issue,
and independent testing used a separate target and exact-baseline worktree.

## Matched release evidence

The clean matched run used `A1/B1/B2/A2` order, release mode with
`allocator-metrics`, one worker, five warmups, and 30 retained samples per
case. Control was exact revision `4ba030c2a6cd68b88e502a117d5730c5a6202fee`;
the candidate was that revision plus the one-file patch SHA-256
`4da5ce10d8f7475279ac4bdef639c3269d0baff30959b67c79b5f8869daf0559`.
The control and candidate binary SHA-256 values were respectively
`c3262985314cfb9e78a9593f80bbffb53043c33f6d40a4aa7c62f95a29cf4ad0`
and
`85252c17b48d5ed98d2383da42fc7c97183fdf18977f2f423c225c9c4da10e1f`.

The deterministic picture-heavy PPTX contains 32 OPC Parts, 51 ZIP members,
and eight direct pictures on the selected slide. Its 16,815,621 archive bytes
have SHA-256
`830928aafdc3ec8a5995a0d84a82ea9e2acb7f190f45008ddc017e2edfbf684b`.
The selected two-MiB payload has SHA-256
`5d6b7fd3f2ed6306e4470510c77d83b051a737e9fbd3776d71340308d4e89063`.

| Selector | A1 to B1 p50 | A2 to B2 p50 | Exact allocation delta |
| --- | ---: | ---: | ---: |
| `pptx_source_backed_images_query` | -9.549% | -10.158% | unchanged |
| `pptx_source_backed_image_query` | -12.434% | -12.404% | -8 calls; -1,831 bytes |
| `pptx_source_backed_read_image_query` | -4.668% | -4.915% | -8 calls; -1,831 bytes |

For `image`, deallocation calls also fall by eight, allocated and deallocated
bytes both fall by 1,831, and reallocation calls remain 121. `read_image` has
the same exact reductions around its selected payload. These allocator vectors
are identical between the two control legs and between the two candidate
legs. The `images` allocator vectors are identical across all four legs.

Logical source evidence is also invariant. Metadata-only `images` and `image`
queries issue five reads returning 706 bytes and read zero selected-media
bytes. `read_image` issues 74 reads returning 2,098,549 bytes, of which 65
reads and 2,097,797 bytes cover the selected compressed media range.

The [retained evidence bundle](../results/change-0393/) contains all four raw
reports compressed without changing their JSON, the capture summary, exact
allocation vectors, percent-delta table, artifact manifest, and final
adjudication.

## Claim boundary and decision

The authorized latency claim is limited to the p50 reduction of
`pptx_source_backed_image_query` on the exact corpus, binaries, configuration,
and ABBA protocol above: 12.404% to 12.434% in the two matched directions.
The exact operation-local allocation/deallocation call and byte reductions are
also accepted for the named `image` and `read_image` selectors.

The favorable `images` timings are observations, not an accepted claim,
because that path retains its original allocation behavior and the selected-
descriptor mechanism does not directly explain the result. `read_image`
latency is also observational. All p95/p99 and mean latency claims are
withheld; in particular, the `read_image` control p99 drift was 5.053%, just
over the uniform 5% caution boundary used for this adjudication. No result is
generalized to other slides, corpora, PPTX operations, cold-cache behavior,
RSS or peak memory, physical I/O, decompression, throughput, parallel scaling,
fixed memory, or general OOM prevention.
