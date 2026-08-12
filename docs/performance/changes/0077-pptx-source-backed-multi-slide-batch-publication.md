# Change 0077: source-backed PPTX multi-slide batch publication

Date: 2026-08-12

Production base: `a82a5d9f3f3e34706a9ee890fd23f48b92718e6f`

Status: accepted

## Hypothesis and implementation

The existing source-backed PPTX editor could publish an atomic text batch only
inside one slide. A common bulk update across a presentation therefore still
required eager ownership and recompression of all 229 ordinary Parts. Eight
existing slides form a bounded same-topology closure: the package catalog,
presentation Part and ordered slide relationship graph stay immutable, while
only the selected slide XML payloads change. Publishing that set through one
multi-Part preservation plan should avoid materializing the other 220 Parts.

`SourceBackedPackage::write_part_overlays_to_stream` is an additive consuming
OPC primitive for at most 64 unique existing Part replacements. It sorts and
deduplicates Part URIs, checks per-Part and aggregate logical/archive limits,
reads every selected source payload to prove framing, compression, declared
size and CRC, validates changed XML, checks signatures, and builds the complete
preservation plan before output. Changed selected members are regenerated;
selected exact no-ops and all unselected members retain their raw local and
central ZIP records. An empty set or an all-no-op set copies the source artifact
byte-for-byte. The earlier one-Part API is retained as a singleton wrapper.

`SourceBackedPresentationEditor::edit_slides` adds a non-cloneable borrowed
batch transaction for at most 32 distinct existing slides. Each slide accepts
one existing same-slide shape-text batch of up to 256 unique nonoverlapping
selectors. The outer batch has the existing 8 MiB aggregate replacement-text
budget, canonicalizes slides by presentation position, and produces an exact
reversible batch snapshot/patch/commit. Publication recaptures the package,
presentation, slide-reference, slide-Part and relationship closure for every
selected position, applies the patch only to that exact set, then consumes the
OPC source through the multi-Part publisher.

Slide topology/order, relationships, content types, layouts, notes, charts,
media, hyperlinks, shape structure and MCE-selected raw XML remain outside the
capability. Changed signed sources, duplicate slides/Parts, stale or foreign
commits, source-version changes, limit violations, unsupported ZIP layouts and
partial sinks retain typed failures. No dependency, runtime, unsafe code,
global cache, relationship edit, topology operation, or iWork/IWA code was
added.

## Matched corpus and protocol

Both controls use the existing deterministic media-rich PPTX archive: 200
slides with eight text boxes each, eight referenced incompressible 2 MiB PNG
Parts, 229 ordinary Parts, and 445 ZIP members. The archive is 17,017,139 bytes
with 17,568,429 logical Part bytes and SHA-256
`61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757`.

Both paths replace all eight text boxes on zero-based slide positions 0, 28,
57, 85, 114, 142, 171, and 199. The eager control performs positional open,
full OPC ownership, one ordinary opened-presentation transaction, and ordinary
sequential publication. The candidate performs positional open, one guarded
source-backed multi-slide transaction, and one eight-member overlay
publication. Both emit the same 17,017,145-byte artifact with SHA-256
`23d6d7b8dd433ff453307ee56485efadc34d6b18b455a6858d9f565c4b1b6cd9`.

The frozen control binary SHA-256 is
`9bee76796aecd22cc881e7d4e5cd2936b031a30d86b1f8e9b7572b49d6ad0993`;
the candidate binary SHA-256 is
`ece3a2e77d591813ae9611a29bc7c62787ff1cdf87dedabeb5713cc968af6e3d`.
Corpus construction, complete PPTX reopen, all-slide text readback,
topology/relationship/content-type checks, exact unselected Part and media
payloads, raw ZIP-record comparison, patch/inverse checks, hashing and
source/sink inspection remained outside timing.

The retained serial sequence on CPU 3 was eager A, source A, source B, eager B.
Each leg used 30 warmups and 200 samples, yielding 400 samples per state. Exact
evidence is in the four `abba-pptx-multi-slide-batch-*` reports and the
[`measurement summary`](../results/pptx-multi-slide-batch-publication-summary.json).

## Results

| Metric | Eager control | Source-backed | Delta |
|---|---:|---:|---:|
| pooled samples | 400 | 400 | — |
| p50 | 331.362 ms | 13.997 ms | **-95.78% (23.67x)** |
| mean | 332.160 ms | 14.291 ms | **-95.70% (23.24x)** |
| mean 95% interval | 331.483-332.837 ms | 14.127-14.454 ms | disjoint |
| p95 | 345.298 ms | 16.505 ms | **-95.22% (20.92x)** |
| p99 | 352.804 ms | 18.518 ms | **-94.75% (19.05x)** |
| semantic Part materializations | 229 | 9 | -96.07% |
| output bytes | 17,017,145 | 17,017,145 | exact |
| sequential writes | 260 | 1,403 | +439.62% |
| largest write | 65,536 B | 32,768 B | bounded |

The source-backed path materializes only the presentation root and eight
selected slide Parts. The other 220 ordinary Parts and all non-Part physical
members are copied from their source spans. Payload reads remain archive-sized
because unchanged compressed members still have to reach the sequential sink.
Between-leg p50/mean drift is 0.31%/0.05% for eager and -1.84%/-3.84% for the
candidate, within the 5% stability policy.

## Allocation, counters and memory

One-sample Heaptrack attribution reports allocation calls 3,277,310 to
2,210,873 (-32.54%) and temporary allocations 627,700 to 207,039 (-67.02%).
Peak heap falls 175.14 to 159.49 MiB (-8.94%). Heaptrack-inclusive RSS is flat
at +0.16%, and two uninstrumented GNU Time processes per state report mean
maximum RSS 145,916 to 146,654 KiB (+0.51%, flat). Both retain 1.78 KiB of
profiler/runtime leakage.

Three `perf stat` repeats per state used two warmups and ten samples:

| Counter | Eager | Source-backed | Delta |
|---|---:|---:|---:|
| task clock | 6,186.41 ms | 1,892.27 ms | -69.41% |
| cycles | 30.256 billion | 9.320 billion | -69.20% |
| instructions | 76.781 billion | 23.348 billion | -69.59% |
| branches | 13.453 billion | 4.007 billion | -70.22% |
| branch misses | 188.295 million | 21.100 million | -88.79% |
| cache references | 1.719 billion | 549.034 million | -68.06% |
| cache misses | 43.374 million | 27.236 million | -37.20% |
| page faults | 424,684 | 299,447 | -29.49% |

Each profiled process had one CPU migration. The complete profiles lost no
samples; cycle-weighted `deflate_medium` attribution falls about 93.08%. The
candidate profile's remaining compression is chiefly deterministic corpus
construction. Whole-process reductions are smaller than scoped latency because
corpus construction and complete untimed verification remain inside each
profiled process.

## Validation and limitations

Passed on the final source:

- focused OPC multi-Part tests for raw selected/unselected member identity,
  exact signed all-noop, duplicate/count/aggregate limits, invalid XML,
  changed signatures, source versions, unsupported layout and partial sinks;
- focused PPTX multi-slide tests for sorting, one operation per slide, exact
  patch/inverse, complete reopen, all selected/unselected semantics, raw ZIP
  identity, media/topology/relationships, signed/no-op, stale/foreign/version,
  MCE/limits and partial sinks;
- matched harness equivalence and CI-pinned corpus, output, sink and 229 -> 9
  materialization assertions;
- complete OPC, PPTX and performance-harness suites plus warning-denied Clippy;
- warning-denied ODF-common all-target/all-feature Clippy and rustdoc,
  revalidating the GenericArray deprecation fix from `1194fbc7f`; and
- formatter, workflow parsing, JSON parsing, ADR digest, whitespace and final
  diff checks.

This is not general multi-slide authoring. It cannot add, remove, reorder or
duplicate slides; edit relationships, notes, layouts, charts or media; change
shape structure; publish MCE-projected slide XML; or mutate signed sources.
OLE2, RTF and ODF production code are unchanged by this batch; iWork/IWA
remains deferred.
