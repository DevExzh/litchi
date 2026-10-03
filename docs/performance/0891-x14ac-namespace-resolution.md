# 0891 x14ac namespace-resolution work elimination

This batch evaluates a narrow XLSX worksheet extension change. `x14ac::capture_inner`
used to call `NamespaceResolver::resolve_event` for every XML event, including
events and element names whose resolved namespace could not affect the modeled
worksheet surface. The candidate keeps the XML reader and structural checks in
place, and removes only those discarded lookups. The seven-gate XLSX quality
receipt, release build, paired capture, reduction, and retained-only packet
verification now pass. Independent raw-data review accepted the frozen primary
gate, and the root decision adopted the candidate.

## Profile motivation

The retained [0886 profile](0886-producer-full-cell-profile.md) identifies
x14ac as a secondary target in the producer XLSX full-cell scan. Its native
sampled x14ac symbol group accounts for **5.69–6.36% of period-weighted self
samples**. A separate Callgrind diagnostic attributes **19.21–19.24% of
collected guest instructions inclusively** to `x14ac::capture_inner`.

Those figures describe different scopes. The native range is self attribution;
the Callgrind row includes nested work and overlaps the MCE processing rows,
including `start`. They are not wall-clock percentages, must not be added, and
do not predict an end-to-end speedup. The profile only establishes why this
small work-elimination candidate is worth a matched measurement.

## Candidate and preservation contract

The base is commit
`3348479699f248223184c7bdbc005879b0153307`. The reviewed candidate is limited
to the production owner
`crates/litchi-xlsx/src/raw/worksheet/x14ac.rs` and the test-only differential
module
`crates/litchi-xlsx/src/raw/worksheet/x14ac/reuse_tests.rs`.

Before the change, every result of `reader.read_event()` was passed to
`resolve_event`. The candidate still reads every event with
`check_end_names` enabled. `NsReader::read_event()` therefore retains its
pending namespace-scope pop, start-tag push, declaration checks, XML parsing,
and reader error behavior. For a Start or Empty event, the candidate performs
the same `resolve_prefix(name.prefix(), true)` lookup only when the parent is
`None`, `Worksheet`, or `SheetData`; these are exactly the states in which
`start()` can recognize `worksheet`, `sheetFormatPr`, `sheetData`, or `row`.
For a `Row` or `Other` parent, the result is necessarily `Context::Other`, so
the name lookup is discarded. End and non-structural events also skip the
discarded lookup while retaining the existing stack, EOF, unexpected-end, and
depth paths.

Structural attribute handling remains in `descent()`: checked attributes,
attribute namespace resolution, value decoding, duplicate detection, lexical
parsing, and `Descent` validation are unchanged. Default and row inference,
limits, namespace rebinding, unknown prefixes, malformed XML, and first-error
behavior therefore remain covered by the original paths. The MCE
`AlternateContent` stream dispatch is outside the production change.

The candidate adds no public API, dependency, unsafe code, persistent cache, or
production allocation. Its source identities are sealed in the
[v3 candidate freeze](results/change-0891/candidate-freeze-v3.json) and
[v3 source review](results/change-0891/source-review-v3.json):

| File | Bytes | SHA-256 |
| --- | ---: | --- |
| `x14ac.rs` | 28,418 | `9f766274df5cc0a6a80291644ceecbc4f9d8afc6dac556af86bd21fbde66d148` |
| `x14ac/reuse_tests.rs` | 22,599 | `bb2ceef1b0593a9523b528a5ccfdc894d6171e2e59140359e5fb0822c6da7f3a` |

The source review finds no blocking semantic issue and makes no performance
claim. Its differential reference is intentionally a test-only copy of the
base scanner; it shares quick-xml and selected project helpers, so it is not an
independent XML implementation.

| Constraint | Evidence planned or retained |
| --- | --- |
| ADR 0001: correctness and safety first | Baseline-versus-candidate values and typed-error comparisons |
| ADR 0002/0024: ownership and dependency direction | Change confined to the XLSX worksheet owner; no dependency or public surface change |
| ADR 0005: measured performance | Matched release protocol and a primary adoption gate; no source-only speed claim |
| ADR 0006: preservation, validation, and bounded refusal | Namespace, malformed-input, depth, row-limit, duplicate, MCE, and error-order controls |

## Correctness and quality status

The focused differential module contains **nine tests**. They compare the base
reference and optimized capture in both row-capture modes, checking successful
values and error discriminants/messages. The controls cover default, strict,
prefixed, and aliased namespaces; rebinding and restoration; unknown prefixes;
malformed end and attribute syntax; invalid and negative descent values;
duplicate and reserved-marker errors; explicit and inferred row limits; the
256-level depth boundary; MCE Choice/Fallback selection; raw errors; trailing
XML; and real worksheet parts.

The real-fixture control covers **five XLSX packages and 13 worksheet members**;
three members contain a `dyDescent` marker. The [fixture census](results/change-0891/fixture-census.json)
was produced by an independent standard-library ZIP walk and does not assume
that the Rust parser accepts any member.

The first quality attempt is retained under `attempts/quality-v1`: it failed
before running the focused tests because eight qualification-test diagnostics
were denied by strict lints. The v2 attempt fixed those test-only lint issues,
passed eight focused tests, and then failed one selected-fallback assertion
that expected an overly specific error string. The v3 test-only correction
checks the typed `Error::Descent` result for the `NaN` fallback instead. The
current focused result is **9/9 passed**. The [aggregate quality receipt](results/change-0891/quality.json)
seals all seven gates: nine focused tests, 1,390 full XLSX library tests, the
all-features check, warnings-denied Clippy, warnings-denied rustdoc, scoped
formatting, and crate-boundary validation. The release build, paired capture,
reduction, and retained-only packet verification also pass; the root decision
adopted the candidate after independent raw-data review.

## Measurement protocol

The packet uses the retained 0885 before binaries and a release candidate
rebuilt from the frozen overlay. The workload has four paired cases:

- producer XLSX full-cell scan (primary);
- ordinary XLSX full-cell scan;
- selected DOCX paragraph;
- selected PPTX slide.

There are **72 fresh children**: 64 native children in eight alternating AB/BA
blocks and eight allocator children in four primary-case blocks. Native
children use three warmups and 50 measured samples; allocator children use
three warmups and 15 samples. Children run sequentially on CPU 12 with a warm
filesystem cache. Whole-child pairs, rather than pooled samples, are the
resampling unit. The reduction uses bootstrap seed 891, 10,000 draws, and the
249/9749 interval indices.

The primary adoption gate is a native primary-case p50 median ratio of at most
0.98 with a 95% bootstrap upper endpoint below 1.0. Tail latency, control
scenarios, p50 spread, allocator counters, and RSS remain review evidence;
allocator and RSS observations do not replace the primary gate.

Each native child contributes 50 measured samples, so the nearest-rank p99 is
the maximum among that child's 50 samples. The eight-pair control-tail
population is therefore small and noisy, especially for the DOCX control; it
cannot support a no-regression claim.

Managed executable copies are byte-verified and reused by their scheduled
children. This controls pathname and inode provenance, but it does not prove
that an RSS difference has a causal relation to the candidate. The warm-cache
run has no exclusive-host guarantee and makes no physical-cold-cache claim.

The final [methodology review](results/change-0891/methodology-review-v4.json)
supersedes the retained v3 review and records the pre-capture correction to the
post-cleanup `verify_commands` adapter. The v3 receipt remains retained as the
earlier methodology record. After capture, the [post-cleanup verification](results/change-0891/post-cleanup-verification.json)
passed 72 semantic command checks and 38 binding checks, reconstructed the
10,766-file source union, and confirmed that temporary managed-copy roots were
removed.

## Results and adoption (adopted)

The sealed [adoption decision](results/change-0891/decision.json) records
`decision: adopt` after the independent [capture review](results/change-0891/capture-review.json)
passed. The [reduction receipt](results/change-0891/reduction.json) reports
`status: pass`, and the [packet verification receipt](results/change-0891/verification.json)
reports `status: pass` after recomputing the reduction and checking 72 semantic
receipts. Ratios are candidate over baseline; p50 intervals are the 95%
bootstrap intervals from eight native process pairs (16 children per case).

| Case | p50 ratio (95% CI) | p95 ratio | p99 ratio | Peak RSS ratio |
| --- | ---: | ---: | ---: | ---: |
| Producer XLSX full-cell scan (primary) | 0.945064 (0.940462–0.950490) | 0.944516 | 0.946277 | 0.959876 |
| Ordinary XLSX full-cell scan | 0.997342 (0.995310–1.009379) | 0.999181 | 0.985992 | 0.959332 |
| Selected DOCX paragraph | 0.995173 (0.976749–1.008767) | 1.004846 | 1.017701 | 0.993056 |
| Selected PPTX slide | 0.974019 (0.969012–0.975964) | 0.976039 | 0.987362 | 0.965182 |

The primary p50 ratio is **0.945064**, a **5.4936% improvement**, with a
95% interval of **4.951–5.954% improvement**; its ratio upper endpoint is
0.950490, so the primary gate passes. The primary p95 and p99 medians also
improve by about **5.55%** and **5.37%**, respectively, and every primary
pair is below 1.0 for those metrics. No primary, RSS, or native p50-spread
review flag fired.

The reduction emitted three control p99 metric flags, covering five individual
pair ratios above the `1.05` review threshold:

| Control | p99 ratio (95% CI) | Flagged pairs above 1.05 |
| --- | ---: | --- |
| Ordinary XLSX full-cell scan | 0.985992 (0.870710–1.056185) | `b04` 1.056185; `b06` 1.080204 |
| Selected DOCX paragraph | 1.017701 (0.962797–1.131528) | `b01` 1.221528; `b07` 1.131528 |
| Selected PPTX slide | 0.987362 (0.934893–1.004585) | `b01` 1.102177 |

These control flags are review evidence and do not establish a candidate
regression. The DOCX tail interval is wide, so it cannot support a claim of no
DOCX tail regression. The PPTX p50 ratio of **0.974019** is retained as a
layout and host control even though the candidate does not own that path; it
does not support a PPTX performance claim.

Allocator counters were unchanged across all four allocator pairs: **25,933
calls** and **37,618,900 allocated bytes** before and after. Region peak live
bytes moved from **9,252,257** to **9,252,252**, a five-byte difference
(ratio 0.99999946) that is not material.

The adopted claim is limited to the fixed 4,096-cell producer XLSX full-cell
scan under the warm-filesystem, CPU-12, Rust 1.95 release, shared-host
protocol. It does not extrapolate to other CRUD operations, cold cache,
concurrency, or general Office performance. Process RSS remains descriptive,
and the control p99 flags cannot rule out tail costs.

| Item | Status |
| --- | --- |
| Seven-gate XLSX quality | Pass: focused 9, full library 1,390, and five static/format/boundary gates |
| Release build | Pass; frozen candidate receipt retained |
| 72-child paired capture | Pass; 64 native and 8 allocator children |
| Reduction and retained-only verification | Pass; receipts linked above |
| Primary native p50 gate | Pass; 5.4936% improvement, 4.951–5.954% interval |
| Control tails, spread, allocator, and RSS review | Reviewed; three control p99 flags cover five pairs, with caveats above |
| Adoption or rejection | Adopted by the sealed root decision; scope and control caveats above remain binding |

The [0891 packet README](results/change-0891/README.md) records the frozen
source, build, managed-copy, capture, reduction, and retained-only replay
workflow. Global performance indices record this bounded adoption below.
