# Change 0079: source-backed XLSX data-validation publication

Date: 2026-08-12

Production base: `98b365af26bb93c1f3741ed63bf21221b51c2559`

Status: accepted

## Hypothesis and implementation

Worksheet data validations already had a complete typed model and a bounded,
byte-preserving XML rewriter for direct core `dataValidations` and Office 2010
`x14:dataValidations`. Publishing a small validation change still required an
eager OPC conversion that inflated and recompressed every media Part.

`litchi_xlsx::data_validation::SourceBackedEditor` now owns one immutable
positional source. Its snapshot binds the exact workbook URI, content type and
XML; the unique package `officeDocument` relationship; selected sheet name,
position and workbook relationship; exact worksheet URI, content type and XML;
and the complete sorted outbound worksheet relationship set.

The isolated editor stages one complete `Vec<Collection>`. `set_collections`
validates before adoption, `update` edits a clone atomically, and `clear`
removes every collection. Commit uses the existing rewriter, consumes its
typed post-write readback, and produces an exact reversible source-specific
patch. Immutable snapshots share that checked collection vector. Publication
revalidates the complete retained source closure without redundantly reparsing
the unchanged validation model, then consumes the accepted one-Part overlay
publisher.

Validation formulas, quoted lists, UIDs, prompts, and `sqref` values remain
inert typed text under this capability. It cannot edit cells, formulas, names,
styles, relationships, Parts, signatures, or workbook topology. MCE-selected
collections, stale/foreign sources, relationship changes, non-worksheets,
malformed or over-limit XML, changed signed sources, unsupported ZIP layouts,
and partial sinks retain typed refusals. Exact no-ops reproduce the complete
source artifact. No dependency, unsafe code, global cache, or iWork/IWA code
was added.

## Matched corpus and protocol

Both controls use one workbook, one normal worksheet, one drawing, and eight
referenced deterministic incompressible 2 MiB PNG Parts. The worksheet starts
with one core whole-number validation and one Office 2010 quoted-list
validation. The corpus has 12 ordinary Parts, 17 ZIP members, 16,783,570
logical Part bytes, and a 16,787,213-byte archive with SHA-256
`55be448c6a1ec7d2aae2a93bcbc9dd714061778be336675dd69f9edf842321b6`.

Both paths replace the complete collection vector with the same changed core
and Office 2010 rules. The eager control materializes all Parts, applies the
same typed rewriter, and uses the ordinary package writer. The candidate
performs one guarded source transaction and publishes one selected worksheet
overlay. Both produce the same 16,787,229-byte artifact with SHA-256
`5109c16e75ea6b6f85b8e3b4b8cdde0ba1994d37e5c94e7a45d6ac46fc9eb264`.
Typed workbook reopen, complete validation equality, calculation metadata,
package topology, relationships, content types, untouched Part/media bytes,
hashing, source counters, and sink bounds remain outside timing.

The control is an independent eager case in the same frozen release binary as
the source-backed case, avoiding codegen/toolchain drift. The binary SHA-256 is
`2492cfd3a7dcb85f621f5bd8c72fad85a2f02eed2105cb04cb1abef851a5ecba`.
The retained CPU-2 ABBA order was eager A, source-backed A, source-backed B,
eager B, with ten warmups and 100 samples per leg (200 per state). Raw reports
are [`before A`](../results/abba-xlsx-data-validation-before-a.json),
[`after A`](../results/abba-xlsx-data-validation-after-a.json),
[`after B`](../results/abba-xlsx-data-validation-after-b.json), and
[`before B`](../results/abba-xlsx-data-validation-before-b.json). Aggregated
evidence is in the
[`measurement summary`](../results/xlsx-data-validation-publication-summary.json).

## Results

| Metric | Eager control | Source-backed | Delta |
|---|---:|---:|---:|
| pooled samples | 200 | 200 | — |
| p50 | 222.945 ms | 5.009 ms | **-97.75% (44.51x)** |
| mean | 223.569 ms | 5.036 ms | **-97.75% (44.40x)** |
| p95 | 231.642 ms | 5.327 ms | **-97.70% (43.49x)** |
| p99 | 235.834 ms | 5.679 ms | **-97.59% (41.53x)** |
| semantic Part materializations | 12 | 2 | -83.33% |
| output bytes | 16,787,229 | 16,787,229 | exact |
| sequential writes | 630 | 547 | -13.17% |
| largest write | 32,768 B | 32,768 B | bounded |

Both same-state legs remain within the five-percent drift policy. The
candidate materializes only the workbook and selected worksheet; the other
ten Parts remain compressed and are raw-copied into the sequential output.

## Allocation, counters and memory

One-sample Heaptrack attribution covers the whole process, including corpus
construction and untimed verification. Allocation calls are 17,381 eager
versus 18,236 source-backed (+4.92%, within policy). Peak heap is 152.84 versus
152.82 MiB (-0.01%). Uninstrumented maximum RSS is 146,044 versus 143,872 KiB
(-1.49%). The allocation guard prompted removal of redundant typed reparsing
and shared ownership of immutable snapshot collections before acceptance.

Three `perf stat` repeats per state used two warmups and ten samples:

| Counter | Eager | Source-backed | Delta |
|---|---:|---:|---:|
| cycles | 20.605 billion | 6.644 billion | -67.76% |
| instructions | 50.837 billion | 13.507 billion | -73.43% |
| branches | 8.646 billion | 1.897 billion | -78.06% |
| branch misses | 187.707 million | 26.876 million | -85.68% |
| cache references | 1.444 billion | 480.104 million | -66.76% |
| cache misses | 30.446 million | 22.407 million | -26.41% |

Latency, materialization, instruction, allocation, peak-heap, and RSS gates all
clear the acceptance thresholds.

## Correctness and regression closure

Focused integration tests cover changed publication, ordinary complete reopen,
exact unselected payload/content-type/relationship preservation, patch replay
and inverse restoration, exact signed no-op output, changed signed refusal,
outbound-relationship conflict, source-version conflict, MCE-selected
collection refusal, and partial-sink failure. Existing codec tests retain core
and Office 2010 collection cardinality, strict/transitional namespaces,
formula and `sqref` validation, quoted-list and UID behavior, XML limits, and
unrelated-byte preservation.

The complete XLSX and performance-harness suites, formatting, and focused
denied-warning Clippy are release gates. The ODF-common GenericArray
deprecation fix is separately revalidated with fully denied Clippy and rustdoc
warnings in this batch.

## Alternatives retained

RTF sparse 1% paragraph editing remains a measured approximately 5.9% next
candidate. ODT count-only paragraph scanning has an approximately 49% scoped
query proxy. OLE2's next DOC owner opportunity remains benchmark-first. None
has this transaction's combination of complete existing semantic closure and
media-rich end-to-end save impact.
