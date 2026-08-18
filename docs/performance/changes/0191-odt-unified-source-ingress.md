# Change 0191: source-backed high-level ODT path ingress

Date: 2026-08-18

## Decision

Route filesystem ODT ownership in `litchi::Document::open(Path)` through the
existing positional `FileSource -> SourceBackedPackage ->
SourceBackedDocument` chain. The prior unified path read the complete package
into a `Vec<u8>` and constructed the eager ODT document.

The unified detector now prepares one ODF package/index, retains its source
identity, and hands that same package to the ODT semantic owner. Content,
styles, metadata, notes, endnotes, footnotes, and hyperlinks retain eager
semantic parity. `Document::from_bytes` remains eager.

OOXML retains precedence over an ODT marker. A package with an OOXML catalog
is probed from the same physical source, malformed OOXML catalogs are not
hidden by ODT fallback, and DOCX read limits do not govern an ordinary ODT.
I/O and source-version failures propagate instead of being discarded during
MIME arbitration. ODS and ODP are not claimed by the document facade.

## Correctness evidence

The focused feature matrices passed with warnings and deprecations denied:

- ODT-only, DOCX-only, DOCX+ODT+Markdown, and ODT+PPTX+XLSX checks;
- eight DOCX+ODT source semantic, cold-media, mutation, suffix, limits, and
  malformed-package tests;
- OOXML precedence without DOCX and ODS non-claim tests;
- the matched high-level ODT harness oracle test;
- strict all-target harness Clippy, formatting, and diff checks.

The full `litchi-odt` all-target suite (537 unit tests plus integrations) and
strict crate Clippy passed on this implementation track. Independent
current-tree production and evidence reviews returned SAFE.

## Measurement contract

All four legs use the same dirty final release binary, SHA-256
`981ed4fbea8625b7d3feb4721262d992400bd522c51ce4be7b41071447129e59`,
built over source revision `61154e014d81b31dfa434f0186d40e4e33868afa`.
Fresh CPU-2-pinned processes ran `A1 eager, B1 source-backed, B2
source-backed, A2 eager`, with 30 warmups and 500 retained samples per case.

The open-only control times `fs::read(Path) + Document::from_bytes`; the
source-backed path times `Document::open(Path)`. Lifecycle cases additionally
include `Document::text`. Corpus construction, file publication, semantic
parity, archive/media identity, and source-range replay are outside timing.

The deterministic package is 16,812,034 bytes with 10,000 paragraphs and
eight 2 MiB opaque `Pictures/*` payloads. The direct runner is warm and
in-process; it does not participate in the fresh-child filesystem cache-state
protocol.

The predeclared p50/mean/p95/p99 same-implementation drift ceilings are
5%/5%/10%/15%. Each statistic is accepted independently only when both paired
directions are lower and both eager/source-backed drifts pass its ceiling.

## Result

For ODT open alone:

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Eager drift | Source drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | 46.07% | 53.65% | 3.46% | 11.07% | reject: drift |
| mean | 42.55% | 55.44% | 7.93% | 16.29% | reject: drift |
| p95 | 32.26% | 62.99% | 25.42% | 31.47% | reject: drift |
| p99 | 37.63% | 55.12% | 17.42% | 15.49% | reject: drift |

Open p50 values are 3.565 ms -> 1.922 ms and 3.688 ms -> 1.710 ms, but no
latency statistic is accepted because same-implementation drift exceeds its
predeclared ceiling in every row.

For open plus full text:

| Statistic | A1 -> B1 reduction | A2 -> B2 reduction | Eager drift | Source drift | Decision |
|---|---:|---:|---:|---:|---|
| p50 | 31.41% | 31.74% | 0.93% | 0.44% | accept |
| mean | 31.35% | 32.44% | 1.23% | 0.37% | accept |
| p95 | 35.36% | 32.77% | 5.43% | 1.65% | accept |
| p99 | 30.02% | 32.50% | 3.51% | 6.92% | accept |

Lifecycle p50 values are 9.239 ms -> 6.337 ms and 9.325 ms -> 6.365 ms. All
four reported statistics are accepted for this exact generated corpus because
both paired directions improve and both implementation drifts remain within
the predeclared ceilings.

In a separate untimed replay, each source preparation issued 30 logical reads
for 29,080 bytes from the 16.8 MB package. It read zero bytes from all eight
`Pictures/*` compressed ranges, and the retained full-text query needed zero
additional source reads. This is logical range evidence only, not physical
filesystem-I/O evidence.

No allocation, RSS, physical-I/O, cold-cache, decompression, throughput,
scaling, edit/save, real-producer, broad ODF, or iWork claim is made.

Artifacts:

- [machine-readable summary](../results/odt-unified-ingress-0199-summary.json)
- [artifact manifest](../results/odt-unified-ingress-0199-manifest.json)
- compressed raw A1/B1/B2/A2 reports listed in the manifest
