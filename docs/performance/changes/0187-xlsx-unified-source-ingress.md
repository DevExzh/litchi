# Change 0187: source-backed high-level XLSX path ingress

Date: 2026-08-18

## Decision

Route `litchi::Workbook::open(Path)` for validated XLSX packages through the
existing positional `FileSource -> SourceBackedPackage ->
SourceBackedWorkbook` ownership chain. The prior path retained the complete
file in a `Vec<u8>` and then eagerly decompressed the workbook package through
`Workbook::from_bytes`.

One source-backed OPC package now supplies both the XLSX workbook owner and
core-properties metadata. Workbook catalog operations remain source-backed;
worksheet payloads are parsed only when selected by a fallible read. The
byte-backed `Workbook::from_bytes`, typed XLSX edit/publication APIs, and
non-positional platforms are unchanged.

The unified detector retains OOXML-before-ODF precedence. Valid ODS packages
still take the existing source-backed ODS path, and other recognized or
disabled OOXML families keep their previous behavior. Hard OPC input limits,
cancellation, source change, execution, I/O, and allocation failures propagate
instead of being hidden by the byte fallback. Names, count, metadata, and text
fence the source version. Fallible XLSX row/cell iterators preserve lazy parse
and source errors; chart/dialog/macro sheet kinds retain the established empty
grid projection.

## Correctness evidence

The focused feature matrices passed with warnings and deprecations denied:

- XLSX-only `litchi` library: 53 tests;
- XLSX+ODS `litchi` library: 63 tests;
- ODS+XLSB disabled-XLSX polyglot filter: four tests;
- strict XLSX+ODS library Clippy;
- strict all-target performance-harness Clippy;
- focused high-level XLSX selector oracle test;
- formatting and diff checks.

Regressions prove source ownership and eager semantic parity for worksheet
names/count, text, metadata, date system, and cells; suffix-independent XLSX
and ODS detection; OOXML/ODF polyglot precedence; deferred failure of an
unselected malformed worksheet; chart-sheet text parity; typed source-change
errors from retained worksheet handles; and hard OPC input-limit propagation
before byte fallback. Two independent current-tree reviews returned SAFE.

## Measurement contract

The control production revision is
`f3a11a9f465a23a5206bc40005296758216a22f0`. The identical committed harness
diff is applied to that control tree, so its raw reports deliberately record
`dirty=true`; control binary SHA-256 is
`a55b4f48e4f410e3ff2390d8fb1e60e6dd51ca4581e4116c4fd3a892bac6936d`.
The clean candidate harness revision is
`26607a1757fc750d94d14437d28b08a289724d14`, candidate binary SHA-256
`3f338fe18768034b93cb7cea9e7e073d8f26eb615dc2619fbd7c11b985690dcc`.

Fresh CPU-2-pinned A1/B1/B2/A2 release processes use 20 warmups and 500
retained samples per case. `xlsx_file_open` times exactly
`litchi::Workbook::open(Path)`. `xlsx_file_open_lifecycle` times that open plus
worksheet names, count, and full text. Corpus construction, temporary-file
publication, eager `from_bytes` oracle construction, metadata/projection
comparison, and source hash verification are outside timing.

The fixed corpus has four 48x48 worksheets, 9,216 logical cells, 17 archive
members, 4,226,429 archive bytes, and archive SHA-256
`dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`.
Deleting revision/dirty state and elapsed distributions produces canonical
non-timing projection SHA-256
`bde33e8db9184bdd99f93f124648bea595ae785c9ce27dcc3be738d918f7793c`
for all four legs.

The predeclared p50/mean/p95/p99 same-implementation drift ceilings are
5%/5%/10%/15%. A statistic is accepted only when both paired directions are
lower and both implementation drifts pass.

## Result

All eight named statistic cells pass the paired-direction and drift gates.
For `Workbook::open(Path)` alone, candidate reductions are:

| Statistic | A1 -> B1 | A2 -> B2 |
|---|---:|---:|
| p50 | 93.10% | 92.98% |
| mean | 92.97% | 93.02% |
| p95 | 91.96% | 92.34% |
| p99 | 91.59% | 92.13% |

The p50 values correspond to 2.092 ms -> 0.144 ms and 2.013 ms -> 0.141 ms,
or 14.49x and 14.25x lower elapsed time ratios.

For open plus worksheet names/count/full text, candidate reductions are:

| Statistic | A1 -> B1 | A2 -> B2 |
|---|---:|---:|
| p50 | 16.92% | 14.76% |
| mean | 16.54% | 14.75% |
| p95 | 15.45% | 14.54% |
| p99 | 18.30% | 14.35% |

The result is limited to warm, in-process elapsed time for this generated
media-rich XLSX corpus. The selectors emit no source-read counters and do not
enter the fresh-child filesystem/cache-state runner. No physical-I/O,
cold-cache, decompression, allocation, RSS, throughput, scaling,
real-producer, broad OOXML, edit/save, or iWork claim is made.

Artifacts:

- [machine-readable summary](../results/xlsx-unified-ingress-0195-summary.json)
- [artifact manifest](../results/xlsx-unified-ingress-0195-manifest.json)
- compressed raw A1/B1/B2/A2 reports listed in the manifest
