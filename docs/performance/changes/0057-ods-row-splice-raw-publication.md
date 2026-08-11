# Change 0057: ODS row-splice raw publication

Date: 2026-08-12

Status: accepted

## Decision

Retain the exact checked row-range publication produced by an eligible ODS
worksheet transaction through package emission. The ODF common package layer
now accepts that provenance-bearing `content.xml` splice and applies the same
raw ZIP preservation gate and logical rebuild fallback as its established
content replacement path.

The change adds one low-level ODF common publication entry point and otherwise
stays private to `litchi-ods`. It adds no archive type to an ODS public API, no
cache, runtime, lock, global state, dependency, unsafe code or persisted index.

## Problem and attribution

The accepted ODS row-local editor already discovered exact source row spans
and serialized only changed rows. It then flattened those edits into a
`String`. The package layer tried to rediscover one maximal source/target diff;
the row replacement was not a compact fragment under that conservative rule,
so publication fell back to a complete package rebuild and recompressed eight
unchanged 2 MiB media members.

On the deterministic media-rich one-cell edit/save control, the matched
profile attributed 42.58% of process self cycles to zlib medium deflate.
37.86% was below
`copy_source_files_from_except -> rebuild_package -> replace_content_xml ->
worksheet::Edit::commit`; the related longest-match frame added another
13.81% in that commit subtree.

## Implementation

- The row-local editor retains each ordered `start..end` range with its compact
  replacement fragment while assembling the same bounded result string.
- Every range is rechecked against an `XmlSourcePart` from the exact package,
  then added to an `XmlSplicePublication` as audited markup or deletion.
- `replace_content_xml_spliced` rejects foreign package provenance, unexpected
  path/content, and the existing 16 MiB content bound before using the shared
  raw member-preservation gate.
- Unsupported ZIP layouts, signatures, encryption-sensitive layouts, and
  every raw-preservation refusal keep the established logical rebuild. Signed
  changed packages still strip stale signatures through that fallback.
- Unified ODS staging records `content.xml` as provenance-spliced only when the
  worksheet commit used this exact path. Structural/table fallback and patch
  application keep the existing authored-part validation path.

## Correctness and safety boundaries

Integration tests prove that untouched row XML, `mimetype`, the manifest and
an opaque media member retain exact bytes while the edited row and
`content.xml` change. They also prove inverse restoration, typed cell readback,
touched-opaque-row refusal, foreign-but-byte-identical package provenance
refusal, expected-content mismatch refusal, and signed-package fallback with
signature removal and media retention.

Compact XML validation, bounded row serialization, complete package reopen,
worksheet snapshot parsing, exact typed-sheet readback, unified package-size
and security checks, patch construction, inverse behavior and independent
media verification remain mandatory. Structural edits and ineligible row
shapes still use the existing full-table/package fallback.

## Measurement method

Base revision: `8756bedff5e4006e86559f4d6968d4e8a278cb44`.

Exact release binaries, built from one identical unchanged harness:

- before: `316dba3dd63112de27f05c8dcaa551f148ad79fd441985ad8dd35c437550da98`
- after: `b381629b2b30a860b3547c310b62cca0e736d7990ac15ceac1196744780b8912`

The primary measurement uses three balanced pairs on CPU 11, alternating the
pair direction. Every leg has 10 warm-ups and 50 timed samples, yielding 300
samples per state. The fixed corpus has 2,048 cells plus eight deterministic
2 MiB incompressible media members: 16,790,689 archive bytes, archive SHA-256
`46b7f61cb74639115f6d120dc6498b97d6b310d51c78c4fb85ac60d6fc758b14`.
Every timed result is fully reopened and checks the edited cell, package/media
inventory, exact media payloads and patch semantics. Pooled raw samples and
distributions are in the
[`primary summary`](../results/ods-row-splice-raw-publication-primary-summary.json).

## Latency result

| Case | Before p50 | After p50 | p50 | Mean | p95 | p99 |
|---|---:|---:|---:|---:|---:|---:|
| Media-rich one-cell edit/save | 287.766 ms | 74.365 ms | **-74.16%** | **-74.17%** | **-74.11%** | **-74.10%** |

The before mean 95% interval is 288.078-289.225 ms; the after interval is
74.363-74.738 ms. Every primary leg improves independently.

## Guard cases

Two direction-balanced pairs cover ordinary ODS open, list, one-cell, cell
sweep, full-cell-text, exact no-op and one-edit/save cases for tiny, medium and
large generated shapes. Tiny and medium use 400 samples per state/case; large
uses 120. The materially timed p50 guard movements are:

| Shape | Open | Cell sweep | Full text | No-op edit/save | One edit/save |
|---|---:|---:|---:|---:|---:|
| Tiny | -2.10% | +1.43% | -7.72% | -1.67% | -7.21% |
| Medium | +1.38% | -2.87% | +1.55% | +1.32% | -5.85% |
| Large | -1.14% | +0.79% | -0.21% | +1.56% | -2.83% |

Nanosecond list/one-cell accessors remain too short for a process-layout claim.
Raw distributions are retained in the
[`tiny`](../results/ods-row-splice-raw-publication-guards-tiny-summary.json),
[`medium`](../results/ods-row-splice-raw-publication-guards-medium-summary.json)
and
[`large`](../results/ods-row-splice-raw-publication-guards-large-summary.json)
summaries.

## Allocation and memory evidence

Matched one-sample Heaptrack processes retain the same 140.05 MiB peak heap
and 1.78 KiB leak report. Allocation calls fall from 387,422 to 382,532
(-1.26%) and temporary allocations fall 4.65%. Heaptrack-inclusive RSS moves
164.26 to 164.41 MiB (+0.09%). Reports are
[`before`](../results/ods-row-splice-raw-publication-before-heaptrack.txt) and
[`after`](../results/ods-row-splice-raw-publication-after-heaptrack.txt).

Four uninstrumented GNU Time runs per state report mean maximum RSS of 156,573
KiB before and 156,724 KiB after (+0.10%, flat at process resolution). Raw
reports are stored under `results/ods-row-splice-raw-publication-time*.txt`.

## CPU evidence

Two matched `perf stat` pairs cover 20 timed edit/save samples per process.

| Counter | Before, mean/process | After, mean/process | Change |
|---|---:|---:|---:|
| Cycles | 42.478 billion | 16.215 billion | **-61.83%** |
| Instructions | 104.941 billion | 32.489 billion | **-69.04%** |
| Branches | 17.996 billion | 4.866 billion | **-72.96%** |
| Branch misses | 330.885 million | 19.617 million | **-94.07%** |
| Cache misses | 104.524 million | 87.918 million | **-15.89%** |

The exact matched profile removes the unchanged-media deflate/rebuild commit
subtree. Remaining zlib deflate is attributed to deterministic corpus creation
outside the timed edit; total sampled cycles fall 59.81%, with zero lost
samples. Reports are
[`before`](../results/ods-row-splice-raw-publication-before-perf-report.txt) and
[`after`](../results/ods-row-splice-raw-publication-after-perf-report.txt);
raw counters are stored under
`results/ods-row-splice-raw-publication-stat*.csv`.

## Validation

Passed on the final source:

- focused raw-member, provenance, signed-fallback and row-local transaction
  tests;
- complete all-feature ODF common and ODS test suites;
- warning-denied ODF common all-target Clippy and rustdoc;
- warning-denied ODS production Clippy, changed integration-test Clippy and
  rustdoc;
- all 33 performance-harness tests and warning-denied all-target Clippy;
- formatter, JSON parsing, link-target, whitespace and final-diff checks.

The ODF deprecation fixed in commit `1194fbc7f` remains clean under the
warning-denied ODF common Clippy and rustdoc gates. The repository-wide ODS
all-target Clippy command still reports the previously recorded unrelated
test-only lints; the changed test target is clean. There is no dedicated ODS
fuzz manifest in the current tree.

## Limitations and next work

This is a generated warm-memory, one-cell, same-topology, compact-row corpus.
It adds no new CRUD category, structural/resource-adding edit, encrypted
publication, real-producer coverage, cold-source measurement or streaming
surface. Broader source-backed OOXML edits, remaining native OLE2 final-owner
work, RTF formatting/media scenarios and ODF structural/real-producer paths
remain separate tranches.
