# Change 0416: local ZIP64 read and preservation compatibility

`performance_claim: none`

`claim_authorized: false`

## Problem and implementation

Python can force ZIP64 local framing for a small entry while retaining ordinary
central size fields and a ZIP32 archive tail. The previous reader selected
descriptor width from central/tail metadata and refused that valid source in
strict borrowed access and source-backed preserving edits.

The archive owner now derives 64-bit descriptor width from validated local
ZIP64 sentinel fields and extras. Local-only framing requires version 45 or
later and both size sentinels. Descriptor-backed entries require canonical zero
placeholders; seekable entries without descriptors require exact local/central
size and CRC agreement. Complete CRC and both sizes select the descriptor
interpretation, including unsigned descriptors whose CRC equals the optional
signature. Slice readers, positional readers, stream verifiers and preservation
indexing share that decision. Ordinary positional reads allocate no new local
extra buffer unless a size sentinel requires it.

Untouched member spans and central records retain their bytes. A targeted OPC
XML edit is checked against a Python source containing Store, Deflate, package
metadata, relationships and opaque content. Exact no-ops retain exact archive
bytes. The private ZIP32-only preservation policy explicitly refuses local
ZIP64 framing; ODF-specific publication and repair boundaries are checked
separately.

The framing follows [PKWARE APPNOTE sections 4.3.9 and 4.5.3](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT)
and [Python's `force_zip64` writer interface](https://docs.python.org/3/library/zipfile.html).

## Contract and coverage

| Requirement | Evidence boundary |
| --- | --- |
| ADR 0002/0010/0011/0024 ownership | ZIP grammar remains in `soapberry-zip`; no public physical type or dependency edge added. |
| ADR 0003/0006 preservation | Raw untouched records and exact no-op bytes; source immutability and before-output/midstream version failures. |
| ADR 0005 resources | Existing metadata/entry/output limits, cancellation and bounded non-seek sink failure checks; no new ambient I/O or workers. |
| ADR 0008 verification | Independent Python fixtures, complete Rust/Python payload verification, malformed fields, public OPC edits and scoped native artifact guards. |

The 13 deterministic fixtures cover small and 256-member archives, Store and
Deflate, signed and unsigned descriptors, a payload with the descriptor-marker
CRC, empty members, seekable local sizes, and central/local ZIP64 with ordinary
or ZIP64 tails. Three existing LibreOffice-resaved DOCX/XLSX/PPTX artifacts are
indexed as common-path guards; no new native application session is claimed.
The global ZIP64-tail fixtures remain correctness cases because the existing
borrowed Store API returns `None` for that archive-level form.

The existing ZIP fuzz target now reaches bounded strict borrowed access and
ReaderAt preservation indexing, in addition to central parsing. ZIP and OPC
sanitizer smoke results, exact flags, seeds and lockfiles are retained.

## Measurements and disposition

The [evidence bundle](../results/change-0416/) records the protocol, identities,
raw ABBA vectors, process RSS, independent fixture verification and check logs.
The control is `4794d2e5e`; the final candidate identity is in `identities.json`.
Warm immutable input is loaded before timing. Borrowed timing includes Store
CRC scans and per-file metadata/overlap validation; indexed timing includes
fresh index/scratch construction and preservation validation. Neither path
decompresses Deflate payloads during the performance guard. The metadata
oracle is prepared outside the timer, with comparisons and digest inside it.
Capability refusals remain separate from successful-work latency ratios.

## Guard results and review

Control `4794d2e5e` and candidate `d18cd04a2` use identical probe source and
locks. Both captures contain 11 rows and 44 processes; the follow-up uses
1,000 samples after 100 warmups. All per-row metadata observations agree.
All seven local-only capability fixtures refuse on control and succeed on
candidate in both modes, in each capture. No refusal latency enters a ratio.

| Fixture / mode | A1 → B1 p50 (µs) | A2 → B2 p50 (µs) | Paired p50 change |
| --- | ---: | ---: | ---: |
| central64-many256 / borrowed | 1277.736 → 1246.600 | 1278.815 → 1246.465 | -2.44% / -2.53% |
| central64-many256 / indexed | 84.680 → 86.070 | 85.760 → 85.710 | +1.64% / -0.06% |
| central64-tiny / borrowed | 0.650 → 0.640 | 0.630 → 0.670 | -1.54% / +6.35% |
| central64-tiny / indexed | 1.710 → 1.710 | 1.710 → 1.700 | +0.00% / -0.58% |
| zip32-many256 / borrowed | 1140.300 → 1179.650 | 1136.380 → 1183.864 | +3.45% / +4.18% |
| zip32-many256 / indexed | 81.710 → 83.940 | 84.490 → 81.950 | +2.73% / -3.01% |
| zip32-tiny / borrowed | 0.610 → 0.620 | 0.610 → 0.620 | +1.64% / +1.64% |
| zip32-tiny / indexed | 1.630 → 1.640 | 1.630 → 1.640 | +0.61% / +0.61% |
| native-docx / indexed | 4.520 → 4.570 | 4.510 → 4.600 | +1.11% / +2.00% |
| native-pptx / indexed | 17.480 → 17.820 | 17.540 → 17.800 | +1.95% / +1.48% |
| native-xlsx / indexed | 5.380 → 5.260 | 5.210 → 5.300 | -2.23% / +1.73% |

The original capture flags one candidate p99 increase, +5.88% for tiny
central-ZIP64 borrowed access. The follow-up flags the same row in the second
pair: p50 630 → 670 ns (+6.35%) and p95 660 → 700 ns (+6.06%).
No other follow-up row crosses the 5% review threshold. Both summaries retain
every p50/p95/p99/RSS comparison and within-revision drift. There is no clean
latency-regression pass and no speedup claim. Follow-up process RSS spans
2,176–2,760 KiB; no RSS comparison crosses 5%. Retain this necessary input and
preservation capability with the small central-ZIP64 tail/median trigger open;
the next performance investigation should isolate that sub-microsecond path
before changing its implementation. No aggregate hides this row.

## Resource diagnostics

All seven matched indexed fixtures have identical logical ReaderAt call and
accepted-byte counts between revisions. ZIP32 tiny uses 14 calls / 650 bytes;
ZIP32 many256 uses 1,030 / 76,890. Candidate local-only many256 uses 1,030 /
84,058. These are metadata-range observations over memory; reads can include
payload bytes incidentally, and they are not physical device or remote calls.

Separate whole-process heaptrack runs on ZIP32 many256 indexed access include
100 samples, 10 warmups, source/oracle setup and JSON output. Both revisions
record 143,646 allocation calls and 530.62 KB peak heap. `heaptrack_print`
reports 28,167 temporary allocations; raw heaptrack stderr reports 28,166.
Both sources are retained, and this one-count provenance difference is not
used as an optimization result. Heaptrack reports 544 retained bytes at process exit for each; that
number is not an object-lifetime leak diagnosis. Its RSS includes instrumentation
and is excluded from normal RSS comparisons. These totals do not establish
operation-local allocation attribution or local-only input memory behavior.

Separate CPU/PMU processes each run 20,000 samples and 100 warmups on the same
ZIP32 many256 indexed fixture. Control/candidate sampling records 337/339
samples with zero lost samples. The retained reports and flame graphs identify
name normalization, preservation indexing and copying as active work. These
short paired diagnostics do not attribute the tiny borrowed guard trigger or
establish a statistically supported CPU improvement. Exact cycles, instructions,
branches, branch misses and page faults remain in the two stat CSVs. Control
and candidate instructions total 33,968,036,179 and 34,955,721,524 (+2.91%);
cycles total 7,603,116,888 and 7,505,690,538. These single whole-process
captures are descriptive, not a hardware-counter improvement claim.

## Validation and limitations

The final ZIP/OPC/ODF all-feature/all-target suite passes 1,246 tests. The 12
ODT embedded-resource GC integration tests and its new local-framing unit test
pass. Both explicitly selected large ZIP64 tests pass; 43 doctests pass and two
are ignored. Scoped warning-denied Clippy and four-crate warning-denied rustdoc
pass. Clippy retains the five existing command-level allowances documented in
0415. Changed-file formatting passes; full workspace formatting reports only
unchanged `litchi-docx/tests/glossary_authoring.rs`. Boundaries, CRUD coverage,
eight strict registered claims and report classification pass. Initial compile,
fixture and offset failures are retained alongside successful final checks.

ZIP and OPC each complete 1,000 AddressSanitizer/sanitizer-coverage fuzz runs
from the 13 fixtures. This is bounded smoke coverage. The generator reproduces
all fixture bytes and its manifest exactly. `verify.py` checks source/binary
capture bindings, corpus hashes, run order and argv, raw-vector summaries,
diagnostic hashes/observations and required check results.


## Remaining limits

The prepared ODF catalog observes central/tail ZIP64 metadata only. A seekable
local-only sentinel is not currently reflected in that catalog fact; this
pre-existing classification gap remains open. This batch does not establish
semantic row/paragraph/slide streaming, all four append scenarios, the complete
failure matrix, cold or remote input, native-device I/O, worker scaling or the
full non-iWork performance program.
