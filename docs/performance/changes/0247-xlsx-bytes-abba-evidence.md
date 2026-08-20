# Change 0247: XLSX bytes-facade ABBA evidence (0230 package)

Date: 2026-08-21

Status: independently audited narrow bytes-facade evidence; only the listed
statistics are accepted, with no broad XLSX performance claim

## Evidence-package identity

This is the XLSX half of the 0229/0230 evidence tranche. The neighboring DOCX
record is [change 0229](0229-docx-text-binding-tracker.md); its corrected
resource report is retained at
`/home/zhuhe/CodeProjects/litchi-perf-resource-0229-docx-c70283f0c-20260821-corrected.json`
with SHA-256
`52e57a41dea6a5fc2ef100d42fe9aec58c89e2db9603e7d9581ee19316dd5525`.

The XLSX evidence package retains its historical `0230-xlsx-bytes` identity;
this documentation record uses change number 0247 because
`0230-operation-write-and-proc-io-metrics.md` already occupies 0230. The
package is retained outside the repository at
`/home/zhuhe/CodeProjects/litchi-perf-evidence/0230-xlsx-bytes-20260821/`.
Its manifest SHA-256 is
`7f84e54f30ef6fb0cfa8159b4ab007113a0d6c18c3c1f628310e7077da02d17e`.
The compressed `summary.json` SHA-256 is
`8642c66d3a7b2dc74f6da8ff85333c2f29dd7725e03e0b8920fd2af893150c3f`; its
canonical report SHA-256 is
`f64217647bd3f5da962bdc01d352bd4dfdfb59e69d5db09b0ca62590f7fbbf21`.

The package summary identifies control revision
`134a1eb4195ab81a01d93912e94ac0a2d3a08cda` for A1/A2 and candidate revision
`93249204898e8d5d5846becfb7a0d79df28e6b93` for B1/B2; every leg reports a
clean worktree on the recorded host. The candidate revision's Git tree is
`8768a608daa7b35acbfeb7f2ea8ccda5328b3cb2`, exactly the tree of landed commit
`9aaf9d1369096d678e2c74a5e94fdffc1037914b`. The candidate commit and landed
commit remain distinct, but their tracked source trees are equal, so the
measured candidate source content is the landed content.

The retained package members have these identities:

| Member | Compressed/file SHA-256 | Canonical SHA-256 |
|---|---|---|
| `0230-xlsx-bytes-manifest.json` | `7f84e54f30ef6fb0cfa8159b4ab007113a0d6c18c3c1f628310e7077da02d17e` | — |
| `summary.json` | `8642c66d3a7b2dc74f6da8ff85333c2f29dd7725e03e0b8920fd2af893150c3f` | `f64217647bd3f5da962bdc01d352bd4dfdfb59e69d5db09b0ca62590f7fbbf21` |
| `a1.json.zst` | `eec1c348e8f0a73c83420b53fdc5950d21d04cfe3cd93bd5478be32c6089be3c` | `2b852dc2f4acc7a7cfba0ac9e473dba83f497e6af8a4e2e6d504682ad9ad8698` |
| `a2.json.zst` | `d905bcc9c9a6cca1a80fa70df0af5fff881fdb792874e595776fe7288c0bb184` | `457e77596bb65872aff9c5a19842cb1e3bf951fafabaa617c3f17d05880d40e0` |
| `b1.json.zst` | `1692dc64977c1ec157d87afafc244036ff2e6aa2594e4b0facfdc462f3d2c012` | `a9fba28db7cf38db9af1aba0ab2230d077a699fa10906b8eeea4a643d78db472` |
| `b2.json.zst` | `9fb2951f68277d8743c0ab8fc07a78cc4909b167eef1a466bff86eae7664255d` | `0dd53f86748cd5c72b278031c033a1c14d65b0b2310042f7bbf6977e7fcb5f67` |

## Strict timing scope

The two opt-in selectors time only an in-memory bytes facade:

| Selector | Timed interval |
|---|---|
| `xlsx_bytes_open` | `litchi::Workbook::from_bytes(Vec<u8>)` construction |
| `xlsx_bytes_open_lifecycle` | That construction followed by worksheet names, worksheet count, and full text |

The typed eager XLSX projection, archive hash, and independently opened
OPC/property metadata digest are outside the timer. The fixed deterministic
corpus is `xlsx-cell-values-medium`: a four-sheet 48×48 scalar grid with media,
17 archive members, 4,226,429 compressed bytes, 4,231,168 uncompressed bytes,
and archive SHA-256
`dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`.

The ABBA order was A1 control, B1 candidate, B2 candidate, A2 control on CPU 2,
with 30 warmups and 500 retained samples per leg. Positive values below mean
lower candidate elapsed time. Drift ceilings are 5%/5%/10%/15% for
p50/mean/p95/p99.

## Exact accepted and rejected statistics

### `xlsx_bytes_open`

| Statistic | A1→B1 | A2→B2 | Verdict |
|---|---:|---:|---|
| p50 | 84.39874499813368% | 84.70055878341924% | rejected: candidate drift |
| mean | 84.45052083066089% | 84.80127971958677% | rejected: candidate drift |
| p95 | 83.30728149016544% | 83.9710534196833% | accepted |
| p99 | 82.43745022283225% | 84.71483001747943% | accepted |

The open selector withholds p50 and mean because candidate same-implementation
drift is −5.587297869314617% (p50) and −5.369007891267394% (mean), exceeding
the 5% ceiling. The accepted p95/p99 candidate drifts are −9.389817282448403%
and −6.134103614007369%, within their 10%/15% ceilings. No statistic has an
adverse-both reading.

### `xlsx_bytes_open_lifecycle`

| Statistic | A1→B1 | A2→B2 | Verdict |
|---|---:|---:|---|
| p50 | 6.079170638931871% | 9.578368272364624% | accepted |
| mean | 5.247091173935794% | 9.610098142558385% | accepted |
| p95 | 3.49949680288066% | 10.739855634989409% | accepted |
| p99 | 5.806086003261778% | 8.731088541948674% | accepted |

All four lifecycle statistics pass both implementation-drift checks. The
candidate drifts are −4.931920220107789% (p50), −4.775438262581311% (mean),
−5.891536252526927% (p95), and −3.689570098077196% (p99); none exceeds its
corresponding ceiling.

These are the complete accepted statistics for this bytes-facade run. The
rejected open p50/mean readings remain visible as drift-rejected evidence and
are not converted into an all-statistics claim.

## Claim boundary

This record is limited to the in-memory `Workbook::from_bytes(Vec<u8>)`
construction and the named in-memory lifecycle above. It makes no claim about
`xlsx_file_open`, source-backed range selectors, filesystem or cold-cache
behavior, physical I/O, source preservation, output-byte identity, sink or
operation metrics, allocations, RSS, edits, saves, other XLSX shapes, or
real-world producers. The lifecycle name means the timed bytes-facade
projection, not a source-file lifecycle.

The package verification confirms case/corpus/configuration and stable
environment identity and recomputation of statistics from samples. Source,
sink, output, and operation-metrics identities are consistently absent, so
the package does not establish those properties. The separate historical
0230 operation-metrics schema record is not revised by this evidence note.

Verification for this documentation-only update is Python hash/link checking
plus `git diff --check`; no Cargo command or benchmark is part of this record.
