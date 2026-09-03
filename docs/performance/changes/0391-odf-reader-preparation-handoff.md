# Change 0391: ODF reader preparation handoff

Date: 2026-09-03

Status: implemented; matched release measurement completed without an authorized latency claim

`performance_claim: none`

`claim_authorized: false`

## Scope and mechanism

Smart detection for builds that contain both ODF and OOXML owners now gives a
canonical ODT, ODS, or ODP reader to one ODF-owned preparation operation. The
operation performs the fixed local `mimetype` probe, takes one bounded owned
snapshot, and builds one strict retained ZIP index. ODF classification and any
required lower-precedence OOXML probe share that snapshot through the same
`Arc<Vec<u8>>` allocation.

The existing indexed central-directory pass retains only opaque catalog facts:
canonical and duplicate-free names, bounded declared local spans (including
directories), data-descriptor, ZIP64, and encryption flags, archive end offset,
and the case-insensitive exact `[Content_Types].xml` marker. A proven ordinary
catalog skips OOXML. A marker or an uncertain fact set gives OOXML its
historical first opportunity, then returns the already prepared ODF format.

Strict-preparation failure is deliberately not optimized. The original `Vec`
allocation is recovered and the historical catalog, OOXML, and ODF byte
detectors run in their established order. This preserves behavior for aliases,
malformed or duplicate records, central/local `mimetype` disagreement, and
other strict-only refusal cases. Extra passes on this fallback do not support a
performance claim. The reader cursor is restored on every ordinary outcome;
a restoration failure stops lower-precedence detection from an unknown cursor.
The lower-precedence iWork reader path is unchanged.

## Correctness evidence

The focused suite proves one snapshot and one index for canonical ODT, ODS,
and ODP packages, exact shared-allocation identity, and no OOXML probe for a
proven ordinary catalog. It also covers the marker and unknown decisions,
polyglot precedence, trailing bytes, descriptors, encryption, ZIP64,
non-canonical and invalid names, duplicate members, declared spans that cross
the central directory (including a directory record), catalog input limits,
short reads, seek failures, cursor restoration, and the recoverable strict
`mimetype` mismatch fallback.

Independent stable Rust 1.98.1 runs passed 323 `soapberry-zip` tests, 287
`litchi-odf-common` tests, and the facade detection matrices for combined
ODF/OOXML, ODF-only, OOXML-only, and representative ODT/PPTX configurations.
One pre-existing `odt,docx` native-policy integration failure was reproduced
outside the changed detector path and is not treated as evidence for this
change.

## Matched release evidence

The exact baseline was `945510529dd7ba7ba45d9fe0cc1e98966fd52fa0`.
The candidate applied only the four production files in this change; that
patch's SHA-256 was
`0947d98ba97ab914d5317ed6cd92d33e6934d2ced0e5eda0bb4bed7ae089a123`.
The baseline and candidate release binaries had SHA-256 values
`a59a4f3ef4ccc3cf2bf3ce42a3648740680797687a70db50696b4fa71ec7e1b4`
and
`3d65777d19927d67ede5fc9daeab548bda711a2bc2987f1f0196aef5429a4763`.
Both used locked Rust/Cargo 1.98.1 builds, one build job and one benchmark
worker, isolated targets, five warmups, 30 retained samples, and an A1/B1/B2/A2
order. All four reports had identical configuration and per-case corpus
identities.

The table gives matched p50/p95/p99 nanoseconds as `A1 -> B1; A2 -> B2`:

| Corpus | p50 | p95 | p99 |
|---|---:|---:|---:|
| ODT tiny, 2,058 B | 1,950 -> 1,910; 1,960 -> 1,865 | 2,060 -> 1,980; 6,180 -> 1,980 | 2,070 -> 2,030; 75,940 -> 1,990 |
| ODT medium, 2,557 B | 1,935 -> 1,900; 1,960 -> 1,945 | 2,030 -> 2,030; 2,060 -> 2,010 | 2,070 -> 2,040; 2,070 -> 2,030 |
| ODT large, 28,329 B | 2,570 -> 2,720; 2,650 -> 2,720 | 2,650 -> 2,850; 2,720 -> 2,820 | 2,670 -> 2,880; 2,770 -> 2,910 |
| ODS tiny, 1,066 B | 1,480 -> 1,725; 1,510 -> 1,660 | 1,520 -> 1,790; 1,570 -> 1,750 | 1,540 -> 1,830; 1,580 -> 1,750 |
| ODS medium, 7,011 B | 1,650 -> 1,890; 1,640 -> 1,860 | 1,720 -> 1,940; 1,700 -> 1,940 | 1,750 -> 1,960; 1,710 -> 1,960 |
| ODS large, 98,941 B | 4,350 -> 3,950; 3,470 -> 3,930 | 4,510 -> 4,080; 3,630 -> 4,060 | 4,541 -> 4,090; 3,810 -> 4,070 |
| ODP tiny, 2,139 B | 1,960 -> 1,950; 1,925 -> 1,935 | 2,030 -> 2,050; 2,000 -> 2,040 | 2,040 -> 2,070; 2,030 -> 2,040 |
| ODP medium, 2,268 B | 2,140 -> 1,965; 2,075 -> 2,005 | 2,210 -> 2,070; 2,130 -> 2,180 | 2,360 -> 2,100; 2,220 -> 2,200 |
| ODP large, 3,335 B | 2,050 -> 2,045; 2,050 -> 2,040 | 2,220 -> 2,100; 2,140 -> 2,120 | 2,330 -> 2,160; 2,150 -> 2,120 |

The result does not pass the host-stability or regression review gates. The ODT
tiny baseline mean drifted by 133.60% because of a 75,940 ns outlier, while the
ODS large baseline p50 drifted by -20.23%. The candidate also regressed both
ODS tiny legs by 9.93% to 16.55%, both ODS medium legs by 13.41% to 14.55%,
and the first ODT-large leg by 5.84%. These individual results are retained as
decision evidence, but their mixed direction and unstable controls authorize no
latency claim.

## Claim boundary

The test-only snapshot/index and OOXML-probe counters establish mechanism and
precedence, not latency, allocation, RSS, physical-I/O, decompression, copy,
throughput, or cold-cache improvement. The matched release result failed the
stability and regression review gates, so this change authorizes no performance
claim.
