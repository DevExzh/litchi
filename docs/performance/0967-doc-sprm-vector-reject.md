# 0967 — DOC vector-free strict SPRM reject

**Decision: reject the candidate; production remains unchanged.** The
independent raw audit found complete custody and semantic evidence, then
applied the frozen performance gates. Nine real `FloatingPictures` latency
cells and four allocator RSS guards fail. The candidate is not adopted and no
CPU, latency, or allocation speedup is claimed. The authoritative v2
post-capture audit and decision are now bound below; they re-audit the retained
capture without rerunning measurements.

## Candidate and source boundary

The candidate is a private borrowed SPRM walk used by
`RevisionEditor::revisions()` for revision enumeration on a strict
no-revision body. It validates
the complete CHPX/PAPX body, records whether a revision trigger was seen, and
avoids constructing the owned `Vec<Sprm>` for a trigger-free body. A
trigger-positive or marker-like body keeps the existing owned parser and
revision metadata path, so the positive path can walk twice. The walk cannot
stop at a trigger; malformed suffixes and existing error precedence remain
observable.

The 0966 census motivated this bounded hypothesis with short nonempty SPRM
groups. It did not establish CPU or allocator ownership, and it supplies no
global speed bound. The source-backed `fkp_has_revision_sprm` refusal route
remains a separate contract. No source-backed admission, raw-group ownership,
public API, snapshot, commit, or package ownership behavior changed.

The sealed external candidate is based on
`893072657ad8b1d726bfce351b5ec2e98f084e94` and changes exactly these three
production files:

```text
crates/litchi-doc/src/sprm.rs
crates/litchi-doc/src/tracked_revision/codec.rs
crates/litchi-doc/src/tracked_revision/package.rs
```

The source tree is `/home/zhuhe/litchi-performance-0967/source-candidate`.
The main checkout has no candidate source change.

## Semantic and build qualification

The independent source review found no candidate blocker. The full DOC library
suite passed **1,146 tests with 2 ignored**. The frozen literal Rust parser
oracle passed its three checks and covers all **65,536 opcodes**, including
strict prefixes, one-byte tails, and variable-length boundaries. Revision
controls passed positive insertion metadata, late malformed-suffix ordering,
duplicate property-marker last-record behavior, and a late error in a
non-marker group.

The matched ordinary/native and allocator control builds used immutable
baseline and candidate source trees, the same harness and lockfile, and
separate target directories. Functional controls covered positive revisions,
dense 1,024-marker fallback, a false marker, and the real
bookmark/delete/redline query and reopen refusal. These controls establish
semantic custody; they do not themselves establish performance.

The known baseline rustdoc failure remains in four pre-existing diagnostic
files (`bookmarks.rs`, `comments.rs`, `rsids.rs`, and `sections/codec.rs`).
Those files are byte-identical between baseline and candidate, and candidate
files have no diagnostics. This retained baseline condition is not a 0967
candidate failure.

## Frozen capture

The v5 protocol measured 24 real cells: two fixtures
(`FloatingPictures` and `NoHeadFoot`) × four edit counts (1, 2, 8, 32) ×
three retention modes (zero, default, release). Each real cell used 12 fresh
counterbalanced native pairs, five warmups, and 80 measured samples. Each
allocator cell used three fresh pairs, zero warmups, and one measured sample.
The four timed controls were generated-positive `e1`, `dense1024` `e32`,
false-marker `e1`, and false-marker `e32`, with 12 native pairs and three
allocator pairs each. Positive-marker and dense controls were mandatory because
the candidate's fallback can pay for a second walk.

The native noninferiority bound was p50 and p95 paired-bootstrap upper CI at
most **1.05** for every real and timed-control cell. The native benefit branch
required an endpoint p50 median at most **0.95** with upper CI below **1.00**.
The allocation branch required endpoint allocation-call ratios at most **0.90**
for `FloatingPictures/e32/default` and **0.97** for
`NoHeadFoot/e32/default`, other real-cell call medians at most **1.01**,
requested/allocated bytes at most **1.01**, and peak live/process RSS at most
**1.05**. P99 was descriptive only; no retry or replacement was allowed.

The root capture terminated with runner exit 0 under execution session 66783.
The independent audit verified **896/896** terminal packets, schedules,
oracles, binary custody, and no-retry policy. It also verified both AA
endpoint readiness cells. AA is a readiness check only and is not pooled with
candidate ratios or used as candidate benefit evidence.

## Hard gate results

The authoritative audit classifies the fixed cohort as
`reject_candidate`. Nine `FloatingPictures` cells fail the native
noninferiority bound. The table gives the p50/p95 bootstrap upper limits; a
cell fails when either exceeds 1.05.

| cell | p50 upper | p95 upper |
| --- | ---: | ---: |
| `e1/zero` | 1.124477 | 1.050874 |
| `e1/default` | 1.118693 | 1.119583 |
| `e1/release` | 1.042763 | 1.101670 |
| `e2/zero` | 1.163094 | 1.144691 |
| `e2/default` | 1.094627 | 1.088374 |
| `e2/release` | 1.066431 | 1.045918 |
| `e8/zero` | 1.009646 | 1.051593 |
| `e8/default` | 1.095644 | 1.137535 |
| `e8/release` | 1.084043 | 1.168135 |

These are confidence-limit gate failures, not a claim that every one of the
nine candidate medians is slower; their medians are mixed. The other 15 real
native cells and all timed-control native cells pass native noninferiority.

Four allocator RSS guards fail:

| scope | guard ratio | bound |
| --- | ---: | ---: |
| `FloatingPictures/e2/default` | 1.069050 | 1.05 |
| `FloatingPictures/e8/release` | 1.081315 | 1.05 |
| `FloatingPictures/e32/release` | 1.050566 | 1.05 |
| `dense1024/e32/default`, control pair 3 | 17,952 / 16,688 = 1.075743 | 1.05 |

The RSS violations are observed hard guard failures. The audit applies the
declared bound without a post-hoc host-noise waiver. Peak live bytes and the
allocation-byte values in these rows do not erase an RSS failure.

The endpoint allocation-call signals are favorable as descriptive measurements:
`FloatingPictures/e32/default` is **0.809241** candidate/baseline and
`NoHeadFoot/e32/default` is **0.946489**. The allocation-call guard also
passes, and requested bytes and peak live values stay within their respective
limits in the endpoint rows. These observations do not qualify the candidate:
the native benefit branch fails because endpoint p50 medians are
`0.961343` (upper `0.975828`) and `0.986274` (upper `0.992034`), while the
real and control RSS/native gates above fail.

## Evidence, reducer history, and disposition

The authoritative v2 raw audit is
[`results/change-0967/review/independent-raw-audit-0967-v2.json`](results/change-0967/review/independent-raw-audit-0967-v2.json),
SHA-256
`2712e39c456d7528c4c9347d6f085dd33253bf43d389d265b979ffee606b29ea`.
The v2 decision receipt is
[`results/change-0967/review/independent-decision-0967-v2.json`](results/change-0967/review/independent-decision-0967-v2.json),
SHA-256
`27ab3f1cdad3262a3d374d0dea83a065c2f15909ede165e96bc06518fc2cb1dd`;
it records `reject_candidate`, complete evidence, and
`candidate_adoption:false`. The separate read-only v2 audit helper is bound
by SHA-256
`8215eaded6f6b0b6b45e42bf2fef603f5566c1ad0ec9c3a2b8fbfeca63451f77`.

The custody incident receipt
[`results/change-0967/review/audit-custody-incident-0967-v1.json`](results/change-0967/review/audit-custody-incident-0967-v1.json)
records that a helper invocation overwrote the derived v1 audit output. The
sealed v1 bytes are unavailable byte-for-byte; only the overwritten derived
audit is retained as an incident artifact. The v2 audit is a read-only audit
of the unchanged 896-packet capture, with no measurement or workload rerun.
The capture terminal receipt and manifest remain bound to execution session
66783 and manifest SHA-256
`273b1de3e0ba8fe53d66a48f6d31a054796deccf404ac048223ad9ad5f6396df`.

The original root reducer output is retained unchanged at
[`results/change-0967/analysis/reduction-v1.json`](results/change-0967/analysis/reduction-v1.json),
176,748 bytes, SHA-256
`f6a77df77fa34b4e811a6281c49341825c660faca4106794224c8ecb5fd044ed`.
Its `pending_incomplete`/`no_decision` classification is a reducer
classification defect: a timed allocator failure was appended to `reasons`,
then `complete = not reasons` made the gate calculations short-circuit and
treated a failed declared performance gate as missing evidence. The raw
24-cell native and 24-cell allocator sets were complete. The independent
read-only audit recomputed the raw packets, preserved the thresholds, and
correctly classified the fixed cohort as rejection. No benchmark reducer,
threshold, or workload change was used for that decision; the v2 helper is a
separate read-only audit.

No source optimization is applied. The rejection is scoped to this candidate,
source revision, fixtures, host, build, and frozen v5 metrics. It does not
establish a global DOC slowdown or a host-noise explanation, and it does not
change the 0966 census interpretation. Any future SPRM work requires a new
source/design review and fresh gates.

## ADR disposition

| ADRs | disposition |
| --- | --- |
| [0001](../adr/0001-priorities-and-api-layers.md), [0005](../adr/0005-io-memory-and-performance.md) | Correctness, preservation, bounded resources, and named uncertainty gates precede the descriptive allocation signal. |
| [0002](../adr/0002-crate-topology.md), [0024](../adr/0024-current-topology.md) | Candidate source ownership stays within `litchi-doc`; no dependency or public-layer change is made. |
| [0003](../adr/0003-snapshots-edits-and-patches.md), [0006](../adr/0006-validation-security-and-compatibility.md) | Full validation, fallback, semantic controls, refusal behavior, and immutable publication contracts remain unchanged. |
| [0008](../adr/0008-migration-and-verification.md) | Failed hard gates remain visible; rejection leaves production unchanged and does not promote the candidate. |
