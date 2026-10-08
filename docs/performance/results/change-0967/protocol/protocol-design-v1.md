# 0967 prospective protocol: borrowed SPRM revision scan

**Status:** revised pre-freeze design, 2026-10-07. No 0967 source build,
benchmark workload, Office capture, or timing result has been run. This file
sets the final-plan inputs and gates; it is not an adoption result.

The candidate is a private borrowed SPRM scanner used from
`RevisionEditor::revisions`. It completely validates every CHPX/PAPX group.
A group with no possible tracked-revision opcode may avoid constructing the
existing owned `Vec<Sprm>`; a possible marker falls back to the existing owned
parser and metadata path. A positive group can therefore be walked twice. The
protocol measures the real lifecycle and allocator behavior; an owned-parser
call count is only an optional diagnostic, not the adoption metric.

## 1. Custody and retained executable inventory

The 0967 base export records revision
`893072657ad8b1d726bfce351b5ec2e98f084e94` in
`/home/zhuhe/litchi-performance-0967/source-head.txt`. The candidate staging
copy is `/home/zhuhe/litchi-performance-0967/source-candidate`. The observed
pre-freeze production changes are confined to:

```text
crates/litchi-doc/src/sprm.rs
crates/litchi-doc/src/tracked_revision/codec.rs
crates/litchi-doc/src/tracked_revision/package.rs
```

Their staging hashes are observations, not final bindings:

| file | staging SHA-256 |
| --- | --- |
| `crates/litchi-doc/src/sprm.rs` | `7b28d3080210184494f1544c19ea715b0eacc57c485db9322e79ad6635c85659` |
| `crates/litchi-doc/src/tracked_revision/codec.rs` | `076c8515b7f9b5d072d324a2e5fa30e85f9578a6dc390514846fab7b38e62e69` |
| `crates/litchi-doc/src/tracked_revision/package.rs` | `9885ba87525f0629d19a2b44daa442501f4ccc0d8bf66a08766dd03eae015a25` |

Before build, root must create immutable baseline and candidate manifests from
the same base export and bind the final allowlist. Isolated test files may be
added only when the sealed coder/reviewer receipt names them. Any other changed
path, source-manifest mismatch, or source revision drift stops the experiment.

The existing `/home/zhuhe/litchi-performance-0967/design-v1.md` is a mechanism
note, not the source binding. Its historical text names an older source head;
the 0967 `source-head.txt` and final manifests are authoritative. The coder
readback and non-XLSX ranking reviews must be sealed before an executable is
accepted.

| role | retained executable | bytes | SHA-256 | source status |
| --- | --- | ---: | --- | --- |
| 0917 ordinary native | `/home/zhuhe/litchi-performance-0917/target/release/doc_retention_probe` | 55,146,832 | `2ae48b4f7548c1468c8e7691d56d89b231c4a63f52ed48383f07fe3609c59a13` | functional reference, source head `bb19ca74bd52044fbd2c1c7cc9875e975bb334c0` |
| 0917 allocator | `/home/zhuhe/litchi-performance-0917/target/release/doc_retention_probe_alloc` | 55,150,520 | `927606880090f9169e222254944df92366dd1cb0e1aa37fc0861234c466ccf1b` | functional reference, same old source custody |
| 0942 native profile | `/home/zhuhe/litchi-performance-0942/target-profile-v2/release/doc_retention_probe_0942_cfb_profile` | 126,326,672 | `92f85ce6019405fa1b0d97ec796947cd7ec240d2b8143f9cf06356d97409e70e` | correct DOC lifecycle/profile scope, source-v2 head `73559c87c37898b675100b02985181a7764c66d2` |
| 0942 allocator | no retained executable | — | — | `src/bin/alloc.rs` exists, but no binary is available |

The 0942 source-v2 listing is
`/home/zhuhe/litchi-performance-0942/source-manifest-v2.jsonl`, SHA-256
`bf99dd9d4ea021a29c3f003d3b38b6642019f70d9e664837ea8c061e040897f8`; its
receipt records `production_changes: false`. The 0917 archive is
`/home/zhuhe/litchi-performance-0917/source-base.tar`, SHA-256
`61df35b45519c95b43fb4194b0f679c69ce17aa8b39e5c112c0aea9b02f95863`.
These are historical custody, not proof that either old executable is a
whole-tree match for 0967. A retained binary may be used only after its full
source and harness manifest matches the frozen baseline; otherwise build a
paired baseline and candidate from the 0967 exports.

The matched 0942 harness files include:

```text
tools/perf-cfb-profile/src/lib.rs       b4893321ac9a9321a9055823a7f05ecb35e59319531913affc49342b8ccd73dc
tools/perf-cfb-profile/src/bin/alloc.rs 2542257a9ab68347d3e8133b5f65a2254cc1e439fef665e327e093470f365f5f
tools/perf-cfb-profile/Cargo.toml       41dee12e4eb7d1e3853a33a7743afdf6b5d6519d5e7456e61ce4978aa9177945
tools/perf-cfb-profile/Cargo.lock       343c5464ae1144cbb72284cbb3e9b52b5dff2b798d09462d3c0d6a74b317bac7
```

The native lane must use the original `tools/perf-cfb-profile` lifecycle, not
the 0966 counter-only executable. The retained 0942 profile binary has a
marker/barrier wrapper and its plan excludes timing claims; it is a scope
reference unless identical wrapper behavior is bound in both binaries. The
0967 five-percent comparison should use matched uninstrumented `run(false)`
binaries. The allocator binary is always a separate lane.

## 2. Candidate contract and correctness controls

The implementation may change only the private revision scan. For every CHPX
group and every PAPX body reached by the current loop, it must:

1. validate the complete SPRM grammar, including opcode, size code, variable
   length, checked extents, operand bounds, and all later records;
2. preserve the owned parser's error type and first-error order;
3. use a conservative possible-marker predicate, including false, malformed, and
   duplicate marker operands, before falling back to `strict_sprms`;
4. preserve metadata precedence, author/time/reason/RSID lookup, CP ranges,
   adjacent merging, move inference, revision limits, and open-time checks; and
5. keep raw CHPX/PAPX bytes authoritative after clone-first mutations.

No persistent parsed-vector cache or public API change is in scope. A positive
marker group may be double-walked only if the end-to-end control below bounds
and reports that cost. If no representative positive control can be sealed, the
candidate cannot be adopted unless it is changed to a verified no-double-parse
borrowed-event design.

The final controls must include these end-to-end routes through the same public
lifecycle (`open_bounded`, one `Edit`, repeated paragraph-0 replacements,
`commit`, output copy, reopen/readback, and `revisions()`):

- a valid revision-bearing control with character and property markers;
- a dense-marker control that exercises many positive groups and exposes the
  fallback double walk; and
- a false-marker control whose marker-like records do not emit a revision but
  remain fully validated.

A corpus DOC or deterministic writer-generated DOC is acceptable. It must bind
input/output bytes, author/timestamp tables, revision kinds and ranges, and
exact readback. Unit tests cover all marker opcode types, duplicate/false
operands, malformed tails, variable operands, truncation, and PAPX style-index
errors; no formal fixture must independently contain all five marker kinds.
A marker followed by a malformed tail must return that tail error before any
metadata lookup.

Every candidate and baseline control must preserve the existing DOC oracles:
fixture identity, stage text changes, final paragraph/public-reader readback,
retention-route output equality, no-op/source sharing, forward/inverse patches,
foreign/stale refusal, same-value and operation-limit atomicity, capacity
ceilings, zero/default/release behavior, untouched streams and directory
metadata, semantic projections, and semantic change counts.

## 3. Frozen workload matrix

The real fixtures are:

| fixture | bytes | SHA-256 | edits |
| --- | ---: | --- | --- |
| `NoHeadFoot.doc` | 26,112 | `45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48` | 1, 2, 8, 32 |
| `FloatingPictures.doc` | 335,360 | `41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d` | 1, 2, 8, 32 |

The real matrix has 24 cells: 2 fixtures x 4 edit counts x 3 retentions
(`zero`, `default`, `release`). The harness author must seal the final matrix
and add the three end-to-end controls above before root authorizes capture.
The retained functional control packets are correctness evidence only. A
separate matched control-profile harness measures the fixed control scope
before the real matrix: generated positive at e1/default, dense1024 at
e32/default, and false-marker at both e1/default and e32/default. Each timed
control cell uses the same 12 counterbalanced native pairs with 5 warmups and
80 samples, plus three allocator pairs with warmup 0 and sample 1. Its p50 and
p95 paired-bootstrap upper bounds must be at most 1.05; allocator
requested/allocated bytes, peak-live bytes, and RSS use the 1.01/1.05/1.05
guards. The real-positive query/reopen route remains a separate correctness
control and is not relabeled as a timed edit route. These timed controls are
required evidence for the double-walk bound and do not replace the 24 real
cells.

### Native lane

Run matched baseline/candidate child pairs for all 24 real cells. Use **12
fresh counterbalanced pairs per cell**: six baseline-first and six
candidate-first. Each child has **5 warmups and 80 measured samples**. The
native binaries emit no SPRM diagnostics, counters, or instrumentation.

For each child, retain elapsed samples and compute p50, p95, and p99 (rank 80,
the maximum of the 80 samples). Pair corresponding baseline/candidate children
by cell and order. Bootstrap the 12 paired child ratios to produce predeclared
95% confidence intervals for the p50 and p95 ratios. The p99/rank-80 value is
reported descriptively only; it is never an adoption gate.

This is 288 native pairs, or 576 native children, for the 24 real cells. The
positive, dense, and false-marker controls are run by the same harness and
bound in the final plan; their correctness does not count as a substitute for
real-cell pairs.

### Allocator lane

Run **three fresh paired repeats per real cell**, warmup 0 and sample 1, with
baseline and candidate allocator binaries. This is 72 allocator pairs and 144
children before controls. Report `allocation_calls`, requested/allocated
bytes, deallocated bytes, peak live/retained values, and process max RSS. The
allocator lane makes no elapsed-time claim.

No retained 0942 allocator executable exists. The retained 0917 allocator is
usable only as a historical functional reference until complete source and
harness equality is proven. Otherwise root must build source-matched baseline
and candidate allocator binaries before capture.

## 4. Noise control and native comparison

The A/A noise control is separate from the candidate comparison and is frozen
before capture. Use the same baseline executable in fresh AB and BA orderings,
with the same child settings. The preferred coverage is all 24 cells; a
minimal predeclared coverage is the two endpoint cells
`NoHeadFoot/e32/default` and `FloatingPictures/e32/default`. The final plan
must choose one coverage before launch; it may not select cells after seeing
results.

For every A/A cell and p50/p95 metric, calculate both directional ratios and
apply the symmetric ratio

```text
symmetric(r) = max(r, 1 / r)
noise = max(symmetric(r_AB), symmetric(r_BA))
```

using the same paired-bootstrap interval. If the symmetric upper bound exceeds
`1.05`, or either direction is incomplete, the baseline/candidate comparison
must not be interpreted. Stop before using candidate results; retain every raw
A/A packet and do not convert this readiness failure into an adoption or
rejection result. This catches either ordering's inflation; a one-direction
pooled ratio is not sufficient. There is no retry or post-hoc noise threshold.

For each candidate/baseline native cell, the upper bound of the paired
bootstrap interval must be **at most 1.05 for both p50 and p95**. This is a
non-inferiority condition for every real cell. A point-estimate median above
1.05 is still a review flag even when its interval upper bound passes. The
p99/rank-80 maximum remains descriptive and cannot be used to waive a p50/p95
failure.

## 5. Adoption gates

All correctness, custody, allocator, and per-cell native non-inferiority gates
must pass first. The A/A run is a separate readiness check only; its points are
never pooled into candidate ratios or used later to qualify or reject the
candidate. Adoption then requires one of these predeclared benefit branches:

1. **Primary allocation goal:** in `FloatingPictures/e32/default`, the
   allocator candidate has at least a 10% reduction in `allocation_calls`
   (paired median ratio at most `0.90`), and in `NoHeadFoot/e32/default` it has
   at least a 3% reduction (paired median ratio at most `0.97`) across the three
   fresh repeats. The 0919 authoritative route census records 19,024 total
   allocation calls and only 1,168 SPRM-subtree events for NoHeadFoot/e32;
   even an ideal removal of every such event cannot produce a 10% whole-
   lifecycle reduction. The 3% gate is therefore a predeclared, feasible
   bounded-opportunity target, not a new measured outcome. The larger
   FloatingPictures opportunity retains the 10% target.
2. **Native benefit:** after all-cell non-inferiority, at least one of those
   two primary endpoint cells has a paired p50 median ratio at most `0.95`
   and a 95% bootstrap upper bound below `1.00`. This combines the practical
   five-percent point improvement with statistical evidence; the upper bound
   is not required to be `0.95`. All other cells still satisfy native
   non-inferiority.

A call-count reduction without either branch is descriptive and cannot qualify
the candidate. No new SPRM counters are required in the native binaries and
no native counter is part of the adoption gate. After correctness, root may
run one small one-cell sidecar diagnostic with counters to demonstrate
marker-free versus positive fallback behavior; it is separately hashed and
cannot alter native samples or rescue a failed gate.

Memory and allocation guards apply to every real cell and every control. Across
the three paired allocator repeats, no candidate ratio may exceed:

```text
requested/allocated bytes: 1.01
peak live bytes:          1.05
process max RSS:          1.05
```

For every cell outside the two primary endpoints, `allocation_calls` must also
have a paired median ratio at most `1.01`. If the native-benefit branch is used,
the endpoint `allocation_calls` ratios must also be at most `1.01`; the endpoint
reductions above are the stricter requirement for the allocation branch. These
are hard guards; pooling cannot hide a repeat that exceeds a memory bound.

## 6. Binding, reduction, and rejection

Before root authorizes a build or workload, the final plan must bind immutable
baseline/candidate manifests, exact changed-file allowlist, mechanism and
coder/reviewer receipts, harness and lockfile hashes, fixture hashes, binary
commands and SHA-256 values, host state, order schedule, control packets,
allocator metric schema, and the chosen A/A coverage and native benefit form.

The reducer must validate every child terminal, source/binary/fixture hash,
cell, order, warmup/sample count, route, retention, correctness oracle, and
metric field before calculating any ratio or bootstrap interval. Failed,
missing, duplicate, timed-out, malformed, or hash-mismatched packets are
retained and stop the run. There is no retry or replacement.

Reject the candidate for any correctness or source-custody mismatch, absent
revision-bearing/dense/false-marker controls, failed native non-inferiority,
any memory/allocation guard violation, missing allocator evidence for a
complete adoption claim, no benefit branch, or an unbounded positive-marker
double walk. An A/A readiness failure leaves the comparison uninterpretable and
produces no adoption or rejection decision; its raw packets remain retained.
A native median above 1.05 remains a review finding even when the confidence
interval passes; it cannot be silently omitted or relabeled. Requested bytes,
p99 rank 80, A/A points, or an optional sidecar counter cannot be used as an
adoption claim.

This revision is a new pre-freeze protocol. It authorizes no build, workload,
or Office capture.
