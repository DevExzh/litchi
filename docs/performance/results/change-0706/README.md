# Change 0706 evidence packet

This packet records the rejected XML attribute-cardinality pilot at revision
`cfd2aa4f3fb051f3914503fce597ef70593a26dd`. Production is restored; the eight
public attribute-cardinality integration tests remain as independent coverage.
There is no accepted 0706 performance claim.

## Decision

The candidate reuses the existing lexical attribute-layout scan to prove
zero/one attributes and skips duplicate-key bookkeeping only for those proven
starts. Starts with multiple attributes retain the checked iterator. The
candidate has no extra probe/replay pass, retained state, public API, new
dependency, or unsafe code.

The frozen native gate requires, for every medium and dense-sparse primary
shape/repeat, at least 2% reduction in total p50 and mean and 5% reduction in
publication p50. Seven of eight rows fail. The allocator gate passes, with
publication allocation calls changing from 19,197 to 365 (medium) and 36,573
to 365 (dense-sparse). Because the native gate fails, the candidate is
rejected and no speedup is admitted.

The native analysis retains 157 absolute-change flags over 5%: 124 paired
timing flags and 33 repeat-drift flags. The paired flags include 53 adverse
and 71 favorable changes. Repeat two is higher in 18 drift entries and lower
in 15; the repeat-drift range is −29.86% to +11.32%.
The baseline A/A diagnostic has a maximum absolute shift of 45.03% and no
fixed pass threshold. The allocator analysis retains 16 over-5% flags, all
favorable publication allocation/deallocation count or byte reductions; its
primary operation-region peak-live-byte changes are below 5%.

This is distinct from rejected 0529: 0706 reuses an existing lexical proof;
0529 added an unchecked attribute probe followed by checked replay. The gate
magnitude is comparable, while the implementation and evidence are separate.

## Packet map

| Evidence | Artifact |
| --- | --- |
| Frozen plan and order | [`plan.json`](plan.json) |
| Candidate hypothesis and ADR boundaries | [`design.md`](design.md) |
| Before/after source and manifest census | [`source-baseline.json`](source-baseline.json), [`source-candidate.json`](source-candidate.json) |
| Build commands, binaries and source receipts | [`build-baseline.json`](build-baseline.json), [`build-candidate.json`](build-candidate.json), build logs and receipts |
| Native timing and paired calculations | `native-*.json`, [`analysis.json`](analysis.json) |
| Allocator measurements | `alloc-*.json`, allocator receipts and [`analysis.json`](analysis.json) |
| Producer control | `producer-*.json` |
| Public differential oracle | `oracle-*-differential.json` and receipts |
| Auditor microguards | `oracle-baseline-A1.json`, `oracle-candidate-B1.json`, `oracle-candidate-B2.json`, `oracle-baseline-A2.json`, [`microguard-analysis.json`](microguard-analysis.json) |
| Oracle and test corrections | [`oracle-exact-correction.json`](oracle-exact-correction.json), [`oracle-preflight-correction.json`](oracle-preflight-correction.json), [`test-preflight-correction.json`](test-preflight-correction.json) |
| Static candidate review | [`review.json`](review.json) |
| Focused restored and candidate tests | `focused-restored.*`, `focused-baseline.*`, `focused-candidate.*` |
| Final scoped quality commands | [`integration/results.json`](integration/results.json) and logs |
| Reproducibility scripts | [`build.py`](build.py), [`capture.py`](capture.py), [`oracle.py`](oracle.py), [`analyze.py`](analyze.py), [`run-integration.py`](run-integration.py), [`run-evidence.py`](run-evidence.py) |

The earlier normalized oracle and pre-correction microguard derivation remain
archived in `normalized-oracle-preflight/` and
`microguard-analysis-preflight.json`. Raw captures were not rewritten. The
corrected microguard analysis compares `baseline-A2` with `candidate-B2` in
the second ABBA pair and reports 60 paired rows with 45 metric-level review
flags over 5%, alongside 14 baseline A/A drift flags over 5%.

## Reproduction and review boundaries

The packet was acquired with separate release native and allocator binaries,
CPU affinity 12, two ABBA repeats, primary 200-sample/20-warmup children,
guard 30-sample/10-warmup children, producer 200-sample/20-warmup children,
and microguard 100-sample/5-warmup batches of 100 calls. The public oracle
contains 4,148 cases and 29,036 calls, with identical baseline/candidate full
JSON outcomes and result digests and zero panics. The corrected exact oracle
recapture is retained alongside the archived normalized preflight.

The native region includes the returned `MultiSnapshot` drop and excludes sink
setup, remaining handle destruction, reopen, and oracle work. Allocator metrics
are operation-region diagnostics. No RSS,
physical cold-storage, native Office, cross-platform, or universal workload
claim is present. The conditional publication-instruction gate was not
required after the primary gate failed.

The focused restored run reports 63 passed and one ignored test; the candidate
run reports 66 passed and one ignored test. The final scoped commands are
`cargo fmt --all --check`, `cargo check -p xml-minifier --all-features
--all-targets --locked`, and the matching `xml-minifier` Clippy command with
`-D warnings`. All six broader evidence gates pass; their commands and logs
are retained in
`evidence/results.json`. The final packet audit and artifact manifest bind the
retained evidence after cleanup.

The owned 0706 target directory, frozen binaries, isolated worktree, and Python
caches were removed. `cleanup.json` retains six exact binary identities;
`audit-before-cleanup.json` and `audit-after-cleanup.json` record passing audits
on both sides of cleanup. From this batch's committed checkout, verify with:

```sh
python3 -B docs/performance/results/change-0706/audit.py
python3 -B docs/performance/results/change-0706/artifact-seal.py --check
```
