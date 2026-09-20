# 0707 current XLSX planning attribution packet

This packet records a current-source diagnostic for the XLSX source-backed
one-percent scalar-cell edit/save route at revision `a664c8539894de38fea670bb34eaa66cc69058ce`.
No production or harness source changed, no candidate was implemented, and
there is no accepted 0707 performance claim.

The measured selector uses an instrumented in-memory positional provider over
medium and dense-sparse deterministic four-worksheet workbooks. Native timing
retains two CPU-12-pinned repeats with 100 samples after 20 warmups. The
allocator lane is a separate binary with two repeats, three samples, and two
warmups. The profile lane is a separate Callgrind capture with two repeats,
one sample, and no warmup; the second repeat reverses shape order. Native
timing, allocator diagnostics, and guest-instruction attribution are kept as
separate evidence types.

## Result

Planning is 35.75–36.88% of the native elapsed interval, publication is
41.06–42.72%, and commit is 21.20–21.59% on the retained shapes. Planning also
has the largest operation-region allocation count and requested-byte volume:
67,845 calls / 12,139,796 bytes on medium and 129,411 calls / 15,850,404 bytes
on dense-sparse. These are current-source observations, not before/after
comparisons and not a prediction of a candidate's speed.

The native A/A analysis retains five repeat-drift flags over 5%: medium
planning p99 −5.68%, and dense-sparse reopen p50 +8.50%, p95 +8.24%, p99
+7.88%, and mean +8.61%. Reopen is outside the timed edit/save interval. No
native total metric has a repeat shift over 5%.

The raw profile captures isolate
`litchi_xlsx::cell_values::source::SourceBackedEditor::edit_sheets`. Three
lifecycle calls and one measured call are retained per profile child, with
selection by positive incoming edge and all numbered dumps preserved. The
independent profile analysis passes the measured/lifecycle selection, output
and source identity checks, and disjoint immediate-child accounting. Its
measured rows put `worksheet_xml_and_parse_source` at 96.36–96.79% inclusive
of the owner and `Validator::observe` at 18.38–18.75%; nested rows remain
diagnostics and are never summed. Callgrind's profile process uses Valgrind's
allocator replacement, so allocator edge counts from these files are not
treated as native allocation measurements or production allocator cost.

The prospective next experiment is a private validator name-storage design:
static modeled names, callback-borrowed Empty names, and owned fallback for
unknown names. It remains a design handoff in
[`next-experiment.md`](next-experiment.md). Exact validation order, close-name
matching, namespace and dialect behavior, copied-subtree preservation, depth
limits, budgets, and fallback must remain unchanged.

## Packet map

| Evidence | Artifact |
| --- | --- |
| Frozen plan, profile scope, and investigation record | [`plan.json`](plan.json), [`profile-plan.json`](profile-plan.json), [`investigation.json`](investigation.json), [`next-experiment.md`](next-experiment.md) |
| Constraints and source identity | [`constraints.json`](constraints.json), [`source-baseline.json`](source-baseline.json), [`host.json`](host.json) |
| Build commands, logs, binary and source receipts | [`build.py`](build.py), [`build-baseline.json`](build-baseline.json), `build-baseline-*.log`, `*-baseline-*.receipt.json`, and per-child source manifests |
| Native samples and independent analysis | `native-baseline-*.json`, [`analysis.json`](analysis.json), and `analyze.py` |
| Allocator samples | `alloc-baseline-*.json`, allocator receipts, and [`analysis.json`](analysis.json) |
| Callgrind profiles and analysis | `profile-baseline-*.callgrind*`, `profile-baseline-*.json`, `profile-baseline-*.inclusive.txt`, `profile-baseline-*.self.txt`, [`profile-analysis.json`](profile-analysis.json), [`analyze_profiles.py`](analyze_profiles.py), profile receipts, and per-child source manifests |
| Symbol and helper custody | [`planning-symbols.txt`](planning-symbols.txt), [`planning-symbols.json`](planning-symbols.json), [`helper-identities.json`](helper-identities.json) |
| Capture correction and reproduction scripts | [`capture-preflight-correction.json`](capture-preflight-correction.json), [`capture.py`](capture.py), [`build.py`](build.py), [`analyze.py`](analyze.py) |
| Evidence-gate and audit helpers | [`evidence.py`](evidence.py), [`audit.py`](audit.py), [`artifact-seal.py`](artifact-seal.py) |

## Reproduction boundaries

The retained native interval includes open, planning, staging plus commit, and
sequential publication. It excludes sink setup, remaining handle destruction,
reopen, and semantic or preservation oracles. The allocator values are
operation-region relative diagnostics; absolute process gauges are not
compared and phase peaks are never added. Callgrind Ir is guest-instruction
attribution only, and Valgrind's allocator replacement means its allocation
edges do not measure native allocator latency. No profile result is used as an
allocator, RSS, or latency claim.

From the repository root, the retained baseline analysis and audit commands
are:

```sh
python3 -B docs/performance/results/change-0707/analyze.py
python3 -B docs/performance/results/change-0707/audit.py
python3 -B docs/performance/results/change-0707/artifact-seal.py --check
```

The build and capture scripts use owned sibling target and binary paths and
the frozen plan. A new acquisition requires a fresh packet directory and a
new source-bound build; these retained results are not a historical speedup
baseline. No RSS, physical cold-cache, native Office, cross-platform, or
parallel-scaling campaign is present. iWork is excluded and the broader
non-iWork GOAL remains active.

All four profile stderr logs report Valgrind `brk segment overflow`. The
children exit successfully and output identities match, but this remains an
instrumentation environment limitation. No native allocator-cost claim or
causal explanation for the call-field discrepancy is drawn from those profiles.
