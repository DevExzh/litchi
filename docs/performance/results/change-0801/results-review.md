# 0801 results review

## Scope and disposition

The 0801 candidate was measured as an isolated helper experiment rebased on
`54aa0f3fb4967d98c923f58b3944911b9b4525ab`. The production sources remained
unchanged. The candidate is rejected by the frozen preflight policy because
the protected `distinct-0` consume case regressed; the production corrected
0800 baseline remains the retained implementation. No public-workflow,
resource, cross-format, or production-adoption claim follows from this packet.

The authoritative result is recorded in `decision.json`, with the full native
table in `summary.md` and the protocol in `plan.json`. Ratios are after/before;
values below 1 are faster.

## Independent raw timing checks

I recomputed the nearest-rank process p50 from each of the six raw native JSON
reports for representative consume cases and paired those six block values.
The values below are in block order 0 through 5; the reported ratio is the
median of the six paired ratios, matching `analysis.json`.

| Case | Before p50 (ns) | After p50 (ns) | Ratio | 95% bootstrap interval |
| --- | --- | --- | ---: | ---: |
| `distinct-0` | 36360, 36351, 36360, 36320, 36340, 36301 | 39120, 38201, 39071, 39070, 39070, 39140 | 1.075420 | [1.062726, 1.077057] |
| `distinct-1` | 109681, 109101, 109030, 114140, 108841, 107690 | 66550, 88390, 66501, 66490, 66510, 66650 | 0.610504 | [0.594645, 0.714536] |
| `distinct-2` | 172781, 172901, 172881, 172721, 180001, 172851 | 139321, 139331, 139451, 139341, 139380, 139360 | 0.806294 | [0.790086, 0.806685] |
| `distinct-16` | 1560228, 1542058, 1542888, 1535237, 1547538, 1555098 | 2588633, 2695174, 2596583, 2610963, 2592463, 2609064 | 1.680343 | [1.667178, 1.724234] |

These spot checks agree with the official rows in `summary.md`,
`root-native-audit.json`, and `analysis.json`. The protected empty case is a
7.542% regression whose interval is wholly above 1, so it alone vetoes
advancement. The two required dominant-class benefits pass: `distinct-1`
improves 38.950% and `distinct-2` improves 19.371%, with both intervals wholly
below 1.

The complete analysis records 13 significant consume regressions and 50
process-p50 spread flags among 156 case/mode/leg groups. The remaining
regressions are retained review triggers; they are not hidden by the single
protected veto. In particular, the 4-, 16-, and 32-attribute distinct cases
show the cost of entering and using the ordered backend.

## Semantic, layout, and counter evidence

The frozen matrix contains 39 cases, both construct and consume modes, six
alternating native blocks, 30 samples per process, three warmups, and 4,096
iterations. The capture contains 936 native children and 28,080 native
samples. The release probe's semantic oracle reports exact first-error,
position, sequence, checksum, repeated-`None`, and clone parity for both
helper legs across the frozen cases.

The isolated helper quality run covers five owners in each leg. The baseline
leg ran 14 tests per owner (70 total), and the candidate leg ran 19 tests per
owner (95 total), for 165 helper tests overall. All passed with zero failures
and zero ignores, and both legs passed warnings-denied Clippy. These are
minimal mirror-crate checks, not full production-crate or whole-workspace
verification; the final receipts are in `quality/complete.json` and the test
counts are visible in `quality/after-1.log`.

The measured `CheckedAttributes` size is 120 bytes for the baseline and 128
bytes for the candidate, stable across all 1,248 native and counter reports.
This is the complete iterator size, not the private phase enum, and it says
nothing about heap allocations made by the lazy map.

The Callgrind packet contains 312 positive and 312 empty-termination reports.
All 312 owners qualify, scalar conservation passes for all 624 dumps and all
five events, and the termination counters are zero. The counters are guest
instruction and branch diagnostics; they do not count native latency or
allocator API calls. The native and counter results are therefore retained as
separate evidence streams.

## Quality corrections and retained failures

The first helper quality attempt passed both test legs but failed candidate
Clippy on `map_or(true, Result::is_err)` and a test identity map. The exact
candidate, patch, logs, and relocation record remain under
`quality-failed-0/`, `candidate-failed-0/`, and `candidate-failed-0.patch`.
The final correction uses `Option::is_none_or` in the five equivalent helper
copies and removes the identity map; it does not change the parser strategy.

The first direct-probe attempt passed formatting, release build, and release
check, then failed probe Clippy because the standalone candidate module's
`unchecked_attributes` method is unused by the checked-only harness. Its
source and binary identity remain under `build-failed-0/` and
`probe-src-failed-0/`. The final fix is a narrow dead-code allowance on that
probe module only; the archived helper source is unchanged, and no failed
attempt was used for native capture.

The final probe formatting, release build, release check, Clippy, fixture
listing, and self-check all pass in `build/build.json`. The final cleanup record
`cleanup.json` reports removal of the owned target and 180,293,462 logical
bytes while retaining the final and failed executable identities.

## Limits and source claims

The candidate source archive describes a no-replay short prefix, raw
quick-xml-compatible duplicate preflight, and lazy ordered-map handoff. The
source review found no source-level correctness blocker, and the semantic
oracle supports those claims for the frozen input matrix. The 0801 candidate
manifest remains a source-only, unadopted archive; the measurements do not
promote it to production.

The direct probe measures named helper construction/consumption with opaque
construction and checksum work included. It does not establish public
workflow speed, resource behavior, tail latency, allocation API counts, or
cross-format impact. The packet retains the 39-case and 18-protected-case
policy, and the `decision.json` disposition is `advance_to_workflow_trials:
false` with `production_adoption: false`.
