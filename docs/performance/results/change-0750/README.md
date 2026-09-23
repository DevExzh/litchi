# Retained evidence — change 0750

Record: [`../../0750-xml-audit-well-formedness-gaps.md`](../../0750-xml-audit-well-formedness-gaps.md).

## Provenance

| | |
| --- | --- |
| base commit | `3174242282` (branch tip with records 0745 and 0747) |
| branch | `perf/0750-xml-audit-well-formedness-gaps` |
| before leg | detached worktree `/home/zhuhe/code/litchi-worktrees/0750-before-src` at the base, root `Cargo.lock` copied in, target `targets/0750-before`; removed after use |
| after leg | the branch worktree, target `targets/0750` (harness, tests, gates) and `targets/0750-census-after` (probes) |
| host | AMD EPYC 9R45, 32 cores, `Linux 7.0.0-1012-aws x86_64`, shared with other agents |
| CPU pin | `taskset -c 24` for every timed and callgrind process |
| harness build | `cargo build --release --manifest-path tools/perf-baseline/Cargo.toml --locked --offline --bin litchi-perf-baseline`, identical for both legs (`scripts/build_legs.sh`; the after leg rebuilt by `scripts/final_pipeline.sh` on the final code); rustc 1.95.0, pinned by `rust-toolchain.toml` |
| probes | `probe/census` and `probe/audit`: the same source built against each leg's `xml-minifier` (`Cargo.toml.example` with the leg's checkout for `<CHECKOUT>`), `--release --offline`, from outside the checkout, so rustc 1.98.1 for both legs |

Binary SHA-256s and compilers are in [`binaries.txt`](binaries.txt).

## Contents

| path | what it is |
| --- | --- |
| [`gaps/verdicts.md`](gaps/verdicts.md), `gaps/verdicts.json` | the 150 probed inputs: `verify_source` on both legs, and whether `verify`, `verify_authored` and `verify_reader` changed (they did not) |
| [`census/summary.txt`](census/summary.txt) | corpus composition, per-policy verdict transitions, newly refused and still refused members, prefix declarations in scope (`scripts/census_summary.py` over the rows) |
| `census/census.jsonl.gz`, `census/corpus.txt.gz` | one row per XML member (12,726) with all eight verdicts; the corpus list and unreadable members |
| [`campaign/totals.txt`](campaign/totals.txt), `campaign/seeds.txt` | the release differential campaign, 8 seeds × 3,000,000 cases |
| [`draft-all-policies/failures.md`](draft-all-policies/failures.md) | the dependents' test failures of the first, all-policy implementation, which set the scope |
| `timing/` | harness ABBA: `status.txt`, raw JSON reports and logs (`raw/*.gz`), `analysis.json`, `flags.json` (every paired process moving more than 5%) |
| `probe-timing/` | audit-probe ABBA: `status.txt`, raw JSON lines (`raw/*.gz`), `analysis.json`, `flags.json` |
| [`callgrind/probe-summary.txt`](callgrind/probe-summary.txt) | instructions per audit on each leg (the audit probe at N iterations minus 0) |
| [`callgrind/pairs-summary.txt`](callgrind/pairs-summary.txt), `callgrind/pairs-runs/` | harness isolation pairs: program instructions per sample on each leg, the auditor entry points' share with their callers (`scripts/cg_attribute.py`), and the harness reports of those runs |
| [`callgrind/tiny-document.txt`](callgrind/tiny-document.txt) | where the source audit's fixed cost goes, by function (`scripts/cg_tiny.py`) |
| [`probe/parts.txt`](probe/parts.txt) | the audit probe's inputs: source member, size and SHA-256 |
| [`probe/pair-verdicts.txt`](probe/pair-verdicts.txt) | what the probe's two pair cases return (one window proof, one refusal) |
| `probe/audit/`, `probe/census/` | probe sources |
| [`gates.txt`](gates.txt) | every gate's command, exit status and tail, and the test-suite totals |
| `scripts/` | every script used (census, part extraction, timing, analysis, flags, campaign, callgrind, gates, pipeline) |
| [`log-sections.md`](log-sections.md) | ready-to-paste sections for `HOTSPOTS.md`, `REPORT.md` and `GOAL_AUDIT.md` |
| [`cleanup.json`](cleanup.json) | what was removed and what was kept |

Not kept: binaries (hashes kept), the probe's inputs (rebuilt byte-identically by
`scripts/extract_parts.py`), raw callgrind profiles (summaries kept), build and
test logs (totals kept in `gates.txt`).

## Replay

```sh
# 1. before-leg source: git worktree add --detach <dir> 3174242282; copy Cargo.lock
# 2. harness legs: scripts/build_legs.sh (identical command)
# 3. probes: probe/{census,audit} with Cargo.toml.example per leg
# 4. census: python3 scripts/census.py <repo> <before census probe> <after census probe> <out>;
#            python3 scripts/census_summary.py <out>/census.jsonl <out>/corpus.txt
# 5. probe inputs: python3 scripts/extract_parts.py <repo> census/census.jsonl.gz <parts>
# 6. campaign: scripts/differential_campaign.sh
# 7. timing: scripts/run_abba.sh; scripts/analyze.py; scripts/flags.py
#            scripts/run_probe_abba.sh; scripts/analyze_probe.py; scripts/probe_flags.py
# 8. instructions: scripts/probe_callgrind.sh; scripts/callgrind_pairs.sh;
#                  scripts/cg_attribute.py <cg-pairs dir>; scripts/cg_tiny.py <cg-probe dir>
```
