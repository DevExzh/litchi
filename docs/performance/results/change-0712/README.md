# 0712 active-offset validation evidence

`performance_claim: none`. [Report](../../0712-docx-active-offset-validation.md).

The current-source probe records 11 cases and two reachable no-anchor MCE
refusals. It also distinguishes nested anchor discovery from outer-range writer
selection. The instruction analysis reuses four historical 0709 profiles with
exact custody; it does not supply current native performance measurements.

| Artifact | Purpose |
| --- | --- |
| `revision.json`, `constraints.json`, `source-final.json`, `source-relation.json` | Current revision, unchanged constraints and source binding |
| `oracle/`, `oracle-freeze.json`, `run-probe.py` | Frozen synthetic generator and source-bound locked build/run |
| `oracle/current/report.json`, `analysis.json`, `analyze.py` | Complete observations and checked counterexamples |
| `oracle-replay.json`, `probe-format.json` | Second-process byte parity and formatting |
| `mce-attribution.py`, `mce-attribution.json` | Historical raw-edge partitions and inlined-search attribution |
| `evidence/results.json` | Six repository gates |
| `final-report-gate.json`, `audit.py`, `audit.log` | Terminal documentation and packet verification |
| `cleanup.json`, `artifact-manifest.json` | Exact binary witness, owned scratch removal and artifact seal |

Replay from the repository root with:

```text
python3 -B docs/performance/results/change-0712/analyze.py
python3 -B docs/performance/results/change-0712/mce-attribution.py --replay
python3 -B docs/performance/results/change-0712/audit.py
python3 -B docs/performance/results/change-0712/artifact-seal.py --check
```

The probe's lexical full-offset set is intentionally distinct from the
production outer-range scanner. Its one-million-offset refusal is a direct
public active-offset call, not a whole-package stress test. Both distinctions
are retained in the report. No production optimization or broad quality-suite
result is claimed. Owned target and binary directories are removed after the
probe and its exact replay finish.
