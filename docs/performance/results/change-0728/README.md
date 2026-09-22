# 0728 — current DOC/PPT save baseline evidence

This packet measures unchanged production source at `1dc122a2bc`. It separates
public DOC/PPT open/edit/commit lifecycles from a common OLE editor control that
uses the exact replacement streams produced by each public edit. Reuse and
Rewrite are existing policies, not before/after implementation variants.

`cases.json` binds the three repository fixtures. `hypothesis.md` and `plan.json`
define the prospective experiment. `freeze.json` binds source, fixtures, probe,
qualified semantic witnesses, analysis scripts, and both release executables.
`captures/manifest.json` retains the full ordered process census, commands,
stdout/stderr hashes, and start/end bindings. `analysis.json` contains every
per-process statistic; `analysis.md` presents ranges across processes.

The current common-editor stage includes stream replacement, rendering, reopen,
recapture and discovery. Its phase fractions are measured within one owner.
PPT's public save uses its own writer; the common-editor PPT route is an
alternative container control. Subtracting separate route medians does not
estimate a nested format phase or an achievable speedup.

Probe oracles run outside the measured intervals. Format-internal validation
remains inside the public lifecycle. Native processes use the system allocator;
separate instrumented processes report allocation regions. Returned editors and
output vectors remain live at the named boundaries. Region-relative peak live
bytes are not RSS, and separate phase peaks must not be summed into a lifecycle
peak. This packet does not claim cold-cache, storage-I/O, remote-source,
concurrent, cross-platform, native Office, or production speedup results.

Replay after cleanup from the repository root:

```sh
python3 -B docs/performance/results/change-0728/source-guard.py
python3 -B docs/performance/results/change-0728/analyze.py
python3 -B docs/performance/results/change-0728/report-tables.py
python3 -B docs/performance/results/change-0728/audit.py
python3 -B docs/performance/results/change-0728/negative-checks.py
python3 -B docs/performance/results/change-0728/artifact-seal.py --check
```

`cleanup.json` substitutes exact executable identity witnesses after removal of
the two owned build/binary roots. Analysis and independent auditing remain
possible without those executables. The artifact manifest covers every packet
file except itself. Rebuilding or collecting a new experiment belongs in a new
packet; do not overwrite this sealed evidence.

Qualification history is retained, including failed build/test/lint attempts,
probe source snapshots, preliminary reviews, and earlier smoke runs. Only the
successful source-bound receipts named by `qualification.json` qualify the
frozen matrix. `preflight.json` is synthetic schema integration, not a timing
measurement. Deliberate live output corruptions are checked by the probe; the
separate Python negative checks exercise acceptance of altered evidence.

The broader non-iWork performance goal remains active. No registered performance
claim or CRUD coverage status is promoted by this unchanged-source baseline.

Final acceptance: `build-5.json`, `quality-3/manifest.json`, and
`qualification-3/manifest.json` qualify the frozen source. All 81 capture
processes, 1,620 native samples, 27 allocation samples and 621 applicable live
corruption checks pass. `audit.json` independently matches analysis;
`negative-checks.json` records 15 evidence controls. `terminal-checks.json`
records eight successful checks with exact post-cleanup replay.

[Source route review](source-review.md), [closed probe review](final-review-resolution.md),
and [result interpretation review](result-review.md) distinguish tested
preservation from reported limitations. The current result and next action are
in the [performance record](../../0728-doc-ppt-current-save-baseline.md).
