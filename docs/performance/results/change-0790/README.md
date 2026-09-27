# 0790 ordinary-fork control

[Report and explicit failed acceptance criterion](../../0790-fork-launcher-calibration.md).
This is a standalone diagnostic, with no library candidate or adoption.

Offline replay, from this batch's repository source revision:

```sh
python3 -B docs/performance/results/change-0790/analyze.py --check
python3 -B docs/performance/results/change-0790/validate.py --require-final-seal
```

A successful replay preserves the experiment's **failed** overall acceptance:
three traced controls violate the frozen universal child/launcher startup
comparison. All 48 direct ordinary matrix controls satisfy the comparison.
The analyzer reports failures as data; neither replay nor cleanup promotes
that result to a full acceptance pass.

Add `--check-workspace` to validation only in the original owner's workspace
to check unrelated file hashes and other worktrees. Normal replay does not
require those untracked user files. Exact binary cleanup witnesses allow
replay after the temporary target is removed. No replay executes a probe.

Root executed these provenance commands once, serially:

```sh
python3 -B docs/performance/results/change-0790/build.py
python3 -B docs/performance/results/change-0790/capture.py matrix
python3 -B docs/performance/results/change-0790/capture.py traces
```

These commands refuse an existing target or capture directory. Do not rerun
them over the retained packet; a fresh experiment requires new output paths,
source/host receipts, and a predeclared plan.

`probe.c` is byte-identical to 0789. The original marker names remain
`RSS0789`/`ACK0789`. `launcher.c` adds an ordinary-fork child and an exact-PID
wait4 usage receipt; the launcher's outer resource usage remains separate.
`plan.json` freezes six cases, two affinities, four launcher/observer positions,
and two reversed-order repeats (96 children), plus four separate traces.
`matrix/` and `traces/` retain every raw report, transcript, usage result, and
compressed proc snapshot. Verbose trace rusage keeps execve abbreviated.

`build.json` and `qualification/` bind three strict builds and seven checks.
The sanitized build is the **launcher**, not the probe. The unchanged probe's
0789 sanitizer checks are inherited by exact source and receipt identity.
`inherited.json` references the earlier source census and accounting review;
`architecture-inputs.json` binds all 35 unchanged goal/taxonomy/ADR inputs.
`launcher-review.md` is an independent source and raw-data review.

`analysis.json` retains all 100 rows, 600 checkpoints, counter differences,
trace comparisons, and the three acceptance failures. This has 312 full and
288 identity snapshots. Two repeats do not justify a confidence interval,
launcher equivalence, or a transient residency peak claim.

`cleanup.json` records verified removal of three owned executables;
`seal.json` binds every payload and six performance documents. The historical
0787 rejection remains unchanged. OLE2/OOXML remain active, ODF deferred,
and iWork excluded.
