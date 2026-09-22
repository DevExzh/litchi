# 0736 sample-order sensitivity and lifecycle audit

Post-hoc analysis of the sealed rejected 0735 experiment; no production changes
or fresh timing captures. The original rejection remains in force.

Run from the repository root:

```sh
python3 docs/performance/results/change-0736/analyze-order.py
python3 docs/performance/results/change-0736/plot-order.py
python3 docs/performance/results/change-0736/independent/independent-order.py
python3 docs/performance/results/change-0736/compare-audits.py
python3 docs/performance/results/change-0736/artifact-seal.py --check
```

The first command verifies the exact 0735 packet before reading timings and
reproduces all full-window paired medians. It writes `order-analysis.json` with
all 36 process records, 18 pairs, six descriptive windows and process-order
strata. The second command needs Matplotlib and draws all 1,800 samples in
recorded order. Neither command executes native code or changes production.

`provenance.json` binds the starting revision, unchanged constraints and restored
source identity. `independent/` contains a separately implemented numerical
audit. `lifecycle.md` maps the probe's measured and unmeasured work and defines
the next control experiment. See [the report](../../0736-ppt-sample-order-and-oracle-lifecycle.md)
for interpretation and limitations.

Windows overlap, were chosen after the original captures, and are descriptive.
They cannot replace the original retention gate. Serialized witness sizes are
not live allocation measurements. No cause of the regression is established.

The independent comparison checks 532 shared scalar statistics within 1e-10.
Analysis and plot regeneration are byte-identical on the recorded environment;
`verification.json` records that check. No build targets or binaries were created.
