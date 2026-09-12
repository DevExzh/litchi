# 0520 source-backed XLSX phase attribution

`performance_claim: none`

The [change record](../../changes/0520-xlsx-source-edit-phase-attribution.md)
states the scope, exact timing boundaries, results, repeat variations and
remaining requirements. This bundle measures existing production/harness
source at `b776ccb2a55c7a2a22018d4498e78676c6648dde`.

## Reproduce

Use the pinned toolchain and repository at that revision. The benchmark is a
standalone Cargo workspace with its own committed lockfile and ordinary
release defaults (no workspace-root LTO override). `plan.json` records the
exact build command; `build.json`, `build.stdout`, and `build.stderr` retain
its result. `source-manifest.json` binds tracked crate and benchmark inputs;
`build-context-review.json` additionally reviews the root toolchain and Cargo
config against that revision. `host.json` records actual machine/tool commands.

```sh
cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml \
  --bin litchi-perf-baseline --target-dir /home/zhuhe/litchi-goal-0520-target
```

For one native child (use a fresh output path):

```sh
taskset -c 2 /home/zhuhe/litchi-goal-0520-target/release/litchi-perf-baseline \
  --warmup 20 --samples 100 \
  --case xlsx_source_backed_cell_values_one_percent_edit_save \
  --xlsx-cell-crud-shape medium --json /tmp/xlsx-medium-reproduction.json
```

`capture.py` implements the recorded four-child order: medium, dense-sparse,
dense-sparse, medium. It checks source and executable hashes around every
capture, refuses an existing receipt, and records terminal status, commands,
timestamps and artifact hashes. For a fresh campaign use a separate evidence
directory and new plan/build receipt; do not overwrite this sealed bundle.

`profile.py` implements `profile-plan.json`: the same binary under Callgrind,
zero warmups and one sample, exact commit-name toggle, zero-before and
dump-after. Each capture produces four numbered operation dumps plus the
final process dump. Only `.callgrind.4` is the selected timed commit; the
first three are untimed lifecycle gates. `.inclusive.txt` and `.self.txt`
are complete annotations of that fourth dump. Annotation replay fixes
`PERL_HASH_SEED=0` and `PERL_PERTURB_KEYS=0` to stabilize equal-cost row order.
Callgrind timings are not
compared with native timings.

Recompute the JSON analyses to temporary destinations from the repository root:

```sh
python3 -B docs/performance/results/change-0520/analyze.py /tmp/xlsx-native-analysis.json
python3 -B docs/performance/results/change-0520/analyze_profiles.py /tmp/xlsx-profile-analysis.json
```

The profile analyzer also regenerates the deterministic annotation files.
Its two unchanged parser helpers are in `change-0519`; their hashes are in
`profile-analysis.json`. `verify.py` validates the seal, source, custody,
serial intervals, raw native sums and isolated profile call accounting, and
requires exact report/annotation replay. The binary/build directory is not
needed for retained-evidence verification.

## Interpretation

The native phase vectors retain acquisition order; `elapsed_ns.samples` is
sorted and `sample_order` supplies the alignment. All 400 rows are checked.
Counters are logical in-memory source/cache observations. Each iteration uses
a fresh editor/cache; this is not a warm semantic-cache reuse benchmark.
Configuration defaults for filesystem cache and simulated ranges are inactive
for this selector. The unmanaged zero budget vectors are not allocation or
peak-memory measurements.

The raw benchmark includes successful lifecycle/oracle execution, not a new
full workspace test suite. Partial-sink evidence is absent for these shapes;
inverse restoration checks semantics and package identity. All four profile
children retain Valgrind `brk segment overflow` notices and complete successfully;
allocation/RSS inference from those runs is excluded. No production speedup,
cold/range, scaling, or native Office-producer claim is made.
