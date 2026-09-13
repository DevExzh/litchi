# 0557 metrics final review

This is a prospective contract review of the 0557 XLSX source-backed cell
value campaign. No Rust measurement build, benchmark capture, candidate output,
or performance result is claimed here. The source revision pinned by the plan
is `0d812aca92b608f9de8f9398fad1c379077ce708`.

The exact reviewed inputs are:

```text
analyze.py          010f44e811d852ade5f5a62fcd3092961815ef0a3e81217697bc170b5e92bbb2
run.py              c9c0db8ac91da2d8604c3a7d70b77f5927f0b6ac9c6fb8066e4b7c350b8b8999
plan.json           d8111317e6d6d44219c844aaf740ea50fc149c361601bb6483a47d33843f88d8
analyzer-notes.md   1cd479751a7348aeb05d3de5661a2f45dadde21e2d7f45ea60773cba32d99481
driver-notes.md     a5975d097fc17707b39ca5033ee321f26ebf36af9de39f6332977079894cb73f
```

The preserved pre-review copies are:

```text
analyze.py.before-final-review           0a5949bd72da7d0e584de4e4a8be7869506ab73aca5987aa03c6b38d2027e6ee
analyze.py.before-metrics-final-review  6ab6bb08d5eb14a3bff247d02613c491013719e6fc53945ae89d7ab08ce4c009
analyze.py.agent-final-before-gate-review b006a4d8be82b984fa33eab5e1eca5ace47b7c7c654ee4831ee52b4063b0713b
plan.before-final-review.json            00b2418b96adc351fc2371a1479b10662a0e754b600e26d0d96d36aef76dfab9
```

The review corrections are:

- `plan_data()` now pins all five source phase vectors and the explicit gate
  lane contract. Native elapsed p50/mean gates cover primary and controls;
  allocator elapsed values remain diagnostics; RSS remains per-child in both
  lanes; allocation vectors are allocator-lane evidence.
- Eager reports with no `source` object become typed unavailable
  `xlsx_source_phase` evidence. The allocation gate accepts a matched pair of
  unavailable eager rows without turning omission into zero-valued evidence.
  Source-backed allocator rows must have measured phase vectors on both sides.
- Elapsed statistics are read from the validated nested `statistics` object,
  while confidence interval endpoints remain under the elapsed envelope. This
  fixes the prior accessor failure during matched comparisons.
- The admission floor is recomputed for each native primary comparison from
  that comparison's matched native baseline p50:

  ```text
  max(1.0, 3 * N, 100 * 50000 / matched_baseline_p50_ns)
  ```

  Noise-pilot floors are retained only as provisional diagnostics. They are
  never used as the matched admission denominator.
- `N > 5.0` is the strict instability stop. The arithmetic preserves an exact
  5.0% boundary, and ordinary analysis returns a stopped candidate gate before
  loading candidate rows when the pilot is unstable. There is no noise rescue.

The `run.py` ABBA custody contract is consistent with the plan. The final
baseline leg uses baseline output stage plus candidate execution stage, and
the receipt binds both manifest hashes and guard scopes, binary and descriptor
custody, lock identities, environment, command, and artifact hashes. No driver
implementation change was needed. `run.py` does not calculate `N` or block a
direct candidate capture invocation; the coordinator must run the noise-only
analysis and enforce its stop before launching a candidate child.

Validation performed without Rust jobs:

```text
python3 -B docs/performance/results/change-0557/analyze.py --self-test
0557 analyzer pure self-test: PASS
```

A synthetic complete matrix passed with 16 native primary gates, 48 native
control gates, 128 per-child RSS gates, and 64 allocator allocation gates.
Primary/control rows were native-only, allocation rows allocator-only, eager
allocation rows were explicitly unavailable, and matched floor records named
`matched_native_baseline`. Synthetic noise checks passed for the exact 5.0%
boundary and the strict greater-than-5.0% stop.

