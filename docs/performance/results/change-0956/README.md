# Change 0956 evidence packet

See [the report](../../0956-xlsx-managed-compression-native-widths.md). All 144 native children completed and passed validation. The authoritative external packet is `/home/zhuhe/litchi-performance-0956`. This bounded packet contains the frozen plan, source/build identities, qualifications, completed capture inventory, comparison summary and reviews. Raw captures, full reductions and executables remain external. There is no production-adoption decision.

## Frozen schedule

`plan-final-v1.json` is the exact authorized plan (schema `litchi.performance.0956.native-width-plan.v3`, SHA-256 `211d4c56fa38eef0b4d3e513a4a79054bff29fe02e0100096c8ba0af8dbd9da0`). It schedules three cases (`medium-one-percent`, `dense-sparse-one-percent`, and `medium-one-edit`), two comparisons (width 1 versus 2 and width 1 versus 4), six ABBA/BAAB blocks per comparison, and four fresh children per block: 144 native children total. Each child has 20 warmups and 200 measured samples. Bootstrap uses seed 956, 10,000 resamples, and geometric means over the six block ratios. The required CPU mask is 12, 13, 14, 15; the process deadline is 900 seconds; retries and replacement children are disabled.

The timed route is limited to the explicit finite managed-worker phase. The plan keeps setup, fixture/oracle construction, admission, sink reservation, verification, and resource checks outside the phase timer. Whole-process RSS and `/usr/bin/time -v` counters are review data. The experiment makes no allocation, OS-thread-count, cold-filesystem, or production-adoption claim.

## Source and build identity

The measurement source is the frozen external `source-v1` snapshot: 10,930 files and 273,393,096 source bytes, based on commit `c8e80705989183d88950e282ef1852f3e3470b10` with approved overlay v3. The five overlay paths and full source census are bound in `source-build-provenance-v1.json`; the 2,266,359-byte census remains external under the recorded path and SHA-256. The checkout may contain unrelated working-tree changes and is not the measurement source.

The release build used `+1.95.0`, `--release --locked --offline`, fat LTO, opt-level 3, one codegen unit, debug information enabled, and incremental builds disabled. The benchmark executable is external at the path recorded in `external-records-v1.json` and is bound to SHA-256 `38381c594b6dd5924e07d3a27f57bf15729811b8f72d7353aac5aac810741304`; no compiled executable is committed.

The 0955 qualification manifest is an external correctness oracle, recorded with its size and SHA-256. Its checks are retained as correctness preflight; they do not supply 0956 timing results.

The retained correctness-only wrapper completed nine successful configurations and recorded a post-workload setup failure while creating its existing-output negative control. The recovery used a unique pre-created child; no successful workload child was retried or replaced. `review/runtime-runner-status-v1.json` preserves that wrapper status, and it is not a formal native-capture result.

## Helpers and relocation

The `tools/` files are byte-identical copies of the frozen external helpers: `capture_0956_v3.py`, its imported legacy validator `capture_0956_v2.py`, `controls_0956_v3.py`, and `reduce_0956_v3.py`. Their external paths and hashes are authoritative. The helpers bind absolute packet/source/binary paths and record the original working directory; a relocated replay must use a separately reviewed adapter and a fresh output directory. The copied scripts alone do not make a moved packet valid.

`completed-packet-manifest-v1.json` seals all 1,154 capture files before reduction. Its SHA-256 is `f923ebe1bb5f8a3a58458f6458b0795ea26071ebd6ea4254c624d769b489ad28`; inventory verification passes after reduction. `validation.json` binds it to the original authorized plan, separately reviewed reducer v4, and final comparison. V4 only adapts the historical omission of the deadline from terminal receipts: start deadlines remain mandatory and contradictory terminal values are rejected. The original failed v3 reduction is retained, with no native retry or raw-data mutation.

`comparison-summary-v4.json` retains every metric/quantile point, bootstrap interval, six block points and review flags. The full reduction is 3,937,963 bytes; its exact path and SHA-256 are in `validation.json` and the summary. `cleanup.json` records removal of owned build/temp caches and unchanged retained executable identities. Earlier provenance files retain their preparation-time status; completion is established by `capture-terminal.json` and `validation.json`. See [`reproduction-v1.md`](reproduction-v1.md) for source and binary regeneration commands.

To verify and reduce the retained original packet into a new output file:

```sh
python3 -B docs/performance/results/change-0956/tools/freeze_packet_v1.py \
  /home/zhuhe/litchi-performance-0956/capture-native-v1 \
  docs/performance/results/change-0956/completed-packet-manifest-v1.json --verify
python3 -B docs/performance/results/change-0956/tools/reduce_0956_v4.py \
  --plan /home/zhuhe/litchi-performance-0956/plan-final-v1.json \
  --packet /home/zhuhe/litchi-performance-0956/capture-native-v1 \
  --output /tmp/litchi-0956-comparison-regenerated.json
```

These commands require the original bound source/helper paths to remain available. Reduction output includes a new creation timestamp; compare all other JSON values for deterministic replay. `tools/summarize_v1.py FULL_JSON NEW_SUMMARY_JSON` regenerates the bounded summary from a full reduction.

Regenerate native evidence with the retained binary and source/helper bindings,
using a fresh capture directory:

```sh
LITCHI_PERF_0956_CAPTURE_APPROVED=1 python3 -B \
  /home/zhuhe/litchi-performance-0956/tools/capture_0956_v3.py \
  --plan /home/zhuhe/litchi-performance-0956/plan-final-v1.json \
  --packet /tmp/litchi-0956-native-recapture --execute
```

That command creates a new experiment; it cannot regenerate identical timings.
Keep its new manifest and reduction separate from the recorded evidence.
