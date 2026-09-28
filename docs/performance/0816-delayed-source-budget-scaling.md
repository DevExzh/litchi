# 0816 — finite execution budgets with delayed caller reads

The reusable execution harness now measures bounded caller-source reads under
finite budgets. On this synthetic corpus, eight requested workers accelerate
delayed large CFB reads by 7.364× and delayed large ordered-Part reads by 5.674×
relative to one worker. Mixed-size batches remain near 1× with one observed
simultaneous read; primed Parts make zero timed source reads and become slower with more
workers. These are descriptive configuration comparisons on unchanged production,
not adopted optimizations or real-network measurements.

The baseline is `c8ff2b9f65`. Only `tools/perf-execution/src/main.rs` and its
README change. The new opt-in `--source-max-read-bytes` and `--source-delay-us`
options cap a returned prefix and sleep after each successful nonempty read.
Zero disables either option; maximum values are 1 MiB and 100,000 µs. OPC
`from_bytes` refuses nonzero settings because it bypasses the source adapter.
No production API, dependency, scheduler, budget, or cache contract changes.

The gap is the combination of delayed short reads, finite hierarchical budgets,
fresh/primed sessions, and CFB/OPC scheduling. Historical
[0498](changes/0498-bounded-source-backed-part-batch.md) already measured delayed
short-read Parts. [0786](0786-finite-execution-budget-scaling.md) established
finite-budget scaling over fully readable memory, and
[0787](0787-cached-part-scheduling.md) tested a cache scheduling hint. This packet
adds the current CFB/Part source-model intersection; it imports no historical timings.

The 72 cases cross CFB bulk reads and ordered OPC Parts, large/fresh,
mixed/fresh and large/primed sessions, requested widths 1/2/4/8, and local,
capped, and delayed sources. Local is uncapped with no delay; capped returns
at most 64 KiB with no delay; delayed adds 250 µs per successful nonempty call.
The corpus has 32 members: large uses 256 KiB each; mixed has 31 large members
and one 4 KiB member. OPC stores Deflate-compressed payloads; CFB uses regular
FAT streams. All member, archive and verification identities match sealed 0786.

The task floor is 64 KiB. Every sample uses finite memory, byte, object, work,
worker, I/O and CPU-task limits. Setup, metadata, priming, verification and
teardown are outside the operation clock. The same source model applies to
setup and priming, but their calls and delays are not included in timed source
counters. Worker reservations may remain with the session after the operation;
worker/I/O counts must be zero after teardown. Primed Parts are a cache-hit
control, while primed CFB still reads payloads.

Root retained all 648 successful child reports and 13,320 measured outputs:
72 qualification samples, six native blocks of 30 samples per case with three
warmups (432 reports/12,960 samples), and two observer blocks of two samples
without warmup (144 reports/288 samples). No measured child was retried or
excluded. Native and observer binaries are separate, and ordinary timing has
no source-counter atomics. CPUs 12–19 are selected on the shared AMD EPYC 9R45
Linux host; its cgroup hierarchy has no CPU quota. This is not exclusive-core
or physical-network isolation. Rust/Cargo 1.95.0 use opt3, thin LTO, one codegen
unit, debug level one, unwind, and no flag override.

The table gives median process p50 in milliseconds at each requested width.
Speedup is the median of six paired width-one/width-eight ratios, not the
quotient of displayed medians. Intervals use 10,000 bootstrap resamples with
seed 816816 and sorted endpoints 250/9749; process quantiles use nearest rank.

| Route | Shape/state | Source | W1 ms | W2 ms | W4 ms | W8 ms | W8 speedup [95% interval] |
| --- | --- | --- | ---: | ---: | ---: | ---: | --- |
| cfb | large/fresh | local | 1.218832 | 1.049075 | 0.677118 | 0.582818 | 2.114723 [2.069453–2.143375] |
| cfb | large/fresh | capped | 1.223846 | 1.079185 | 0.684724 | 0.579988 | 2.110567 [2.075099–2.122116] |
| cfb | large/fresh | delayed | 40.047576 | 20.883247 | 10.569489 | 5.440918 | 7.363746 [7.342908–7.403719] |
| cfb | mixed/fresh | local | 1.172416 | 1.178716 | 1.174506 | 1.174141 | 0.996558 [0.988577–1.001613] |
| cfb | mixed/fresh | capped | 1.184641 | 1.178931 | 1.177691 | 1.179196 | 1.001469 [0.982039–1.022217] |
| cfb | mixed/fresh | delayed | 39.114396 | 39.113316 | 39.105257 | 39.112992 | 1.000266 [0.999577–1.001977] |
| cfb | large/primed | local | 1.157226 | 1.023211 | 0.648223 | 0.501462 | 2.295949 [2.228124–2.369609] |
| cfb | large/primed | capped | 1.167346 | 1.032955 | 0.599698 | 0.521038 | 2.243116 [2.231994–2.275779] |
| cfb | large/primed | delayed | 40.017100 | 20.773585 | 10.506229 | 5.346567 | 7.489920 [7.472460–7.501877] |
| parts | large/fresh | local | 0.777309 | 1.359962 | 0.816079 | 0.785309 | 0.989491 [0.975479–1.008093] |
| parts | large/fresh | capped | 0.774904 | 1.395822 | 0.824049 | 0.785209 | 0.986909 [0.967706–1.002872] |
| parts | large/fresh | delayed | 10.556259 | 6.410318 | 3.342457 | 1.860380 | 5.674399 [5.497420–5.726218] |
| parts | mixed/fresh | local | 0.759499 | 0.758108 | 0.758664 | 0.759634 | 0.999023 [0.994705–1.004592] |
| parts | mixed/fresh | capped | 0.759504 | 0.758339 | 0.757939 | 0.759444 | 1.000395 [0.997037–1.001623] |
| parts | mixed/fresh | delayed | 10.531064 | 10.530964 | 10.529077 | 10.530473 | 0.999988 [0.999621–1.000295] |
| parts | large/primed | local | 0.002995 | 0.194146 | 0.181671 | 0.234101 | 0.012739 [0.012280–0.013038] |
| parts | large/primed | capped | 0.003010 | 0.189311 | 0.181616 | 0.238317 | 0.012749 [0.012240–0.012953] |
| parts | large/primed | delayed | 0.003035 | 0.215102 | 0.201316 | 0.247712 | 0.011908 [0.011652–0.013056] |

The mixed request has a below-floor member and stays on the serial route even
at width eight; all observer samples show a maximum of one simultaneous read. In delayed large
CFB and fresh Part cases, the observer reaches eight simultaneous reads. That is an
observation from the instrumented source build, not a measured native worker
occupancy. Large primed Parts show zero timed active/source reads in every model;
their requested-width overhead cannot be attributed to delayed provider reads.

The independent source-counter checks distinguish logical calls from returned
bytes. CFB large reads return 8,388,608 bytes: uncapped uses 32 calls and zero
short reads, capped/delayed uses 128 calls and 96 short reads. Mixed CFB returns
8,130,560 bytes in 32 uncapped or 125 capped calls, with 93 capped short reads.
Fresh large/mixed Parts use 64 logical calls and return 146,041/143,782 bytes,
with no short reads under the 64 KiB cap. Primed Parts return no source bytes.
The histogram groups requests of zero through 64 bytes together. Logical call
totals therefore do not identify the number of nonempty calls that slept, and
requested delay is not a measured sleep duration or network RTT.

Logical payload throughput describes bytes represented by returned members; it
is not compressed-source bandwidth or memory bandwidth. In particular, a cached
Part can return shared ownership without moving its logical payload. Process
CPU/wall ratios use a slightly wider CPU interval. RSS covers the whole child,
including corpus generation, setup, priming, verification and drops. Tail and
spread observations remain diagnostics; no general tail, RSS, or allocation
improvement is claimed. Apparent Amdahl fractions are descriptive fits; invalid
fractions and residuals must not be interpreted as measured serial phases.

Across the 72 width rows, 21 show negative scaling, 40 carry the p99/p50
tail flag, and 54 have at least one block-spread flag above five percent.
Eight of the 18 unconstrained Amdahl fits fall outside the admissible [0, 1]
serial-fraction range. Primed Parts produce fractions around 92–101, which
illustrate the model failure; clamped fits and residuals remain available.

The [replayable scaling tables](results/change-0816/scaling.md),
[machine-readable curves](results/change-0816/scaling.csv),
[observer rows](results/change-0816/observer.csv), and
[raw analysis](results/change-0816/analysis.json) retain all widths, source
controls, latency quantiles, CPU, throughput, RSS, and fit limitations.

The first numerical outputs and reader snapshots are preserved under
[analysis-attempt-0](results/change-0816/analysis-attempt-0/recovery.json).
Offline review corrected the bootstrap lower endpoint from index 249 to the
frozen index 250 and rejected out-of-range Amdahl fits. The endpoint correction
changes no reported numerical value; fit validity changes in eight families
(32 width rows). No raw measurement or frozen driver changed, and no workload
was rerun. Root independently cross-checked all 18 table families.

All six harness quality gates pass: formatting, all-feature/all-target checking,
seven tests, warning-denied Clippy and rustdoc, and the full crate-boundary
checker. Tests cover capped reconstruction, empty/EOF/offset behavior, CLI
bounds/duplicates, the bypassed OPC route, observer counter conservation, and
public-route output/resource checks. All 35 previously read normative inputs
retain their hashes. Production is unchanged across its 9,196-file census.

[Source review](results/change-0816/source-review.md),
[protocol review](results/change-0816/protocol-review.md), and
[independent results review](results/change-0816/results-review.md) delimit the
evidence. The [packet README](results/change-0816/README.md) records execution
and replay. Reproduce offline checks with
`python3 -B docs/performance/results/change-0816/validate.py --final`.

Root removed the owned build directory after verifying both captured binaries: 2,113 files / 808,118,216 logical bytes. Production and the three unrelated files remained unchanged. The pre-existing tool lockfile was preserved. Post-cleanup replay passes after
two reader-interface corrections documented in the
[execution notes](results/change-0816/execution-notes.md); numerical outputs
remain unchanged.

This closes the selected low-level CRUD category-15 provider/scaling evidence
gap for these cases. It does not establish native-format CRUD, physical cold
behavior, real-network performance, cross-session contention, allocation
savings, or universal scaling. The queued [next corpus step](results/change-0816/next-step.md)
uses tracked Office inputs with fresh selector admission and output checks.
The broader OLE2/OOXML goal remains active; ODF is deferred and iWork excluded.
