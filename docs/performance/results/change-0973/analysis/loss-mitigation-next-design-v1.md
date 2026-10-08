# 0973 future capture loss mitigation note

Status: read-only post-failure design note. It does not alter the sealed 0973 plan, runner, artifacts, or gate. No workload, build, capture, or probe was run for this note.

The retained child-01 capture used `/usr/bin/perf 7.0.14` with `instructions:u`, period `1,000,000`, `dwarf,32768`, and outer `taskset --cpu-list 31`. Its command omitted an explicit `-m/--mmap-pages`, `--aio`, `--mmap-flush`, and `--no-buffering`. The capture wrote 218.216 MB after 873 writer wakeups. It contains 6,931 sample records of 33,008 bytes each and two `PERF_RECORD_LOST` records on CPU 31/event id 5177, each losing 4 records, at timestamps 300158.557290 and 300158.582288. The terminal `PERF_RECORD_LOST_SAMPLES lost samples:8` is the same ID|LOST counter trailer, so the loss total is 8, not 16. The child exited successfully, but the zero-loss gate failed; no replacement child or all-12 conclusion is admissible.

The direct evidence proves event loss in the retained perf stream. A ring-buffer/drain burst is a plausible mechanism because the command left the data-ring size implicit, each sample carries a 32 KiB user stack, and the recorder and workload were launched under the same CPU-31 taskset. The retained artifacts do not expose the effective default ring size, kernel data-head/tail history, or scheduler contention, so buffer pressure and recorder/workload contention remain hypotheses rather than proven causes. Replaying stock, noop, and corrected decoders produces the same loss because they consume the same raw capture; those are not independent loss observations.

## Smallest future diagnostic change

Keep period 1,000,000, `--call-graph dwarf,32768`, CPU 31, serial 12-child order, the existing marker gate (at least 100 marker samples, positive marker period share, zero lost events), and stop-without-retry policy unchanged. Add one explicit perf option immediately before `--call-graph`:

```text
-m 16M
```

Thus the only proposed argv delta is:

```text
... perf record -e instructions:u -c 1000000 -m 16M --call-graph dwarf,32768 ...
```

This gives perf an explicit 16 MiB data mmap size per ring (nominally up to 512 MiB if 32 CPU rings are allocated) to absorb short bursts while preserving the event, sample payload, affinity, and gate. It is a diagnostic mitigation, not a guarantee that the next capture will be loss-free. Record the exact argv and stderr again; accept a child only when `PERF_RECORD_LOST` and the terminal lost-sample counter are both zero. Do not use `--overwrite`: perf documents that it discards the oldest records when full, which would violate the zero-loss gate while hiding the problem. Leave `--aio`, `--mmap-flush`, and `--no-buffering` unchanged for this smallest first test; combining writer-policy changes would prevent attribution of any improvement.

## Bound inputs and tool facts

| item | bytes | SHA-256 |
|---|---:|---|
| retained command JSON | 1,185 | `1d40d7d72eb8c8f56e0f6bf8edbcac96041253a8f816e909b48dc51d11982fac` |
| retained raw perf.data | 228,856,640 | `4140b417724ea9606b6637ae6aade3a874a682bd9baf84d6fa0f6fdb716acbaa` |
| retained perf stderr | 192 | `8305826c550c725e0ba1884ef31f3fbe6ce81659a9623e95e5ba26fca1f42ffa` |
| retained perf header | 1,919 | `3d04b0f539690b173184dfe1fb15e9aee770f29481f001ca0d0c454bd1b027ca` |
| retained child packet | 36,286,196 | `dc89d24f10589ba6978262f691b78fd39f66ef0dc0627ca8ef9a9aa2f4859397` |
| failed-pilot audit | 32,316 | `d5602129520f63a9ce62612d12b9833572f37ce861ac48038873cbeb0e1228c3` |
| frozen 0973 plan | 14,249 | `12bd803923586cf73c0881bfd3ef7e3cbc7a1e4853867c91d422b6333fa2e3b4` |
| 0973 runner | 20,510 | `f4f7af4ec8b8bb1818f3ce25dec046c8b8ad7c23d0a2611c76e24bf10b6aef5d` |
| `/usr/bin/perf` | 25,221,960 | `174229fb48b8557aafa821dd828d883dce47c35c31067e1f31e4a01a91c6c98b` |
| captured local perf help | 48,884 | `e6a80fc6e7074fcf71eeafd9efab50ed3fe74e6feaf55f175d652256a6b98612` |

The local help records `-m/--mmap-pages` as the data-mmap size option, `--aio` max 4 control blocks, `--mmap-flush` as the output extraction threshold, and `--overwrite` as an overwritable ring that drops old records.