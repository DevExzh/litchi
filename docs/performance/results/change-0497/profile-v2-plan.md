# 0497 syscall profile v2 disposition

The retained `strace1` attempt is immutable and failed before the benchmark
child launched: `/usr/bin/strace` rejected the host-invalid `fstatat` filter.
There is no child output, benchmark execution, or performance observation in
that failed attempt.

`profile_v2.py` is copied from the frozen `profile.py` helper. Its only
tracing-filter change removes `fstatat`, which this x86_64 `strace` rejects,
while retaining `newfstatat`. The v2 helper preserves the same report, path,
warmup, ordering, custody, and no-claim validation. Run it as a new attempt
so the helper hash and argv are sealed in fresh `strace2` receipts:

```sh
flock -x /home/zhuhe/.cache/litchi-goal-0484/cpu.lock -- \
  python3 -B docs/performance/results/change-0497/profile_v2.py capture strace2 \
  --build docs/performance/results/change-0497/build-after-<attempt>.json \
  --binary /home/zhuhe/.cache/litchi-goal-0497/retained/after/normal/docx_replayable_tail_append \
  --modes counting,atomic --samples 1 --warmups 1 \
  --source-count 8192 --authored-count 256 --chunk window --text near \
  --provider file-store --replay-sync data --replay-max-bytes 67108864
```

After capture, verify only the new attempt:

```sh
python3 -B docs/performance/results/change-0497/profile_v2.py verify strace2
```

The v2 profile remains diagnostic syscall evidence. It makes no speedup,
latency, throughput, or operation-attribution claim, and it does not repair or
reinterpret the failed `strace1` attempt.
