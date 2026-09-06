# 0445: Plain-source OPC Part-addition calibration

The matched plain source exposes the large cost of the existing read observer.
Large normal median is about 18.9 ms plain versus 62.1 ms observed. **This is
observer calibration, not a production optimization.** Both selectors execute
one shared timed publication body, with identical fixtures and hashing sinks.

- [All 24 reports, comparisons and repeat review](measurements.md)
- [Plain profile and next allocation candidate](hotspot-review.md)
- [Exact timing/data boundary](data-path.md), [validation](validation-notes.md)
- [Frozen protocol](protocol.json), [summary](summary.json), [stack summary](profile-summary.json)

One build, observed/plain ABBA, 720 retained samples, three sizes, normal/allocator,
CPU 2, one worker, 30 samples/three warmups per report. The one observed-allocator
tiny p99 repeat flag is retained. Allocation calls/requested bytes/above-entry
peaks match between modes. Plain source metrics are unavailable with no values;
its actual sink/process/allocator observations remain represented. Six exported
ZIPs exactly match 0444 and pass independent payload/XML/raw-record verification.

After cleanup, portable verification from this directory needs Python 3 only:

```sh
python3 -B verify.py --sealed --cleanup
python3 -B derive.py --check
python3 -B profile-summary.py --check
python3 -B measurements.py --check
```

These commands validate the sealed inventory, source/binary/protocol/argv custody,
fixture/report oracles and derived statistics without Git, original binaries,
Cargo, perf or workload execution. Logs and profiles may be deterministically
compressed. Original absolute binary paths remain identities, not dependencies.
Portable probes exercise an exported bundle and reject custody/argv/seal changes.

To reproduce, use the exact build receipt/environment with the retained source
manifest and candidate files. Rust 1.98.1, jobs 4, incremental disabled, release
symbols and frame pointers are pinned. From the repository root:

```sh
taskset -c 2 tools/perf-baseline/target/release/litchi-perf-baseline \
  --case opc_part_add_plain_lifecycle --semantic-shape large --workers 1 \
  --samples 30 --warmup 3 --json /tmp/part-add-plain.json \
  --corpus-manifest /tmp/part-add-plain-catalog.json
```

Use `opc_part_add_lifecycle` for observed mode and the `litchi-perf-baseline-alloc`
executable for allocator mode. Both source modes bind to the same two executables.
Use a fresh evidence directory; retained capture/profile attempts are not
rewritten. The fixture exporter takes one explicit nonexistent output directory.

Profiles have whole-process and run-frame scope; run frames include setup/probes
outside elapsed time, and inclusive rows overlap. No exact timed-region, native,
cold/range, scaling or bounded-total-memory claim follows. Source review points
to content-type ownership transfers for the next measured production candidate.
Only three standalone harness Rust files change. The full non-iWork goal stays
active. Cleanup preserves both Cargo target directories and user-owned GOAL.md.
