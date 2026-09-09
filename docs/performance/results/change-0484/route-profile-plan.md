# Route profiler helper

`profile_routes.py` is an opt-in diagnostic driver for the frozen 0484 route
protocol. It profiles the normal copied route executable in fresh child
processes and writes to `route-profiles/<attempt>/`. Its default inventory is
the source-heavy `s131072-a64-short-c64` case and the authored-heavy
`s64-a16384-short-c64` case across `deterministic`, `memory_store`, and
`file_store`. The tool selection defaults to `perf stat`, `perf record`,
`strace`, and `heaptrack`; `--tool`, `--case`, and `--route` select subsets.

The helper requires the frozen `route-protocol.json` and the matching normal
route build receipt. It verifies the copied executable's path, size, hash,
source custody, and protocol hash before starting a profile. The route argv is
constructed by `measure_routes._route_argv`; a successful child report is
checked with `measure_routes.check_route_report` before its receipt is marked
`ok`.

The coordinator runs the helper through `gate.py` when source snapshots and
the CPU lock are required. The helper does not invoke `gate.py`, build a
binary, freeze a protocol, or make a timing or speedup claim. External
profiler observations cover the whole child process and include profiler
overhead. Missing tools, unsupported events, and permission-denied counters
are retained as explicit `unavailable` receipts; absent counters remain
`null`.

For a file route, every profile receives a new scratch replay directory and
the frozen route ceiling and sync policy. The directory is removed only after
the child report passes the route oracle and the directory is verified empty.
Any failed child or cleanup leaves its scratch files and raw streams in place.
Every profile receipt records the exact wrapped command, route argv, binary and
protocol bindings, environment, raw stdout/stderr/resource/report artifacts,
profiler artifact size/SHA-256 metadata, and the profile driver's own
content hash. A timeout starts each profiler in a private process group and
retains the termination signals and group status before the parent is reaped.
Daemonized or escaped-session descendants are outside this helper's
process-group custody.

Example root invocation, with the gate supplied by the coordinator:

```text
python3 -B gate.py --attempt PROFILE_ATTEMPT route-profile-perf-stat \
  python3 -B profile_routes.py \
  --binary /path/to/normal/docx_replayable_tail_append \
  --attempt PROFILE_OUTPUT_ATTEMPT \
  --build-attempt FROZEN_BUILD_ATTEMPT --tool perf-stat
```

The gate attempt, profile output attempt, and frozen build attempt are
independent path-safe tokens. `--attempt` names the new output tree and
`--build-attempt` selects the existing `route-attempts/<attempt>/build-normal.json`
record; when omitted, the build attempt defaults to the profile output attempt.
The build attempt must contain the matching frozen normal route build receipt.
Existing profile attempts are never replaced.
