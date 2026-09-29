# 0843 current-source many-Part OPC publication diagnosis

Diagnostic only, based on `03fa423d73`. No production edits or optimization adoption.
The immutable `admission.json` binds the source, final binary build description,
probe, readers, schedule, eight ZIP qualifications and two observer preflights.
The release binary is removed only after closure checks; `cleanup.json` retains
its verified identity. Rebuild using `build_v3.py` in a fresh isolated packet;
do not overwrite the retained capture directories.

Replay retained evidence from the repository root:

```sh
python3 -B docs/performance/results/change-0843/reader_tests.py
python3 -B docs/performance/results/change-0843/heap_reader_tests.py
python3 -B docs/performance/results/change-0843/close_v2.py verify
```

The final schedule is eight qualification processes, 48 native processes with
30 samples each, eight one-sample Callgrind processes, and eight one-sample
Heaptrack processes: 72 formal reports / 1,464 samples. Two one-sample observer
compatibility preflights are separate. Native runs have three untimed warmups;
instrumented timings are never pooled with native timings.

`source-map.md` maps the exact public route and validation boundaries.
`analysis.json` keeps native, guest-instruction and allocation observations
separate. `tables.md` exposes all native metrics and all observer totals.
`crosscheck.json` independently replays 56 native arithmetic rows.

The first two probe Clippy attempts failed before qualification. `initial/`,
`preflight-v2/` and `preflight-repairs.json` retain the source/error history;
`probe-v3-*` are the passing final probe gates. `qualification-pre-final.json`
is a superseded reader result, retained before strict schema/logical crossbinding
was finalized; the final `qualification.json` reuses exactly the same archives
and reports. `archive_check_v2.py` adds changed-target stable metadata checks to
the original independent ZIP checker. No workload was replaced for these fixes.

Production matches 0842 exactly. `production-gate-reuse.json` binds its 6,931
passing affected-owner tests and other applicable gates, with a fresh replay
receipt. Baseline-confirmed iWork example/rustdoc failures remain excluded.
Fresh probe formatting, Clippy, tests and release compilation pass in this packet.
The normative ADR/checklist files and unrelated workspace files retain their
recorded hashes. Cleanup removes only the two `.owner-0843.json` roots.
