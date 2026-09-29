# 0845 Callgrind side population

`profile.py` is a packet-local diagnostic runner. It reuses the already frozen
0845 before/after release binaries and the existing qualification ZIPs; it does
not invoke Cargo, rebuild the probe, or alter `plan.json`, `build-*.json`, or
the main `runs/` tree.

The frozen `profile-plan.json` contains 16 publish reports: shared and
deep-unique paths at 64 and 4,096 parts, two repeats, both legs, one sample,
and zero warmups on CPU 12. Each case has one before-first and one after-first
order across the repeats. Output is kept under `profile-runs/`.

Before capture, the runner requires the passing admission inventory and binds
both source/build freezes, binaries, probe files, the main plan, qualification
summaries, and its own plan/script descriptors. Every profile has an exact
command and artifact inventory. The analyzer checks receipt order and rejects
overlap with Cargo, native probe, or Heaptrack commands. Qualification reports
are independently re-read so every profile's source and expected output
identity matches the corresponding main qualification result.

Each Callgrind pair uses the 0843 owner and options. The 0844 reader must find
exactly one positive call to `profile_publish`, an owner inclusive cost equal
to the scoped summary, conserved function self costs, and a zero-Ir terminal
dump. The full report and Callgrind checks run before the analyzer drops
unrelated functions. `profile-analysis.json` retains raw Callgrind hashes,
owner costs, and all matching function/edge rows for topology-limit, physical
lookup, metadata, EntryId, and member-name lookup symbols. A symbol may be
inlined or absent in one leg; that absence is recorded rather than treated as
a production failure.

The resulting Ir values describe guest instruction work inside the scoped
owner. They do not measure native cycles, elapsed phase time, or a summed phase
cost: selected regions are inclusive and can overlap. The profile population
stays separate from the 512-report native/Heaptrack matrix.

Run after admission and before cleanup:

```sh
python3 -B docs/performance/results/change-0845/profile.py capture
python3 -B docs/performance/results/change-0845/profile.py analyze
```
