# 0705 current-head XLSX edit/save diagnostic

This packet refreshes the existing source-backed one-percent scalar-cell
edit/save selector on medium and dense-sparse deterministic workbooks. It
does not compare current timings with historical binaries or change
production code. iWork is excluded.

`build.py` records host information and checks an exact Rust/manifest census
before and after building the existing native harness. `plan.json` freezes
the selector, two shapes, CPU affinity, twenty warmups and one hundred samples
in each of two fresh children per shape. `capture.py` reverses shape order
in the second repetition and binds each result to the binary, source census,
plan and acquisition script. `analyze.py` independently checks phase sums,
statistics, corpus identities and correctness/resource evidence.

Run from the repository root:

```sh
python3 docs/performance/results/change-0705/build.py
python3 docs/performance/results/change-0705/capture.py
python3 -B docs/performance/results/change-0705/analyze.py
python3 -B docs/performance/results/change-0705/producer.py
python3 -B docs/performance/results/change-0705/allocations.py
python3 -B docs/performance/results/change-0705/rss.py
python3 -B docs/performance/results/change-0705/profile.py
python3 -B docs/performance/results/change-0705/analyze_allocations.py
python3 -B docs/performance/results/change-0705/analyze_controls.py
python3 -B docs/performance/results/change-0705/analyze_profiles.py
```

The named target directory is owned scratch and is removed after evidence
capture. Reproduction requires rebuilding it. Original receipts remain
historical records. For a new acquisition, copy only the scripts and plan
files to a fresh sibling packet directory, set `plan.json` revision to the
checked-out `git rev-parse HEAD`, and preserve the original packet. The build
records a new source census; a different census is a new baseline, not a
reproduction of these source-bound results.

The timed interval sums open, planning, staging plus commit, and sequential
publication, including the returned publication snapshot's drop. Sink setup,
other handle destruction, output reopen and semantic/preservation oracles
are outside that sum. The provider is an instrumented in-memory `ReadAt`;
logical read counters do not measure physical device I/O or decompression.
Fresh editor/cache instances do not constitute a physical cold-cache test.

`producer.py` measures the separate producer Edit variant's planning-sheet
numeric edit, with identical output verification. `allocations.py` builds the
isolated allocator binary; its times are excluded from native conclusions.
`rss.py` observes whole-child RSS including setup and oracles. `profile.py`
collects guest instructions only within the publication method; its scope
excludes the returned snapshot destruction included in native publication.

`initial-build.*` retains the first full build. The final `build.*` receipt is
an incremental recheck with root workspace manifests/configuration explicitly
hashed before and after the command. Both binaries have the same digest.
The source census covers Rust files and Cargo manifests/locks under crates
and the standalone harness; `workspace-inputs.json` adds the root inputs.

The audit can be replayed with `python3 -B .../audit.py` after evidence gates
finish. It independently recomputes the native, allocator, producer/RSS and
profile conclusions. `artifact-seal.py --check` verifies every retained packet
file. The seal excludes only its own manifest.
