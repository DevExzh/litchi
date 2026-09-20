# 0724 prospective attribution experiment

The rejected 0723 checkpoint removed substantial late-target FAT traversal but
regressed early/missing replay and second-query index construction. Its combined
change does not identify which component accounts for those regressions.

Build four variants against 0d943df447: baseline; only the larger query-cache
layout and logical charge; that layout plus the target-checkpoint selection code
with no checkpoint constructor; and the complete archived 0723 production change.
The last variant retains both the construction and populated target checkpoint.
The independent variable is source, with the same unchanged native/repeat probes.
Each build binds its full CFB/XLS Rust/TOML census and archived changed files.
All library source is restored before native acquisition.

The full-minus-selection comparison combines construction cost and populated
checkpoint effects. The layout contrast includes logical capacity and compiler
layout effects, not only the number of struct bytes. The selection contrast may
also change code generation. These observations cannot isolate an instruction
or prove a general causal cost on other machines. The experiment is diagnostic;
no variant is eligible for production retention from this packet alone.

All observations remain, including adverse tails and drift. Source, fixture,
binary, probe, plan and analyzer identities are frozen before acquisition. The
0723 gate rejection and all prior packets remain unchanged. No iWork work,
public API, dependency, unsafe code or retained production change is planned.
