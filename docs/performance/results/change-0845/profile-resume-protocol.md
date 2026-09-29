# Additive profile resume protocol

`profile_resume.py freeze` records the immutable state after the first four
Callgrind commands. It binds the frozen plan, both source/build/binary legs,
the eight selected qualification identities, the original failed
`profile-capture` command, every prefix artifact and command descriptor, and
the v2 reader descriptor. The failed command remains an explicit controlled
failure; `capture-profile.json` is never synthesized.

The freeze also binds `reader-repair-spec.json`, its installed Callgrind format
reference, and the completed v2 regression result. The bounded repair accepts
Callgrind `fe` source switches and `cfl` callee-file aliases while retaining
owner and cost-conservation checks.

`profile_resume.py resume` rechecks that freeze and runs reader v2 over all four
existing reports before writing the missing fourth receipt. It then runs plan
rows 4 through 15 in order, refusing existing row or command directories and
sealing each receipt only after report, owner, terminal-dump, conservation,
command, source, binary, and artifact checks pass. Native and Heaptrack
commands remain separate and are checked for interval overlap.

`profile_resume.py analyze` replays every row with reader v2 before compacting
raw hashes, the complete selected function/edge rows, and owner costs into
`profile-analysis-v2.json`. Inclusive edge costs remain concrete rows because
selected regions may overlap. Missing or inlined selected symbols are valid;
file and line fields are retained as parser context only. The result makes no
native-cycle, phase-time, or pooled profile claim.
