# 0547 baseline OLE2 attribution protocol

The previous goal turn made progress: 0546 retained and committed measured
OOXML/XLSX gains. This batch follows the ranked OLE2 opportunity with a fresh
baseline-only attribution; ODF remains deferred and iWork excluded.

The frozen plan binds the current committed source, CPU 2, one normal release
binary, and four constructor-selected Callgrind children: XLS owned-source open
and CFB tiny, many-small and few-large. Each child has zero warmup and five
measured samples. CFB setup opens are retained but excluded by positive incoming
owner ancestry, not merely dump ordinal. Termination dumps are retained.
Instruction addresses must map to the measured binary's exact disassembly.

The coordinator executes one release build, profiles and assembly commands
serially. The benchmark's existing correctness oracles and deterministic corpus
manifests remain enabled. No runtime source or harness change is permitted.
Each command binds its source manifest, executable hash, timestamps, full raw
output and artifact hashes. No shared /tmp storage is used for the build;
the sole owned target contains its own tmp and retained executable directories.

Analysis separates constructor inclusive Ir, collector self Ir, direct callee
Ir and collector inclusive Ir. Disjoint instruction categories partition self
Ir; direct children are never added to those categories or to the constructor
as if independent work. Setup/collection-off call and jump metadata are not
operation-local execution counts. One baseline repeat ranks mechanisms; it does
not admit a candidate, prove native speedup, or characterize tail variability,
allocation, RSS, cold caches, I/O or scalability.

The exact source manifest is compared with the sealed 0546 tested candidate.
That permits reuse of nine final quality commands and 1,306 prior passed tests;
0547 does not count them as new executions. New analysis and finite graph-model
checks receive their own receipts after profiling terminates. The model can
prove results only within its declared finite domain; allocation failure and
Rust integration remain explicit future proof obligations.

Before committing, replay all analysis and custody checks, bind the retained
binary hash, confirm no accessible live process references the owned tree,
remove it, seal all evidence, and run strict verification. Preserve failed
attempts and any analysis-only adaptations. Any later production candidate
requires a new frozen two-repeat ABBA native/allocation/profile protocol with
malformed-input guards; this diagnostic cannot stand in for those gates.
