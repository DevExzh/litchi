# 0812 — resolve the current scanner's concentrated native leaf

This batch follows committed 0811 at 2741474cd5. The previous turn made
progress by sealing fresh native controls and samples. Production and the 35
previously read architecture/taxonomy/goal inputs remain unchanged.

Question: which instructions correspond to scanner offset 0x24d, which appears
151/217 and 223/254 times among the scanner self-leaf samples in the two sealed
0811 repeats? Those preliminary counts came from the retained frame text; the
final reader must independently validate the exact owner and sealed hashes.

Rebuild only the original frame-pointer probe, using the same source, probe,
Cargo arguments, flags, and original target path. The absent 0811 target is
temporarily recreated for deterministic reconstruction and owned by this batch.
Do not mutate any sealed 0811 artifact. Compare the complete rebuilt binary's
SHA-256 and length to the sealed binary witness before mapping historical IPs.
If they differ, historical instruction mapping is unproven; retain that result
and do not imply code identity from a matching symbol name or offset alone.

On exact identity, retain the scanner's nm symbol and bounded objdump output,
join exact-owner scanner leaf offsets to instruction boundaries, and verify
all counts against both sealed readers. Static source reviews independently
consider event payload movement and namespace/attribute traversal feasibility.
No workloads, new timing samples, candidate, production change, or adoption
claim belongs to this batch. Sampling skid and frame-pointer perturbation
(+6.925% large capture in 0811) prevent assigning causal instruction costs or
ordinary-build phase fractions. This narrows the next measured experiment.

Root alone runs the rebuild and binary tools. Readers may inspect completed
artifacts afterward. Recheck the 0811 seal after removing only the recreated
target, preserve unrelated workspace files, and seal/commit this packet with
its report and performance-index updates. No fresh production correctness
claim is needed for an unchanged-source disassembly diagnostic.
