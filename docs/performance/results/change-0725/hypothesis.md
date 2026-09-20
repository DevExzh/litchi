# 0725 revised XLS checkpoint hypothesis

The 0723 candidate is rejected; 0724 confirms useful late-target work elimination
alongside missing-query and publication costs, without isolating instruction
costs. This fresh candidate retains the worksheet-start/SST checkpoints and one
40-byte fixed target checkpoint, but removes the linear post-scan slot search
and duplicate path-vector allocation. It records the first matching raw frame
in transient sink state and constructs the checkpoint after successful scan
fences while the existing borrowed path remains available.

Indexed missing queries resolve their empty slot range after the existing entry
execution check and worksheet lookup. They skip unused path/resolver/hint setup,
then perform the same final execution and source-current checks. Stored replay
continues decoding all matching frames in existing duplicate/error order. No
source bytes, decoded values or errors are cached. Original earlier-target
fallback remains; later targets can walk the remaining suffix.

The unchanged native A/A+ABBA, repeated-loop, source-count and budget matrices
from 0723 are frozen afresh against ee5e0b0650. The same 5% timing gates and
limited native warm-query 10ns exception remain. Both late-target native and
loop p50/mean must improve at least10% in both pairs. q2 allocation allowance
shrinks to40bytes and zero newcalls; missing q3/q8 must remove exactly the prior
16byte path allocation/deallocation/peak, with no retained-byte change. Other
strict, zero-budget and typed-refusal regions remain exact. These rules precede
any main timing capture. Retain only if all gates pass; otherwise preserve all
observations and restore baseline. No iWork or unrelated document is changed.
