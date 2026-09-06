# 0441: Share the immutable preservation projection

Current source 25b5516b0 clones validated slides into a detached staging Vec,
then clones that Vec again as the pristine source comparison model. Only
read access is used for the latter. The snapshot already owns the identical
validated projection behind Arc. Share that immutable authority while keeping
the draft Vec detached. Keep page coverage refusal, source-index mapping,
content comparisons, all XML parsing and publication validation unchanged.

The retained 0440 candidate profile places 30.56% of weighted sampled periods
under transaction, but only 0.266% under visibly symbolized clone frames there.
Inlining and missing frames limit attribution. This motivates ownership and
allocation measurement; it does not predict a large latency improvement.
The expected gain is one fewer deep source-model copy and a lower operation
peak. The lifecycle still materializes the complete source and candidate.

The frozen A1/B1/B2/A2 protocol uses fresh binaries, 24 reports, 720 samples,
three shapes and normal/allocator modes, CPU 2 and one worker. Baseline A1
precedes source edits. Capture/profile/oracle/protocol scripts are frozen before
builds. Root CPU work is serialized; source changes occur only between terminal
jobs. Four whole-process profiles follow the main matrix. At least 5% fewer
medium/large allocation calls or requested bytes in both repeats is required,
or at least 5% lower normal p50 in both repeats. Review all adverse >5% flags;
no automatic acceptance. Native breadth, cold/range, scaling, bounded existing
append and the complete non-iWork goal remain open.

Accepted ADRs were read in earlier goal turns and the unchanged tree remains
c950b6c8be822561b498d7bbe87c460873dcbf49. ADR 0003 requires immutable sharing and
isolated edits; 0005 requires explicit ownership and measured memory scope;
0006 requires exact preservation and fail-closed coverage/readback. The change
stays inside the ODP owner under 0002/0023/0024. No public API, unsafe code,
dependency, ambient I/O or validation exemption is introduced.
