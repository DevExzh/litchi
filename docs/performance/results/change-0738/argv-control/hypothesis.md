# Separately planned argument-parsing control

The main96-process matrix did not restore secondary timing equivalence by
restoring Sample fields. It is retained unchanged. This follow-up holds the
original/prior binaries fixed and changes startup arguments only. The original
probe gains a redundant `--operation format` pair; the prior probe omits the
explicit default `--lifecycle legacy`. Both pairs have option length11 and value
length6; effective PPT operation and lifecycle are unchanged. All full oracles
remain mandatory. This tests sensitivity to the argument-parsing/setup path;
it does not uniquely identify allocator, cache or executable-layout mechanisms.

Before any follow-up capture, freeze72native and24allocation processes, both
fixtures, sameCPU12 serialexecution,9native and3allocationrounds. Native50/3,
allocation1/0. All raw samples retained; originalgate/5%flags/pairedbootstrap
seed7338 with10000resamples unchanged. No pooling with the main matrix and no
candidate reinstatement. Cross-binary matched-argument comparisons remain
confounded by the other binary changes, so do not call them equivalence proofs.
