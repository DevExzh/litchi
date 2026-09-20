# Baseline build interruption during patch preparation

The first native baseline compile observed literal patch text in
`opened/cross_copy_plan.rs` and failed before producing its native executable.
The failed compiler log is retained here. When root inspected the failure,
all 602 production source hashes again matched the recorded baseline.

Root stopped the policy preparation agent, prohibited temporary production
writes by every candidate lane, and regenerated the policy patch entirely in
memory from unchanged source. `git apply --check` validated that artifact
without applying it. Root then rebuilt all three baseline executables with
source checks before and after each build. No native executable or timing
sample from the failed attempt is used in the comparison.

The exact transient source file was not captured; the compiler log records the
observed syntax. `policy-initial.patch` preserves the original hand-authored
patch artifact, and the packet's root `policy.patch` is the checked replacement.
