# 0362ul137c — DrawingML v2 protocol and probe binding

This source-only follow-up preserves `972e33220fa360cefb1ca9c7365a0d482941b4a9`
and closes the deployment-boundary findings.  It does not run Cargo, a native
probe, an allocator capture, or a performance workload.

Materialized v2 profiles now resolve the validator, generic runner, and
compile-driver contract from the explicit `$LITCHI_NATIVE_QUAL_ROOT` protocol
checkout.  The paired runner must itself execute from that root, and the
generic runner rechecks the seven trusted protocol-tool identities before
importing the v2 policy module.  Candidate source roots therefore cannot
replace a validator or activate their pending `CANDIDATE_BINDING`.

The v2 compile adapter now validates the requested probe root against all eight
prepared probe-copy identities before invoking the legacy harness writer.  Its
compile receipt records the probe root and eight identities, checks the normal
and allocator adapter hashes, and receipt consumption repeats those checks.

The source-only suite passes 44 tests, including distinct protocol/candidate
root and arbitrary-probe rejection coverage.  No source pins or historical
measurements were changed.
