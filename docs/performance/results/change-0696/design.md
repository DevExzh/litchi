# 0696 — skip empty local namespace installation

performance_claim: none

The sole production candidate adds an empty-list guard around
`c.ns = c.ns.with_local(local_namespaces, lim)?` in MCE `start`.
`with_local` already returns `Ok(self.clone())` immediately for an empty list;
the assignment then releases the previous equivalent namespace owner. Skipping
that operation preserves the namespace head, binding count, all inherited
pointer identities and every later QName/directive decision.

The guard follows the existing attribute decode and namespace syntax checks.
A declaration with an empty default namespace still creates a nonempty entry
and takes the original installation path. Nonempty declarations still perform
duplicate detection, binding-limit checks and parent-linked layer creation.
The empty branch has no recoverable error; no error is moved or cached.

The measured basis is 0695's current-source MCE attribution: the five repeated
presentation calls account for only about 3.3% of isolated cycle/time weight,
while shared `start` has a 32.22% isolated self profile. Namespace clone/drop
annotations support investigation but do not themselves prove machine-level
savings. Baseline/candidate native assembly is retained separately.

No streaming-parser path, public API, persistent representation, retained cache,
resource policy or parallel behavior changes. Existing immutable Arc scope
ownership remains. All 33 previously-read goal/accepted-ADR constraints are
hash-verified in baseline.json. ADR 0001/0003 preservation and snapshot rules,
ADR 0005 measured performance and resource rules, ADR 0006 validation/error
contracts, ADR 0008 verification and parent-linked namespace architecture, and
ADR 0024 ownership remain applicable. No new architecture exception is needed.

Both native and separately instrumented allocation probes are frozen before
the production edit. The baseline refusal and shared-MCE oracle binaries are
also frozen before source changes, using their final locked dependencies.
Baseline A/A, allocation, refusal and native profile runs precede the guard.
Candidate measurements use the same probes, fixtures, marker counterfactual
and parameter choices. Keep every >5% metric as a review trigger and reject a
candidate lacking a practically useful representative result.
