# 0437 ODP evidence draft

This directory is a planning handoff for the ODP fresh creation measurement.
It is intentionally not a runnable protocol or verifier until the new harness
selectors and the external ODP oracle have settled their report fields.

The outer matrix can reuse the reviewed 0435 shape, subject to the final
source and oracle bindings:

- selectors: `odp_buffered_create` and `odp_streaming_create`;
- source field: `odp_slides`;
- slide counts: tiny 64, medium 4,096, large 8,192;
- normal and allocator modes, 30 samples, 3 warmups, two repeats, CPU 2,
  one worker;
- phase order: A1, B1, C1, C2, B2, A2, with normal lanes before allocator
  lanes within each declared phase;
- 36 reports and 1,080 retained samples, plus six whole-process large-normal
  profiles (stat and record for each of the three roles).

The ODP oracle must freeze the actual report schema before adapting the outer
drivers.  Its required identity should cover the deterministic slide title and
body projection, frame geometry, semantic digest, immutable styles and meta
digests, manifest/member topology, output and sink identities, and any source
or operation fields emitted by the new harness. Same-role archive bytes and
cross-role semantic/title/body/geometry/style/meta/topology rules must be
validated according to that oracle. XML lexical framing may differ only where
the oracle explicitly permits it.

The outer implementation should then be copied from the reviewed 0435
capture/build/profile/verify/summary machinery with these constraints:

1. Every path is derived from the evidence-bundle root or an explicit
   descriptor; no inherited 0436 cleanup path or temporary binary path is
   hard-coded.
2. The copied oracle and its protocol digest are recorded in every report
   binding. The oracle lives outside the harness source and is treated as an
   input artifact, not regenerated during a workload.
3. Before-build/preparatory profiles remain separate from the six formal
   profiles. Formal profiles include setup, corpus generation, warmups, timed
   calls, output hashing, and oracle work in their whole-process scope.
4. Summaries retain available and unavailable PMU fields, RSS scope, lost
   samples, symbolization warnings, and source/binary/artifact hashes. They do
   not authorize a speedup, causal hotspot, fixed-memory, or 10x claim.

No 0437 protocol, report schema, capture command, build descriptor, or
measurement artifact is frozen by this draft.

## Lifecycle handoff contract

The separate lifecycle scaffold is the owner of the final cleanup and replay
contract.  Before these drivers are used, the frozen protocol must carry
`cleanup_paths` and an identical `cleanup_expected_paths` allowlist, the
ordered three-role matrix, and explicit receipt lists rather than relying on
directory globs: 18 pilot receipts, six formal profile receipts, two
preparatory profile receipts, and the `replay_drivers` list.  The paths named
by those lists must match the receipt paths emitted here.  In particular, the
capture phases emit six lane receipts each under `runs/<phase>/<attempt>/`,
and profiles emit one receipt under `profiles/<role>/<kind>/` (or under the
preparatory attempt directory).  This draft does not create or infer the
lifecycle lists.

`profile.py --preparatory` is deliberately independent of the formal build
descriptors: it requires the passing `checks/before-build.json` and
`before/binary-copies.json`, validates their source, log, executable, and
revision bindings, and records both receipt hashes.  A draft protocol without
an `oracle` stanza uses the copied bundle-root `verify-report.py` and records
its hash; every formal profile requires the frozen protocol oracle binding.
