# Harness review

The independent read-only review found no blocker in the frozen Rust harness.
Normal and allocator dispatch reach the same standalone module. Lifecycle mode
uses the original five public calls within one clock; phase mode measures those
same calls separately and checks their sum against the enclosing interval.
Hashing for result checks, report assembly, final drops, and fixture preflight
remain outside both timing modes. The publication sink's own hashing remains
inside the publication call in both modes.

Warmup results are checked and discarded; measured indices start at zero.
Corpus construction establishes patch replay, inverse, stale-source, no-op,
manifest, and preservation gates. Per-row checks establish actual source,
candidate, and sink identity plus changed/non-noop status. Preflight patch gates
are not represented as per-row patch replay.

Allocator phase boundaries are nonnested. Earlier phases' retained objects stay
live at later phase entry, and retained report rows cause absolute baselines to
grow across iterations. Region peaks must therefore be interpreted relative to
each entry; adding phase peaks would be invalid. CLI limits, checked iteration
counts, create-new output, and error propagation were also reviewed. This
review did not run builds or tests and makes no performance claim.
