# 0807 protocol review

Independent evidence review and root source/driver inspection support fresh
CPU attribution. Current production matches all 9,196 entries in the 0806
before-source manifest. The 0791 CPU profile predates the retained 0792
empty-attribute-tail shortcut and 0800 duplicate-error parity correction.
The 0806 heaptrack evidence measures allocation paths, not CPU cost; its
rejection does not establish the next CPU bottleneck.

The exact 0791 probe, original 0784 owner and 0780 fixture marker are reused.
Thirty-six alternating native control/wrapper children (1,080 measured
outputs), six owner-scoped Callgrind children (six outputs), and two separately
built frame-pointer native profiles (200 measured outputs) remain distinct.
Native controls quantify wrapper perturbation, not an optimization benefit.
The bootstrap seed is fixed to 807080. All workload execution is root-owned
and serial, with no concurrent heavy offline replay.

Require immutable production/probe/plan/tool/host manifests and exact binary
identities before capture. Callgrind must have one positive numbered dump,
one owner invocation, an empty terminal dump and conservation of all self
costs and owner-plus-immediate-child costs. Native sampled ancestry counts
only stacks containing exactly one public capture owner; warmups remain in
sampled stacks and unknown interior frames remain visible. Nested costs and
sample counts overlap and must not be summed as independent shares.

No production optimization, historical timing comparison, allocation benefit,
RSS improvement or broad CRUD claim is authorized by this diagnostic. Independent
raw frame and function-row readers, exact fixture/readback checks, replay
after cleanup, and a seal over the final packet and six documents close the
batch. Cleanup must verify all three executable identities before removing
the owned target. Unrelated files and existing worktrees remain untouched.
