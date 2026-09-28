# 0828 source and scope review

Base: `c990602492106d968898b7310bdd4bc17f3e4fcb`. Production Rust and the ordinary
performance harness remain unchanged. The 35 normative inputs are revalidated
against 0827. The input and complete-output oracle remain the admitted real
`shapes.pptx` and sealed 0821 default output; no historical timing is reused.

0827 measured a 2.494% full-lifecycle improvement and 11.524% edit improvement
for this file after the 0824 compaction proof. Its default/full publication
retains file and directory synchronization. The next compatible investigation
is remaining CPU work inside the edit, not weakening durability.

`opened/model.rs::package_fingerprint_with_memo` must visit all logical payloads
on a cold snapshot to establish the complete-package revision. Names, content
types and relationships are fed separately from the payload-digest memo.
Allocation identity alone cannot authorize reusing a whole-part digest.
`litchi-opc::DeferredPayload` uses a shared once cell; cloning a package does not
by itself prove repeated inflation. Neither initial complete-source validation
nor final dependency-closure validation is dispensable on existing evidence.

The public edit has three useful call boundaries: capture a transaction root,
set the selected shape text, then commit/apply the result. The replacement path
reads the original Scene, finds the raw span, rewrites text, and reads the staged
Scene. Commit compaction can reuse the staged validity proof when the complete
root and outside non-text bytes remain identical. Those checks are unchanged.

The packet-local probe adds named noinline phase wrappers under its whole-edit
wrapper. Direct mode follows the same public sequence without these wrappers.
Each iteration opens a fresh owner before the clock. Output serialization,
full byte/hash comparison, semantic reopening and owner destruction are outside
the clock. The publication snapshot returned by apply is dropped inside the
publish phase, matching the ordinary-save edit helper. No snapshot capture is
performed before timing to warm the owner's digest memo.

ADR 0003 source checks, atomic publication and patches; ADR 0005 bounded work;
ADR 0006 preservation/refusals; ADR 0030 lazy materialization; and ADR 0032
allocation-owned small derived-value memos remain in force. No production memo,
validation shortcut, runtime instrumentation or architecture change is proposed.
