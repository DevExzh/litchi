# Change 0432 ADR and measurement boundaries

The accepted ADR tree is unchanged at
`c950b6c8be822561b498d7bbe87c460873dcbf49`; the root previously read its accepted
records and rechecked the tree before this batch. This tool-only change follows
ADRs 0001/0002/0010/0011/0024 by retaining semantic writer use and keeping all
instrumentation in the standalone performance package. It changes no facade,
container, source policy, production runtime or dependency.

ADRs 0003/0006: generated workbook correctness remains checked through the
separate ordinary Workbook reader; every timed archive digest and sink count
must match that exhaustive untimed artifact. Normal save/preservation behavior
and existing snapshot/patch semantics are unchanged.

ADR 0005: an explicit single-worker context, finite row/cell/output/work
limits and caller-owned non-seek discard sink remain in use. The existing
allocator executable owns its isolated instrumentation; normal binaries never
install it. The observed allocation region includes context/writer setup,
caller row-text generation, compression and finalization, but excludes the
materializing correctness oracle. The row-window model, callback-order heap
peak and setup-inclusive process RSS must be reported separately. No scalar
window value or endpoint RSS alone establishes total memory boundedness.

ADR 0008: targeted harness/allocator tests, strict lint review, documentation,
formatting and relevant writer regressions are required. Machine-readable
capture identities, raw observations, verifier mutation checks and portable
replay must support the final result. This observation extension does not
claim a production speedup or completion of the whole non-iWork program.
