# 0512 final profile and evidence review

Independent read-only review finds the profiles valid for the narrow claim:
aggregate simulated instruction references during three active dense-wide
`Edit::commit` bodies. Both raw runner-to-commit edges have three calls, and
their aggregate equals the retained profile total. The four selected immediate
commit children have six calls each and sum to about 99.13% of the collected
cost. They form a disjoint major-branch partition, including descendants;
they are not exclusive leaf costs.

The active-ancestor 100% rows after `--zero-before` are not evidence that setup
was collected. Raw direct edges establish the operation boundary. Nested
parser/scanner/address/tag/MCE diagnostics overlap and must not be summed.
Descendant call counts may include collection-off calls; only the selected
direct edges prove measured invocations.

The two `brk segment overflow` warnings are material instrumentation caveats.
Both workloads complete with full summaries and exit 0. This supports the
scoped instruction diagnostic but no allocator, RSS, native-cycle or memory
claim. Whole-child hardware counters are separately labeled and fully covered.

The verifier passes source/binary/fixture/tool custody, raw durations, corpus
and sink identities, exact profile selectors/edges, recomputed attribution,
hardware group quality and the negative short-vector case. Replay matches
before and after owned scratch removal. No new production test suite is claimed.

There is no source or completed-evidence blocker to retaining this profiling
batch. No production optimization is admitted. The next parser/snapshot design
still requires operation-local allocation and phase evidence, MCE/original-byte
identity, exact error-order proof and bounded overlap. Generic commit selectors
currently lack allocator observations; save requires a post-oracle boundary.
Further ODF work remains deferred until the full OLE2/OOXML goal is complete.
