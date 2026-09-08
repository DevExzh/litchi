# Remaining work after central-directory spooling

The central-directory spool is one prerequisite for bounded-memory fresh
streaming creation. It does not complete `docs/GOAL.md`.

The next production step is a checked generated-name capability for the public
PPTX streaming writer. The ZIP Office wrapper retains normalized names, and
OPC retains full names and ancestor/descendant conflict indexes. Those owners
still grow with part count. The [ownership audit](ownership-audit.md) identifies
the exact containers and the proof obligations for removing them from the
closed generated namespace. Arbitrary-name APIs must keep exact duplicate,
case-equivalence and derived-name validation. An unchecked bypass flag would
not satisfy the requirement.

Integrate explicit caller scratch through the semantic fresh-creation API
without introducing ambient file creation or ZIP types into ordinary CRUD
signatures. Prove the generated topology before output, enforce its sequence,
and test malformed plans, overflow, skipped/repeated members, scratch failure,
and accepted output counts. Then repeat the complete public PPTX operation
measurement over several slide counts with full semantic reopen oracles.
Keep live heap, allocated bytes, caller scratch, output storage and process RSS
separate. The low-level ZIP member corpus here differs from the previous PPTX
slide corpus and cannot close that public-path measurement gate.

The file-backed spool makes one provider write and one temporary record
allocation per finalized member. If the measured cost matters to the public
workload, evaluate explicit buffered append or reusable record scratch under a
separately frozen comparison. Preserve the exact byte quota, short-transfer
behavior and fail-closed publication. Do not silently inflate the configured
replay window or move costs outside the timed lifetime.

Fresh creation, logical append to an existing structure, adding a package Part,
and arbitrary editing followed by repackaging remain separate scenarios. The
full native-producer, source-variant, cold-cache, semantic CRUD, cancellation,
feature-breadth and bounded-worker scaling requirements remain open. Follow
the program audit and required scenario taxonomy; these targeted transport
changes do not replace that scope. iWork remains excluded by the user.
