# 0439 baseline scope and selection

The preceding goal turn made verified progress: 0438 rejected a fixed-markup
optimization under its predeclared practical-gain gate, retained its evidence,
and restored production. The current checkout starts at `a31fc506e` with only
the user-owned `docs/GOAL.md` untracked. The accepted ADR tree is unchanged at
`c950b6c8be822561b498d7bbe87c460873dcbf49`.

The coverage audit prioritizes the existing-document append lifecycle. The
old `odp_semantic_one_edit_save` opens a Presentation and obtains a Snapshot
before its timer. The timer covers transaction creation, one appended slide,
commit, and copying the committed bytes. That cannot establish an open-through-
publication lifecycle baseline, and fresh streaming creation is a different
append category in the required CRUD taxonomy.

The new opt-in selector is intended to time owned snapshot opening, transaction
creation, appending exactly one title/body slide, commit (including its required
internal readback), and sequential caller-sink output of the committed bytes.
Input cloning, sink setup, independent oracle work, digest finalization and
destruction stay outside the timer. Source and commit owners remain alive at
the operation allocator endpoint; a positive live-byte delta is retained state,
not a leak or a bounded-memory result. The public editing API does not accept
an ExecutionContext, so this baseline cannot claim cooperative cancellation.

The deterministic fixture will include ordinary source slides and an opaque
package member. Required checks cover complete slide order/text, exactly one
append, immutable source bytes, untouched decoded member identities, manifest
consistency, deterministic output, exact no-op and patch forward/inverse/source
guards. Physical compressed preservation must be observed separately and never
inferred from a decoded digest. This is owned input and materialized commit
output; writing those bytes to a non-seek sink does not make commit streaming.

This batch establishes evidence rather than a production optimization. It has
no before/after speedup claim. A new selector does not promote any other CRUD
category, default benchmark row, source-backed path or native support claim.
The complete non-iWork goal remains open.

The source inspection also identified a future shared ODF hypothesis:
`GeneratedXmlReader::fill_fragment` audits and then parses each fragment again
for shape restrictions. A conservative classification collected during the
audit may remove that second traversal. It must preserve declaration/comment/PI
and reference restrictions at every depth and retain audit-first error ordering.
No implementation or performance claim for that hypothesis is included here.
