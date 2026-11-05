# 0360ada — Return XLSX raw batch admission and policy gaps

The coordinator returns raw slicer/timeline candidate `3f0bc5c51c` for the
following repairs. Its componentwise intersection with retained package limits,
exact relative-target length calculation, and preservation of caller-selected
native IDs and part paths are useful and must remain. The passing tests and
single-candidate allocator diagnostic in 0360ad do not close these gaps.

`timeline::Edit::insert_caches` serializes each value during preflight, then
clones the entire batch before the storage layer checks collection size,
identity grammar and duplicate identities. Preflight currently checks an ID's
length but not its grammar. Admit these inexpensive batch invariants before
serialization or model cloning, and avoid serializing the same model again
when the admitted bytes can be carried into staging.

Owner edit verbs also take a per-call `&SourceLimits` on this candidate.
Record 0351 §2.5 requires owner configuration through retained state or
`edit_x_with_options(EditOptions)`. Preserve the ability to tighten a raw
insertion through that options entry; the raw parser/record functions retain
their required explicit limits argument. A raw argument type does not make its
owner edit method a raw parser.

The limited raw path retains the caller/package intersection, but public load
helpers and the final graph validation use default limits on reparsing. Thread
the effective policy through these reads. An opened package's tighter policy
must not disappear during initial or repeated family operations.

The bounded XML writers stop at `max_part_bytes`, then compare the completed
allocation against `max_total_part_bytes`. A tighter total ceiling therefore
refuses too late. Bound each sink by the available aggregate budget as well as
the per-part ceiling. Batch admission must account for the final package's part
count and total bytes. Similarly, `check_relationship_insertions` compares an
affected owner's relationship count to the global relationship ceiling without
including other owners. Preserve global accounting without adding a complete
archive scan for every cache. New native part names also require admission
against the retained archive-member-name ceiling, independently of their
relative relationship-target spelling.

Several load/storage/graph helpers iterate all package parts, including paths
inside cache processing. Distinguish existing cross-owner integrity checks
from newly repeated work, retain the integrity proof, and reuse one bounded
catalog/accounting pass where possible. The record's owner-local measurement
claim needs correction: one affected relationship plan does not establish
that the complete operation avoids unrelated package scans.

Repair record 0360adb must cover tight retained policies, malformed/duplicate
identities, oversized batches, aggregate part/relationship budgets and early
allocation refusal. It must preserve native identity capabilities, replay,
inverse replay and unchanged parts, and provide matched allocation/instruction
measurements. No capability or performance exception is accepted.
