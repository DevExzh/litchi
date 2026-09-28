# 0803 exact-empty check placement control

Before is the exact rejected 0802 helper source, not production. After changes
only the location of the exact-empty conditional: construction always stores
First, and the first next call checks the same lengths and sets Done if equal.
The linear array, ordered map, duplicate preflight, names, offsets, clone
implementation and shared tests remain unchanged. This is a diagnostic
comparison; it cannot establish an improvement over current production or
qualify a candidate for workflow trials. The 0802 rejection stays in force.

An empty iterator is internally First until its first request instead of Done
at construction. Both yield None on every request, and clones must retain
that observable iterator behavior. Debug formatting of private state can differ;
no production API change is adopted. Whitespace and malformed tails still pass
through lexical processing. The existing empty/clone/error regression tests
are identical between legs.

The patch uses archive-relative names and is checked against candidate/before.
It is not a patch against current production. Candidate metadata describes
source handoff; actual quality and measurement results belong to receipts.
