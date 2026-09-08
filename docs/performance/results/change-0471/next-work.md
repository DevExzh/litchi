# Next measured snapshot representation lead

The retained 0468 profile places the snapshot scan at 24.41% of inclusive
commit sample weight, `Scanner::cell_address` at 4.93%, and `wire::tag` at
3.69%. These contexts overlap and are not additive or a current speedup claim.
The 0470 changes removed ordinary web traversal without changing this path.

Each ordinary numeric cell currently owns its tag name, attribute name, decoded
coordinate value and attribute slice. Consider retaining source-borrowed names
and unnormalized values within the ephemeral Layout. Entity/whitespace-normalized
values still require ownership. Quick-XML event accessor lifetimes must be
handled explicitly: an event-local `&str` cannot be returned merely by changing
the field to `Cow`. Original-byte offsets or a source-bound parser seam must
prove the borrow. Check object-size growth as well as removed allocation calls.

A full eager-parser/snapshot fusion has additional costs: eager parsing uses
MCE-processed XML while Layout addresses original bytes; source Store parsing
also happens before exact no-op detection. Building Layout during it may add
work to no-ops and overlap raw semantic cell storage with Layout. Do not retain
a persistent Layout cache or grow the bounded Store handoff to sidestep these
constraints. Measure the common tag representation before that larger fusion.

A second independent review recommends first removing the duplicate checked
cell-attribute traversal while retaining owned Tags. `cell_address()` currently
scans and decodes `r`; `wire::tag()` then scans every attribute again. A local
bounded scratch collection could share those raw attributes while retaining
coordinate-before-other-value error ordering and all duplicate checks. This
avoids the source-borrow redesign but removes less allocation. Neither option
is implemented or measured here. Profile the total expected saving, ownership
size and scratch overflow behavior before choosing; do not add retained Layout
state merely to reuse an event-local borrow.
