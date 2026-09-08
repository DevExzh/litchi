# Next work after shared DOCX publication

The 0480 handoff removes the overlay's duplicate XML owner. The current large
operation still has a 35,371,102-byte incremental peak and spends most of its
time outside the removed allocation. It remains a materialized random-access
transaction, not the required explicit-window append capability.

## Address the remaining measured scanner and snapshot work

The 0479 control profile identifies the source-backed scanner as 68.037% of
sampled lifecycle periods; 0480 changes no scanner code. Repeated full scans,
per-event name clones and exact-one-element range growth remain candidate
costs. Use current phase data and the same corpus/control protocol to evaluate
reuse of validated projected ranges and immutable target bytes. Source
recapture and source-version/artifact checks must remain in place.

`Patch::apply` still copies its immutable `after` XML and calls `with_xml`,
which rescans it. Sharing that already-owned Arc can remove another duplicate
payload, but it must retain the source snapshot's exact limit/readback rules.
A cached layout requires a proof that the after bytes, paragraph ranges and
all relevant event/depth/byte limits remain valid under the applying snapshot's
policy. Durable decode, inverse direction, no-op and different finite-limit
snapshots need independent tests; prior validation under a looser limit is not
permission to skip the applying limit. Prefer removing duplicate work before
low-level tuning. Do not combine changes with unrelated format admission.

Requested reallocation bytes remain a separate accounting issue: exact-one
range growth requests very large cumulative capacities without demonstrating
physical copy traffic. A checked growth policy or borrowed event names can be
measured independently, including any spare-capacity peak/RSS tradeoff.

## Keep bounded-window logical append as a required capability

The current snapshot exposes complete XML and paragraph ranges; current
reversible patches retain complete before/after XML. An explicitly bounded
tail API needs its own format-owned admission and owner-lifetime design under
ADR 0005, plus the source-checked reversible publication obligations of ADR
0003. Design around stable positional input, finite event/byte/work limits,
explicit caller scratch/replay if needed, sequential output and exact unchanged
compressed-member preservation. No ambient spool or weakened validation is
permitted. Source validation may require a separate immutable replay before
output to preserve preflight refusal semantics.

Count library owners and caller scratch separately. Prove malformed-tail,
source-change, cancellation, partial-sink, section placement and inverse
behavior. The current ordinary copy API refuses section properties and permits
one copy per edit; do not relabel it as a multi-append streaming capability.
Repeated 64/256 appends require separately measured lifecycles or a reviewed
new capability.

The checklist's fresh creation, existing-structure logical append, new Part
addition and arbitrary edit/repackaging remain distinct. Native/cold/range
sources, atomic replacement, broader CRUD and explicit-worker scaling remain
open parts of the full non-iWork goal.
