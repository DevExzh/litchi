# Source review and measured next step

Two independent read-only reviews inspected the public edit route. The initial
review identified worksheet lossless layout construction, MCE/x14ac processing,
metadata cloning and publication catalog work as plausible contributors. These
were hypotheses, not allocation attribution. Root's two fresh Heaptrack runs
then localized 90.448939% of requested bytes to three column maps, changing the
priority relative to those broad hypotheses.

`column.rs:420-429` constructs `Assignments<T>` with 32,768 nodes. The parser
creates `Assignments<Properties>` at `raw/worksheet/codec.rs:1008-1017`; both the
initial store (`transaction.rs:1877-1884`) and required post-write readback
(`:2309-2313`) reach it. The observed allocation is 1 MiB for each invocation.
`raw/worksheet/edit/validation.rs:54-98` builds `Assignments<usize>` at line 68,
including when the column-action map is empty; this allocates 512 KiB.

The first candidate is `if actions.is_empty() { return Ok(()); }` immediately
after the existing protected-sheet check. For the measured value-only A1 edit,
there is no column action. The existing protected check only refuses when an
action key exists, and no later action can query the owner tree. The candidate
removes that unused allocation without moving a real column-action refusal or
changing range-assignment behavior. It still requires fresh correctness and
matched performance verification before adoption. This packet does not edit
production or claim the predicted saving as a measured improvement.

A broader predicate could skip ownership construction when no action sets a
style without also setting width. That depends more tightly on the current
`StyleNeedsWidth` predicate and is not the first proposed change.

The two parser maps need separate design work. `column.rs:482-524` propagates
inherited values across overlapping ranges. ADR 0008's column contract requires
last-matching-record semantics, omitted attributes remaining omitted and bounded
work under malicious wide overlap. A sparse representation must preserve all
three operations (`assign`, `get`, `into_ranges`), typed allocation failure and
error order. Neither deleting the post-write parse nor reviving the rejected
0471 buffer-lifetime / 0514 parser-fusion experiments follows from this profile.

Root owns all executions and Git. Agents prepared the probe, strict reader,
numerical analysis and reviews without running workloads or build gates.
