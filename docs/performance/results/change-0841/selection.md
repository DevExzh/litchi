# Selection and diagnostic boundary

The previous turn made authoritative progress in commit `307434b454`: fresh
CFB output avoids initializing gaps over large streams. The current checkout
matches that commit outside the three preserved unrelated files. All 35
previously read normative inputs and all three unrelated hashes match the
preceding packet. The broad goal remains incomplete; OLE2/OOXML retain priority,
ODF remains deferred under 0758, and iWork is excluded.

The PPT fresh writer still clones each plain text box into private
`UserShapeData` immediately before serialization. Historical 0753 identified
that conversion as material remaining work. Two subsequent measured attempts
are already rejected: 0781 used `Cow<str>` and 0782 used `Option<&str>`. In 0782,
the narrower representation removed the expected allocation events and helped
ASCII payload writing, but many-short and Unicode write/lifecycle cases crossed
the frozen latency guard. The existing whole-process counters include expensive
setup/readback and cannot attribute these operation regressions.

The next action is attribution, not another representation spelling. Reapply
the exact archived 0782 patch only for a fresh current-source diagnostic. Both
legs use the same public probe, corpus generation, semantic and raw text checks,
compiler configuration and equally long executable paths. All ten existing
write/lifecycle cases remain in the native matrix. Four write cases (many,
Unicode, ASCII payload and rich) receive separate owner-scoped Callgrind runs.
A shared non-inlined wrapper surrounds `Writer::write_to`; construction, input
generation, readback and returned-object destruction are outside that owner.
Its emitted symbol and actual call/return are admitted before capture.

Callgrind guest instructions are a distinct diagnostic, not native instructions,
cycles, physical copies or wall-time fractions. Wrapper self plus immediate
child inclusive costs must reconstruct its incoming total; summed function
self costs must reconstruct the positive publication. Descendant inclusive
costs overlap and cannot be summed as independent phases. Every profile must
have one positive owner publication and an empty termination publication.

The candidate is restored regardless of the diagnostic outcome. This plan does
not authorize adoption, infer a native cause from layout, or pool old timings.
A useful result narrows which operation and generated-code differences deserve
a subsequent measured change. Current output artifacts, source and compiler
bindings, raw profiles, all failed commands, per-case distributions and
independent reader checks remain retained.

Relevant accepted constraints are ADR 0001 correctness priorities, 0002/0024
ownership, 0003 immutable borrows, 0005 explicit I/O and measured performance,
0006 deterministic preservation and typed refusals, and 0008 verification.
Private rich paragraphs keep their owned mutation path; the public writer owns
its input strings. The diagnostic does not change an API, dependency, durable
wire, concurrency policy, filesystem durability or unsafe boundary.
