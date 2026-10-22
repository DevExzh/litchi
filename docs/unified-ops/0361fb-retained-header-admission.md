# 0361fb — Admit retained IWA headers under the requested limits

- Parent: `2f8fa9534b`, the unaccepted constructor candidate.
- Status: correctness validated; fresh allocation and native qualification required.

An archive parsed under permissive limits may retain unknown or noncanonical
header bytes for exact publication. Later validation previously checked only
the decoded model. A seven-byte archive containing one identifier and two
unknown scalar fields reproduced the defect: fresh ingress rejected a
one-field ceiling, while object validation, archive validation, encoded length
and publication all accepted it. The failing regression is preserved with exit
101; it was an assertion failure, not a compilation failure.

All four paths now admit the retained wire header under the requested profile
when canonical comparison selects it for publication. Header, field, nesting,
message-count, metadata, memory and object-size limits therefore cover the
bytes that are actually retained. If the decoded model changes, the existing
canonical output rule discards the old header, which then consumes no output
budget. Payload admission also precedes conversion to wire-length integers.

Ordinary profiles use conservative constant-time bounds derived from the
private admitted header's byte length and fixed schema depth. Each field
needs tag/value or tag/length bytes; nested allocation charges and packed
scalar charges are bounded by their respective wire bytes. If a bound does
not fit or arithmetic overflows, the existing exact wire preflight supplies
the actual typed count. The bound never directly refuses an input. This adds
no owner fields or default-path allocations. Performance remains unqualified.

Boundary regressions cover all four public paths, exact and one-under header,
object and decoded-memory ceilings, stricter unknown-field admission, exact
unchanged bytes, and publication after changing the model. All 115 core tests
(76 unit, six clone and 33 framing tests), formatting and all-target/all-feature
Clippy with warnings denied pass. [The validation receipt](results/0361fb/validation.json)
records exact source hashes, commands and external log sizes and SHA-256s,
including the initial failing reproduction and both successful validation runs.
The final run includes the payload-admission ordering correction.

The protected v4 constructor binaries and interrupted timing attempts in
0361fa describe their original source, not this repair. This candidate needs a
fresh matched build and retained-header fidelity/allocation routes before
native qualification. No performance exception, merge acceptance or push is
recorded here.
