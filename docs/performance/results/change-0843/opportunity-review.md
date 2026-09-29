# Independent opportunity review

The read-only source/profile reviewer recommends missing-ancestor staging in
`PartNameIndex::insert` as the smallest next experiment. At 4,096 Parts,
Part-name insertion accounts for 20.39–20.45% of owner guest instructions.
Both allocation captures identify the temporary ancestor Vec and String sites;
`ancestor-allocation.json` independently selects their full-stack rows.

Look up existing ancestor keys by borrowed string, stage only missing keys,
and complete every fallible staging/reservation operation before mutation.
Retain `or_insert(at)` semantics, the first descendant representative, and
current duplicate/equivalent/derived conflict ordering. Keep randomized hashing
and all source/admission validation. Test shared ancestors, distinct/deep
ancestors, insertion order, rejection atomicity and current error behavior.
The current fixture suggests avoidable allocation events but does not predict
native latency benefit. A controlled current-source comparison remains required.

Whole-source index retention/reuse is deferred: the open-time index is currently
dropped, so retaining it would create a parsed metadata memo requiring ADR 0032
review. A replacement-only validation bypass would require separate proof for
constructor coverage and error ordering. The focused local staging experiment
avoids those broader architectural questions. Preservation-index ownership and
topology-limit scans remain measured follow-ups.

No production files or benchmark workloads were changed by this review.
