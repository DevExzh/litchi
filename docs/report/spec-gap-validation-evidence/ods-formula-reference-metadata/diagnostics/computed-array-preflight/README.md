# Superseded computed-array preflight snapshot

This snapshot passed seven isolated gates (1,647 tests) and produced one complete
4,650-sample performance capture. Subsequent semantic review found that metadata
shape probing rejected computed matrix IF arguments as reference operators.
These receipts describe the saved selected sources only and are diagnostic,
not acceptance of the corrected implementation. The original capture is retained
in full; its SUMIFS +5.865% timing observation is not erased by later captures.

The first attempted correction caught the public
`EvaluationFailure::Unsupported(ReferenceOperator)` variant. A transient
provider test disproved that approach: projected `SHEET(IF(SUM(A1)=1;A1;A2))`
returned `Array(1,1)` after retrying the failed read. A second focused test
found the same problem in range geometry with projected `ROW(A1:A2)` and a
transient `sheet_count` failure. The accepted direction uses a private
planner outcome for internal deferral and propagates provider failures.
Neither intermediate correction received a performance capture.
