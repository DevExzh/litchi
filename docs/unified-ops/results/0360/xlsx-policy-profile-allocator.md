# XLSX policy-profile allocator receipt (superseded)

The current source-reader and tab-state policy-profile evidence is recorded in
[xlsx-reader-state-repair.md](xlsx-reader-state-repair.md) as record
0360ae. That record supersedes the earlier 39bd753/f75bf4 reader-state
receipt and carries the c8ef binary hashes, matched Callgrind distribution,
default/managed/custom allocator rows and exact source-test results.

The earlier raw reports remain outside the checkout at
(deleted-coord-artifact)
and
(deleted-coord-artifact).
The earlier single matched Callgrind pair was +248,268 Ir; it is retained as
historical diagnostic evidence and is not an acceptance result. The earlier
candidate's +64 lazy reader-state calls are superseded by the c8ef result,
which has identical ordinary-reader calls to e63 and lower byte totals.
