# Log sections for change 0652

## For `HOTSPOTS.md`

## 0652 — the ten decision rows of the queue are decided; the third wave may implement what the second wave priced

Record: [0652](0652-owner-decisions-for-the-third-wave.md). The owner decided
every row of the refreshed queue that waited on a human: the MCE codec's
namespace re-declaration may change its consumers' public API (row 1); the
publication audit of original part bytes loosens and accepts non-compact XML
(row 2); ADR 0030 (lazy OPC part decode) and ADR 0031 (execution-context
budgets) are accepted (rows 3 and 6); the memoized PPTX revision proof lands
with its format bump and the old patches invalidated (row 4); the cross-copy
candidate is retained under a new budget (row 5); the XLSX value editor widens
its admission surface by 0602's D4 (row 7); the ineligible-read gate may move
its refusal's timing (row 8); the DOC fence gets its second public `litchi-cfb`
entry point (row 9); DOCX compaction and OLE2 sector reuse become
policy-controlled with preservation-by-default (row 10). Three standing
trade-offs bind the wave: breaking changes are acceptable in early alpha,
correctness and safety come before performance, and most inputs are benign.
The queue table itself is unchanged by this record; the third wave's closing
record will refresh it.

## For `GOAL_AUDIT.md`

## 0652 — the owner's trade-offs read against the goal: nothing in the definition of done moves, the blocked rows do

Record: [0652](0652-owner-decisions-for-the-third-wave.md). The three
trade-offs restate `docs/GOAL.md`'s first rule (correctness, lossless
preservation, bounded resources and safety over speed) and add two things the
goal did not say: that public API may break while the crates are 0.0.x, and
that the common path is the benign one, so a defence may move off it as long
as the malicious minority is still refused with the same typed error before
any partial result reaches a caller. No clause of the definition of done is
weakened; ten rows that waited on a human move from "blocked" to "authorized",
each with the invariants its implementing record must prove written beside
the decision.

## For `REPORT.md`

## 0652 — a decision record, no numbers

Record: [0652](0652-owner-decisions-for-the-third-wave.md). This record
measures nothing and claims nothing. It records the owner's decisions of
2026-09-16 on the ten rows of the queue that waited on a human, and the three
trade-offs that bind later waves. The figures those decisions unblock are the
priced ones of the records they cite (a real-deck PPTX edit −94% behind the
codec row; −26% of an opened PPTX lifecycle behind the revision proof; the
apply phase of a cross-copy −72%; a DOC open −20%; 94 of 95 real packages
admitted to source-backed publication), and none becomes a result until an
implementing record measures it on landed code beside a floor.

## For `ADR_COMPLIANCE.md`

## 0652 — ADR 0030 and ADR 0031 accepted; three amendments assigned to the records that implement them

Record: [0652](0652-owner-decisions-for-the-third-wave.md). Proposed ADRs 0030
and 0031 are Accepted as of this record and move into the accepted table of
the ADR index; the rule that no code cites a proposed record is unchanged, and
until this commit nothing cited them. Decisions 2, 4 and 5 will amend ADR 0006
(the original-bytes compactness statement) and ADR 0005 (retained state: a
facade-carried memo, a budgeted retained candidate) in the records that
implement them, each naming 0652. Decisions 8 and 9 move a refusal's timing
and a fence's read count within ADR 0003 and ADR 0006 as those records read
them; decision 10 adds two public policies whose defaults are
preservation-by-default (ADR 0006). The trade-off that breaking changes are
acceptable does not touch ADR 0001's layering or ADR 0005's leakage rules.
