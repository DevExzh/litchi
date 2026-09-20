# ODS byte-position text performance evidence

This directory owns the reproducible performance harness for FINDB, LEFTB,
LENB, MIDB, REPLACEB, RIGHTB, and SEARCHB. The semantic profile is pinned by
the parent [contract](../contract.md), which selects UTF-8 octets, complete
Unicode scalar output, backward snapping of interior starts, and the shared
bounded evaluator resource rules.

The [plan](PLAN.md) records the matched baseline, controls, workload lanes,
exact read bounds, and capture protocol. `harness/` is an isolated process
benchmark with a counting allocator and borrowing resolver. `run_profile.py`
performs contract and lock checks, candidate freeze verification, independent
preflight, matched baseline capture, and fresh-child timing. `verify.py` is the
fail-closed receipt checker; `summarize.py` emits the p50 comparison after a
successful capture.

The source closure includes the byte production modules, scalar and value
dispatch, ordinary text support used by controls, all three byte integration
tests, `FEATURE_MATRIX.md`, the reviewed contract/oracle/goldens, and native
profile evidence. The authoritative dependency lock is the frozen gate copy;
the ambient workspace lock is retained separately in the parent evidence.

Current state at source handoff: harness smoke and all 84 candidate/control
preflight cases pass on the frozen candidate. The final baseline/candidate
timing capture is a separate, hash-locked operation; its report and raw
receipts are added only after the capture and fail-closed verification finish.
