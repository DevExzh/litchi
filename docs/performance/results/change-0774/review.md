# Integration review and decisions

Read-only review of original 0766 against current production found no ordinary
compile/correctness blocker, but identified two issues requiring action.

1. The formula-token map retains derived bytes and its table capacity without
   a configured memory bound, release control or observability. ADR 0032 covers
   small immutable snapshot memos and does not supply a policy for this mutable
   authoring writer. Rather than expand retained unbudgeted state, integration
   commit `3376806c11` removes that map and reuse path. Early validation and the
   static function table remain. Existing writer-wide budget migration remains
   outside this change and is not declared complete.
2. External-workbook registration checks individual inputs, but cumulative
   ExternSheet entries can exceed the writer's single-record bound and be
   rejected only at serialization. The integration adds aggregate preflight
   and failure-atomic boundary coverage; see the final source/gate receipts.

The initial eight gates and 1,024-case property run bind `b5739a89e8` and must
not be represented as validation of later edits. Final-source gates are retained
separately. The first performance probe build failed on a Rust formatting-string
capture in the argument parser; its log and source are retained under measure-0.
It produced no measurements. The corrected probe is built for the final source.

Final read-only review of `836120efe7` found no concrete defect. The helper
matches the stream predicate: defined names select one internal entry per
worksheet; records/pivots without defined names select one workbook-wide entry;
otherwise internal count is zero. External sheets, first add-in marker and each
DDE/OLE link contribute one-for-one. Checks precede mutations, and the BIFF
writer independently validates the same aggregate before emission.

The added test covers aggregate N−1/N/N+1 and refusal byte equality. Dedicated
near-limit tests for every individual internal-activation transition remain a
coverage improvement; the reviewer did not find an implementation blocker.
