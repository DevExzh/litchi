# 0497 development attempts

- `validation/focused1.json`: terminal exit 101; 31 tests passed and the new cancellation assertion failed. The OPC boundary maps execution cancellation to `OpcError::Cancelled`; the test had expected `OpcError::Execution(ExecutionError::Cancelled)`. Production was unchanged. The corrected assertion retains the typed cancellation requirement.
- An early root invocation of `test_measure.py` while that helper was still being written failed at import because it referenced a nonexistent `capture.py`. No capture or protocol was created. The final helper test gate must supersede this development attempt.
- The initial read-only rustfmt check found formatting differences in the new DOCX integration tests. Root applied rustfmt 1.98.1 only to that owned test file; the final formatting gate must verify the frozen checkout.

These attempts are development evidence, not passing validation. All final claims require the retained terminal gates.

- `candidate1-docx-default` failed the existing section event-limit test; `baseline-sections1` reproduced it on the original before checkout. A test-only finite-budget correction preserves both the typed memory refusal and event ceiling. `candidate2-docx-default` then caught an unnecessary-qualification lint introduced by the test import. The redundant import was removed; candidate3 is the next full gate attempt.
- `replay-failure1` passed all 33 tail-stream tests, including the added post-prepare authored-replay failure through an actual atomic destination.

- `candidate3-opc-preservation` stopped during compilation because the sparse clones omitted the 13 tracked change-0416 ZIP corpus files embedded by existing tests. Root copied their exact 66af baseline bytes into both clones and extended the supplemental asset manifest. No production source changed; passed DOCX and OPC atomic gates remain applicable. Candidate4 resumes at OPC preservation.

- `candidate4-workspace-check` reached pre-existing iWork library-test compile errors. The user explicitly excludes iWork; no iWork source was changed. Candidate5 runs the workspace/all-target check with the 16 iWork packages excluded, retaining the broad non-iWork scope. OPC all-feature preservation and DOCX/OPC doctests passed in candidate4.

Candidate5 passed non-iWork workspace check, production Clippy/rustdoc, allocator harness suite, harness Clippy/rustdoc. Workspace formatting encountered an existing Keynote formatting difference; candidate6 selects non-iWork workspace packages explicitly without modifying Keynote.

Candidate6 fmt selector initially omitted glob workspace members, resulting in an empty -p list and default workspace formatting; candidate7 expands crates/* and rejects an empty inventory. No Rust source changed.

Candidate7 non-iWork workspace fmt passed. Harness fmt --all also traversed local Keynote dependencies; candidate8 selects the harness package explicitly. No Rust source changed.
