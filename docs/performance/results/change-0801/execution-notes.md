# 0801 retained execution corrections

The first isolated helper attempt passed both test legs and baseline Clippy.
Candidate Clippy rejected `map_or(true, Result::is_err)` and an identity map in
a candidate test. The failed quality receipts, candidate archives, patch and
minimal test projects remain under the `*-failed-0` paths, with a relocation
record. Root changed all five methods to `is_none_or(Result::is_err)` and
removed the identity map. The parser strategy and decision policy did not
change. All six subsequent helper gates passed: 70 before and 95 after tests,
zero failures or ignores, and warnings-denied Clippy for each leg.

The first direct-probe attempt passed formatting, release build and release
check, then Clippy refused the unused `unchecked_attributes` method in the
candidate module. The direct probe exercises only the checked API, and unlike
the baseline the candidate does not call that trait method internally. The
failed build receipts and exact probe sources are retained under
`build-failed-0` and `probe-src-failed-0`. The built executable was copied to an
owned target subdirectory before rebuilding; its original and relocated
identities are retained in the relocation record. The fix is a narrow
`allow(dead_code)` on the probe's candidate module, leaving its exact helper
source unchanged. No failed-attempt native capture was run.

The frozen source handoff describes a "likely 128-byte phase". The measured
120/128-byte values in this packet are sizes of the complete `CheckedAttributes`
iterators, not the private phase enum. The source-only manifest status describes
that handoff; final quality and measurement receipts record the later execution.
