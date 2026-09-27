# 0773 integration review

Target: `339572acbf570b2614cfa346e73b8f8d6da358f4`, integrated onto
`6074e10e57`. This is a read-only source and evidence review; no Cargo or
native gate was run by this reviewer.

## Result

No source-level integration blocker was found in the reviewed durability
paths. `Durability::Full` remains the default. All three levels keep the
sibling temporary file, complete staging and flush, route validation, and one
same-directory replacement. Only the requested file and parent-directory
synchronizations are gated. The CFB shared publisher closes handles before
cleanup, marks the temporary name published immediately after replacement,
and reports a parent-sync failure as `Committed`; the OPC path has the same
post-replacement distinction. The existing sequential CFB cancellation
checkpoints, source-backed preflight checks, and bounded planning paths remain
outside the durability switch.

The public entry-point coverage is complete for the 0761 OLE2/OOXML scope:
OPC atomic/package APIs, DOCX, XLSX package/workbook, PPTX plaintext package,
XLSB, all four CFB publishers, DOC/XLS/PPT writers, and the listed
source-backed DOC/XLS/PPT overlays. Encrypted filesystem saves and the DOCX
tail-append publisher remain Full-only by design. ODF and iWork are outside
this change's stated scope.

## Evidence blockers

The original 0761 record must be treated as an archival draft until it is
repaired or superseded by this 0773 packet:

* `docs/performance/0761-save-durability-policy.md` contains the literal
  `@HARNESS@` placeholder. Its verification link points to
  `results/change-0761/gates.txt`, which is absent from the packet. The linked
  `results/change-0761/cleanup.json` is also absent.
* The packet README claims `measure/`, `tables.md`, and the gate/cleanup
  artifacts, but the durable packet contains only the probe traces, scripts,
  build logs, and metadata. The corresponding files exist only in the
  untracked 0761 scratch directory and cannot substantiate a published
  record.
* The scratch gate summaries report the harness as `553 passed, 2 failed, 1
  ignored` in both runs. The failures are
  `pptx_native_image::tests::shapes_original_and_resaved_keep_their_typed_behavior`
  and
  `tests::fresh_writer_corpora_are_deterministic_and_identify_the_packaged_stream`.
  Scratch notes call these base-known, but no durable base-side harness
  comparison is in the packet; therefore the 0761 record cannot claim a
  passing harness gate. The same summaries record the known eight base
  all-target Clippy lints and the base `non_iwork_gate` failure.
* `trace/windows.json` is a useful raw result: its verifier booleans are all
  true, and its 24 old-branch windows retain rename while Full/FileOnly/NoSync
  show the expected 2/1/0 `fsync` counts. It was built from the old
  `1d1044e3ac` base and 0761 production sources, not from integrated HEAD
  `339572acbf`; it therefore does not establish the syscall contract after
  integration. It also cannot establish cancellation, failure atomicity,
  typed post-rename errors, or API coverage.

## Required final 0773 gates

1. Finish the current post-integration check, tests, Clippy, docs, dependent
   crates, facade, and harness runs. Preserve each command, exit code, source
   manifest, and any known baseline failure. At review time the fresh quality
   attempt had passed the all-ten-owner all-features/all-targets check, the
   core/OPC/CFB tests (`1738` passed, `2` ignored), and the seven
   `save_durability` suites (`8` passed). The owner-run Clippy and docs gates
   also passed, and the filtered harness `save_durability` test passed. This
   is not a full harness run: the old full harness failures above remain
   unsupported claims until a full current/base comparison is captured.
2. Rebuild the syscall probe from integrated HEAD and the exact `6074e10e57`
   source pair, with source/lock/binary hashes recorded. Re-run all twelve
   routes over existing and absent destinations and verify default versus base,
   `save` versus `Full`, the `2/1/0` sync counts, one rename at every level,
   and normalized syscall subsequences.
3. Exercise the test seams on the integrated tree: Full parent-sync failure
   returns the route-specific `Committed` variant and leaves the new bytes at
   the destination; FileOnly and NoSync do not call a skipped failing sync;
   pre-replacement stage/flush/file-sync/identity/cancellation failures leave
   the old destination and clean owned temporary files; source fingerprint
   changes remain pre-replacement refusals. Verify the OPC/CFB conversions to
   `litchi_core::Error::Committed` and the high-level nested error variants.
4. Keep the final packet free of placeholders. Either copy the raw gate and
   cleanup evidence into it or label the old 0761 record explicitly archival;
   do not promote the old scratch claims as current integrated evidence.
