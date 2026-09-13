# Verifier custody audit

Audit date: 2026-09-13. This is a read-only review of `verify.py` against the
0552 evidence currently retained. No build, test, or capture command was run
for this audit. The review file is the only file added by this audit.

The verifier's existing performance-stage checks are structurally compatible
with the frozen plan, receipt schemas, and the eventual two-stage capture
layout. The correctness custody model has moved since the verifier was written,
however. The current bundle now contains a cumulative `candidate-attempts/draft-06`
application and a five-step `public-exact-test-sources` correction chain. Those
artifacts are not read by the verifier, and one legitimate restored-source
state is rejected as a result.

## Findings

### P0: candidate attempt custody is completely outside verification

`verify.py` has no reference to `candidate-attempts` (the only candidate source
check is `validate_source()` around lines 705-733 and 2956-2969). A future
candidate stage can therefore be captured from source that is unrelated to the
reviewed `draft-06` patch while `candidate/source.patch` and its manifest still
look internally consistent.

The current `candidate-attempts/draft-06/inputs.json` binds a 16-file cumulative
patch (`a009dfc5…`), all 16 retained source snapshots, the restored-HEAD base,
`draft-05` as logical parent, the baseline restore record, and the prior
baseline-public-exact-05 receipt. None of those claims is checked.

Add a `validate_candidate_attempts()` component and call it from candidate
correctness and terminal disposition validation. It should, at minimum:

* verify each `inputs.json` and `candidate.patch` hash, safe path inventory, and
  every retained snapshot hash;
* verify tracked `changes.*.baseline` values against `plan.revision`, and
  `changes.*.candidate` against the corresponding snapshot;
* verify the incremental draft chain 01→05 and the cumulative draft-06 patch
  against the explicitly named `restored HEAD baseline` base;
* accept that draft-06 base only when
  `baseline_restore_sha256` matches `baseline-restore-for-public.json`, the
  restore paths/hashes are internally consistent, and its logical parent is
  draft-05;
* bind the final candidate stage manifest's 16 implementation/test entries to
  draft-06 snapshots. Public test files may be checked by their separate bundle
  validator rather than being incorrectly required to appear in the draft
  snapshot; and
* require named passing candidate correctness receipts, including the
  compact-only proof tests and the full source-backed integration test. Merely
  retaining arbitrary successful check receipts is insufficient.

The special restored-HEAD case matters: a validator that assumes every
incremental `application_base_attempt` is another draft would reject the
legitimate current draft-06 custody record.

### P0: the exact public worksheet oracle is omitted from correctness evidence

`validate_public_tests()` (lines 649-699) only validates
`public-test-sources/source-hashes.json` and its two-file `public-tests.patch`.
It never visits `public-exact-test-sources`, its root patch, or attempts 02–05.
Likewise, `validate_baseline_correctness()` (lines 2253-2328) hard-codes only
`baseline-public-01` and `baseline-owner-tests-01`. The passing
`baseline-public-exact-05` receipt, its two tests, and the preserved failures and
corrections are consequently not a mandatory input to any verifier gate.

Add a dedicated `validate_public_exact_tests()` validator. Its custody checks
should bind the root source snapshots and patch, then validate the sequential
attempt chain:

* root parent `b9a7ba85…` → `e32fac36…` and child `18c3e957…`;
* attempt 02 `18c3e957…` → `352e820d…`;
* attempt 03 `352e820d…` → `f2696dfc…`;
* attempt 04 `f2696dfc…` → `20b29175…`; and
* attempt 05 `20b29175…` → `2cb65491…`.

For every step, verify the retained source snapshot and patch hash, apply the
patch against the previous snapshot/base, and bind the recorded prior or failed
check receipt to the expected named check. Require the terminal
`baseline-public-exact-05` receipt to be successful with the exact two-test
command and source manifest. Preserve attempts 01–04 as historical evidence,
including their expected failure/pass status; do not replace them with only the
latest source.

The existing old public bundle remains useful and should stay independently
validated. The exact oracle is an additional bundle, not a replacement for the
five original public guards.

### P1: a valid rejected restoration retaining the exact tests is rejected

`validate_final_source()` (lines 2528-2539) allows a rejected final source to be
either the bare baseline or `public_augmented_manifest()`. That helper (lines
2507-2515) includes only the old two-file public bundle. The current exact-test
bundle adds the updated parent test and `public_exact_output.rs`, so a rejected
candidate whose final checkout retains the correctness tests has a legitimate
manifest that is outside the allowed tuple.

Extend the restoration manifest calculation to include the latest exact public
bundle (old guards + compact child + exact parent + exact child), or define a
separate explicitly named `public_oracle_augmented_manifest()`. Validate its
patch replay and live-source equality just as for the existing public bundle.
The bare frozen baseline must remain an allowed restoration form for a cleanup
that removes all injected tests.

### P1: baseline correctness metadata is stale relative to the current public
oracle

`baseline-correctness.json` and its verifier contract describe only
“baseline-compatible public guards.” That was sufficient before the exact
whole-worksheet oracle was introduced. It is too weak for the current
candidate: the latest source-backed integration result includes both exact
whole-worksheet byte preservation and the explicit formatting-whitespace
publication refusal, but no machine-readable record binds those checks to the
baseline correctness gate.

Either extend the baseline correctness schema with an exact-oracle check row and
the augmented source manifest, or add a separate immutable
`public-oracle-correctness.json` record and require it from candidate/final
correctness. In both designs, the verifier must check the actual receipt hash,
exact command, exit status, test count (`2 passed` for the latest exact child),
and source manifest containing the latest parent/child hashes. Do not silently
change the meaning of the existing historical record.

### P1: candidate correctness receipts are generically validated but not
semantically required

`validate_all_check_attempts()` (lines 2227-2239) validates receipt shape,
source stability, and hash custody for every retained check, but it accepts any
set of labels, commands, and outcomes. `validate_baseline_correctness()` is
the only place that names expected correctness checks, and it names baseline
checks only. Thus a terminal bundle could omit candidate proof/integration
checks while retaining unrelated successful receipts and still reach the
performance/quality disposition path.

Add an explicit candidate correctness manifest or named expected rows for the
passing proof, integration, and relevant preflight checks. Keep generic
validation for exploratory failures, but require the named final candidate rows
and their expected source-manifest hashes before adoption or rejection is
certified.

### P2: the current bundle already contains a verifier bytecode cache

`docs/performance/results/change-0552/__pycache__/verify.cpython-314.pyc`
currently exists. `validate_seal()` explicitly rejects any `__pycache__` (lines
2938-2945), so final sealing will fail until this owned cache is removed. The
final cleanup/seal procedure should remove it and invoke Python with `-B` or
`PYTHONDONTWRITEBYTECODE=1`; the verifier should not create a replacement while
being replayed.

## Positive compatibility checks

The following current design points should be preserved while making the
changes above:

* `validate_source_patch()` correctly accommodates untracked Rust snapshots by
  requiring a retained byte witness, which is needed for the new source-proof
  files before they are committed.
* `validate_check_attempt()` accepts failed development attempts while still
  requiring source stability, controlled target/TMPDIR/job settings, and exact
  artifact hashes. The new public correction history can use this mechanism,
  provided its named receipts are additionally bound to the chain.
* `capture.py` intentionally records profile receipts with
  `execution_stage == "candidate"` for both baseline and candidate profile
  binaries; the profile validator's corresponding expectation is compatible
  with that driver and should not be changed while adding custody checks.
* The draft-06 resource fields (8-byte slots, 8 MiB source cap, 131072-event
  cap, and 2 MiB logical proof heap cap) are currently retained in its inputs;
  a candidate-attempt validator should verify those concrete values rather than
  re-infer them from source or discard them.

## Recommended sequencing

Before candidate performance capture, add the two custody validators and bind
the final candidate correctness rows. After the current integration receipt is
terminal, replay the public exact chain and candidate draft chain against the
retained snapshots. Then update the restoration manifest allowance and run the
verifier's component checks. Remove the bytecode cache only immediately before
the final seal, because any subsequent Python import can otherwise make the
seal stale.
