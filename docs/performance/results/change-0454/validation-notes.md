# Validation history

All root CPU commands are serialized. Rust source epochs are held fixed during
owned builds, tests, inventories, and measurements. Receipts retain exact source
manifests before/after and preserve failed attempts.

- The initial unnamed-slide negative control failed to compile because the local
  RawMember test helper lacks PartialEq. Fieldwise comparisons corrected the test.
- The corrected negative control failed on the original production name guard,
  demonstrating the intended missing-destination-name refusal.
- Baseline corpus probing found no successful copies. The name-only candidate
  changed four of 189 outcomes to publication-time noncanonical relationship
  refusal, with the other 185 outcomes identical.
- The first integrated focused compile found a borrowed namespace comparison and
  an incorrect eager OpcPackage test accessor. Both were corrected after the
  command handle reached terminal status.

An independent external-harness review withdrew its initial claim that InputBytes
and OutputBytes should return to zero: those resources are consumed cumulative
counters. Only Memory, Objects and Depth are live gauges. The reviewer found no
remaining API, lifetime, timer-scope or preservation-oracle defect.

- `checks/integrated-focused-r3.json` is a retained failed receipt with 10
  passing and 1 failing test and an unchanged source manifest. The failure was
  in the new default-namespace fixture: its absolute
  `https://before.invalid/` target was declared as an internal relationship by
  omitting `TargetMode="External"`. The OPC opener correctly returned
  `InvalidPackUri` before the append path ran. The fixture was corrected; no
  production parser change was needed.
- `checks/integrated-focused-r4.json` then passed all 11 focused unit tests but
  exposed an obsolete integration assertion that rejected every noncanonical
  relationships member. Append-only lexical preservation is now intentionally
  admitted, so `source_backed_topology.rs` was updated to assert exact original
  bytes plus the generated child before the root close. A separate internal
  replacement test retains the typed `SourceBackedOverlayUnavailable` refusal
  and zero-output guarantee.
- `checks/integrated-focused-r6.json` is the corrected focused integration
  receipt: 2 tests passed, 0 failed, exit status 0, and the source manifest was
  unchanged. The intermediate r5 compile receipt records only an incorrect
  `Result<PackURI>` assertion in the test and is superseded by r6.

The next draft source epochs completed the focused implementation checks before
the final source freeze:

- `checks/draft-provenance-strict-r1.json` through `r4` retain the successive
  compiler and strict-lint corrections in `xml_splice.rs`, `source_backed.rs`,
  and `source_cross_copy.rs`: private source access and missing XML-version
  imports, borrowed attribute-key lifetimes, strict qualification and checked
  arithmetic, namespace-attribute ownership, the large topology payload enum,
  and the PPTX classifier error type. `r5` passed strict Clippy with an
  unchanged source manifest.
- `checks/draft-harness-strict-r1.json` records the standalone harness API
  failures: the SHA-256 digest was formatted with an unsupported `LowerHex`
  implementation and `PreservationIndex::new` received a `ZipSliceArchive`
  where the current API requires a `ZipArchive`. `r2` passed after those
  archive and digest API corrections.
- `checks/draft-xml-unit-r2.json` records 5 passing XML splice unit tests with
  0 failures. `checks/draft-xml-integration-r2.json` records 6 passing public
  source-XML integration tests with 0 failures, covering source formatting,
  provenance, stale and overlapping ranges, destination limits, cancellation,
  and managed gauge release.

The first final-gate epoch is retained as historical evidence. The initial
`checks/final-strict.json` and retry `final-strict-r2.json` stopped on Clippy's
`manual_bits` diagnostic for the metadata reservation calculation. After that
memory-accounting correction, `checks/final-strict-r3.json` passed. The next
`checks/final-strict-r4.json` stopped on a moved `SourceBackedPackage` in a
managed relationship-limit test; `checks/final-strict-r5.json` passed after
that test-lifetime correction. `checks/final-harness-strict.json` and
`checks/final-harness-strict-r2.json` passed with unchanged source manifests.
`checks/final-pptx.json` then passed 854 tests with 0 failures and 2 ignored.
`checks/final-opc.json` completed 347 tests but had 2 failures in the prefixed
noncanonical relationship append cases: the closing QName comparison treated
the prefixed closing element as outside the OPC namespace, and the related
limit assertion consequently observed the wrong error. Root corrected that
parser comparison after this receipt.

The subsequent post-fix receipts are separate from that historical epoch:
`checks/final-pptx-r2.json` passed 854 tests with 0 failures and 2 ignored, and
`checks/final-opc-r2.json` passed 496 tests with 0 failures and 1 ignored. The
standalone `checks/final-harness.json` run reached the test binary and recorded
many successful cases, but ended with exit `-2` while interrupted; it has 0
parsed pass/fail counts and an unchanged source manifest. The run stopped at
that harness result, so its downstream document, workspace, format, and
boundary receipts remain pending.

The next frozen pre-format source epoch is recorded by manifest
`a00209c4c550da09e7cc2862328dd1afb10238222fe226f999ed6983a372212f`:

- `checks/final-strict-r6.json` and `checks/final-harness-strict-r3.json`
  passed with unchanged sources.
- `checks/final-opc-r3.json` passed 497 tests with 0 failures and 1 ignored;
  the additional passing case is the managed source-fragment capacity
  regression. `checks/final-pptx-r3.json` passed 854 tests with 0 failures and
  2 ignored.
- `checks/final-harness-r2.json` passed. Its primary `litchi-perf-baseline`
  suite reported 343 passed, 0 failed, and 1 ignored (344 cases); the receipt's
  aggregate across binaries is 381 passed, 0 failed, and 1 ignored.
- `checks/final-doc.json` and `checks/final-workspace.json` passed with
  unchanged sources.

`checks/final-format.json` is the only failure in that pre-format sequence. It
reported rustfmt diffs in four of the ten source-file inputs
(`source_backed_cross_copy.rs`, `litchi-opc/src/lib.rs`,
`source_backed_topology.rs`, and `pptx_native_copy_probe.rs`); no compiler or
test failure was involved. Root formatted all ten inputs, and the direct
format check now passes. The updated final ordering starts with the new
`final-format-r2` receipt, followed by strict-r7, harness-strict-r4,
OPC-r4, PPTX-r4, harness-r3, doc-r2, workspace-r2, and boundaries.

The earlier `checks/final-candidate-build.json` entry recorded the candidate
build while the release-gate receipts were still pending. It is superseded by
the completed `final-candidate-build-r2` and native-inventory receipts below.

The final release-gate epoch is now complete under manifest
`654d4be5da2408ce3ac62ce665322ae11ac1651f1d312327bfb2200bfa7d0dbf`:

- `checks/final-candidate-build-r2.json`, `final-format-r2.json`,
  `final-strict-r7.json`, and `final-harness-strict-r4.json` passed.
- `checks/final-opc-r4.json` passed 497 tests with 0 failures and 1 ignored;
  `final-pptx-r4.json` passed 854 tests with 0 failures and 2 ignored.
- `checks/final-harness-r3.json` passed with 381 aggregate tests, 0 failures,
  and 1 ignored (the primary perf-baseline suite accounts for 343 passed and 1
  ignored); `final-doc-r2.json`, `final-workspace-r2.json`, and
  `final-boundaries.json` also passed.
- `checks/oracle-boundary-regressions.json` passed one valid insertion and
  rejected all three preservation mutations. `checks/final-native-inventory-r2.json`
  passed with four published probe outputs and 185 unchanged outcomes across
  the final native corpus. External pilots 0 and 1 passed after three earlier
  oracle-failure epochs were archived as historical evidence.

The fuzz execution gate subsequently completed through an isolated standalone
target amendment:

- The original `checks/fuzz-build.json` failed with two Rust borrow/move
  diagnostics (`E0505` and `E0382`) in the standalone `parse_opc` harness. Its
  production source manifest was unchanged.
- `fuzz-source-amendment.json` records the only follow-up source change:
  `crates/litchi-opc/fuzz/fuzz_targets/parse_opc.rs`. It bridges the borrow
  lifetime in that standalone harness only, from the release manifest
  `654d4be5da2408ce3ac62ce665322ae11ac1651f1d312327bfb2200bfa7d0dbf` to
  amended manifest
  `c03d99b31629b3b16d46202a46836c9cdaa8b30799cf133e3e181d60f64885b5`.
  The measured production source epoch remains the former manifest.
- `checks/fuzz-lock-r2.json`, `fuzz-build-r2.json`, `fuzz-smoke.json`,
  `final-fuzz-strict.json`, and `final-fuzz-format.json` passed under the
  amended standalone target. The smoke run used the ASAN-instrumented binary,
  seed 454, and 1,000 runs. `checks/fuzz-artifacts.json` also passed. These
  receipts validate the amended fuzz harness and do not create a new
  production-source acceptance epoch.

The formal capture passed all 18 retained lanes: 16 provider lanes plus 2
external lanes, each with 30 retained samples, for 540 samples. The release
source manifest for those measurements remains
`654d4be5da2408ce3ac62ce665322ae11ac1651f1d312327bfb2200bfa7d0dbf`; no
timing or raw report inputs were changed by the later custody correction.

The capture coordinator's Markdown-renderer correction temporarily rewrote the
`protocol_sha256` field in all 18 formal receipts. The exact reconstruction is
recorded by
`custody-corrections/renderer-r1/restoration.json`: it restores the captured
protocol hash and each receipt, while retaining every intermediate hash and
disclosing that this is reconstructed custody rather than uninterrupted
receipt custody. `derivation-amendment.json` retains the original `derive.py`
hash `316841a5e46e248c486bb2f7b4c16ae17eabc051539bdd2a67332ea2f4da4c3f` and
records the separate corrected `derive-final.py` hash
`2be514cfb04c4e21d80881a605040943da608347555611c70e9b213ec1999797`.
That derivation amendment changes only external counter-key display lookup;
sampling, statistics, timing inputs, and raw timing reports remain unchanged.

The final verifier and cleanup sequence are still pending; the receipts above
must not be read as a completed precleanup or sealed-custody verdict.

### Post-cleanup verifier corrections

The first replay failed on a temporary output path resolved against the process
cwd. Subsequent attempts exposed the precleanup attestation being selected as a
cleanup-action receipt, a stale verifier-amendment hash, and a shadowed cleanup
receipt path. The retained `checks/after-cleanup-r1.json` and
`checks/after-cleanup-r2.log` through `after-cleanup-r4.log` record these failures.
The outer verifier now uses an absolute temporary replay path, excludes the
precleanup attestation from cleanup actions, and retains a dedicated raw-output
cleanup receipt path. `verifier-adapter-amendment.json` binds the final verifier
to the precleanup verifier hash and describes all changed components. Two
intermediate verifier snapshots are retained; the original precleanup source
was not retained separately, and its recorded hash is not represented as a
recovered source snapshot. The frozen external oracle and capture receipts were
not edited by these fixes. Post-cleanup replay validates retained attestations;
it does not reparse the deleted PPTX outputs.

The complete post-cleanup replay then passed (`checks/after-cleanup-r5.log`,
`checks/postcleanup-verification.json`). A first separate-copy run correctly
refused the newly added worktree source attestation because its `final-` filename
reserved it for command-gate receipts. That attestation was renamed to
`checks/worktree-source-attestation.json`; its contents and verifier logic were
unchanged. `checks/portable-verification-r1.json` preserves the failed run.

The corrected separate-copy replay passes (`checks/portable-verification.json`),
with all 14 required gates and 18 formal lanes verified. Its temporary directory
is absent. The final worktree source attestation checks all 6,672 recorded files.
