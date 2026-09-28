# 0806 retained setup and reader repairs

Production source remained at `e3ff267ee3` throughout baseline preparation.
The cross-format baseline build and eight-row qualification succeeded; all
eight corpus identities match their sealed 0794 references.

Three main-probe setup versions were retained before accepting qualification:

- Setup 0 built all three binaries. Formatting passed; the all-features test
  target passed 28 tests and failed seven. Its process-global allocator counted
  real test-harness allocations in a synthetic-counter test; the first failure
  poisoned a shared lock and caused six subsequent failures. Test-only allocator
  registration isolation fixed this without changing runtime instrumentation.
- Setup 1 built successfully and passed formatting and all 35 tests. Clippy
  rejected unused copied support APIs, documentation continuation and test-module
  placement. Narrow non-test dead-code allowances retain those APIs and their
  layout; documentation and test placement were repaired.
- Setup 2 built successfully and passed all three probe quality gates, including
  35 tests. All eighteen qualification children exited successfully and passed
  their preservation checks. Independent review and root replay nevertheless
  rejected the three new rows because Serde emitted `valid4-attr` while the plan
  and CLI require `valid-4attr`. An explicit variant rename and a six-shape
  round-trip test repair that report identity. The test's deserialization and
  equality derives are enabled only in test builds.

`setup-attempt-{0,1,2}.json` names the immutable source, build and quality
archives and the nine retained executable identities. The eighteen unaccepted
setup-2 reports remain in `qualification-setup-2/`; they are separate from the
qualified measurement counts and are never pooled into timing results.
Original receipts retain original absolute paths. The archive metadata maps
them to the retained directory and binary identities.

The repaired baseline passed formatting, all 36 tests and warnings-denied
Clippy, then completed eighteen fresh qualification children. Root's first
replay of those fresh reports found an offline-reader error: its expected
attribute-name list omitted the namespace prefixes, although the frozen
generator, review and reports consistently use `lx1:probeOne` through
`lx4:probeFour`. That repair belongs in the readers; it does not require a
source, fixture, report or binary change. After correcting those constants,
root replay verified all eighteen strict report contracts and independently
matched the original fifteen source, output, full-text and readback identities
to the sealed 0792 qualification. Independent review still precedes oracle
freezing.

No main native, allocation or heaptrack lane ran during these setup repairs.
`probe_amendment_audit.py` checks the narrow source changes against setup 0 and
retains evidence that counter layout, counter logic and fixture logic remain
unchanged. The after-build wrapper keeps the original frozen build driver
unchanged and writes the qualification/application/source witness before
invoking after compilation.

## First production quality attempt

After independent qualification review passed, root froze
`qualification-four-attr.json` and applied the original six-file candidate.
`application.json` retains that exact source census and original candidate
manifest/patch identities. `quality-0/` passed formatting and failed the
offline locked all-features/all-targets check: `litchi-sign` denies the now
unused private `unchecked_attributes` method. The original candidate's
constructor duplicated that method's attribute-iterator initialization.

The repair will reuse the existing inline helper from the constructor rather
than suppress the production lint. Its separate archive and application
witness preserve the original candidate and failed attempt. The amended
source requires fresh protected native preflight against production-before,
followed by the production and public workflow gates. No after build or main
timed capture preceded this failure.

The separate `candidate-quality-amendment/` archive contains the mechanical
constructor repair for all five helper copies, leaving shared tests unchanged.
Independent source review approved it for protected preflight. The supplemental
mirror quality run passed 70 baseline tests, 100 amended tests and both
warnings-denied Clippy checks. Both native builds completed successfully.
They intentionally produce the same executable identity: the probe contains
both implementations, and each child selects its measured leg explicitly.
The original production application and failed quality attempt remain intact;
the amendment has not yet been applied to production.

The supplemental native lane then completed all 936 children and 28,080
samples. Its first primary replay rejected an offline-reader expectation
that confused the Cargo binary name with the copied executable basename;
the frozen probe correctly reports `before-native` or `after-native`.
After correcting that expectation, primary replay passed and found no
protected consume regressions, with both required one- and two-attribute
benefits. The first independent replay found another reader mismatch:
long inputs use a hex representation in the case source identity, while
the fixture archive retains the decoded UTF-8 input. These reader repairs
do not change source archives, binaries, fixtures or captured results.

The preflight preparer produced the raw audit and initial final decision after
capture. Root's writer refused to overwrite the already present witness;
root then independently reproduced it with `--check` and reran finalization.
The first preparer replay also caught a stale probe-inventory key in the
offline reader before writing a witness. No benchmark child was rerun.
The decision advances the amendment to workflow trials; it does
not adopt production code. The constructor amendment was applied with its own
immutable source witness. The next production attempt, `quality-1/`, passed
formatting but failed the all-features/all-targets check: the original
candidate had accidentally narrowed the OLE helper's public `BytesStartExt`
trait and `CheckedAttributes` type to crate visibility. That breaks existing
`litchi-crypto` imports. The other four helpers retain their baseline visibility.

This failure corrects the earlier source review's visibility assessment.
Restoring those two OLE declarations is a compatibility repair; the OPC
implementation measured by the supplemental probe does not change. A separate
visibility archive and application witness will preserve this source step,
and the full production gates must run again before any after workflow build.

The visibility archive passed independent review and root's exact two-token
comparison. Root applied it through `apply_visibility_amendment.py`, verified
the complete source census and recorded `visibility-amendment-application.json`
before starting `quality-2/`.

`quality-2/` passed all six production gates. Its 500 test suites report
12,885 passed, zero failed and 89 ignored; Clippy, warning-denied rustdoc and
the crate-boundary checker also passed. Root wrote `quality-summary.json`,
froze `after-build-inputs.json` against the visibility-restored application,
then started the serial after-build, after-probe-quality and cross-after-build
sequence. No main timed capture preceded that completed quality gate.

The main after native, allocation and profile builds and all three after-probe
quality gates completed. Pre-capture replay verified both build records, all
18 qualification reports, the frozen extension oracle, production quality and
after-build chronology. It also exposed a stale offline warning expectation:
the repaired allocation builds correctly emit no warnings, whereas the reader
required a nonempty warning list for every build. The reader now accepts the
zero-warning allocation case and checks the exact eleven retained standalone
native/profile warning headlines on both legs. No source or binary changed.

The cross-format after build also completed successfully; root verified its
binary identity and exact visibility-restored source census. Independent
`after-build-review.md` checks the final application chain, six-file scope,
quality receipts, frozen qualification and after-build chronology. Root then
started the frozen main native, cross native, allocation and profile sequence
serially. Offline numerical replay waits until the capture sequence ends.


The frozen capture sequence completed: main 306 reports/6,714 samples,
cross 13/2,888 and profile 4/4 with eight decodes. Primary and independent
analysis found no eligible 3% capture/lifecycle benefit, no main latency or
resource veto, no cross veto, and four qualified profiles. Root rejected the
candidate and restored all six production files exactly through
`restore_candidate.py`; `restored-source.json` verifies the full baseline census.

After verifying all 19 copied binary identities, root ran `cleanup_all.py`.
The owned target was removed in one operation, with separate main, cross,
setup and supplemental cleanup witnesses. Post-cleanup main replay exposed
one stale live-binary requirement in the supplemental preflight reader. That
reader now checks the exact two build descriptors, native completion witness,
cleanup schema and target absence; live hash checks remain before cleanup.
No measurement or decision changed. Root replayed main and independent main
analysis, quality summary, profile qualification, official cross analysis and
independent cross audit, supplemental raw audit and retained failure audit.

Generated Python caches were removed before sealing, including the sole extra
unsealed cache in 0805; historical sealed payloads were unchanged. Final report
corrections distinguish main nearest-rank p50 from cross midpoint p50 and the
three added valid-4attr rows from the eighteen-row full matrix.


Final root `validate.py --check-workspace` passes after three strict reader
corrections: the original candidate has six archive names (the five helpers
plus shared OPC tests), probe-quality inputs embed their full source census
rather than a path descriptor, and independent allocation results contain
both blocks for each of eighteen cases. The validator now requires all 36
(shape, mode, block) identities exactly. These changes align the reader with
the immutable producer schemas and do not modify reports, measurements,
source archives or the rejection. The report's eighteen timing rows were
also compared programmatically with retained analysis; all local links resolve.
