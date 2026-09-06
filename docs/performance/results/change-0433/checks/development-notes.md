# Development receipts

The initial release harness check (`before-harness-focused`) compiled against
production revision `1518a09017c085329a6d5b5c2181f63d012c5d50` plus the new
baseline harness. It ran three tests: two passed and the scalar corpus acceptance
test failed with `buffered ODS scalar text cell differs from specification`.
Source custody remained unchanged throughout the command.

The corpus deliberately includes escaped XML characters in its text field.
Source inspection found that both standalone and fused ODS worksheet parsers
ignore `Event::GeneralRef`. This makes a semantic creation benchmark invalid:
a package can be authored successfully yet reopen with different cell text.
The corpus is retained. A narrow reference-decoding correctness repair is a
prerequisite to freezing the performance baseline; it is not presented as a
performance optimization or a measured speedup. The initial failure is retained
as evidence, and the formal before revision will identify the repair explicitly.

The original proposed large corpus (131,072 physical rows) remains covered by
an expected refusal test against the existing XML publication safety limit.
The matched performance shapes are 64, 8,192 and 32,768 rows.

Independent source review also found a harness metadata mismatch before capture:
`uncompressed_payload_bytes` described authored XML while `sink.input_bytes`
described the canonical scalar projection. Both logical-input fields will use
the canonical projection byte count; target/authored XML bytes remain separate.
No formal measurement was captured with that mismatch.

The first complete ODS crate run passed 301 tests before stopping at an existing
sequential-text assertion that explicitly expected `&amp;` to disappear. The
expectation was corrected to retain the ampersand. The complete rerun passed
433 tests. The new independent reference integration checks also pass, including
CR versus literal-CR normalization, hyperlink ranges, single decoding, source
byte preservation, invalid references, and strict numeric-reference grammar.
The buffered release harness rerun passed all three focused tests.

The full ODS all-target/all-feature warning-denied Clippy command stops at the
preexisting `large_enum_variant` finding for `ArchiveReaderKind` in
`crates/litchi-odf-common/src/package/model.rs`. That reader is unchanged from
the initial baseline. A separate diagnostic command exempting only that lint
category passes; it is not described as a passing full strict gate.

The provider committed the initial reference repair as `d1b18749d` before the
root gates ran. Root retained that history, tightened numeric syntax (Rust's
integer parser alone would accept a leading plus), and ran the recorded checks
before committing the follow-up and integration coverage. No performance
measurement is attributed to the unvalidated intermediate revision.

The first owned-file formatting check reported two newly added match arms in
`worksheet/codec.rs`. Pinned rustfmt corrected their layout, and the second
check passed before the baseline commit and build. This was a formatting-only
follow-up; it did not alter the reference-decoding behavior.

The frozen baseline build at `be82ef53dc191a64bc3c3365502bf44599528577`
passed with frame pointers and release debug symbols. Separate normal-tiny,
allocator-tiny and normal-large pilot processes each completed 30 samples and
three warmups. Pilots are retained as protocol/verifier development evidence,
not included in the formal repeated comparison matrix.

The formal before role completed all twelve process lanes with unchanged
source custody and passing report verification. Each workload's untimed Rust
oracle checks the exact three-member package, MIME and manifest entries, and
reopens the complete scalar sheet. The retained semantic catalog reports
member listing as unavailable, and the bundle does not retain generated archive
bytes. Portable Python verification therefore binds the mandatory Rust oracle
flags and corpus/output hashes to recorded source and binary identities; it is
not an independent archive parser or a standalone archive-byte readback.

The first generated-XML release compilation stopped before executing tests: an
ignored successful XML audit report violated `unused_must_use`, and a test
sink lacked `Debug` required by `unwrap_err`. Root explicitly bound the audit
report and derived `Debug` for the sink. The second compilation exposed the
corresponding missing derive on its control field; that was corrected too.
Both failed receipts/logs retain unchanged source custody and zero test counts.

The ODS integration compile exposed an internal lifetime tie between cell text
and the execution context. Root separated context ownership from the per-call
row/cell bounds. The next compile caught stale test helper calls (borrowed Copy
limits and Results passed where an extracted error was required); these were
corrected without changing the refusal assertions.

The first executable ODS integration run passed eight of eleven tests. Three
assertions incorrectly required compressed output after an input/resource
refusal to equal a byte prefix of a successful full-document archive. The
Deflate encoder can finish its partial input during unwinding, so its compressed
block bytes may differ from those of the longer successful input. The intended
failure contract is accurate accepted sink progress and an incomplete archive
that cannot be finalized. Tests for semantic/input/audit refusal use that
contract; successful exact-limit identity and sink/output refusal byte checks
remain separate.

The same integration sequence first exposed a valid-small-output preflight
configuration error: the fixed metadata ceiling exceeded the requested output
ceiling (`max_metadata_bytes exceeds max_output_bytes`). The candidate ODS
limit wiring now caps that metadata reservation to the output ceiling, so the
exact-output case can reach the intended publication assertion. This is a
correctness/limit repair, not a performance result, and the candidate source
remains outside the `be82` commit provenance.

The initial warning-denied production Clippy run also found two candidate-side
issues: two `io_other_error` constructions in the generated-XML reader and a
`field_reassign_with_default` test setup. They were corrected; the final
scoped production check passes with only the unchanged
`ArchiveReaderKind` `large_enum_variant` category exempted. The unscoped final
receipt therefore still records that pre-existing strict debt.

The first complete performance-harness run retained 308 passing tests, one
ignored test, and one failure because `selectable_case_count_matches_current_enumeration`
expected 430 while enumeration returned 431. The selectable count was repaired
and the retry records 314 passing tests, no failures, and one ignored test.
That broad candidate-harness result is separate from the later 12-test
post-lint ODS-focused subset; the latter is not a replacement for the full
suite. A harness formatting command also spilled into
`filesystem.rs`, `pptx_slide_boundaries.rs`, and `xls_numeric.rs`; the recorded
spill receipt shows all three files restored to their exact pre-format hashes.

The bounded scalar streaming implementation was subsequently committed as
`f5bf1696192f0007db56554aa5b16719cdf1950b`. The first warning-denied harness
Clippy retry found one candidate helper using `ordinal % 2 == 0`; it was changed
to `ordinal.is_multiple_of(2)`. The remaining 29 findings were checked against
the retained 0429 baseline: the strict-debt comparison reports 17 unique
pre-existing groups and zero changed harness findings. The post-lint ODS
focused harness run passed 12 tests, and the post-lint harness documentation
and owned-file format checks passed as well. These are correctness and debt
receipts, not performance results.

The release harness build for the committed revision was still running at the
latest retained `after-build` receipt. No timing, allocation, RSS, or speedup
outcome is attributed until that build and the independent performance capture
complete.

The production strict-review derivation initially collided with its command
receipt filename. It refused to overwrite the running receipt; root retained
that failed receipt unchanged under `production-strict-review-driver-first.json`
with an explicit relocation record, then used a distinct driver tag. The
subsequent derivation passed.

All six profile workloads passed. The first portable profile report reader
incorrectly treated the frozen verifier's `VALID\n` stdout marker as JSON
because its retained filename ends in `.json`. The reader now requires the
exact marker bytes and still binds all artifacts and reruns report validation;
`profile-report-v2` passed. Captures and frozen drivers were not changed.

A development verification invocation through `check.py` encountered its own
still-running receipt and refused the incomplete record. The failure is
retained in `portable-development-first`; terminal portable replay uses the
external-log `replay.py` driver so inventory verification observes no running
receipt and writes its result only after the verifier exits.

The first sealed replay exposed the same profile marker in the whole-bundle
expected-check scan. That scan now skips only the six known profile marker
paths with exact `VALID\n` bytes. The second replay copied the bundle outside
the checkout and exposed an eager repository-parent lookup in the strict-debt
reader. Repository discovery is now confined to generation-only operations;
portable check mode uses retained artifacts exclusively. Both failed receipts
remain sealed. `portable-development-sealed-v3` passed copied-bundle replay
and all mutation checks with the corrected readers.

Final staging initially rejected whitespace in the retained historical
formatter patch. The patch is now stored with deterministic gzip and verified
against its original decompressed SHA-256. An initial compression helper used
a relative path where an absolute path was required; the original bytes were
preserved and the compression mapping was completed after a round-trip check.
The failed staged-check receipt/log and its exact original receipt bytes remain
retained with an explicit relocation record. The canonical staged check and
`final-portable-v3` pass; the two earlier final replay failures remain retained.
