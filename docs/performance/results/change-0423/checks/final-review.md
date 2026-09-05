# Final evidence review

This was a read-only review of the 0423 protocol, verifier, report guards,
summary and portable replay. No Cargo command, test, profiler or workload was
run. The formal capture has now completed with `status: pass` and 16/16 rows,
and `checks/summary-final.log.gz` records a passing 16-run summary after the
corrections below. The post-cleanup portable replay and copied-tool binding
receipts also pass.

The evidence contract is appropriately conservative. The fixed corpus and
selector identities are checked per report; journals bind the report,
catalog, protocol, build, binary, source revision, verifier output and raw
`time -v` artifacts; the summary revalidates all 16 runs and the O/S/S/O
order. The report verifier always emits `claim_authorized: false`, and the
protocol and documentation withhold cross-role speedup, memory reduction,
physical-I/O, zero-copy and post-drop-retention claims. Portable replay uses
the four hash-pinned shared validators and then replays every retained report.

The two earlier publication issues are resolved. The historical
`checks/script-syntax.json` is now superseded by passing
`checks/script-syntax-final.json`, whose hashes bind the final scripts. The
validation index explicitly identifies `validation-ready.json` and
`validation-lint-final.json` as authoritative and labels
`validation-final.json` and `validation-corrected.json` as superseded failed
attempts. The replay-tools manifest also matches its four pinned modules.

The portable mutation guard intentionally runs the R1 report in each of the
eight lane/corpus/role combinations. The summary's full 16-run replay remains
the all-row integrity check; publication should describe mutation coverage as
R1-only rather than as eight guards per repeat. SHA-256 custody here proves
internal consistency of the exported bundle, not an external signature or
independent attestation. The protocol's `baseline_revision` is a contextual
ancestor; the measured source identity remains the 0423 implementation
revision.

The terminal capture, corrected summary and post-cleanup portable replay now
pass. I found no evidence-script logic that authorizes an unsupported
scientific claim and no production correctness blocker.

## Summary correction addendum

The first summary attempt exposed a transient-state bug: `validate_report`
read `_journal_sha256` before `validate_journal` had computed it. The fix leaves
journal hashing at the journal-validation boundary, after the report proof is
complete, and the final row still carries the computed journal hash. A second
attempt exposed a field-name mismatch in table rendering (`count` versus the
shared statistics result's `sample_count`); the renderer now uses the latter.

The retained `checks/summary.log.gz` and
`checks/summary-corrected.log.gz` are failed historical attempts.
`checks/summary-final.log.gz` records the passing 16-run summary, and the
generated table contains 100 samples for normal rows and 30 for allocator
rows. The media-rich owned normal repeat is outside the stated drift ceilings
and is marked `False` in the table; its statistics remain descriptive and are
not promoted to an acceptance claim. These corrections introduce no new
evidence or scientific claim. The post-cleanup portable replay and copied-tool
binding checks both have passing terminal receipts.
