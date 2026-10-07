# 0947 — scattered cell edits justify testing multiple replacement windows

The medium and dense-sparse one-percent workloads change **93 / 178 cell
records**, occupying **2,665 / 5,021 replacement bytes**. Every corresponding
worksheet byte gap outside those cells is identical to the original. Yet the
current auditor completely scans seven of the eight changed worksheets on the
replacement side. This supports implementing and qualifying a bounded
multi-window candidate next, while retaining the complete original audit.
It does not establish a safe multi-window proof or a performance improvement.

## Evidence and current behavior

This follows [0946](0946-xml-reader-attribution.md), which located substantial
XML token-reading work but could not separate original and replacement work.
The fixture pairs already exported by
[0904](0904-xlsx-attribute-census.md) are reused after verifying every ZIP and
XML digest, each XML sidecar against its ZIP member, and the exact changed-member
set. Their complete source/output ZIP hashes match the adopted 0945 candidate's
medium one-percent, dense-sparse one-percent, and medium one-edit reports.
Ten focused current source files, including the auditor, OPC publication,
XLSX writer, snapshot, exporter, and standalone harness lockfile, match that
candidate. This is focused source and artifact correspondence, not a fresh
end-to-end benchmark or full source-tree requalification.

A small standalone Rust probe calls the current public `verify_source` and
`verify_source_replacement` APIs with default limits on all twelve pairs.
Both sides pass; the ten `Complete` and two `Window` outcomes agree with 0904.
The new evidence is changed-span locality and coverage, rather than another
claim that complete auditing occurs. Full audit reports retain attributes,
events, maximum depth, and text bytes. Those full-artifact counters do not
measure work on the window path. The debug probe also performs extra complete
audits; its execution is not a release timing or scanner-work measurement.

| Case / worksheet | Changed cells | Original cell bytes | Replacement cell bytes | Current proof |
| --- | ---: | ---: | ---: | --- |
| Medium 1%, sheet 1 | 24 | 642 | 642 | Complete |
| Medium 1%, sheet 2 | 23 | 677 | 677 | Complete |
| Medium 1%, sheet 3 | 23 | 670 | 670 | Complete |
| Medium 1%, sheet 4 | 23 | 676 | 676 | Complete |
| Dense-sparse 1%, sheet 1 | 164 | 4,600 | 4,601 | Complete |
| Dense-sparse 1%, sheet 2 | 11 | 330 | 330 | Complete |
| Dense-sparse 1%, sheet 3 | 2 | 60 | 60 | Complete |
| Dense-sparse 1%, sheet 4 | 1 | 30 | 30 | Window, 14 bytes |
| Medium one-edit, sheet 1 | 1 | 22 | 22 | Window, 8 bytes |

An independent namespace-aware Expat analysis matches cells by worksheet,
row, and address, rejects duplicate keys, and records complete balanced cell
spans on both sides. It compares every corresponding outside gap separately,
including lengths and hashes. All nine worksheet pairs pass. The generated
fixtures contain no MCE content; they do not qualify general producer markup.
The workbook has a separate 103-byte `calcPr` insertion and no cell spans, so
it remains a complete replacement audit in the model below.

Six worksheets have a common-prefix/suffix difference enclosure larger than
half the replacement. The current auditor therefore stops searching for a
single window before the original scan. This is a direct source predicate
checked against the retained bytes, not inferred from aggregate Callgrind
calls. Small individual changes spread across a document can trigger that
predicate even when their total byte coverage is small.
The remaining dense-sparse sheet 3 has differences at offsets 2,755 and
5,839: its difference enclosure is only 3,085 bytes, but the enclosing
`sheetData` element spans `[159, 8103)`, or 7,944 of 8,115 bytes. The later
half-payload check rejects that candidate. The workbook insertion has no
covering non-root element. These are source-derived fallback explanations;
the public probe observes `Complete`, not an internal fallback-reason event.

## Conservative coverage model

The model substitutes complete changed-cell spans only for worksheet pairs
currently using `Complete`. It keeps the complete workbook replacement audit
and both existing smaller `Window` proofs. It always keeps the original audit.

| Case | Full original bytes retained | Current selected replacement bytes | Hypothetical selected replacement bytes |
| --- | ---: | ---: | ---: |
| Medium 1% | 269,182 | 269,285 | 3,192 |
| Dense-sparse 1% | 508,857 | 502,623 | 5,532 |
| Medium one-edit | 63,718 | 535 | 535 |

The primary cases would omit 266,093 / 497,091 replacement bytes from the
selected audit ranges if every proposed cell window could be proved. These
are coverage counts, not measured byte traffic, parser instructions, native
latency, or a speedup bound. Comparing unchanged gaps, tracking source state,
reserving range metadata, cloning namespace context, and starting many small
parsers all have costs. The one-edit control offers no additional coverage
saving and must retain its current behavior.

## Candidate boundary

The next candidate should carry exact source/output cell-span pairs from the
value-only writer into an operation-local replacement audit. Writer hints
must be treated as untrusted range metadata: verify order, disjointness,
bounds, and exact outside-gap identity. Use the mandatory complete original
audit to establish balanced source boundaries and inherited namespace/parser
state before proving replacement windows. Re-total all removed and added
aggregate counters over
the whole replacement, including event, attribute, text, depth, and byte limits
as applicable. Keep limits, cancellation, freshness, and typed refusal ordering.

Any ambiguous mapping, unsupported shape, reservation failure, or failed
window proof must retain the existing canonical complete-audit behavior.
Replacement failures must still report the full auditor's exact error and
offset. No trust in planning-parser success, semantic readback, disabled parser
checks, or cross-operation parsed-XML cache is introduced. The 0902 rejection
remains in force. ADRs 0003, 0005, 0006, 0011, 0030–0032 continue to constrain
publication and ownership; no ADR change is requested.

The diagnostic reconstructs spans after publication. It does **not** measure
writer metadata allocation or prove that current writer metadata already
contains these pairs. Production plumbing, differential/adversarial tests,
and a fresh prospective native/allocation comparison are still required.

## Validation and retention

The [evidence packet](results/change-0947/README.md) contains input identities,
probe source and results, independent span analysis, and the reproducible
coverage model. Production source is unchanged. This is exploratory diagnostic
work, not a prospectively registered optimization experiment. The broader
Office performance goal remains active; iWork is excluded.

The isolated probe passed formatting, a locked offline build, and warning-denied
Clippy. Independent span analysis and five rejection controls passed. A rebuilt
probe at the fully verified archive extraction reproduced the proof output,
analyzer output, and coverage model exactly. Owned build/replay scratch was
removed; binaries and large fixtures remain external. The minimal source export
omits a test-only OPC module, so its broader formatting attempt failed; no
workspace-wide check or Rust test pass is claimed. Intermediate exploratory
analyzer sources were not all retained; final v2 is the reproducible result.
