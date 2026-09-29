# 0841 additive reader review

This is an independent, read-only review of the additive reader packet. The
frozen `readers.py`, original admission inputs, source, binaries, schedule,
native captures, and workload outputs were not changed by this review. No
build or workload command was run here.

## Disposition

The strict report and Callgrind checks are sound on the retained evidence.
There was one concrete command-binding failure in the first additive
qualifier. It is repaired by the recorded v3 qualifier and must be included
in the final additive custody set.

## Command binding finding

All 16 retained numbered Callgrind files use Valgrind's actual header form
`cmd:  /path/to/probe ...` (two spaces after `cmd:`). `readers_v2.py` keeps one
of those syntax spaces in the parsed value. Consequently,
`qualify_v2.py:57` compared a value beginning with one space with the joined
child `argv` beginning at the binary path. The retained
`commands/paired-analysis/receipt.json` records this exact failure with exit
code 1; the 20 qualification reports alone did not exercise this profile
path.

`qualify_v3.py` normalizes only leading header whitespace before comparing the
target command, while `validate_profile()` still requires the numbered and
termination headers to agree and checks the launcher `argv` for all five
Callgrind controls. `analyze_v3.py` then passes all 156 retained reports and
3,636 samples. The 16 current profiles have matching normalized target
headers, and each recorded launcher contains `--tool=callgrind`,
`--collect-atstart=no`, `--toggle-collect=ppt_profile_probe::profile_write`,
`--zero-before=ppt_profile_probe::profile_write`, and
`--dump-after=ppt_profile_probe::profile_write`. The paired v3 receipt and
`analysis-repair.json` report unchanged metrics and zero replacement
workload runs. Final admission must bind `qualify_v3.py`, `analyze_v3.py`,
`analysis.json`, and `analysis-repair.json`; `qualify_v2.py` is not a passing
profile validator for these retained raw files.

The synthetic reader fixture intentionally uses a child-only command header,
which is correct for Callgrind output, and the original reader rejects it.
The fixture uses one syntax space, so the real two-space Valgrind form is
covered by the retained profile receipts and v3 repair rather than by the
21-check receipt (twenty malformed-input checks and one reproduction of the
original reader rejecting a valid child-only header).

## Strict report and artifact evidence

`readers_v2.py` requires the current 0841 schema, exact `profile_owner`,
sealed 0782 source/generator identity, and sealed 0782 output identity for
each report. Artifact-bearing reports additionally require the report's
artifact identity to equal the bytes on disk and pass the hash-bound 0840
independent CFB parser. The retained 156 reports have no schema/owner or
artifact-byte mismatches in a read-only inventory; the 20 qualification
artifacts are all independently byte-identified and structurally parsed.

`reader-tests-v2.json` is a passing 21-check receipt, but its `reader` field
accidentally inherited the original reader label. The test implementation
actually imports `readers_v2.py`. The additive `reader_tests_final.py` verifies
that imported module, replays all twenty malformed-input checks and the one
valid-header rejection reproduction, and binds the reader, test implementation
and runner hashes in `reader-tests-final.json`. The original mislabeled receipt
remains intact. `qualification-v2.json` and `analysis.json` also bind the
corrected `readers_v2.py` identity
(`6302b6656dc43e4f490b14e161846f7206bdd4ae20aaf757a44dcfd936a0c672`).

## Callgrind edge review

The corrected parser rejects a pending `cfn` when a new `fn` starts and again
at EOF. The four retained fixtures are `orphan-eof`, `orphan-next-fn`,
`cfn-eof`, and `cfn-next-fn`; all are listed as rejected in the 21-check
receipt. A read-only scan of all six retained 0784 numbered profiles found no
incomplete edge record. The current 16 profiles resolve compressed function
names, have one positive owner call, and satisfy both conservation equations:
owner self cost plus immediate-child inclusive cost equals the owner summary,
and all function self costs equal the summary. No live-profile conservation
failure was found.

## Review conclusion

Use the v3 command-normalization repair and its custody receipts for the
final additive admission. With that binding, the reader evidence is suitable
for the diagnostic-only 0841 result; it does not authorize production
adoption or alter the frozen measurements.
