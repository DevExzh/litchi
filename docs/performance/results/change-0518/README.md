# 0518: retained DOCX source snapshot publication

Baseline revision: `afb62ab7a70859dad4a4b9c8eea91402d1ef4052`.
OLE2/OOXML optimization remains active; ODF is deferred until that goal is
complete and iWork is excluded. This batch does not establish completion of
the program or close historical 0499/0500 flags.

## Reproduction

`run.py` builds the unchanged managed DOCX example at the current source;
`--variant candidate` selects the candidate receipt directory. Recreate the
two owned temporary directories from `start.json` before reproduction. Every
capture creates its files exclusively; use a fresh evidence directory for a
new campaign. Toolchain, environment, commands, source manifests, binary hashes,
and artifact hashes are retained in the build/capture receipts.

Run `capture.py preflight`, `r1`, `r2`, `profile-preflight`, `profile-r1`,
`profile-r2`, and `hardware` on the baseline, then `after-r1`, `after-r2`,
`profile-after-r1`, `profile-after-r2`, and `hardware-after` on the candidate.
Run these serially. The frozen plan specifies the 24 native cases and six
profile/counter cases. Native cases each contain two internal repetitions of
three warmups and 30 samples; two fresh campaigns reverse case order.

`allocations.py build baseline` builds the isolated instrumented probe;
`allocations.py capture baseline --repeat 1` and `--repeat 2` capture all 24
cases in separate children. Repeat with `candidate` after the source change.
The same generated probe, standalone lockfile, and canonical allocator wrapper
are bound to both builds. `allocator-probe/generate.py` documents every change
from the unchanged benchmark source.

## Measurement boundaries

Native publication time includes destruction of the returned snapshot.
Callgrind and allocator regions cover only the publication method, before that
caller destruction. Fixture creation, edit/commit, and output verification are
outside those method regions. All instrumented timings are excluded from
native latency claims. Hardware counters and GNU time RSS cover the whole
child, including setup and verification; they are not publication-local peaks.
Allocator region peaks are absolute callback-ordered live bytes, distinct from
RSS; incremental peaks subtract live bytes at region entry.

The allocator probe has its own frozen dependency resolution and Cargo's
default release profile. The native workspace uses LTO and `panic = "abort"`;
the standalone probe does not inherit those settings. Both allocator variants
use the same probe lockfile and profile, so their comparison isolates this
source change within that instrumented build. Its absolute allocation counts
are not measurements of the normal native executable.

Same-API baseline/candidate comparisons are separate from batch-versus-repeated
API-choice ratios. Source I/O, exact output, selected and untouched semantics,
source revisions, forward/inverse preflight, and released memory/object gauges
remain checked. Work charges describe budget accounting, not instructions or
allocator calls. Every phase/RSS regression and repeat-drift flag remains in
the native reports; confidence intervals from internal samples are descriptive
and do not establish host-generalized uncertainty.

## Admission

The source and resource contracts are recorded in `adr-review.md`,
`snapshot-reuse-design.md`, and `opc-hint-review.md`. The implementation is retained for the measured same-API improvement:
48 comparisons reduce lifecycle p50 by 1.42–34.59% and publication p50 by
43.75–67.42%. All 69 short-phase flags remain visible; no matched lifecycle,
publication, or RSS threshold regression is flagged. The all-features OPC,
DOCX, XLSX, PPTX, and XLSB suite plus the independent ZIP64 test passed
5,004 executed tests. All nine quality gates passed.

`fallback-profiles/fallback-proof.md` records two debug test profiles showing
full XML validation for foreign and derived hints. They are guard evidence,
excluded from speedup comparisons. `metadata-correction.json` in that
directory records a receipt-only correction without recapturing raw data.

`replay_reports.py` recomputes five JSON/Markdown report pairs exactly.
`verify.py --require-complete` checks source/build/receipt custody, capture
oracles, method scope, report replay, quality gates, and the strict leaf
inventory in `SHA256SUMS`; `verify_test.py` exercises tampered evidence.
`verification-before-cleanup.json`, `cleanup.json`, and `verification.json`
retain the final audit and removal of the two owned temporary directories.
Historical limitations and the broader OLE2/OOXML goal remain open.
