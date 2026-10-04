# 0904 — exact XLSX attribute census

The selected XML artifacts for the protected one-percent XLSX workloads are
overwhelmingly zero/one-attribute tags: over 99.9% have at most one attribute
and no `xml:space`. The one-edit control has approximately 99.87–99.89%.
This independently qualifies the coverage of the bounded summary hypothesis
from 0903. It does not prove semantic equivalence of an implementation or a
native speedup. Production parsing and benchmark timers remain unchanged.

## Reusable fixture exporter

The standalone benchmark now offers `xlsx_attribute_fixtures`, an untimed
example using its existing deterministic corpus builder, updates, source-backed
editor, commit and publication APIs. It exports medium/dense-sparse one-percent
cases and optionally the medium one-cell control. The manifest binds source
and output ZIPs, selected original/replacement XML copies, Rust quick-xml
attribute counts, and the source-replacement auditor's proof kind and ranges.
The exporter verifies existing cell-value/package identity oracles. It does not
add a dependency or production API outside the standalone benchmark tool.

All three source and output archive hashes exactly match the retained 0902
baseline reports. The independent verifier confirms that the selected member
set equals the changed logical ZIP member set: workbook plus four worksheets
for each primary, workbook plus sheet 1 for the control. Workbook recalculation
invalidation changes its XML. This is not a new compressed-member preservation
measurement; the fixed artifact identities and existing oracles remain the
basis for the corpus correspondence.

## Full-artifact structural census

Namespace processing is disabled in the independent Python Expat parser so
namespace declarations remain counted as raw attributes. Expat counts match
all seven Rust quick-xml census fields for every XML artifact. The census
rejects DTD/entity declarations and verifies ZIP/member/sidecar hashes.

| Case and side | Tags | Zero attributes | One attribute | Two or more | Eligible zero/one without xml:space |
| --- | ---: | ---: | ---: | ---: | ---: |
| Medium 1%, original | 18,642 | 9,221 | 9,416 | 5 | 18,637 |
| Medium 1%, replacement | 18,643 | 9,221 | 9,416 | 6 | 18,637 |
| Dense-sparse 1%, original | 35,906 | 17,797 | 18,104 | 5 | 35,901 |
| Dense-sparse 1%, replacement | 35,907 | 17,797 | 18,104 | 6 | 35,901 |
| Medium one-edit, original | 4,665 | 2,306 | 2,354 | 5 | 4,660 |
| Medium one-edit, replacement | 4,666 | 2,306 | 2,354 | 6 | 4,660 |

No selected tag has `xml:space` or more than 32 attributes. These are counts of
full artifacts, not native timing shares, dynamic call counts or general
Office-producer coverage. Each case includes its own original XML; totals
across cases therefore include repeated source data.

All five medium one-percent replacements use `Complete` proofs. Dense-sparse
has four `Complete` proofs and a sheet-4 `Window` with original/replacement
ranges `[1494, 1508)`. The control has a complete workbook audit and sheet-1
window `[190, 198)`. Full replacement census counts must not be read as the
number of tags re-audited in either window case. Python records those ranges
and hashes but does not independently replay the Rust window proof. The raw
Callgrind call-count discrepancy from 0903 remains unexplained; no causal
claim follows from this census.

## Candidate boundary and next action

The source-only [semantic review](results/change-0904/semantic-review.md)
supports testing a summary returned after the entire existing `check_start`
validation succeeds. Zero attributes must leave state untouched; one raw
non-`xml:space` attribute must charge the aggregate attribute budget at the
existing tag offset after the full scan. Per-element limits, namespace checks,
error precedence, streaming and window accounting remain mandatory. Multiple
attributes and all `xml:space` cases retain the existing iterator.

The next batch should implement and differentially qualify that bounded
candidate against unchanged slice/stream/replacement behavior, including
malformed tails, all relevant limits and the 32/33-name handoff. Only then
should it face fresh paired native primary/control captures. ABI/code-layout
costs can offset avoided traversal; the 0902 benefit gate remains unchanged.
No speedup is adopted here.

## Validation, failures and cleanup

The [packet](results/change-0904/README.md) retains the frozen plan, exact
artifacts, build/quality receipts, census, independent review and replay.
Two builds and exports produced the same 30 artifact files. Final warning-denied
Clippy, formatting and rustdoc checks pass, as does the crate-boundary audit.
Seven malformed-artifact mutation checks pass. Canonical offline replay is
byte-identical, and both exporter and reducer refuse existing destinations
without changing their contents. No new performance capture or broad Office
correctness-suite result is claimed for this harness-only addition.

The first census invocation used an unsupported CLI flag and exited before
analysis; the corrected invocation used the same helper/artifacts. Initial
Clippy rejected two needless borrows; the original exporter source and log are
retained, and the corrected exporter reproduced every artifact byte. A source
inventory expectation initially omitted the harness README; the corrected
check finds only that README and lib shim changed among 10,771 prior files,
with two new exporter files recorded separately. Production source is unchanged.

The duplicate export/replay outputs and the owned build directory were removed.
No worktrees were created; unrelated worktrees and the two unrelated untracked
files were preserved. OLE2/OOXML work remains active, ODF is deferred, and iWork
is excluded from this goal.
