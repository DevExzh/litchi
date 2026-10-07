# 0946 — token reading dominates the remaining XML audit

A deeper reduction of the four retained 0944 profiles assigns about **54% of
XML-audit guest instructions to quick-xml token reading**, including its scanner
and tag-emission callees. This narrows the next investigation: measure why
scattered edits take the complete replacement-audit path and how much validated
replacement coverage a bounded multi-window proof could avoid. No optimization
is implemented, no workload is rerun, and no native speedup is claimed.

## Evidence boundary

This is an exploratory offline decomposition of the known
[0944](0944-xlsx-post-attribute-publication-profile.md) Callgrind data, after the
[0905](0905-xml-attribute-summary.md) attribute-summary optimization. Function
selection followed inspection of those profiles. The input plan binds this
analysis; it is not a prospectively registered performance experiment.

The four measured part-6 dumps are medium, dense-sparse, dense-sparse, medium.
They profile one invocation of
`SourceBackedEditor::publish_multi_commit_to_stream`, with five changed XML
parts. Publication totals remain 92,990,939 / 167,532,948 / 167,533,924 /
92,989,937 guest instructions (Ir), exactly matching the earlier reduction.
All incoming and immediate outgoing edges, including zero-cost and zero-call
records, are retained for twelve functions: the publication and topology
owners, the OPC XML wrapper, five XML audit functions, QName scanning, and
quick-xml reading/start/end emission.

The executable is the 0944 native build, SHA-256
`63cbffd8c17299a4e2d1096e4ffa2a2d34a549cc452dfdf8ba5691b8ab7e75a3`,
from source `933ee45cbb5d2ef9dc29583f7ec890ad234e04fb`. It predates
[0945 payload sharing](0945-xlsx-shared-replacement-payloads.md). Current
`audit.rs`, `audit/wellformed.rs`, and `audit/namespaces.rs` match the profiled
source byte for byte, but that does not make these profiles a measurement of
the newly adopted caller or a current native-time fraction. Quick-xml 0.41.0
reader sources were checked against the cached package whose checksum matches
the standalone lockfile.

## Reconciled costs

First repeat of each shape; the retained analysis includes both repeats.
Inclusive rows below are **nested and must not be added**.

| Function | Medium self Ir | Medium inclusive Ir | Dense-sparse self Ir | Dense-sparse inclusive Ir |
| --- | ---: | ---: | ---: | ---: |
| OPC replacement XML wrapper | 195 | 49,735,942 | 195 | 95,298,865 |
| `verify_source_replacement` | 6,371,741 | 49,735,747 | 12,289,280 | 95,298,670 |
| `verify_with_policy` | 4,939,236 | 24,150,916 | 9,420,890 | 45,994,993 |
| `check_start` | 5,608,051 | 8,096,154 | 10,885,237 | 15,612,704 |
| `inspect_attributes` | 1,253,997 | 1,268,204 | 2,398,365 | 2,412,577 |
| `check_end` | 1,160,828 | 1,160,828 | 2,220,317 | 2,220,317 |
| `scan_qname` | 2,469,853 | 2,469,853 | 4,710,273 | 4,710,273 |
| `Reader::read_event_impl` | 9,473,741 | 26,896,423 | 18,146,510 | 51,486,033 |
| `ReaderState::emit_start` | 2,877,580 | 3,485,106 | 5,507,726 | 6,660,019 |
| `ReaderState::emit_end` | 3,018,708 | 3,395,276 | 5,784,291 | 6,503,601 |

Reader inclusive cost is 28.92% / 30.73% of the publication owner, or about
54% of the nested XML wrapper. `inspect_attributes` now contributes much less
than in the pre-0905 graph examined by
[0903](0903-xml-audit-hotloop-attribution.md); this is a historical diagnostic
comparison, not a new paired optimization estimate.

Within the reader alone, the following immediate partition is additive:

| Reader component | Medium Ir | Dense-sparse Ir |
| --- | ---: | ---: |
| Reader self | 9,473,741 | 18,146,510 |
| All immediate memchr scanner edges | 10,542,010 | 20,175,642 |
| Start emission, inclusive | 3,485,106 | 6,660,019 |
| End emission, inclusive | 3,395,276 | 6,503,601 |
| Question-mark emission | 290 | 261 |
| **Reader total** | **26,896,423** | **51,486,033** |

These scanner edges describe parsing work, not removable work. The start/end
rows include required token construction and balanced-name checks. The graph
aggregates shared callees across callers; it cannot assign their exclusive
self costs to an original-versus-replacement audit side. Raw call counts differ
between callers at similar instruction costs, as already observed in 0903.
They remain provenance and are not token counts or cost-per-call denominators.
Relative raw source positions are retained without claiming independently
resolved callsite or branch coverage. Prior debug attribution and Valgrind
limitations remain applicable.

## Next qualification and constraints

`verify_source_replacement` always fully audits the original first. Its
existing proof searches for one balanced element covering all changed bytes;
if the difference is too broad, or the window cannot be proved, it fully audits
the replacement. The multi-edit extension deferred by
[0747](0747-xlsx-publication-audit-reuse.md) remains a distinct hypothesis.
Current value-only writers retain unchanged output spans and know changed-cell
source spans internally, but their published metadata is not yet sufficient
source/replacement span-pair evidence.

Before implementation, collect an untimed, per-part diagnostic: original and
replacement lengths, common prefix/suffix, actual `Identical`/`Window`/`Complete`
proof, changed source/output span pairs, balanced-window eligibility, bytes and
events covered, namespace/markup-compatibility boundaries, and aggregate
counters. Distinguish full-artifact coverage from bytes actually re-audited.
That evidence must establish whether several bounded windows save meaningful
replacement work before selecting an implementation and freezing fresh gates.

Any candidate must keep the complete original-first audit, exact diagnostics
and offsets, freshness checks, limits, namespace state, preservation, and typed
refusal with no output. Windows must prove byte identity outside the changed
ranges and compatible parser state at every boundary; aggregate budgets must
be re-totalled for the whole replacement. Ambiguous spans, validation failures,
or unproved boundaries must use the existing complete audit. Planning parser
state and successful readback do not replace publication validation, and ADR
0032 does not authorize a parsed XML cache.

Disabling quick-xml end-name checking is excluded: the auditor relies on it,
and quick-xml still stores opened names even with the flag disabled. The
[0902 stopped-observer candidate](0902-xlsx-upfront-stopped-audit-candidate.md)
remains rejected. No previous rejection or safety condition is waived by this
diagnostic. The [source audit](results/change-0946/source-audit.md) records the
specific implementation boundaries and required discriminator.

## Validation and retention

The [packet](results/change-0946/README.md) contains the final reducer, bounded
analysis, input identities, and independent verification. Each selected function
conserves self plus all direct-child Ir against all incoming Ir. The complete
raw function self sum, measured owner, summary, and prior publication totals
also agree. No Rust source changed, so no compilation or Rust test run is
claimed. The broader Office goal remains open; iWork is excluded.

The final v2 reducer replayed byte-identically from independent copied inputs
and from the fully hash-verified external archive. Six negative controls
rejected corrupt inputs, duplicate schedule indices, traversal, descriptor
corruption, symlinks, and output overwrite. Preparation history is incomplete:
a draft helper and a provisional plan were overwritten and cannot be recovered.
The retained v1 output is historical only; reproducibility applies to final v2.
