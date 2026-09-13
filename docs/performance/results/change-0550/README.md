# 0550: current source-backed XLSX commit attribution

`performance_claim: none; diagnostic baseline only`

Fresh exact-owner profiles identify worksheet layout scanning as the leading
commit cost: 53.24–54.94% of instruction references across the four shapes.
XML validation accounts for 30.40–35.11%; provenance merge for 4.64–6.03%.
These nested shares are not elapsed shares or removable fractions. Production
and harness source remain unchanged at `090b15b64ae52da2f8bf765cbb745ef76122792e`.

The next task is a source-bound layout proof that could reuse facts produced
by the existing eligible source traversal. It must eliminate work across
planning and commit, preserve all scanner facts and authoritative fallbacks,
and pass a fresh matched campaign. See [next target](next-target.md) and
[source review](source-review.md). No candidate is admitted in this batch.

## Current native baseline

Two serial repeats each contain 30 measured samples after 3 warmups, on CPU 2.
The 480 native samples are descriptive and do not authorize a registered
latency claim. Values below are p50 milliseconds. The commit phase includes
staged `edit.set` work; the exact Callgrind owner excludes that staging.
Reopen/verification is separate from the measured edit/save workflow.

| Shape | Update | Workflow R1 / R2 | Commit phase R1 / R2 |
| --- | --- | ---: | ---: |
| medium | one cell | 5.3261 / 5.3528 | 1.9140 / 1.9421 |
| medium | 1% | 20.2578 / 20.4271 | 8.0325 / 8.1321 |
| dense-sparse | one cell | 34.8638 / 35.0691 | 12.6746 / 12.8335 |
| dense-sparse | 1% | 38.8424 / 38.9032 | 15.0454 / 15.1103 |
| noncompact | one cell | 5.9176 / 5.8737 | 2.4798 / 2.4671 |
| noncompact | 1% | 22.7799 / 22.0884 | 10.0758 / 10.0668 |
| vendor-extension | one cell | 5.3998 / 5.4006 | 1.9635 / 1.9729 |
| vendor-extension | 1% | 20.2894 / 20.1492 | 7.9992 / 7.9619 |

All 22 native repeat-drift observations above 5% remain in the machine-readable
[metrics analysis](metrics-analysis.json) and [individual review](adverse-review.json).
Six small immediate-child instruction drift rows are also reviewed.
No allocation repeat drift was observed in the captured metric vectors.
The shared host does not justify a stability, tail, physical-cold, provider,
cache-event, or scaling claim.

## Exact commit profiles

Each of eight profiles has one measured `MultiSourceEdit::commit` invocation.
The 26 lifecycle dumps and eight zero-cost termination dumps are independently
classified and excluded from measured costs. The one-cell case also uses the
multi-sheet API with one selected sheet; neither selected benchmark exercises
`SourceEdit::commit`. See [scope review](scope-review.md).

| Shape | Commit Ir R1 / R2 | Layout scan Ir R1 / R2 |
| --- | ---: | ---: |
| medium | 133,269,922 / 133,280,376 | 70,977,241 / 70,990,409 |
| dense-sparse | 258,256,329 / 258,284,803 | 137,507,562 / 137,536,341 |
| noncompact | 165,898,276 / 165,902,662 | 91,137,124 / 91,137,400 |
| vendor-extension | 133,297,977 / 133,269,518 | 71,003,018 / 70,977,289 |

XML-validation attribution includes workbook checks as well as worksheets.
Reduced-readback and reduced-parser source-level names have no positive
out-of-line edge; this is missing separate attribution, not zero executed work.
The raw profiles, complete annotation trees, exact owner partitions, nested
rows, and repeat drift are retained in [profile analysis](profile-analysis.json)
and [profile review](profile-review.md). Instrumented elapsed is excluded.

## Allocation evidence

The separate allocator binary supplies 480 operation-region samples. Every
reported allocation metric is constant within each case/shape and equal across
repeats. The following one-percent commit-region values include staging and
commit, not just the profiled function. They establish a current baseline,
not a memory improvement or whole-document peak.

| Shape | Allocation calls | Allocated bytes | Incremental region peak bytes |
| --- | ---: | ---: | ---: |
| medium | 91,391 | 12,436,307 | 2,344,427 |
| dense-sparse | 172,946 | 18,258,492 | 4,467,508 |
| noncompact | 146,707 | 14,176,758 | 2,506,631 |
| vendor-extension | 91,391 | 12,436,307 | 2,344,427 |

Planning and publication regions, one-cell rows, and all source/sink/output/
semantic identities are retained in the metrics report. Normal binaries
report allocation evidence as unavailable. Copied/reduced-byte and fallback
route counters were not added; they remain measurement gaps.

## Validation, custody and reproduction

Both builds, all 8 preflights, 16 native jobs, 8 successful profile jobs,
16 allocator jobs, and four metadata checks passed. The metadata checks cover
workspace/harness formatting, crate boundaries, and the strict claim registry.
The metrics validator also rejects real-report copies with a short elapsed
vector, short commit vector, or wrong binary hash; the valid original control
passes. The probe's initial external-path setup error is preserved separately.

The full 8,593-file source inventory equals the prior 0549 final manifest.
Prior quality contains 15 checks and 4,382 OLE test executions; it is not a
fresh XLSX test suite. No new broad test, sanitizer, fuzz or native Office
validation is claimed. The existing harness's semantic, lifecycle and
preservation oracles run for every applicable child. All accepted ADR/index
hashes and metadata checker sources match their retained baseline identities.
See [ADR compliance](adr-compliance.md) and [protocol](protocol.md).

One profile setup failed before measurement when Valgrind gdbserver could not
initialize shared memory. The original receipt and artifacts remain retained;
a separately frozen amendment disables the unused debugger and keeps its
paths inside the owned target. No passed capture was replaced. Agent preflight
analysis and annotations are preserved and exactly reproduced by root
canonical analysis receipts; see [execution custody](analysis-execution-note.md).

Strict verification replays source, capture, analyzer, review and cleanup
custody. Run `python3 -B verify.py --strict` from this bundle after sealing.
Fresh captures require an isolated checkout of the recorded source revision
and fresh evidence/target paths; do not rerun captures into this sealed bundle.
Both binary hashes are checked before removing the sole owned target.

OLE2/OOXML remain the priority. ODF is deferred until their optimization goal
completes, and iWork is excluded. The broader performance goal remains open.
