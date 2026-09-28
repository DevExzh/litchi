# 0817 — real-file save timing withheld by admission

The independent artifact audit rejected admission, so this batch adds no
performance baseline. All six corpora were exported across five publication
policies, but no timing case ran. This follows the real-file step queued by
0816. Production and
runtime harness code remain at `953866d382`. The only Rust change corrects an
integration test to expect the instrumentation label selected by the existing
`ordinary-save-process-metrics` feature. No optimization or historical speedup
comparison is introduced.

The three named inputs are DOCX `documentProperties.docx` (23,503 bytes, 12
members), XLSX `dateAutofilter.xlsx` (8,435 bytes, 10 members), and PPTX
`shapes.pptx` (68,822 bytes, 48 members). Their repository paths and exact hashes
are retained in [the input inventory](results/change-0817/input-inventory.json).
Tracked Office interoperability provenance describes earlier edits and an
external LibreOffice resave; it does not certify these newly generated outputs
or establish a Microsoft producer for DOCX/PPTX.

The original audit reports 22 errors. Fifteen arise from applying real-file
archive accounting to generated semantic-corpus metadata. The remaining seven
concern DOCX relationship ordering and XLSX calculation properties. These error
counts are not counts of production defects. The failed audit, its exact script,
and all outputs remain immutable under
[admission-0](results/change-0817/admission-0/artifact-audit.json).

| Real input | Exact observation | Admission consequence |
| --- | --- | --- |
| DOCX | `rId4` (font table) moves from the first relationship to between `rId3` and `rId5`; the relationship edge set is unchanged. | Decoded relationship bytes and order differ outside the paragraph edit. |
| XLSX | `calcPr calcId="152511"` becomes `calcId="0"` with full-recalculation flags. | The frozen audit omitted calculation invalidation from the allowed edit closure. |
| PPTX | The selected shape edit passes the frozen local ZIP/XML checks. | This does not admit the all-three-input timing matrix by itself. |

XLSX's observed flags are `fullCalcOnLoad="true"`, `calcCompleted="false"`,
`calcOnSave="true"`, and `forceFullCalc="true"`; the generated XLSX control adds
the same calculation properties. The DOCX error labeled “relationship graph”
compares an ordered edge list and must not be interpreted as a lost relationship.
Neither deterministic output nor successful reopening overrides the frozen
preservation gate. No case was substituted or omitted to obtain timing results.

The edit appends a DOCX paragraph, replaces XLSX first-sheet A1, or replaces
the first admitted PPTX slide/shape text with the harness marker. Artifact
export precedes admission and includes three generated controls alongside the
three real inputs. The independent ZIP/XML reader checks the intended edit,
non-target content, member inventory, content-type assignments, and relationship
graph. Five publication policies per case must produce identical bytes.
Structural XML comparison is a local oracle, not independent Office application
validation; compressed payload and selected ZIP metadata comparisons remain
separate evidence.

All twelve planned timing cases were withheld. No qualification, native,
observer, latency, allocation, or process-counter report was collected. The
[unchanged plan](results/change-0817/plan.json) retains the intended protocol;
its 108 reports and 2,244 samples are planned counts, not completed evidence.
The phases would have separately measured lifecycle, edit, default full-durability
atomic publication, and counting publication. PPTX counting publication
materializes `to_bytes` and would not establish streaming memory behavior.

The release builds used Rust/Cargo 1.95.0, opt3, thin LTO, one codegen unit,
debug level one, unwind, and two build jobs. All three executables built and the
untimed exporter succeeded. The existing standalone harness lock differs from
the root workspace lock; [exact lock differences](results/change-0817/lock-parity.json)
are retained without changing dependencies. This record imports no timings.

All six quality gates passed with an explicit test recovery. The original full
test command recorded 640 passes and one ignored test in 26 successful suites,
then failed the stale instrumentation assertion. Exact source and log checks
justify reusing those unchanged suites. The corrected integration passed under
all features and allocator-only; the doctest command passed with zero doctests.
Formatting, all-target checking, warning-denied Clippy, rustdoc, and the crate
boundary checker passed freshly. This does not relabel the original failed full
test command as successful.

[Execution notes](results/change-0817/execution-notes.md) retain the earlier
input-freeze coordination failure and pre-Cargo build-driver field lookup
failure, with immutable witnesses. The latter changed no Rust source, feature,
profile, dependency, or quality result. [Source review](results/change-0817/source-review.md)
and [build recovery review](results/change-0817/build-recovery-review.md) explain
those boundaries. All 35 normative inputs and 9,196 production files retain
their recorded identities. The broader OLE2/OOXML goal remains active; ODF is
deferred and iWork excluded.

The [independent results review](results/change-0817/results-review.md) and
[artifact diagnosis](results/change-0817/admission-diagnosis.md) preserve the
exact distinctions above. The earlier graph-based oracle's tolerance for DOCX
normalization does not override this plan's untouched-byte/order requirement
or the ordering requirement in `docs/GOAL.md`.

The [offline analysis](results/change-0817/analysis.json) and
[admission summary](results/change-0817/admission-summary.json) replay the failed
outcome and explicitly require absence of all timing lanes. Two reader defects
were corrected with failed logs retained: the flattened archived test path and
a missing local custody variable. Neither correction changed measurements,
outputs, or the frozen admission oracle.

Root verified the three captured executable hashes before removing both owned
build and scratch directories: 9,063 files / 8,591,960,251 logical bytes. Post-cleanup
validation passes. Reproduce with
`python3 -B docs/performance/results/change-0817/validate.py --final`.
The three unrelated workspace files remain unchanged. The immediate
[next step](results/change-0817/next-step.md) resolves current admission before
larger-corpus timing; no performance improvement is claimed.
