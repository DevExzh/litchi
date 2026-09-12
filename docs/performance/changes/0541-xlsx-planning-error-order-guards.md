# 0541: public XLSX planning error-order guards

Six public integration tests establish the error precedence that a shared
worksheet event stream must preserve. This closes a prerequisite identified
by the fresh 0540 profiles; it does not change production or claim a speedup.
The tests call `SourceBackedEditor::edit_sheets`, including the selected
multi-sheet loader measured by that profiling campaign.

The 37 named fixture configurations cover 19 independent failures, nine
combined raw/validation cases, three valid namespace/PI controls, three MCE
cases, two later-selected-sheet failures, and one cross-sheet ordering case.
They match typed errors and exact stable messages. Parser diagnostics with
library-specific wording retain the error variant and a scoped message check.
Independent numeric, boolean, style, formula, coordinate and entity controls
make the competing error owners observable before they are combined.

Within a worksheet, complete value-only validation wins over a parser error
that appears earlier in the bytes. MCE preprocessing precedes raw parsing,
but a later validator refusal still overrides both. Across worksheets, the
loader validates and parses each selected sheet in workbook order: a later
sheet's validation error cannot replace the first sheet's raw error, even if
the caller supplies selectors in reverse order. Failed plans are retried,
source bytes remain unchanged, and valid unselected owners remain usable.
The later-sheet cases first load a valid sheet, then refuse the faulty sheet
without returning a partial multi-sheet transaction.

Adversarial worksheet XML is inserted through the existing raw ZIP writer
used by neighboring source-backed tests. This is necessary because the OPC
authored-XML writer correctly rejects duplicate/malformed XML before an editor
can read it. Package/catalog members remain valid so assertions reach the
intended worksheet boundary. Production validation is not bypassed or changed.

The [evidence bundle](../results/change-0541/README.md) retains all five source
attempts, exact test copies and patches, source-bound command receipts, and
independent review. Attempt 1 failed compilation on the test-module path.
Attempt 2 passed three tests and failed while building a malformed fixture.
Attempt 3 passed four focused tests, 1,296 full-suite executions and all six
quality checks, then was superseded for the review additions. Attempt 4 also passed six focused tests, 1,298 full-suite executions,
workspace check, Clippy and rustdoc, but failed formatting because its temporary
draft used the default configuration. Attempt 5 applies only the five required
match-arm commas and passes six focused tests, 1,298 full-suite executions,
workspace check, targeted Clippy, rustdoc, formatting and crate boundaries.

This is a test prerequisite, not admission of the future parser change.
The next candidate still needs its own frozen native/allocation gates,
matched baseline and candidate captures, bounded provisional parser state,
source/MCE/x14ac fallback proof, exact no-op and invalid-input guards, and
full correctness checks. Existing tests retain resource, cancellation,
source-version and raw extension coverage; the new matrix is not exhaustive.
OLE2/OOXML remain first, ODF is deferred, and iWork is excluded.
