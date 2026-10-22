# 0362rga — DOCX raw-source preflight and error taxonomy repair

This repair closes the DOCX review findings on the redaction, sanitization,
and story-hyperlink families. It is based on `4f3b589e54` on branch
`uo2/0362rg-docx-redact-sanitize`.

The source-backed raw paths now preflight declared main-document and story
sizes before payload reads. Story capture discovers the complete relationship
graph and checks per-part and aggregate byte and relationship budgets before
allocating any story XML or relationship vector; the load phase verifies the
declared sizes again after the source read. The main hyperlink paths check the
declared main-document size before creating the semantic snapshot. Invalid
relationship diagnostics use fixed reasons and do not echo relationship IDs.

The public error surface follows the shared contract:

| Condition | DOCX error | Facade code |
| --- | --- | --- |
| Raw or source-backed detachment base mismatch | `PatchConflict` | `invalid_input` |
| Main or story redaction base/version/lineage/fingerprint mismatch | `PatchConflict` | `invalid_input` |
| Inverse of a forward-only redaction on a foreign package | `Refused(UnsupportedOperation)` | `unsupported_capability` |
| Document-variable source mismatch | `PatchConflict` | `invalid_input` |
| Document-variable markup cannot be preserved | `Refused(UnsupportedContent)` | `unsupported_capability` |

The former public family-specific conflict and preservation variants were
removed. Consumers, tests, documentation, and the facade classifier use the
shared variants. A source-backed detachment replay test mutates an unrelated
package part and confirms that the complete-artifact proof rejects it before
output; the original owner/share proof remains available for exact matching
packages.

The approved allocator comparison is implemented in the existing
`litchi-perf-baseline-alloc docx-raw-codec` dispatch. It compares one
detach/redact raw-codec lifecycle with the corresponding pre-canonicalization
plan lifecycle; fixture construction, source opening, and semantic output
verification are outside each measured region. The report records allocation
calls, bytes, and region peak for each lifecycle, and treats the pre-existing
redaction rewrite clone as a baseline operation rather than a regression
introduced by this repair. The external receipt will be added after the
coordinator releases the lane for the capped build.

This checkpoint has passed `cargo fmt --all -- --check` and `git diff --check`.
The focused tests, all-target Clippy, and the allocator receipt remain pending
the released lane; no performance conclusion is made at this checkpoint.
