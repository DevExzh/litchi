PPTX read-only recommendation
=============================

Adds typed inert p1710:readonlyRecommended inspection and source-checked reversible CRUD at its presentation-properties extension owner. Exact no-ops preserve source bytes; mutations retain unrelated extensions, comments, namespaces and lexical context. Package validation checks the owner relationship, content type and duplicate/orphan cases. Signed no-ops remain valid; changing signed content requires the existing explicit signature policy. The recommendation is an inert document hint, not a host editing-policy implementation.

The final candidate fixes a review-discovered XML attribute scanner defect. TAB, LF, CR, CRLF and mixed separators now support exact no-op, scalar splice, reopen and inverse restoration. Independent crypto_native_vectors review cleared the original normative/API scope and the final two-file whitespace correction.

Root final-source validation passed 894 ordinary tests and six doctests (two doctests ignored), all-target/all-feature Clippy with warnings denied, warning-denied documentation, and scoped formatting. Source hashes, exact commands and compressed raw logs are retained here; these are crate gates, not whole-workspace certification.
