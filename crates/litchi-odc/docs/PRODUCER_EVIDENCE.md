# Standalone ODC/FODC producer evidence

Checked on 2026-08-10 for the ODC remediation review.

No genuine standalone producer-created `.odc` or `.fodc` is available in the
checked-in corpus, repository history, or the local user filesystem. The local
environment also has no `libreoffice` or `soffice` executable, so it cannot
create and resave a standalone chart through LibreOffice for this review.

The repository does contain genuine LibreOffice chart subdocuments inside
producer-created `.fods` and `.fodt` files. Those embedded XML fragments are
useful interoperability inputs for the shared chart reader, but they are not
standalone ODC/FODC packages and are deliberately not copied, repackaged, or
renamed as standalone fixture evidence.

A future fixture qualifies only if its provenance records a native producer
that saved a standalone ODC/FODC artifact. The corresponding test should open
the original artifact, publish a changed file, fully reopen it, and include a
current native-application resave before claiming changed-file interoperability.
