# 0362pa — Primary replay diagnostics and consumer repairs

The coordinator returns `1c7eb1e84c` for an incorrect RTF replay diagnostic
and three unmigrated unified test consumers. The repairs are included in
the 0362u integration candidate. This is correctness and consumer evidence;
native performance qualification remains open.

An appended RTF paragraph adds 26 bytes in the regression fixture. Applying
the patch's inverse removes those bytes, but `Commit::from_replay` reports
26 inserted bytes because it ignores the patch direction. The
[regression](results/0362pa/regression.patch) fails with observed 26,
expected 0. The repaired diagnostic reports zero for removal and the exact
payload length for append. It also verifies source/output sizes, restored
bytes and no-op owner identity. Paragraph/run counts explicitly describe
caller-staged input; replay supplies no such staging input.

The broader all-target/all-feature check exposes three test consumers in
`litchi-ops`: DOCX source package adoption, PPTX promoted-source replay and
PPTX inverse-source adoption. They now use `into_package()`. An XLS source
error constructor also uses `io::Error::other` to satisfy the workspace
Clippy policy without changing its error kind or fixed reason.

The [receipt](results/0362pa/review.json) retains the failing regression,
locked passing rerun, both failed Clippy attempts and the successful
nine-package Clippy/formatting run. The first regression invocation stopped
before testing because the ignored workspace lock lacked the incoming ops
SHA-256 dependency; its log is retained separately from the actual failure.
Full merged tests are tracked by 0362u. Replay's newly retained Patch also
needs matched allocation/native review; no cost increase is accepted here.
