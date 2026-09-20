# Frozen lookup integration gates

All seven gates passed for the 62 selected inputs in `freeze.json`: ODS tests,
strict all-target Clippy, strict rustdoc, package formatting, selected-source
formatting, crate boundaries, and diff checks. The standalone verifier confirms
1,700 passed tests, zero failures, and zero ignored tests across 127 summaries.
Source manifests before and after execution match the staged checkout.

`selected-source.tar.gz` retains every selected input from the isolated checkout,
with each member verified against the freeze. SHA-256: `219db7325dbd7b216e43e867fcd314b826afa88e6c4c5691106dc8a71a6f0b39`.
Reconstruct the candidate by overlaying these files on baseline commit
`635fd2e1348b621426b50909cbd5765c91837306`. The retained `Cargo.lock` is the gate dependency input;
the ambient workspace lockfile is not part of this receipt.

Attempts 1 through 6 are retained in diagnostic directories because subsequent
review or source changes superseded their successful runs. They do not certify
the current candidate. Performance acceptance is recorded separately.
