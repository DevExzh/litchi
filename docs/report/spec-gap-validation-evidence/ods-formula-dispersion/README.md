# ODS variance and standard deviation evidence

This batch implements `VAR`, `VARA`, `VARP`, `VARPA`, `STDEV`, `STDEVA`,
`STDEVP`, and `STDEVPA` in the explicit read-only formula evaluator. Workbook
formula caches remain inert during document CRUD.

The normative [contract](contract.md), [semantic review](review.md), and
[resource review](resource-review.md) accept the frozen implementation. All
seven integration gates and independent numerical, native, and performance
verification pass against the recorded source hashes.

The baseline commit and pre-existing unrelated edits are recorded in
[baseline.json](baseline.json). Isolated gates use the retained
`gates/Cargo.lock`, inherited from the preceding statistical-reducer batch;
the ambient workspace lockfile is not their dependency input.

The native receipt currently retains 48 cached numeric observations, six per
function, from eight SHA-256 pinned LibreOffice FODS files. Root independently
refetched the inputs and reproduced identical JSON; see
[native-reproduction.json](native-reproduction.json). Fifteen exclusions are
explicit. These particular native closures cover numbers, text, and empty
cells; logical-value admission is covered by authored tests and the independent
oracle rather than attributed to these native fixtures.

The independent numerical oracle retains 512 observations over 56 typed
fixtures, each evaluated in scalar and matrix value modes. Exact rational
variance and 600-digit Decimal square roots supply the reference values.
Ordinary comparisons allow eight ULP; the adversarial fixtures use the
existing database profile's `2e-12` relative tolerance. Exact positive zero
is required where expected. These are corpus acceptance thresholds, not a
universal error bound or correct-rounding guarantee.
[Oracle review](oracle-review.md) records the initial four-ULP failures and
observed error maxima, including order-sensitive majority-equal fixtures.

The isolated gates pass 1,378 ODS tests with zero failures and zero ignored
tests, all-target Clippy with warnings denied, rustdoc with warnings denied,
crate and selected-file formatting, crate boundaries, and diff whitespace.
The complete compiled source closure is unchanged before and after the gates.
[Independent integration verification](integration-verification.json) checks
the gate logs, frozen inputs, numerical generator, and native reproduction.
The new focused suites comprise 9 semantic, 11 resource, 8 oracle, and 2 native
tests; existing statistical and database tests remain part of the full gate.

The final [performance report](performance/results/performance-report.md)
retains 510 baseline and 3,090 candidate samples: seventeen matched controls
and 103 candidate workloads, both evaluate and parse/evaluate phases, fifteen
fresh child processes per group, and three warmups. Matched allocation calls,
requested bytes, and peak live bytes are unchanged. Unrounded median time
shifts range from −3.24% to +2.35%; RSS shifts range from −2.52% to +4.22%.
No positive matched metric exceeds the 5% review threshold. These are
single-host observations, not causal or cross-platform speedup claims.
Nested VAR, VARA, and STDEVP scaling reads 512, 2,048, and 8,192 cells for
64, 256, and 1,024 fixture rows respectively: one outer condition scan and
one cached inner scan. Refused reference-list and resource cases read no cells.

The first capture was discarded because isolated staging omitted the
noncompiled oracle generator; compiled Rust and golden JSON already matched.
The complete capture was repeated after correcting staging, with no selective
metric reuse. See [custody correction](capture-custody-correction.json).

From a repository checkout matching the selected frozen sources, run
`python3 docs/report/spec-gap-validation-evidence/ods-formula-dispersion/verify.py`.
The verifier checks all seven gate receipts, regenerates the independent
oracle, checks native reproduction and source identity, hashes every retained
performance artifact, and independently recomputes raw medians and read bounds.
See [root verification](root-verification.json) and [cleanup](cleanup.json).
