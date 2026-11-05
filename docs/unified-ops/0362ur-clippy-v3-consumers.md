# 0362ur — Workspace Clippy v3 consumer repairs

This source-only repair starts at integration commit
`355cfe414118f0a76dc907a33893e938bfb0bebe`. The completed all-workspace,
all-target, all-feature Clippy run with warnings denied reported diagnostics in
the XLSX, DOCX and XLSB consumers. The fixes remove redundant closures,
needless struct updates and borrows, simplify option predicates, and replace
test-only `expect` calls with `unwrap` under the existing test-module lint
policy. They also move a glossary test module after the production helpers and
group glossary graph budgets in a typed value so its public constructor stays
within Clippy's argument limit.

The three large-enum repairs are representation improvements, not lint-only
work: OPC relationship tokens carry a complete relationship index and are
uncommon sidecars, so XLSX form-control payloads, DOCX durable proof parts and
XLSB theme owners keep that index behind a pointer instead of inflating their
common enum footprint. Existing matching, serialization, validation and
source-retention behavior is unchanged.

The immutable terminal log is retained outside the checkout at
`(deleted-coord-artifact)/check-0.log`
(35,425 bytes, SHA-256
`cfdaf7c207148f82599774af1a3409b34bba4506f5e50af704683a1d5cd6be4f`). The
command was
`cargo clippy --offline --locked --workspace --all-targets --all-features --keep-going -- -D warnings`;
it exited 101 on the unmodified integration source after 275.42 seconds.
Cargo was not run in this isolated leaf.

## Validation

- `rustfmt --edition 2024` and `git diff --check` passed.
- The repository artifact guards pass for the staged source and record files.
- A coordinator rerun of the affected workspace Clippy and tests is required
  after this source-only commit is integrated.
