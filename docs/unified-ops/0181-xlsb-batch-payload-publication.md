# 0181 — XLSB batch payload publication

## Contract

XLSB structural visibility publication stages changed workbook/worksheet view
payloads and installs them through `OpcPackage::replace_part_payloads`.
The native edit still validates sheet kinds, active-view relocation, worksheet
content, physical source limits and signatures before publication. The OPC
boundary resolves all targets and decodes existing payloads before changing any
payload. It rejects duplicate resolved targets, including case aliases, and
reserves its staging allocations fallibly. Empty input retains exact-source
authorization.

For built-in parts, one exact-source revocation records the current signature
graph before the replacements. Ordinary payload setters invalidate XML source
provenance and compressed payload captures. Dynamic final signature detection
and the sticky disposition requirement remain intact. Packages containing
custom parts are refused explicitly because caller-defined setters cannot
provide atomic failure. The existing individual mutable-part API is unchanged.

This removes N complete signature-graph scans from N changed worksheet views.
The workbook metadata replacement before view staging still uses the ordinary
single-part path; a constant number of full-graph policy scans remain. No native
visibility grammar, facade capability, property spelling or OfficeCLI behavior
changes.

## Verification and evidence

Independent implementation, review, test and benchmark audits were used.
Signing-enabled full suites passed: 1,160 OPC and 1,150 XLSB tests/doctests;
the combined DOCX/XLSX/PPTX/XLSB/encryption facade suite passed 1,230.
After the final custom-part refusal refinement, all six OPC batch regressions
passed again; all fifteen native sheet-visibility tests passed with the added
multi-target opaque-record regression. That regression independently checks
both worksheet payloads, selection flags and exact whole-package inverse.
Other focused coverage includes missing targets, duplicate aliases, deferred
refusal, custom-part refusal before setters, exact empty batches, XML provenance
invalidation and low-level signature injection with sticky policy.

Final combined all-target Clippy and affected-crate rustdoc passed with warnings
denied. Scoped formatting, diff checks, the crate-boundary audit (66 packages,
268 internal declarations, 11 existing debts) and generated OfficeCLI parity
check passed. Verification used an external copied lockfile with cached cc
1.6.0; repository lockfiles are unchanged. The base branch has no missing
commits. Temporary worktrees, build outputs and logs were removed after
recording reduced evidence. No Python, CLI or parser-fuzz completion is claimed
by this native publication change.

## Paired measurements

The unchanged direct-command harness ran 1, 10, 100, 1,000 and 10,000 commands
for eager and source owners, three ABBA rounds (six samples per leg), pinned
to CPU 0 without concurrent build/test/audit jobs. Corpus generation, owner
opening and command preparation precede timing. Source promotion is inside
the source timing; final output serialization is outside both legs. All cases
have byte-exact warm equivalence, one native/facade commit and empty failure
matrices. Source/build/binary identities and raw samples are retained in
[before](results/0181-xlsb-sheet-batch-before.json),
[after](results/0181-xlsb-sheet-batch-after.json), and
[repeat](results/0181-xlsb-sheet-batch-repeat.json).

| Owner / commands | Before native / facade ms | After native / facade ms | Repeat native / facade ms |
|---|---:|---:|---:|
| eager / 1,000 | 144.754 / 145.836 | 65.185 / 67.876 | 65.316 / 66.929 |
| eager / 10,000 | 9022.079 / 8727.435 | 925.405 / 952.887 | 919.513 / 942.929 |
| source / 1,000 | 177.613 / 179.324 | 99.041 / 101.516 | 97.926 / 100.083 |
| source / 10,000 | 10533.611 / 10288.915 | 2627.316 / 2648.130 | 2608.256 / 2647.467 |

At 10,000 commands native eager publication improves about 90%, and the source
leg improves about 75%. After/repeat facade overhead at 1k/10k stays below 4.2%,
well within the scoped 2x batch comparison bar. Native 1k-to-10k growth falls
from 62x to about 14x eager and from 59x to about 27x source. These results do
not establish linear scaling: source promotion remains a clear follow-up, and
the remaining eager slope also needs qualification. General sheet-command
coalescing and the full all-format/Python/CLI goal remain open.

A source audit identifies a remaining quadratic promotion loop:
`SourceBackedPackage::materialize_opc_package_with_accounting` calls
`try_add_source_part` for each catalog part, and each insertion's
`validate_new_part_name` scans all existing part names. The source catalog has
already validated its namespace with `PartNameIndex`. A private batch insertion
boundary that preserves that proof, or preflights once with the index, is the
next candidate. Its share of remaining runtime still needs measurement; the
public single-part insertion checks should remain intact.
