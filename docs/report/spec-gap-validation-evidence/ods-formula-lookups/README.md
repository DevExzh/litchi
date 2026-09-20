# ODS lookup and reference functions

Status: implemented in `6fa3b8af6a`. Semantic/resource review, all seven isolated
gates, and retained evidence verification pass. See the
[disposition](completion.md) for the accepted measured performance limitations.

The batch covers ADDRESS, CHOOSE, HLOOKUP, INDEX, INDIRECT, LOOKUP, MATCH,
OFFSET, and VLOOKUP in the existing explicit, read-only formula evaluator.
The baseline is `635fd2e1348b621426b50909cbd5765c91837306`.

See [integration-plan.md](integration-plan.md) for ownership and acceptance,
[baseline.json](baseline.json) for source custody, and
[oracle-native-plan.md](oracle-native-plan.md) for independent evidence design.
[adr-review.md](adr-review.md) records accepted design constraints and
[resource-plan.md](resource-plan.md) records the independent resource design.
[contract.md](contract.md) records implementation contract v2, accepted by the
independent [specification review](spec-review.md). Its original pending-status
header is retained to preserve the reviewed bytes. The current
[semantic review](semantic-review.md) and [resource review](resource-review.md)
are bound to the 62-input freeze by [review-receipt.json](review-receipt.json).
[Gate evidence](gates/README.md) records 1,700 passed tests, zero failures or
ignored tests, strict Clippy/rustdoc, formatting, boundaries, and source stability.
The focused suites contain 34 semantic tests and 9 resource tests; independent
evidence contains 127 oracle observations and 32 native rows.
[coverage-requirements.json](coverage-requirements.json) binds the executed
evidence to the immutable requirements. [verification.json](verification.json)
records the complete retained-only verification after temporary checkout cleanup.
GETPIVOTDATA, MULTIPLE.OPERATIONS, dependency recalculation, pivot calculation,
external-source execution and formula-cache publication remain open work in the
broader audit.

The retained isolated gate lockfile is distinct from the ambient root lockfile;
its exact source and hash are recorded in `baseline.json`. No gate should replace
the root lockfile. Performance captures require a frozen source handoff and must
retain all failed attempts, superseded captures, and measured review flags.

Computed CHOOSE probes retain the projection demand as part of their cache key.
If a selected dynamic reference widens that demand, the evaluator may revisit
selector inputs while preserving position-sensitive computation. The direct
INDIRECT width-growth test fixes its own read count; it does not establish a
single-read guarantee for every computed CHOOSE/reference composition. Final
performance claims must retain this distinction.

The immutable [coverage scope](coverage-scope.json) is frozen alongside the
implementation inputs. The separate [coverage bindings](coverage-requirements.json)
are completed after execution because they contain hashes of gate and performance
receipts. Verification compares their requirement lists exactly and checks the
frozen scope hash; completing receipts cannot remove or weaken requirements.

Search key-array lifting starts only after ForceArray data admission. Invalid
data can therefore produce one global `#VALUE!` before any key-cell reads.
An already-known scalar formula-error key keeps its original error identity on
that refusal path. Once data is admitted, lifted controls and keys retain their
per-coordinate results and errors.
