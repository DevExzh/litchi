# ODS lookup and reference functions

Status: active implementation and validation work. This bundle does not yet
claim production support or passing gates.

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
header is retained to preserve the reviewed bytes; implementation remains on
HOLD pending source fixes and validation. [coverage-requirements.json](coverage-requirements.json)
lists required coverage, with evidence still pending. Neither document is a
passing receipt. Final source reviews will be linked here after source handoff.
GETPIVOTDATA, MULTIPLE.OPERATIONS, dependency recalculation, pivot calculation,
external-source execution and formula-cache publication remain open work in the
broader audit.

The retained isolated gate lockfile is distinct from the ambient root lockfile;
its exact source and hash are recorded in `baseline.json`. No gate should replace
the root lockfile. Performance captures require a frozen source handoff and must
retain all failed attempts, superseded captures, and measured review flags.
