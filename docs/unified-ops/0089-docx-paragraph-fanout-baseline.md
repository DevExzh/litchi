# 0089 — Paired DOCX paragraph text fanout baseline

A separate frozen harness now compares the unified paragraph text fanout from
0087 with the native bulk paragraph replacement API. This establishes a paired
baseline, not a before/after speedup. No production API changed in this record.
The full unified operations goal remains active.

## Evidence

Two independent invocations each ran nine native/unified/unified/native rounds
on CPU 0 (AMD EPYC 9R45), using prepared eager and source-backed memory owners.
Each successful scenario checked exact output bytes and both leaf ledger counts
before collecting 18 samples per leg. The unified operation count must be one.
The five generated corpora have frozen sizes and hashes enforced at startup.
The input replaces one-run paragraph text with text containing tab and LF.

Ratios below are the range of the two per-invocation upper-median ratios:

| Paragraph targets | Eager unified / native | Source unified / native |
|---:|---:|---:|
| 1 | 1.428–1.431 | 1.256–1.276 |
| 10 | 1.327–1.328 | 1.242–1.248 |
| 100 | 1.285–1.296 | 1.231–1.244 |
| 1,000 | 1.273–1.284 | 1.235–1.238 |
| 10,000 | Typed refusal; not timed | Typed refusal; not timed |

The one-paragraph case exceeds the goal's 10% single-write review trigger.
The unified scalar path and native bulk path are equivalent in this corpus,
but profiling and additional representative small edits are needed before
choosing an optimization. No allocation or RSS measurement was taken.

The native bulk call refuses 10,000 replacements with
`LimitKind::CompositeParagraphReplacements` (maximum 4,096). The unified call
returns `Code::LimitExceeded`. The harness asserts both failures and does not
split this one operation into multiple transactions. This leaves successful
10k fanout scaling unmeasured. A one-query fanout also does not establish the
goal's command-batch requirement, including per-item results and errors.

Opening, query compilation, properties/native target construction, serialization
and commit destruction are excluded. Query execution, target proof, property
admission, staging, diagnostics and commit are included. The output commit is
retained until timing stops on each leg. Source mode uses memory owners;
filesystem revision checks are outside this workload. Full performance
acceptance, other workloads/formats and Python remain open.

## Reproduction and validation

See [harness instructions](../../tools/perf-baseline/unified-docx-text-fanout/README.md)
for frozen route names, commands and measurement boundaries. Machine-readable
samples, independently recomputed summaries and source/binary provenance are
under [unified-ops-0089](../performance/results/unified-ops-0089/).
The intentional harness lockfile is retained. The binary is not committed;
its name, size, SHA-256 and build command are in provenance.json.

Release build with the locked dependency graph and all-target release Clippy
with warnings denied passed. Both final invocations passed byte/ledger checks,
frozen-corpus checks and typed refusals. The Python summarizer validates sample
counts, corpus identities, medians, ratios and refusal status across both runs.

Scoped formatting and diff hygiene passed, and the rewritten base remains an
ancestor of HEAD. Both required staged evidence checks printed nothing; an
audit of 382 changed/untracked paths found no prohibited evidence or binaries.
Staging failed because the linked Git metadata is read-only at
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`; no commit
was created. The checkpoint's temporary build output, logs and probe were
removed after validation. No temporary worktree was created.
