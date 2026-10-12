# 0243 — Atomic RTF run text replacement

## Contract

This change adds run
`text` to unified set, including omission-preserving character formatting in the
same native transaction. Source-ordered `RunTextReplacement` batches use typed
paragraph/run positions, preflight all targets and budgets before staging, and
publish one candidate. One text-plus-format target consumes one native operation.

Empty text removes the selected semantic run while preserving its paragraph.
Unicode encoding may fragment a replacement into multiple native runs; semantic
text and formatting coverage define preservation, not a permanent run identity.
Tabs and hex-encoded literal CR/VT/NUL remain run text. LF (including CRLF)
would insert an RTF paragraph break and remains explicitly refused here. Resource edits and run structural verbs
remain follow-ups.

## Preservation and admission

The native replacement reuses the proven run source map, retains unrelated
bytes and wraps canonical replacement encoding in a group with `\uc1`. This
preserves inherited formatting and the source Unicode fallback policy outside
the replacement. Candidate readback compares paragraph properties, complete text
and character-format coverage even when the parser fragments Unicode sequences.
Replacing the final unterminated paragraph's last run with empty text emits an
explicit terminal paragraph delimiter, retaining its paragraph and properties.

Limits and ordered selectors are admitted before staging. Replacement byte
charges accumulate across calls and late failures leave earlier staged work
intact. Changed-source validation runs at publication, so exact no-ops bypass
source-closure restrictions after normal target and budget checks. Source-bound
patches and history preserve exact endpoints; durable new-operation replay
remains refused. The facade bounds fan-out text allocations before cloning
native updates and keeps property/format context on invalid values.

## Verification

The native suite passes 1,513 tests/doctests, full-feature facade 1,727 and
RTF-only facade 592. Seven native run-text tests cover preflight atomicity,
compound formatting, Unicode/escaped/tab text, empty terminal paragraphs,
replacement budgets and exact inverse/history. Two new doctests enforce the
public API and typed selectors. Nine facade tests cover native byte parity,
original-position fan-out, invalid values, source closure, exact restricted-source
no-ops, batching, limits, immutable history and origin-preserving atomic saves.

Warning-denied full/native and minimal Clippy, warning-denied rustdoc, no-feature
compilation and both scoped glossary JSON audits pass. Tooling passes 1,166 tests;
boundaries pass for 66 packages and 268 internal dependencies with 11 existing
debt items. The generated OfficeCLI delta is unchanged. Scoped formatting and
diff checks pass. All nine RTF-filtered frozen harness tests pass. Deterministic parser smoke
passes: 792 query inputs / 748 compiled family plans, 909 path inputs / 435
accepted and 776 value inputs / 627 successful coercions. These smoke programs
exercise the generic parser/engine and do not claim exhaustive fuzz coverage.
Warning-denied standalone benchmark Clippy also passes.
Python bindings are not changed or qualified by this record; full-goal Python
coverage remains outstanding.

## Paired measurements

Three CPU-0-pinned trials each cover twenty cases: 1, 10, 100, 1,000 and
10,000 groups of three mixed runs, either in separate paragraphs or in one dense
paragraph. The last-path route replaces one middle run; the query route replaces
all italic runs. Each replacement applies ASCII text and four formatting facets
in one native operation. Unicode preservation is covered by the integration tests.
Two warm-up and nine measured ABBA rounds yield eighteen samples per route.
Owners, selectors and replacement values are prepared outside timing; staging
and publication are timed. Exact paired bytes, reopened text/formatting and
unchanged source bytes are checked outside timing for every sample.

Ranges below span the three trial medians; durations are microseconds.

| Shape / groups / route | Native µs | Unified µs | Unified/native |
|---|---:|---:|---:|
| Paragraphs / 1 / last | 5.83–5.99 | 6.73–6.90 | 1.13–1.16× |
| Paragraphs / 1 / all | 5.78–5.81 | 8.19–8.27 | 1.41–1.43× |
| Paragraphs / 1,000 / last | 2,034–2,098 | 2,059–2,161 | 1.001–1.045× |
| Paragraphs / 1,000 / all | 2,828–2,886 | 3,267–3,351 | 1.16–1.17× |
| Paragraphs / 10,000 / last | 38,580–39,223 | 38,889–39,692 | 1.000–1.029× |
| Paragraphs / 10,000 / all | 51,537–52,380 | 56,990–58,173 | 1.11–1.12× |
| Dense / 1,000 / all | 2,551–2,604 | 2,901–2,975 | 1.14–1.16× |
| Dense / 10,000 / last | 26,460–26,939 | 26,934–27,822 | 1.010–1.051× |
| Dense / 10,000 / all | 38,369–39,429 | 41,959–42,991 | 1.09–1.11× |

Direct large-document updates remain within the 10% reference. Generic fan-out
exceeds it in several cases; tiny updates retain fixed adapter overhead. From
1,000 to 10,000 groups, unified all-target time grows 17.36–17.45× for separate
paragraphs and 14.45–14.47× for dense runs. The source and candidate cursors are
monotone, but these timings do not prove linear end-to-end scaling. Profiling
parsing/source mapping and generic target resolution, allocation measurement and
repeated-command coalescing remain follow-ups. Full performance acceptance is open.

Receipts: [trial 1](results/0243-rtf-run-text/trial-1.json),
[trial 2](results/0243-rtf-run-text/trial-2.json),
[trial 3](results/0243-rtf-run-text/trial-3.json). Each contains all samples,
source commit/tree, compiler/CPU, binary size/SHA-256 and regeneration commands.
Measured implementation: `ecb72aa1b5`. No binary is committed.

Builds use external copied lockfiles with cached `cc 1.6.0`; repository lockfiles
are unchanged. The owned agent worktree/branch and temporary build artifacts
were removed. The base checkout remains at
`32ad64e3207655b4ba558fc2793d9a7396655cdf`; this branch is rebased onto it.

## Remaining work

Run font/color resources and structural verbs, text containing structural breaks,
broader mixed-operation composition, repeated-command coalescing, durable new-op
replay, allocation/scaling qualification, broader RTF elements, remaining format
and glossary work and Python bindings remain required by the full goal.
