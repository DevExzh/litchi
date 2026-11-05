# 0242 — Atomic RTF run character formatting

## Contract

The unified RTF run setter delegates compound character updates to the native
primary transaction. `RunFormatPatch` carries omission-preserving local facets;
`RunFormatUpdate` selects a source paragraph and run with two typed `Position`
values. The complete ordered batch is admitted before staging and published
once. One selected run consumes one native operation even when multiple facets
change. Exact no-ops preserve the immutable source artifact.

Unified paths and queries can set bold, italic, exact underline style, size,
strike, double strike, hidden text, all capitals, small capitals and outline.
The static run registry supplies coercion and capability metadata. Run text,
font/color resource editing and structural run verbs remain explicit follow-ups.
Run edits participate in unified immutable results, history, atomic and
best-effort batches, and origin-preserving atomic saves.

## Preservation and admission

The native editor maps ordinary body runs to retained lexical source ranges,
inserts changed and restoring character controls, and reparses one candidate.
It validates every run's text and complete local character format, paragraph
properties and opaque preservation before publication. Metadata, header bytes,
resource tables, unrelated character facets and unselected runs remain intact.
Unicode fallback widths, UTF-16 surrogate pairs, escaped braces/backslashes,
hexadecimal text and control-symbol text are covered. Surrogate fallback debt
is reset after the high surrogate's fallback has been consumed.

Changed edits refuse compressed transport, dependent body destinations, review
content, inline line breaks and paragraph style/list references. Mixed native
structural/property work is refused; facade sequential commands publish through
its existing transaction contract. Invalid late batch selectors never stage
partial work. Source, semantic operation, expanded-target, staging and snapshot
limits remain enforced. Exact no-ops preserve protected source bytes. Native
source-bound patches and history retain exact endpoints; durable replay of the
new semantic operation is explicitly refused.

Font size uses the shared font-size coercion policy and must be a positive exact
half-point: `12.5pt` succeeds and `12.25pt` fails without rounding. Coercion errors
retain property and format context. All nineteen underline tokens round-trip.

## Verification

The native suite passes 1,504 tests/doctests, the full-feature facade passes
1,718 and the RTF-only facade passes 583. Five new native integration tests and
two API doctests cover typed selection, atomic admission, omission preservation,
exact no-ops, source-bound inverse/history and mixed-operation refusal. Thirteen
new unified tests cover compound/fan-out updates, coercion, all underline styles,
limits, batches, source policy, Unicode, omitted facets and atomic saves.

Warning-denied native/full-facade and RTF-only Clippy, standalone benchmark Clippy, warning-denied rustdoc,
no-feature compilation, scoped glossary JSON audits and scoped formatting/diff
checks pass. Tooling passes 1,166 tests; boundaries pass for 66 packages and 268
internal dependencies with 11 existing debt items. The generated OfficeCLI delta
is unchanged. All nine RTF-filtered frozen performance-harness tests pass.
After adding the explicit run source-closure capability gap, both run integration
suites pass again. No Python binding or parser-fuzz implementation changed;
Python and fuzz qualification of the full goal remains outstanding.

## Paired measurements

Three CPU-0-pinned trials cover 1, 10, 100, 1,000 and 10,000 groups of
three mixed-format runs. Groups occupy separate paragraphs or one dense
paragraph. Each case updates the last group's middle run or all italic runs,
with four facets in one native update per target. Two warm-up and nine ABBA
rounds yield eighteen samples per route. Prepared owners, selectors and property
values are outside timing; staging and publication are timed. Exact paired
bytes, reopened text/format and source immutability are checked outside timing.

Ranges below are across the three trial medians; durations are microseconds.

| Shape / groups / route | Native µs | Unified µs | Unified/native |
|---|---:|---:|---:|
| Paragraphs / 1 / last | 5.90–5.98 | 6.64–6.79 | 1.11–1.15× |
| Paragraphs / 1 / all | 6.02–6.04 | 8.41–8.49 | 1.39–1.41× |
| Paragraphs / 1,000 / last | 1,993–2,052 | 1,984–2,045 | 0.995–0.996× |
| Paragraphs / 1,000 / all | 2,831–2,883 | 3,277–3,322 | 1.15–1.16× |
| Paragraphs / 10,000 / last | 37,864–38,413 | 37,544–38,645 | 0.992–1.009× |
| Paragraphs / 10,000 / all | 47,849–48,542 | 54,700–55,097 | 1.13–1.15× |
| Dense / 1,000 / all | 2,477–2,535 | 2,827–2,900 | 1.13–1.15× |
| Dense / 10,000 / last | 23,346–23,744 | 23,145–23,805 | 0.991–1.007× |
| Dense / 10,000 / all | 35,201–35,996 | 36,935–37,081 | 1.03–1.05× |

Direct large-document updates stay near the native route. Generic query fan-out
exceeds the 10% reference budget in several cases, and tiny updates retain fixed
adapter overhead. From 1,000 to 10,000 groups, unified all-target time grows
16.5–16.7× for separate paragraphs and 12.8–13.1× for dense runs. Monotone native
selection/publication cursors avoid repeated per-target scans, but these timings
do not establish linear end-to-end scaling. Follow-ups are profiling retained
parsing/source mapping and generic predicate/target resolution, allocation
measurement and repeated-command coalescing. Full performance acceptance is open.

Receipts: [trial 1](results/0242-rtf-run-format/trial-1.json),
[trial 2](results/0242-rtf-run-format/trial-2.json),
[trial 3](results/0242-rtf-run-format/trial-3.json). They retain all samples,
compiler/CPU, source commit/tree, binary size/SHA-256 and regeneration commands.
Measured implementation: `27287f69ea`. No binary is committed.

Builds use external copied lockfiles with cached `cc 1.6.0`; the three repository
lockfiles are unchanged. The owned worktrees/branches and temporary build
artifacts were removed. The base checkout remains at
`32ad64e3207655b4ba558fc2793d9a7396655cdf`; this branch is rebased onto it.

## Remaining work

Run text/resource/structural writes, mixed text/layout/run command composition,
repeated-command coalescing, durable new-operation replay, broader RTF elements,
allocation and performance qualification remain. Other delayed formats, the
remaining glossary migration and Python bindings are still required by the
full goal.
