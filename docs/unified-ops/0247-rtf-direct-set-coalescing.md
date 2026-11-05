# 0247 — RTF direct-set coalescing

## Contract

Consecutive compatible direct paragraph and run setters share existing native
ordered batch APIs. Publication, source preservation, candidate validation and
history remain native-owned. The facade retains bounded planning state and
source-relative selectors, without a new public transaction family. Its temporary
index is independently bounded by the staging profile; when the index cannot
fit, the existing scalar resolver handles the mutation without narrowing its
admission. Authored payload/actions retain their separate staging charge.
A full staging/target window publishes before the next individually admissible
mutation; operation-call and cumulative snapshot limits remain transaction-wide.

Reads, queries, raw paths, structural operations, repeated or descending targets,
and incompatible native families close pending work before resolving the next
operation. Text that can change paragraph/run axes requires a boundary so later
paths retain sequential semantics. Stable nonempty printable ASCII run text
coalesces within a paragraph; empty, escaped/control or Unicode text conservatively
closes before later same-paragraph paths. Changed paragraph alignment and text
are separate boundaries because native text editing can refuse the resulting
mixed-format source. Even no-op alignment needs a boundary before structural
paragraph text, since native staged operations conflict. Text-only and format-only
run families currently remain distinct. Single-request native conflicts remain typed
refusals. Atomic failure attribution and best-effort isolation remain required.

The batch item-count guard now flushes pending work before reporting a later
limit violation, preserving the earlier native failure's originating item.

## Verification

At implementation commit `05ea532864`, the full facade passes **1,764**
tests/doctests and the RTF-only facade passes **629**, including **18** focused
coalescing regressions. Both all-target Clippy profiles and rustdoc pass with
warnings denied; the no-feature check, RTF facade glossary JSON gate, three
parser smoke targets, scoped formatting and OfficeCLI generated delta check pass.
Crate boundaries remain **66** packages, **268** dependencies and **11** tracked
debts. The frozen RTF performance harness passes all nine tests. Repository
lockfile hashes remain unchanged.

The focused cases cover eager/path owners, native bulk bytes and sequential
semantics, homogeneous setters, family/read/query/raw/structural boundaries,
Unicode/empty-run retargeting, per-item failure ordering, atomic/best-effort
isolation, no-ops/inverse/source preservation, staging rollover and cache fallback.
Native formatting bulk publication differs from repeated scalar formatting in
redundant control prefixes; it is checked byte-for-byte against the native bulk
oracle and semantically against scalar commands.

## Paired measurements

The standalone `unified-rtf-batch-set-baseline` harness compares native ordered
bulk calls with a unified direct-set loop at 1, 10, 100, 1,000 and 10,000 targets.
It covers paragraph text, paragraph alignment/layout, run text, and compound run
formatting, including separate paragraphs and dense runs. Prepared selectors,
properties and native update vectors stay outside timing. Native operation
counts, one unified publication, exact bytes, reopened semantics and source
immutability are checked outside timing. Two warmups and nine ABBA rounds yield
18 samples per route. Three CPU-pinned trials pass all **90** case-level output
checks, with one unified publication per case:
[trial 1](results/0247-rtf-direct-set-coalescing/trial-1.json),
[trial 2](results/0247-rtf-direct-set-coalescing/trial-2.json),
[trial 3](results/0247-rtf-direct-set-coalescing/trial-3.json).

All 18 measured 1,000-operation cases meet the 2× facade/native target
(ratios **1.015–1.542**). Ten of the 18 single-operation cases exceed the
10% overhead target; their full ratio range is **1.076–1.257**. The
10,000-operation results are:

| Family / shape | Native median, ms | Unified/native ratio | Native growth, 1k → 10k |
|---|---:|---:|---:|
| Paragraph text | 895.5–935.9 | 0.958–1.003 | 108.4–113.3× |
| Paragraph alignment/layout | 23.44–23.65 | 1.361–1.372 | 16.1–16.3× |
| Run text / separate paragraphs | 45.93–46.66 | 1.026–1.039 | 17.1–17.7× |
| Run text / dense runs | 34.54–34.89 | 1.062–1.066 | 15.0–15.2× |
| Compound run / separate paragraphs | 58.51–59.13 | 1.081–1.088 | 15.5–15.8× |
| Compound run / dense runs | 41.79–42.36 | 1.166–1.168 | 13.1–13.2× |

The paragraph-text native path is nearly quadratic over this interval; the
linear-scaling requirement remains open. Profile and fix that native batch
path next, then remeasure both legs. Other families also need scaling
qualification. These are local three-trial measurements, not full-corpus
acceptance or a before/after regression claim.

Receipts bind the measurements to implementation `05ea532864`, its source tree,
binary size/SHA-256, compiler and machine. Reproduce with the build/run commands
in each receipt and the [harness README](../../tools/perf-baseline/unified-rtf-batch-set/README.md).
The local offline build used an external copied lockfile with cached `cc 1.6.0`;
repository lockfiles were unchanged. No binary is committed.

First direct-set planning still indexes the complete native projection, and
text can be copied into both the plan and native staging. The logical bounds
are not heap-allocation measurements; allocation qualification remains open.

## Tooling catalog repair

Expanded tooling discovery found stale selector pins after the earlier PPT
aggregate-v2 migration. The current v2 catalog now accounts for all **633** live
selectors, including the 28 PPT aggregate routes; historical v1 evidence is
unchanged. Both affected tooling modules pass **78** tests, and the complete
tooling discovery passes **2,045** tests with **20** skips. Three LibreOffice registry tests
require the base checkout's external filter sources; a temporary read-only
symlink supplied those inputs and was removed after verification.

## Remaining work

This closes compatible direct-set publication batching, not the full goal.
Broader RTF CRUD, run move/swap, other operation-family coalescing, durable run
replay, allocation/scaling qualification, ODF/iWork coverage, Python, glossary
migration and complete acceptance evidence remain required.

## Cleanup

All owned implementation worktrees and branches were removed after integration.
The external build target, copied lockfiles, measurement scripts and temporary
logs were removed after the reduced receipts were recorded.
