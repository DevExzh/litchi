# 0174 — Reuse the admitted XLSB cell publication projection

## Change

Cell-value publication already parses and validates its candidate workbook.
It now retains that projection privately in the resulting `Package`, so the
facade's consuming `Package::into_workbook()` can reuse it. The public
`Commit::{package, patch, diagnostics, into_package, into_parts}` contract is
unchanged, following ADR 0034's package-owned publication vocabulary.

Publication still serializes through the bounded writer, physically admits
the exact output bytes under the source limits, and replaces the candidate
workbook's OPC graph with that admitted graph before retaining the projection.
The package's bytes and the returned workbook therefore share the exact
`patch.after()` image. Native worksheet validation, signature checks,
shared-string closure and retained-patch budgets still run.

Only this cell-publication constructor retains the optional `Arc<Workbook>`.
Ordinary workbook conversion and other package constructors remain uncached;
other mutations construct a fresh package without carrying a stale projection.
A unique consuming projection uses `Arc::try_unwrap`; shared projections fall
back to the existing parser. Borrowed `Package::workbook()` still reparses.
The retained workbook includes its already bounded worksheet cache. A
package-only caller can retain more semantic state than before until it drops
or consumes the package; this change makes no allocation or RSS claim.

Five native regressions cover exact output bytes, cloned packages and a later
structural rename, exact no-ops, tolerant cached-number policy, source limits
and skipped-drawing policy. The full native suite passed 1,133 tests/doctests.
Independent review audited all package constructors and mutation paths and
found no stale-cache or policy-retention issue.

## Validation and measurements

The combined `docx,xlsx,pptx,xlsb,encryption` facade suite passed 1,194
tests/doctests. Native all-target Clippy and native rustdoc JSON passed with
warnings denied; the full XLSB glossary profile passed all 26 registered name
bans. Scoped formatting and diff checks passed. Builds used offline resolution,
external targets and an external root-lock copy selecting cached `cc` 1.6.0;
both inherited repository lockfiles remain unchanged.

[Before](results/0174-xlsb-cell-values-before.json),
[after](results/0174-xlsb-cell-values-after.json) and
[repeat](results/0174-xlsb-cell-values-repeat.json) each contain nine ABBA
rounds, 18 samples per leg, CPU 0, source-file hash manifests and binary/build
provenance. No Cargo build or test job ran during timing. All 36 case results
have exact warm-up ZIP-byte equality, empty failure matrices and one native
commit per case. The unchanged harness compares eager native `Package` with
eager facade `Workbook`; both source routes include promotion. Caller input
preparation and serialization are outside the clock, while facade admission
and property coercion are inside it.

| Shape | Cells | Owner | Before unified/native | After | Repeat |
|---|---:|---|---:|---:|---:|
| set | 1 | eager | 1.104 | 1.067 | 1.065 |
| edit | 1 | eager | 1.114 | 1.060 | 1.070 |
| set | 1 | source | 1.038 | 1.021 | 1.001 |
| edit | 1 | source | 1.042 | 1.021 | 1.039 |
| edit | 10 | eager | 1.182 | 1.103 | 1.125 |
| edit | 10 | source | 1.065 | 1.045 | 1.047 |
| edit | 100 | eager | 1.318 | 1.339 | 1.287 |
| edit | 100 | source | 1.198 | 1.167 | 1.170 |
| edit | 1,000 | eager | 1.579 | 1.594 | 1.607 |
| edit | 1,000 | source | 1.497 | 1.475 | 1.505 |
| edit | 10,000 | eager | 1.623 | 1.624 | 1.646 |
| edit | 10,000 | source | 1.621 | 1.609 | 1.634 |

The measured one-cell scenarios are below the 10% overhead threshold in both
after runs. Eager one-shot native/unified medians are 106.140/113.280
microseconds, then 106.581/113.541; one-operation transactions are
105.921/112.261, then 105.301/112.700. Against the before run, every native
median changes by less than 5% (largest increase 2.34% after, 3.73% repeat).
The observed ratio improvement is therefore accompanied by lower absolute
eager unified one-cell timings, rather than a large slowdown of the native
comparison.

The 1k/10k ratios remain below 2x. Unified 10k/1k growth is 9.94x/9.19x
(eager/source) after and 9.93x/9.09x in the repeat, consistent with approximately
linear scaling for this workload. This closes the specific measured small-edit
follow-up from 0172–0173; it does not establish full corpus, memory or all-format
performance acceptance.

Reproduce using `tools/perf-baseline/unified-xlsb-cell-value/README.md` and
`taskset -c 0 .../unified-xlsb-cell-value-baseline --rounds 9`, with build and
machine provenance environment values recorded in each JSON file.


## Remaining scope

Pre-edit worksheet selection still calls `Package::workbook()`, and cloned
publication packages still use the parse fallback. The generated numeric
benchmark does not cover formula-group-heavy or shared-string-heavy scaling.
XLSB formula authoring, formatting, structural facade CRUD, creation and durable
cell history remain open, as do the other format and Python requirements of
the full goal.


The branch already contained the rewritten base at turn start. Temporary
implementation worktrees and branches were removed after integration. Scratch
builds, copied locks, patches and logs were removed after recording the reduced
evidence. No binary or large evidence artifact is committed.
