# 0178 — XLSB catalog validation indexes

## Change

Record 0177 exposed superlinear sheet-visibility fanout. Two catalog checks
compared every sheet with every preceding sheet: native workbook-name
admission and eager/source physical-target admission. Each structural
publication repeats these checks while validating its workbook candidate.

Both checks now use private insertion-only hash indexes. Workbook names use
exactly the previous flattened Unicode `char::to_lowercase` sequence, including
multi-character expansions and without normalization or full case folding.
Physical targets use ASCII lowercase, exactly matching
`PackURI::is_equivalent_to`. Names, relationships and tab states remain in their
original ordered vectors; the indexes never determine output order. Duplicate
errors retain their types, messages and position before catalog insertion.
Index capacity and name-key capacity are admitted with typed allocation errors.
No public API, property registry, capability or OfficeCLI behavior changes.

The regressions cover exact and ASCII-case-variant duplicate physical targets
for both eager and source owners, valid case-variant targets, Unicode lowercase
expansion (`İ` versus `i` plus combining dot), and distinct canonically
equivalent names (`é` versus `e` plus combining acute).

## Paired evidence

The unchanged `tools/perf-baseline/unified-xlsb-visibility` harness ran on CPU 0,
without concurrent build/test jobs, with nine ABBA rounds per case. The
[before](results/0178-xlsb-catalog-before.json),
[after](results/0178-xlsb-catalog-after.json), and
[repeat](results/0178-xlsb-catalog-repeat.json) matrices each cover eager/source
owners and 1, 10, 100 and 1,000 tabs. All native/unified output comparisons and
property-alias comparisons are exact byte matches, with no failures. Evidence
records retain raw samples, source commit/file hashes, corpus identities,
compiler, build command and binary SHA-256; compiled binaries are not committed.

| Owner, 1,000 targets | Before native / unified ms | After native / unified ms | Repeat native / unified ms |
|---|---:|---:|---:|
| Eager | 312.73 / 389.53 | 143.45 / 142.98 | 141.70 / 141.98 |
| Source | 400.45 / 480.30 | 175.39 / 176.86 | 173.68 / 176.45 |

Native medians improve 54–57%, and unified medians improve about 63% in this
workload. After/repeat unified-to-native ratios at 1,000 targets are
0.997/1.002 eager and 1.008/1.016 source. One-tab ratios are 1.065/1.061 eager
and 1.041/1.042 source, below the measured 10% review threshold.

This does not establish linear scaling: unified 100→1,000 growth remains
18.94–19.00× eager and 17.27–17.32× source. Further publication profiling,
10,000-tab evidence and separate sheet-command coalescing remain required.
The comparisons describe one query fanout, not a batch of 1,000 facade commands.

## Verification

Verification passed: 1,161 native XLSB tests/doctests with
`--no-default-features --features sign,vba-inspection`; 1,218 facade
tests/doctests with `docx,xlsx,pptx,xlsb,encryption`; warning-denied native
all-target Clippy; warning-denied native rustdoc; scoped rustfmt; the repository crate-boundary audit; and the generated OfficeCLI delta check.
Cargo used an external copied root lockfile with cached `cc` 1.6.0 through
`resolver.lockfile-path`. The root and shared performance lockfiles are unchanged.
The implementation and independent review were delegated in parallel; the
implementer used an isolated worktree. The base branch has no commits missing
from this branch. The full unified-operations goal remains active.
