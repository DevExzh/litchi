# 0362 — W8: the unified `File` on the record 0351 contract

- Status: Part A integration candidate on `uo2/coordinator-w8-part-a`,
  based on `b34cb1b104`. Not accepted: current-source validation, matched
  performance qualification and the final W4/W4b integration remain required.
  The historical worker implementation was based on `f2bbbc8d07`.
- Date: 2026-10-21.
- Input: owner decision [0350](0350-owner-api-debt-decision.md); contract
  [0351](0351-api-contract.md) §2.5 rule 7, §4.3 (W8 row), §9; the W8
  notes of records 0356b, 0358a and 0359a in `PROGRESS.md`.
- Owned: `crates/litchi-ops/src/{transaction,batch,import,merge,refresh}.rs`
  and their submodules, `crates/litchi-ops/src/detect/**`, the DOC refusal
  arms of `crates/litchi-ops/src/error.rs`, `crates/litchi/src/**`,
  `crates/litchi-cli/**`, `crates/pyo3-litchi/**`, every consumer of the
  changed names (ops tests, examples and fuzz targets, facade tests,
  `tools/perf-baseline`), the `ops-*`/`facade-*` profiles and the new
  `ops-all` profile of `tools/api_glossary.json`, the W8 line of
  `docs/API_GLOSSARY.md` §2a.
- `performance_claim`: unqualified; §5 retains historical measurements.
  [0362ha](0362ha-file-performance-qualification.md) defines the replacement
  qualification and records the integration provenance.

## 1. Summary

- **Configuration.** `litchi_ops::EditOptions` (also `litchi::EditOptions`)
  carries the `SemanticLimits` of one `File` transaction (`new`,
  `semantic_limits(limits) -> Self`, `semantic_limits_value()`; `Copy`,
  default = the old default limits). `File::edit_with_options(options, f)`,
  `batch_with_options(items, mode, options)`,
  `import_with_options(target, source, &import::Options, options)`,
  `merge_with_options(&Record, &merge::Options, options)` and
  `refresh_with_options(&refresh::Options, options)` replace the five
  `*_with_limits` forms, which are deleted. `batch::Options` (the
  command-channel envelope of `apply_with`/`batch_with`) holds an
  `EditOptions` set by `edit_options(..)`; its `limits` setter is deleted.
- **Commit accessors.** `transaction::Commit<T>` and `BatchCommit` have no
  public fields. `Commit`: `file()`, `into_file()`, `patch()`,
  `diagnostics()`, `value()`, `into_value()`, `into_parts() -> (File, Patch,
  T)`. `BatchCommit` uses the same `value()`/`into_value()` accessors for its
  per-item outcomes and `into_parts() -> (File, Patch, Vec<ItemResult>)`.
  `Diagnostics` is `Copy`, so `into_parts` leaves it to `diagnostics()`.
  `file()` is the facade's spelling of the triad's `package()` (0351 §2.5
  rule 7).
- **Detection.** Every public raw path, byte and reader detector in
  `detect::*`, including the `_with_tail` variants, takes a required
  `&detect::SourceLimits`. This detection-owned type is available wherever detection is enabled
  and contains a common input ceiling plus optional policies for
  each enabled container family. Unified owner opens derive it from their
  `OpenOptions` through the internal `detection_limits()` bridge. The facade
  detection exports and all raw consumers pass this policy explicitly.
  The pure OPC probes in `detect::ooxml` take the OPC `&SourceLimits`.
  The redundant `detect_ooxml_format` and `detect_ooxml_format_from_bytes`
  aliases and all public detector `*_with_limits` names are deleted.
- **Record parsers.** `batch::from_bytes(bytes, &batch::SourceLimits)`,
  `merge::Data::from_json(source, &SemanticLimits)` and
  `merge::Data::from_csv(source, &import::Options, &SemanticLimits)` take
  their limits as a required argument (0351 §2.5 rule 3); `from_bytes_with_limits`,
  `from_json_with_limits` and `from_csv_with` are merged into them.
  `format::docx::Docx::section_inventory(&section::Limits)` likewise
  (it was the one remaining public `*_with_limits` on the ops surface).
- **Refusal codes.** The DOC `body_text::Refusal` and
  `source_backed::Refusal` gain `kind() -> RefusalKind`; litchi-ops
  classifies every DOC refusal by kind except the arms that are not
  refusals of a well-formed request (§3).
- **Consumers.** litchi-ops tests (169 targets), examples and fuzz targets,
  the facade and its tests, litchi-cli, pyo3-litchi (the Python surface is
  unchanged: `Commit.file`, `.patch`, `.results`, `.value` stay Pythonic
  properties over the new accessors, same codes), `tools/perf-baseline`
  and its standalone `unified-*` harnesses that used the changed names
  (batch-closure, docx-batch, docx-text-fanout, pptx-chunk-publication,
  pptx-command-batch, xlsb-cell-value, xlsb-remove, xlsb-sheet-batch,
  xlsx-batch; each checked against the head, lock files unchanged),
  `docs/OFFICE_API_GUIDE.md`, `crates/litchi/README.md`.
- **Gate.** New `ops-all` profile (litchi-ops with all 30 features; the
  build budget allowed it, 0351 §9) bans the `_with_limits` suffix on the
  whole ops surface plus the merged names (`from_json_with_limits`,
  `from_csv_with`, the `*_with_limits_and_tail` detectors, the two
  `detect_ooxml_*` aliases, `batch::Options::limits`); `ops-rtf`,
  `ops-odf`, `ops-iwork`, `facade-iwork` and `facade-sheet-text` gain the
  suffix ban.

## 2. Decisions and deviations

1. **Detection correction.** The original implementation took
   `&OpenOptions` and left some raw classifiers without a policy. Review
   returned that deviation. The current candidate applies the detection
   portion of `84a64e8285`: its own feature-independent `detect::SourceLimits`
   satisfies the required raw API without a feature-dependent signature.
2. **`BatchCommit::value`.** A batch commit's value is its per-item results;
   `value()`/`into_value()` therefore follow the same canonical accessor shape
   as `transaction::Commit<T>`. The CLI JSON and Python `Commit.results`
   property retain their semantic `results` spelling over this Rust value;
   there is no Rust alias.
3. **`into_parts` returns the value.** The core triad's
   `into_parts() -> (P, T)` has no closure value; the `File` commit returns
   `(File, Patch, T)` so a caller can take every result without cloning, as
   the old destructurable fields allowed.
4. **W4 consumers.** Ops XLSX/XLSB unit and integration tests migrate the
   commit accessors, `_with_options` names and required detector policies in
   this same candidate. The W4/W4b implementation integration is still
   pending; its final shared-surface validation remains required.

## 3. Error codes across formats

The `Classify` arms in ops `error.rs` were reviewed for arms that map a
refusal by message or by a format-specific variant. None maps by message.
Leftovers by variant:

- **DOC** (`body_text::Error::Refused`, `source_backed::Error::Refused`):
  each variant had its own code. Both refusal enums now have `kind()`, and
  the default arm is `refusal.kind().classify()`. The arms kept explicit
  are not refusals of a well-formed request, and their codes are what the
  other formats report for the same condition: a missing position
  (`ParagraphNotFound`, `TargetNotFound`, `SectionNotFound`) is
  `out_of_range`, as DOCX `OutOfBounds` and the RTF positional errors are;
  `ResourceCollision` is `duplicate_name`; the budgets `OperationLimit`/
  `ReplacementLimit` are `limit_exceeded`; the typed `SoleParagraph` refusal
  remains `invalid_argument`, while the PPTX `FinalSlide` refusal is a
  separate shared `InvalidValue` kind and reports `invalid_value`; the
  protected kinds keep their specific suggestions. Visible code change:
  refusals whose kind is `UnsupportedTarget` (`EmptyParagraph`)
  report `unsupported_type` (was `unsupported_capability`), as the DOCX and
  PPTX `UnsupportedTarget` refusals do. The OfficeCLI corpus and the CRUD
  checklist are unchanged by it.
- **RTF** (`edit::Error`): unsupported well-formed requests now use
  `Refused(Refusal)` and classify through `RefusalKind`, alongside the typed
  limit, positional, parse, publication and `PatchConflict` variants.
  Protected requests retain `RefusalKind::Protected` and therefore report
  `document_protected`; non-refusal validation failures keep their specific
  codes. The same refusal shape is used by RTF tail-append edits.
- **PPTX** `SlideOrderRefusal::AmbiguousBindings` → `corrupt_file` remains the
  documented exception; `SlideRemovalRefusal::FinalSlide` now follows its
  shared `InvalidValue` kind as `invalid_value`. XLSX/XLSB arms are W4's.

## 4. Fidelity

The accessors return the former fields, and `EditOptions::default()`
preserves the former default limits. The historical worker reported unchanged
capability and registry goldens, corpus and CRUD checks, and 301 passing Python
tests on its earlier source. Those reports do not validate this assembled
candidate; fresh validation is recorded separately in 0362ha.

## 5. Historical performance (not acceptance)

`crates/litchi-ops/examples/file_w8_probe.rs` times the `File` routes of
0351 §5 for eleven formats with API that is the same on the base and the
head (commits are dropped, never read): open+get, one `set`, a 1,000-item
atomic batch (DOC: 256), import (workbooks) and merge. One source built on
the base (trunk `f2bbbc8d07`, `git archive`) and the head; release, pinned
to CPU 29, 11 rounds per scenario, A-B-B-A twice
([`results/0362/abba.sh`](results/0362/abba.sh), `analyze.py`,
`summary.json`, `run1/`, `run2/`; binary hashes in `binaries.sha256`,
binaries not committed). Each cell: before µs, after µs, after/before
(median of the after legs / median of the before legs); run 1 | run 2.

| Route | Run 1 | Run 2 |
|---|---|---|
| docx/open_get | 972.2 978.0 1.006 | 965.4 976.8 1.012 |
| docx/set | 66.7 65.8 0.987 | 67.5 65.4 0.968 |
| docx/batch1k | 5491.2 5459.8 0.994 | 5492.4 5512.3 1.004 |
| docx/merge | 846.9 865.5 1.022 | 839.2 836.7 0.997 |
| odt/open_get | 141.0 140.3 0.995 | 142.0 139.9 0.985 |
| odt/set | 424.2 422.8 0.997 | 420.9 422.5 1.004 |
| odt/batch1k | 1650.9 1641.4 0.994 | 1642.4 1644.6 1.001 |
| odt/merge | 469.6 473.4 1.008 | 463.7 472.3 1.018 |
| doc/open_get | 759.4 771.6 1.016 | 778.8 762.4 0.979 |
| doc/set | 1592.8 1526.1 0.958 | 1642.6 1616.5 0.984 |
| doc/batch1k | 297869.5 289325.5 0.971 | 300593.1 295342.7 0.983 |
| doc/merge | 3022.9 3296.0 1.090 | 2994.3 3007.4 1.004 |
| rtf/open_get | 201.9 209.4 1.037 | 203.8 205.6 1.009 |
| rtf/set | 323.2 320.1 0.990 | 328.6 329.0 1.001 |
| rtf/batch1k | 5631.1 5622.5 0.998 | 5668.9 5648.1 0.996 |
| rtf/merge | 706.5 711.1 1.006 | 698.6 718.1 1.028 |
| xlsx/open_get | 481.8 471.0 0.977 | 480.5 468.8 0.976 |
| xlsx/set | 1139.4 1127.4 0.989 | 1143.6 1149.4 1.005 |
| xlsx/batch1k | 2599.2 2529.2 0.973 | 2596.2 2552.2 0.983 |
| xlsx/import | 7792.3 7740.5 0.993 | 7752.5 7766.8 1.002 |
| xlsx/merge | 1162.6 1148.2 0.988 | 1146.2 1146.7 1.000 |
| xlsb/open_get | 154.9 159.6 1.030 | 158.7 157.7 0.993 |
| xlsb/set | 1370.7 1395.6 1.018 | 1388.3 1393.7 1.004 |
| xlsb/batch1k | 2244.9 2274.8 1.013 | 2252.7 2269.8 1.008 |
| xlsb/import | 6446.7 6368.8 0.988 | 6473.8 6387.8 0.987 |
| xlsb/merge | 1443.9 1454.8 1.008 | 1445.6 1451.7 1.004 |
| xls/open_get | 142.0 139.2 0.980 | 141.5 138.4 0.978 |
| xls/set | 369.8 368.9 0.998 | 364.7 369.2 1.012 |
| xls/batch1k | 4430.7 4195.9 0.947 | 4332.4 4230.6 0.976 |
| xls/import | 339914.6 337786.3 0.994 | 339588.8 338313.2 0.996 |
| xls/merge | 422.4 420.8 0.996 | 416.0 421.9 1.014 |
| ods/open_get | 47.2 47.8 1.012 | 47.6 48.0 1.009 |
| ods/set | 408.0 411.5 1.009 | 408.1 413.2 1.012 |
| ods/batch1k | 1211.0 1226.3 1.013 | 1224.7 1241.2 1.013 |
| ods/import | 23093.9 23553.3 1.020 | 23338.3 23354.1 1.001 |
| ods/merge | 387.0 398.1 1.029 | 389.4 394.7 1.014 |
| pptx/open_get | 436.7 454.0 1.040 | 439.3 449.3 1.023 |
| pptx/set | 1925.9 1961.9 1.019 | 1919.5 1948.3 1.015 |
| pptx/batch1k | 12206.0 12488.5 1.023 | 12264.2 12419.6 1.013 |
| pptx/merge | 1831.9 1864.9 1.018 | 1829.9 1859.4 1.016 |
| ppt/open_get | 116.1 119.2 1.027 | 117.0 118.7 1.014 |
| ppt/set | 116.1 116.9 1.007 | 117.6 115.6 0.983 |
| ppt/batch1k | 1679.4 1666.5 0.992 | 1691.0 1669.0 0.987 |
| ppt/merge | 144.6 147.1 1.017 | 147.0 145.6 0.991 |
| odp/open_get | 280.8 280.9 1.000 | 281.2 281.3 1.000 |
| odp/set | 750.4 753.2 1.004 | 754.2 757.2 1.004 |
| odp/batch1k | 1537.5 1528.2 0.994 | 1550.9 1553.9 1.002 |
| odp/merge | 780.9 781.8 1.001 | 785.8 781.0 0.994 |
These historical values are retained without an acceptance verdict. Review
returned the blanket 1.05 threshold and the earlier code-placement explanation:
neither establishes preservation of this candidate's performance. The
`doc/merge` rerun (`results/0362/run3-doc-merge/`) is also historical and does
not replace the adverse first run. The corrected 0362ha protocol requires
matched leaf/facade workloads, source and dependency provenance, guarded
native measurements, and separate allocation evidence. Until those pass,
W8 performance remains unqualified.

Finding (not caused by this wave, same on both binaries): XLS import is
superlinear, 100 rows take 0.34 s and 1,000 rows 13.6 s, so the probe
times XLS import on 100 rows. Follow-up for the XLS owner.

## 6. Gaps and follow-ups (Part B)

- Patch replay by content identity for PPTX `source::cross_slide_copy`,
  ODP `tail_append`, XLS and RTF source-backed patches (the coordinator
  decision), with tests per format. This source work is integrated in
  `0362m`, `0362u`, `f965b0f9c1` and `0641d926aa`; the historical checklist
  remains only for fresh qualification evidence.
- The inventory re-run and final glossary/qualification cells remain. RTF
  `edit::Error` classification and source-backed replay are integrated
  (`c48e382321`, `e66b418b46`, `cbf4e30276`); only fresh profile/native
  validation remains. DOCX `apply_durable_ink` already returns its family
  `Commit` (`3c1d25de24`); its native/allocation qualification remains open.
  Design §16.4 is already closed by the implemented `litchi-sheet-text` leaf
  and `litchi::sheet::text` re-export; it is not a remaining W8 task.
- PPT first-commit cost is resolved by accepted record 0362j, already in
  this candidate base; its measured production source is unchanged here.
- After W4/W4b merge: rerun the gate's `ops-all` profile against their
  surface.

## 7. Gates

Historical worker reports for `f2bbbc8d07` plus the original Part A changes
follow. They are retained for provenance, not claimed as fresh checks of the
assembled candidate. Current-source receipts belong in 0362ha.

| Gate | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-ops -p litchi -p litchi-py -p litchi-cli -p litchi-doc --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features [--features f] --all-targets -- -D warnings`, f = none and each of docx, pptx, xlsx, xlsb, doc, xls, ppt, rtf, odt, ods, odp, odg, odc, odi, odb, oth, odm, odf-formula, pages, keynote, numbers | 0 each |
| `cargo clippy -p litchi --no-default-features --features f --all-targets -- -D warnings`, f = docx, ods, rtf, pages | 0 each |
| `cargo test -p litchi-ops --all-features --lib` (953) and `--doc` (216) | 0, 0 |
| `cargo test -p litchi-ops --all-features --test …`, all 169 targets in six chunks of 30 (executables deleted per chunk); the DOC, CRUD and corpus-fidelity targets rerun after the DOC classification change | 0 each |
| `cargo test -p litchi-doc --all-features` | 0 |
| `cargo test -p litchi --all-features`, `cargo test -p litchi-cli --all-features` (incl. `officecli_corpus`, unchanged) | 0, 0 |
| capability/registry goldens, all features and all-but-`encryption` (`--test capabilities --test registry`, no re-bless) | 0, 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-ops -p litchi -p litchi-doc -p litchi-cli` | 0 |
| `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` | 0: 49 profiles, 3,279 banned names, 498 families, 78 immutable owners, family methods in 30 profiles |
| `crates/pyo3-litchi/tests/run_tests.sh` (301) | 0 |
| `tools/perf-baseline`: `cargo check --locked --all-targets`; the nine changed `unified-*` harnesses `cargo check --all-targets` (lock files restored) | 0; 0 each |
| `crates/litchi-ops/fuzz`: `cargo check --features apply --bins` | 0 |
| tool unit tests (the five required modules, `test_check_api_glossary`, `test_unified_bench_matrix`) | 0 |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check` (gains `p0362`), `check_feature_matrix.py` (31 ok) | 0 each |
| `check_crate_boundaries.py`, `non_iwork_gate.py verify` | 0, 0 |
