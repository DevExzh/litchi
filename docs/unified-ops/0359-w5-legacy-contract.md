# 0359 — W5: DOC, XLS, RTF and the shared legacy crates on the record 0351 contract

- Status: implemented on `uo2/w5-legacy`, cut from `feat/unified-ops-api` at
  `184dd4d908`, merged with trunk twice (`515facc3c0` at `94ccc434c8`, `99eb713051` at
  `f8899dd2cb`, which carries record 0358a).
- Date: 2026-10-20.
- Input: owner decision [0350](0350-owner-api-debt-decision.md); contract
  [0351](0351-api-contract.md) §2, §3.4, §4.3, §5 (wave W5); worked
  examples 0355/0355a, 0356, 0357.
- Owned: `crates/litchi-{doc,xls,rtf,ograph,ole-common,crypto}/**`; the
  DOC/XLS/RTF hunks of `crates/litchi-ops/**` (format adapters, tests,
  examples), `crates/litchi/examples/rtf_*`, the property-set hunks of
  `crates/litchi-ppt`; `tools/perf-baseline`, `tools/perf-cfb-profile`,
  `tools/native-resave`; the `doc-*`, `xls-minimal`, `rtf-minimal` profiles
  and the new `ole-common`, `crypto`, `ograph` profiles of
  `tools/api_glossary.json`; the W5 line of `docs/API_GLOSSARY.md` §2a;
  `results/0289/fidelity.tsv` (re-blessed, §8).
- `performance_claim`: measured (§7).

## 1. Summary

- **Owners.** DOC `raw::Package` is `raw::Reader` (a cursor reader, like
  PPT's); RTF `raw::Document` is `raw::Tree` (the type itself is renamed,
  and the hidden root `RtfDocument` re-export is gone). `Package`,
  `Document`/`Workbook` stay the immutable owners in the gate.
- **Families.** Every DOC and XLS package-level scoped edit is a family
  `x::{Edit, Patch, Commit, Diagnostics, Error, LimitKind}` reached through
  `Package::{edit_x, edit_x_with_options, apply_x}`; the commit is
  `litchi_core::edit::Commit<Package, Patch>` and publishes a new
  `Package`. DOC: route slip, envelope, captions, command bars, annotation
  bookmarks, subdocuments, property sets, tracked revisions, embedded
  objects (plus the existing `package::edit`/`package::metadata`). XLS:
  sheet visibility, comments, user-name publication, charts, OLE objects,
  pivot views, property sets. The byte-level part editors (`Editor`,
  `RevisionEditor`, `embedded_object::Editor`) are crate-private.
- **Codec-level edits** (no package: ole-common streams, crypto data spaces,
  XLS record codecs, DOC smart tags, OLE controls, DOFR, saved selection,
  the DOC part semantic layers in `x::record`) are `Edit` with `staged()`,
  `Commit = litchi_core::edit::Commit<Snapshot, Patch>` and
  `Patch::is_noop`.
- No public `Transaction`, `rollback`, `is_changed`, `into_snapshot`,
  `Mutable*` or `*_with_limits` name remains in the six crates; the gate
  bans them in all six profiles.
- **Refusals.** XLS `UnsafeEdit(String)` and ograph `UnsafeEdit` are gone:
  `Refused(Refusal)` with static reasons, `PatchConflict { operation,
  reason }` and `Internal(&'static str)`; every message is content-free.
  Every litchi-ops error code is unchanged (§8).
- **Options.** `<noun>_value()` getters; DOC `EditOptions` noun setters
  (`revision_limits`, `transaction_limits`, `protection_policy`).
- `Package::apply`/`apply_metadata` return a `Commit` (DOC, XLS).

Fidelity: every round-trip, CRUD, corpus and no-op test passes with
byte-identical outputs; consumers were migrated in the same commits.

## 2. Commits

| Commit | Scope |
|---|---|
| `d86b3b6d8b` | ole-common, crypto: codec edits on the 0351 shape; limit twins folded |
| `97ebefcf93` | ograph: `Refused`/`PatchConflict`, `chart::Model`, `chart::Edit` |
| `44b35f482e` | XLS: content-free typed refusals replace `UnsafeEdit(String)` |
| `e187766b22` | XLS: options accessors, limits through options, `apply` returns a commit, property-set family |
| `1402afb221` | XLS: codec-level commits |
| `6d5bf2f598`, `bb79eeaa6b` | XLS: sheet visibility, comments, user names, charts (`builder::Chart`), OLE objects, pivot views as package families |
| `163c5692da` | DOC: `raw::Reader`, options accessors, no limit twins |
| `c5e9b976e6` | DOC: six part families (`scoped.rs` `family!`) |
| `9e2bc1a0a1` | DOC: codec-level smart tags, OLE controls, DOFR, saved selection |
| `e4a1b97c69` | DOC: property-set family, shared package commits, authorized replays |
| `b74abc8c93`, `e170ffb046` | DOC: tracked-revision and embedded-object families |
| `90e93585a6` | DOC: body-text commits on the shared triad; source aliases dropped |
| `789100cfa3` | RTF: `raw::Tree`, limits through options, redact commit triad |
| `2eca78b99d`, `032eea8f42` | XLS/ole-common: commit accessors, alias removal, `property_set::Edit` |
| `720041f8bb` | gate registrations |
| `515facc3c0`, `99eb713051` | trunk merges |
| `137461188f`, `c1cbb72cbc`, `94ab9057e7` | XLS-only build fix, probe, glossary line; fidelity re-bless; `cell_values::Limits`, native-resave |
| this record | record, ABBA evidence, bench-matrix parser, typed property-set `into_parts` |

## 3. DOC

- `scoped.rs`: `Base::capture` takes the exact package bytes and the edit
  options the package retains (or explicit ones, constrained by them);
  `Base::publish` republishes under the source's `OpenOptions`, keeps the
  granted options as the new package's edit profile, returns the source
  unchanged for an exact no-op and refuses a changed signed package
  (`Error::Signed`). `apply_x` replays the exact whole-package patch, refuses
  another source (`PatchConflict`) and authorizes a changed replay under
  the receiving package's protection policy, as a staged edit is.
- Part families wrap the crate-private byte editor (generated forwarders;
  verbs that returned an intermediate commit now stage). The semantic
  layer is public as `x::record::{Snapshot, Edit, Commit, Patch, Error}`
  (record `Edit::apply` still takes a staged record edit).
- `smart_tags::Editor` (a mutable wrapper over `Snapshot`) is removed:
  `Snapshot::parse*` + `edit` + `commit` + `Patch::apply` cover it.
- `package::property_set`: `Package::{property_sets, edit_property_sets,
  edit_property_sets_with_options, apply_property_sets}`; the policy
  variants of `Patch::apply/revert` give way to the retained or explicit
  edit options. `package::metadata` gains `edit_metadata_with_options`.
- `tracked_revision::Edit` (`add`, `add_text`, `replace`, `remove`,
  `accept`, `reject`) and `embedded_object::Edit` (lifecycle, metadata,
  storage, presentation/native edits, `add_mtef_equation`) replace
  `RevisionEditor` and `embedded_object::Editor`; `Package::tracked_revisions`
  and `Package::embedded_objects` (read view `embedded_object::Snapshot`)
  are the reads. Aliases (`replace_metadata`, `update`, `source`,
  `transaction`) are removed.
- `body_text::Commit` is the shared triad; `body_text::source::Transaction`
  is `source::Edit`; its alias types and methods are removed.

## 4. XLS, RTF, shared crates

- XLS: the families of §1; `chart::Edit` with the draft in `builder::Chart`;
  `cell_values::Commit::{package, into_package}`, `chart_area::Edit::staged`,
  `cell_values::Limits` re-exported; the `Range` and
  `SourceBackedNumericPlanCommit` aliases are removed.
- RTF: `Package::*_with_limits` and `validation::*_with_limits` removed (the
  `*_with_options` forms take `OpenOptions::from(limits)`);
  `transport::decompress(data, DecompressionLimits)` and
  `tail_append::Edit::publication_plan(_with_context)(limits, ..)` take
  their limits as a required argument; `redact::Commit` is the shared triad.
- ole-common: `object::Edit`, `property_set::Edit` (was `Editor`), the
  stream aliases (`OLEPresentationStream`, `PresentationStream`,
  `OLENativeStream`, `NativeStream`) removed. ograph: `chart::Model`,
  `chart::Edit`, `ItemsMut`. crypto: `spaces::Edit`.

## 5. Gate (§4.3)

Profiles `doc-minimal`, `doc-full`, `xls-minimal`, `rtf-minimal` and the new
`ole-common`, `crypto`, `ograph`: bans `Transaction` (exact, suffix),
`rollback`, `is_changed`, `into_snapshot`, `_with_limits` (suffix),
`UnsafeEdit`, `Mutable` (prefix); `Package` in `litchi_doc::raw`,
`Document` in `litchi_rtf::raw`, `RtfDocument`; the old options getters and
DOC `with_*` setters; `Editor` (DOC root and the part modules, XLS suffix,
ole-common exact, crypto/ograph suffix); `Chart` in `litchi_ograph::chart`.
Families: 12 DOC module paths (the nine package families and their facades)
and 7 XLS. The debt inventory (`tools/api_debt_inventory.py`) of the six
crates went from 554 rows to 44 (§9).

## 6. Fidelity

Unchanged outputs: every crate test, the OfficeCLI corpus, the capability
and registry goldens. The corpus fidelity report (`results/0289/
fidelity.tsv`) changed only in the 20 XLS refusal texts (now
`edit cells refused (kind): reason`); every code is the same (the
position-locked drawing anchor refusal is `UnsupportedContent`, keeping
`unsupported_capability`).

## 7. Performance

`crates/litchi-ops/examples/legacy_w5_probe.rs` times the §5 W5 routes with
only the `litchi_ops` API, so one source builds on the base (trunk
`94ccc434c8`, `git archive`) and the head; release, pinned to CPU 29, 31
rounds per scenario, run A-B-B-A twice (`results/0359/abba.sh`,
`analyze.py`, `summary.json`, second run in `run2/`; binary hashes in
`binaries.sha256`, binaries not committed). Ratio = median of the after
runs / median of the before runs (run 1; run 2 in parentheses).

| Route | Before (µs) | After (µs) | After/before |
|---|---:|---:|---:|
| DOC open + get (1k paragraphs) | 894.1 | 773.9 | 0.866 (0.916) |
| DOC get | 0.2 | 0.2 | 0.986 (0.992) |
| DOC set | 1811.8 | 1696.6 | 0.936 (0.983) |
| DOC batch, 256 sets | 694003.0 | 657679.0 | 0.948 (0.971) |
| XLS open + get (100×10) | 155.5 | 151.3 | 0.973 (0.983) |
| XLS set | 98.0 | 99.6 | 1.016 (1.024) |
| XLS batch, 1k sets | 4458.0 | 4435.9 | 0.995 (1.000) |
| RTF open + get (1k paragraphs) | 206.4 | 202.5 | 0.981 (1.012) |
| RTF set | 329.2 | 322.5 | 0.980 (1.008) |
| RTF add | 556.4 | 544.4 | 0.978 (0.995) |
| RTF batch, 1k sets | 5557.9 | 5506.0 | 0.991 (0.948) |

Every route is within the 1.05 regression budget. XLS set is +1.6 to
+2.4 % (inside the budget, not profiled). DOC open is 8–13 % faster on both
runs; the probe does not isolate why (the wave changed the DOC open path
only by renaming and folding the limit twins), so no improvement is
claimed. The DOC batch uses 256 sets because the DOC body transaction's
default bound is 256 operations (base and head refuse 1,000 alike); its
0.65 s per call is the existing route's cost. Host load 5.7–7.4 (other
agents on other CPUs). `tools/unified_bench_matrix.py` gains `p0359`.

## 8. Gates

Re-run in full on the head after the final trunk merge (`99eb713051`), one
cargo command at a time; the working tree was clean afterwards.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-{doc,xls,rtf,ograph,ole-common,crypto,ppt} -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features doc`, `xls`, `rtf`, `ppt` | 0 each (`xls` after `137461188f`) |
| `cargo test --no-fail-fast --all-features -p` each of the seven crates (doc 1,488, xls 1,871, rtf 1,555, ograph 70, ole-common 310, crypto 88, ppt 1,570) | 0 |
| `cargo test -p litchi-ops --all-features` (lib, doc, 167 integration targets in six chunks; 2,640) | 0 after re-blessing `corpus_fidelity` (§6) |
| `cargo test -p litchi --all-features` (209); `cargo test -p litchi-cli --all-features --test officecli_corpus` | 0; 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test capabilities --test registry`, and with every feature but `encryption` | 0, 0 (no change) |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p` the six crates, `litchi-ops`, `litchi` | 0 |
| `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` (47 profiles, 3,195 bans, 483 families, 71 owners) | 0 |
| `check_crate_boundaries.py`; `non_iwork_gate.py verify` | 0; 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `tools/perf-baseline`: `cargo check --locked --all-targets`; `tools/native-resave`: `cargo check --locked --all-features` | 0; 0 (after `94ab9057e7`) |
| tool unit tests (the five required and `test_check_api_glossary`; 1,296) | 0 |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check`, `check_feature_matrix.py` (31 builds) | 0 each |
| `officecli_capability_delta.py --officecli …` (regenerated) | 0, no change |

## 9. Gaps and follow-ups

Items 1–4 are closed by [0359a](0359a-w5-legacy-review.md) (F1, F2 and F4
implemented; F3 accepted under 0351 §2.1 rule 6).

1. **Record-level engines.** `doc::body_text` (Snapshot, `Edit`,
   `PreparedEdit`, `History`, three-way plans) and `xls::cell_values`
   (Snapshot, `Edit`, `SemanticPatch`, `ThreeWayPlan`, `History`) stay public
   whole-file engines whose commits publish a snapshot; `Package::edit` is
   the package route. Making them crate-private needs routes for their
   composition, history and transfer APIs (follow-up 0359-F1).
2. **Source-backed publishers** keep their own commits: DOC
   `body_text::source::Commit` (`snapshot()`, `write_to`/`save`), XLS
   `cell_values::{SourceBackedCommit, SourceBackedPlanCommit}`,
   `comments::SourceBackedCommit`, `sheet_visibility::SourceBackedCommit`,
   RTF `TailAppendSourceEdit` (0359-F2, the W1b/W2 precedent).
3. `raw::Reader` (DOC, XLS) takes `&mut self` for its cursor reads and
   outputs; RTF `raw::Tree` keeps its `&mut self` model setters for the
   writer (`RtfWriter::write_document`), not reachable from a `Package`
   (0359-F3).
4. `DofrEdit` and `SavedSelectionEdit` keep their module prefix because
   `parts::dofr`/`parts::saved_selection` name every type that way (0359-F4).
5. XLS set is +2 % (inside budget); DOC open is faster but unexplained (§7).
