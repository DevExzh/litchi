# 0359a — W5 review: no public engine, deferred source-backed owners

- Status: implemented on `uo2/w5-legacy` after record
  [0359](0359-w5-legacy-contract.md) (tip `7afba5f146`, accepted), merged
  with trunk once (`2f3c75bf8b` at `c3abe0ecc4`, which carries the 0359
  merge and record 0361).
- Date: 2026-10-20.
- Input: the coordinator's review of 0359 (F1–F4, the DOC 1k batch); the
  source-backed family shape of [0355a](0355a-w1-docx-gaps.md) §1 and its
  ODF twin [0358a](0358a-w6-odf-review.md) §2; contract
  [0351](0351-api-contract.md) §2.1 rule 6.
- Owned: as 0359, plus `crates/litchi-ops/examples/legacy_w5_source_probe.rs`,
  the DOC body-run hunk of `crates/litchi-ops/src/format/doc/{edit,write}.rs`
  and `docs/unified-ops/results/0359a/`.
- `performance_claim`: measured (§7); every changed source-backed route
  writes byte-identical output.

## 1. Summary

| Review item | Done |
|---|---|
| F1 — no public snapshot-publishing engine | `doc::body_text::{Snapshot, Edit, Commit, Patch, Composition, ThreeWayPlan, History, PreparedEdit, OpenOptions}` and `xls::cell_values::{Edit, Commit, Patch, History}` (and the snapshot constructors and `edit`) are crate-private; `Package::edit` is the only staging route. Every capability the engines had beyond the primary verbs is a `Package` route or an `Edit`/`Commit` verb (§3) |
| F2 — source-backed publishers on the 0355a §1 shape | DOC `source_backed::paragraph`, XLS `source_backed::{cells, comments, sheet_visibility}`, RTF `source_backed::tail_append`: `edit_x()` on a source-backed owner → `Edit` → `commit()` validates and writes nothing, returning `litchi_core::edit::Commit<Owner, Patch, Diagnostics>` whose owner holds the checked plan; `write_to`/`save` on the owner are the only output, byte-identical to the former publishers (§2, §7) |
| F3 — mutable `raw` codec models | accepted under 0351 §2.1 rule 6 (§4) |
| F4 — unprefixed DOFR and saved-selection edits | `parts::dofr::{Edit, Commit, Patch}` and `parts::saved_selection::{Edit, Commit, Patch}`; the prefixed root re-exports are gone |
| DOC 1k batch, operation bound raised through `EditOptions` | 1,000 `bold` sets: 0.91–1.04× the leaf edit in one process, 0.72× each alone (budget ≤ 2.0×), after the DOC body run coalesces paragraph toggles; on the 0359 tip 6.6 s at 15.7 GB peak, 6.0× the leaf (§5) |

Commits: `cfdeed88d4` (F4), `4394b957a5` (DOC F1), `fc45dbf787` (XLS F1
and F2, harnesses, gate), `29abcb4af5` (DOC and RTF F2, gate),
`6e90981ce5` (DOC body-run toggles, probe), `2f3c75bf8b` (trunk merge),
`e3e805f2dd` (record, evidence, `p0359a`), and this record's gate table.

## 2. Source-backed families (F2)

Every family follows the 0355a §1 shape:

```text
<owner>::edit_x(&self) -> x::Edit                  // shares the owner
x::Edit::{verbs..}(&mut self, ..)                  // staging only
x::Edit::commit(self) -> Result<x::Commit>         // validates, writes nothing
x::Commit = litchi_core::edit::Commit<Owner, x::Patch, x::Diagnostics>
<owner>::apply_x(&self, &x::Patch) -> Result<x::Commit>
<owner>::write_to(&self, impl Write) / save(path)  // the only output
```

| Format | Owner | Family | Former publisher |
|---|---|---|---|
| DOC | `litchi_doc::source_backed::Package` (was `body_text::source::SourceSnapshot`; `OpenOptions`, `Limits`) | `source_backed::paragraph` (`Package::edit_paragraph(position)`, `apply_paragraph`) | `body_text::source::{Edit, Commit}` |
| XLS | `litchi_xls::SourceBackedWorkbook` | `source_backed::cells` (`edit_cells`: `set_number`, `set_numeric`, `set_value`), `source_backed::comments` (`edit_comments`: `replace`, `replace_many`, `set_visibility`), `source_backed::sheet_visibility` (`edit_sheet_visibility`: `set_visibility`, `hide`, `very_hide`, `unhide`, `set_visibility_batch`) | `cell_values::Edit::commit_source_backed{,_plan}`, `comments::Edit::commit_source_backed`, `sheet_visibility::Edit::commit_source_backed` |
| RTF | `litchi_rtf::source_backed::Package` (bound at open to its parsed document and source, under a retained `ExecutionContext`; `OpenOptions` with the publication windows) | `source_backed::tail_append` (`edit_tail_append{,_with_options}`: `append_paragraphs`, `append_text_paragraphs`, `append_runs`) | `tail_append::{TailAppendSourceProof, TailAppendSourceEdit, TailAppendSourcePublicationPlan}` |

How a committed owner works:

- **Commit validates, output streams.** `commit()` runs every check the
  former publisher ran before output (DOC: the paragraph proof, the
  equal-length splice plan and the candidate readback; XLS: fixed-width
  fields, same-length `BoundSheet8`/`TXO` splices, protection, signature,
  macro and freshness refusals, the complete reader's readback; RTF: the
  bounded tail plan against the proven source) and returns a new owner
  holding the checked plan. `write_to`/`save` stream that plan: the same
  `ValidatedOverlayPlan` (DOC, XLS) or tail splice (RTF) the former
  publisher wrote, so the bytes, the windows and the memory bound are the
  former ones (§7: byte-identical on every route).
- **Reads and chained edits.** The DOC and XLS committed owners read
  through the plan's composed source (no materialization). The RTF owner
  retains only its proof and source; editing a committed RTF package again
  materializes its output once (O(artifact) for that package only) and
  proves the new tail against it.
- **An opened owner** writes its source: DOC and RTF through the checked
  identity stream the former no-op commit wrote, XLS through its package
  view.
- **Patch.** XLS and RTF patches are the exact package pair: they replay
  onto a share of the owner they were committed on, their inverse onto a
  share of the commit's owner, and refuse any other (`PatchConflict`). The
  DOC patch keeps its replay rule (source version and fingerprint, target
  fingerprint checked): `apply_paragraph` replaces `Patch::apply`, and the
  inverse restores the source fingerprint (the existing tests).
- **Diagnostics** keep every former accessor (splices, replacement bytes,
  changed spans, fingerprints, the XLS operation shape) and gain
  `changed()`. XLS `write_to`/`save` return
  `Result<Option<PublishReport>, source_backed::OutputError>` (`None` for an
  opened workbook); RTF returns `source_backed::Report`, with
  `source_backed::OutputError` for `save`.
- **Removed:** `SourceSnapshot::{edit, transaction}` (aliases of
  `edit_paragraph`), `Patch::{is_empty, apply}` (DOC), the root RTF
  re-exports of the four `TailAppendSource*` types, XLS
  `{cell_values, comments, sheet_visibility}::SourceBacked*` and
  `commit_source_backed*` (crate-private).

## 3. Engines behind `Package::edit` (F1)

DOC (`4394b957a5`): the body-text engine is crate-private and
`Package::edit` stages it. New `Package` routes (`package/body.rs`): reads
`body_paragraphs`, `story_paragraphs`, `table_cells`, `field_results`,
`saved_selection`, `dofr_records`; composition `compose` →
`JoinedEdits::{join, commit}` with `prepare_{replace, bold, text, format}`;
`plan_three_way` → `ThreeWayPlan::{conflicts, is_conflict_free, commit}`;
`history` → `History<Package>`; `apply_durable`; transfer planning
`plan_{text, embedded, picture}_transfer_from`; and, under
`performance-diagnostics`, `edit_profiled` with
`Edit::commit_profiled{,_with_cfb_observer}`. Every commit is the primary
`Commit` (a `Package`, its `Patch`, `Diagnostics`). Seven suites that drove
the engine moved to `tests/internal/` (compiled as unit tests through
`src/internal_tests.rs`, crate named `litchi_doc` by `extern crate self`).

XLS (`fc45dbf787`): `cell_values::{Edit, Commit, Patch, History}`,
`Snapshot::{from_bytes*, from_shared_ole_file, edit, history}` and
`Package::edit_snapshot{,_with_options}` are crate-private;
`Package::cell_values()` is the read view (worksheets, cells, styles, the
Workbook stream). New routes: `Package::{apply_semantic, plan_three_way,
plan_transfer, execute_transfer, history}`, `Edit::join`,
`Commit::record_in(&mut History)`, `Patch::semantic()`; the
`copy_cells_from` alias is gone. Three suites moved to `tests/internal/`.

The harnesses (`tools/perf-baseline` `xls_numeric`, the comments and
visibility controls, the semantic XLS cases and the security corpus;
`tools/native-resave`; `tools/perf-cfb-profile`) and the litchi-ops XLS
and DOC adapters use the public routes. In the numeric harness the
source-backed and plan-only selectors both time the deferred cells family
(nothing is materialized at commit), and the signed/encrypted security
refusals are matched on the source-backed open (`SourceBackedError::Cfb`,
`EncryptedUnsupported`).

## 4. Mutable `raw` codec models (F3)

Accepted under 0351 §2.1 rule 6: a `raw` module may keep a mutable codec
model under a non-owner noun. DOC and XLS `raw::Reader` take `&mut self`
for their cursor reads and outputs, and RTF `raw::Tree` keeps its
`&mut self` setters for the writer (`RtfWriter::write_document`); neither is
reachable from a `Package`, and neither carries an owner name. No change.

## 5. The DOC 1k batch

`legacy_w5_source_probe` (§7) times one `File::edit_with_limits` of 1,000
`bold` sets on a 1,000-paragraph DOC (`SemanticLimits::max_operations(1_000)`,
which litchi-ops passes to the leaf as `EditOptions::transaction_limits`)
against the leaf `Package::edit_with_options` with
`TransactionLimits::new(1_000, ..)` and 1,000 `set_paragraph_bold` calls.

On the 0359 tip each `bold` set was its own isolated leaf commit retaining a
complete snapshot: the batch is refused at the default 256 MiB retention
bound (item 306), and with the bound raised it takes 6.59–6.66 s at a
15.7 GB peak, 5.99–6.04× the leaf. `6e90981ce5` lets direct `/body/p[N]` `bold` and `italic`
sets join the DOC body run that already coalesced `text` sets: one leaf
edit stages every toggle (a toggle the paragraph already holds stays a
no-op, as in the isolated setter), one commit publishes, and a paragraph
takes one kind of change per run (a toggle after a text replacement of the
same paragraph, or the reverse, closes the run). The batch now takes
1.34–1.57 s at a 117–133 MB peak (the whole probe process): **0.91–1.04×
the leaf** in the same process, and 1.23 s against 1.70 s (0.72×) with
each alone in its process (`results/0359a/doc-batch-isolated.txt`); budget
≤ 2.0×. The 0359 `doc/batch256` route (256 `bold` sets) falls from 0.63–0.65 s
to 0.087 s (0.13×). The new
`doc_paragraph_toggles_coalesce_into_one_commit` test pins one commit, the
isolated setter's readback, the inverse and the held-toggle no-op.

The leaf edit itself costs 1.1–1.7 ms per toggle at 1,000 paragraphs
(1.10–1.71 s per batch, depending on the process's heap state, §7); that is
the body transaction's existing cost, outside this record's budget
(follow-up 0359a-F1).

## 6. Gate

Families `litchi_doc::source_backed::paragraph`,
`litchi_xls::source_backed::{cells, comments, sheet_visibility}`,
`litchi_rtf::source_backed::tail_append` and the owners
`litchi_doc::source_backed::Package`, `litchi_xls::SourceBackedWorkbook`,
`litchi_rtf::source_backed::Package` are registered. New bans: the DOC
`body_text` engine names (direct children), `body_text::source`,
`SourceSnapshot`, `SourceBackedOptions`, `source_backed::Package::{edit,
transaction}`, `paragraph::Patch::is_empty`; the XLS `cell_values` engine
names, `Snapshot::{edit, history, from_bytes*, from_shared_ole_file}`,
`{comments, sheet_visibility}::{SourceBackedCommit, SourceBackedDiagnostics,
commit_source_backed}`, `Package::edit_snapshot*`, prefixed names directly
under `source_backed`; the four `TailAppendSource*` names in RTF.

## 7. Performance

`crates/litchi-ops/examples/legacy_w5_source_probe.rs` times each changed
source-backed route end to end (open, stage, commit, write the complete
output to a `Vec`) and the DOC 1k batch; only its `routes` module differs
between the base build (record 0359 tip `7afba5f146`, the former
publishers) and the head. `legacy_w5_probe` (record 0359's probe, unchanged)
re-times the 0359 `litchi_ops` routes. Release, pinned to CPU 29, 31 rounds
(7 for the DOC batch), each process in a 20 GB memory scope, A-B-B-A
twice (`results/0359a/abba.sh`, `analyze.py`, `summary.json`; second run in
`run2/`; binary hashes in `binaries.sha256`, binaries not committed).
Every route's output SHA-256 is the same on both legs.

Ratio = median of the after runs / median of the before runs, run 2 (run 1
in parentheses; run 1's binaries lack only the output digests and the
`LITCHI_PROBE_ONLY` filter). Host load 3.3–12.5 (other agents on other
CPUs).

Source-backed routes (family route vs former publisher):

| Route | Before (µs) | After (µs) | After/before |
|---|---:|---:|---:|
| DOC paragraph, 1k paragraphs | 3517.1 | 3524.8 | 1.002 (1.000) |
| XLS cells, 1 cell | 239.7 | 209.3 | 0.873 (0.875) |
| XLS cells, 100 cells | 602.6 | 265.0 | 0.440 (0.437) |
| XLS comments, 1 of 256 | 1510.6 | 1214.3 | 0.804 (0.791) |
| XLS comments, 256 of 256 | 2013.0 | 1732.7 | 0.861 (0.854) |
| XLS visibility, 7 of 8 sheets | 76.9 | 66.4 | 0.863 (0.874) |
| RTF tail append, 1k paragraphs | 298.9 | 299.6 | 1.002 (0.979) |

The XLS cells family replaces both former publishers: against the
materializing `commit_source_backed` it is 0.44–0.87×; against the
forward-only `commit_source_backed_plan` (no patch, no reads) it is 1.054
(1.047) for one cell and 1.044 (1.037) for 100 — the cost of reopening the
committed workbook over the plan's composed source at commit, which makes
the committed workbook readable and its patch exact (an explained residual).
The DOC and RTF routes are at parity (the same plan, now held by the owner).

DOC batch and the 0359 routes (0359a head vs 0359 head):

| Route | Before (µs) | After (µs) | After/before |
|---|---:|---:|---:|
| DOC 1k `bold` batch, `File::edit_with_limits` | 6662052.5 | 1338031.8 | 0.201 (0.238) |
| DOC 1k `set_paragraph_bold`, leaf edit | 1103361.3 | 1475024.5 | 1.337 (1.375) |
| DOC open + get | 847.4 | 832.5 | 0.982 (1.188) |
| DOC get | 0.2 | 0.2 | 0.960 (0.986) |
| DOC set | 1727.9 | 1729.1 | 1.001 (1.071) |
| DOC batch, 256 `bold` sets | 654585.5 | 87368.3 | 0.133 (0.137) |
| XLS open + get | 167.9 | 149.7 | 0.892 (0.933) |
| XLS set | 99.8 | 98.3 | 0.985 (1.082) |
| XLS batch, 1k sets | 4456.0 | 4364.5 | 0.979 (1.001) |
| RTF open + get | 207.0 | 200.1 | 0.966 (0.987) |
| RTF set | 328.6 | 323.4 | 0.984 (0.977) |
| RTF add | 560.3 | 547.7 | 0.977 (0.981) |
| RTF batch, 1k sets | 5544.7 | 6186.9 | 1.116 (1.005) |

Four routes exceed 1.05 in one run and not the other (DOC open + get and
DOC set and XLS set in run 1, RTF batch in run 2). Callgrind counts the
same instructions on both builds for each (`results/0359a/callgrind-totals.txt`;
five-round differences: DOC open + get 410,197,855 vs 410,185,117; DOC set
780,058,621 vs 780,099,731; XLS set 55,724,837 vs 55,633,446; RTF batch
681,985,042 vs 681,849,942), so these are layout and host noise, as in
0358a. The leaf DOC batch row is process state, not the leaf edit: the
probe runs it after the ops batch, which on the base faults in about 15 GB
first; with each run alone in its process the leaf takes 1.684–1.688 s on
the base and 1.669–1.690 s on the head (`results/0359a/doc-batch-isolated.txt`).

`tools/unified_bench_matrix.py` gains `p0359a` (the rows above, the
plan-only comparison and the DOC batch's unified/leaf row under the 2.0×
batch budget).

## 8. Gates

Re-run in full on the head after the trunk merge (`e3e805f2dd`), one
cargo command at a time (`.coord` gate script; logs not committed); the
working tree was clean afterwards.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-{doc,xls,rtf,ograph,ole-common,crypto,ppt} -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings`, and with `--features doc`, `xls`, `rtf`, `ppt` | 0 each |
| `cargo test --no-fail-fast --all-features -p` each of the seven crates (doc 1,488, xls 1,871, rtf 1,562, ograph 70, ole-common 310, crypto 88, ppt 1,570) | 0 |
| `cargo test -p litchi-ops --all-features` (lib 953, doc 217, the integration targets in six chunks 1,473; 2,643) | 0 |
| `cargo test -p litchi --all-features` (209); `cargo test -p litchi-cli --all-features --test officecli_corpus` (5) | 0; 0 |
| `LITCHI_BLESS=1 cargo test -p litchi-ops --all-features --test capabilities --test registry`, and with every feature but `encryption`; `git diff --exit-code -- crates/litchi-ops/tests` | 0, 0; 0 (no change) |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p` the six crates, `litchi-ops`, `litchi` | 0 |
| `RUSTC_BOOTSTRAP=1 tools/check_api_glossary.py` (47 profiles, 3,257 bans, 498 families, 77 owners) | 0 |
| `check_crate_boundaries.py`; `non_iwork_gate.py verify` | 0; 0 |
| `crates/pyo3-litchi/tests/run_tests.sh` | 0 |
| `tools/perf-baseline`: `cargo check --locked --all-targets`; `tools/native-resave`: `cargo check --locked --all-features`; `tools/perf-cfb-profile`: `cargo check --all-targets` | 0; 0; 0 |
| tool unit tests (`test_officecli_corpus`, `test_officecli_capability_delta`, `test_no_3rdparty_references`, `test_check_crate_boundaries`, `test_non_iwork_gate`, `test_check_api_glossary`, `test_unified_bench_matrix`; 1,304) | 0 |
| `officecli_capability_delta.py --check`, `officecli_corpus.py --check`, `unified_bench_matrix.py --check`, `check_feature_matrix.py` (31 builds) | 0 each |
| `officecli_capability_delta.py --officecli …` (regenerated) | 0, no change |

## 9. Gaps and follow-ups

1. **0359a-F1** — the DOC body transaction costs 1.1–1.7 ms per staged
   toggle at 1,000 paragraphs (§5) and is sensitive to heap state; not
   profiled here.
2. **0359a-F2** — the XLS and RTF source-backed patches are exact pairs
   only; a reopened copy of the before bytes is refused rather than
   replayed (0355a's content-compared replay is not implemented for them).
3. The XLS numeric harness's source-backed and plan-only selectors now time
   the same deferred route; their evidence fields
   (`target_artifact_materialized_at_commit`, the materialization byte
   counts) report the deferred owner (§3).
4. **0359a-F3** — the XLS cells family reopens the committed workbook at
   commit (1.04–1.05× the former forward-only plan publisher, §7); opening it
   on first read would remove that cost for callers that only write.

## 10. Log entry

- 0359a (W5 review): `doc::body_text` and `xls::cell_values` are crate-private behind `Package::edit` and new `Package` routes (DOC `compose`/`plan_three_way`/`history`/`apply_durable`/transfer plans and body reads; XLS `apply_semantic`/`plan_three_way`/`plan_transfer`/`history`, `Edit::join`, `Commit::record_in`).
- Source-backed publishers are deferred families on source-backed owners: DOC `source_backed::paragraph` (`source_backed::Package`), XLS `source_backed::{cells, comments, sheet_visibility}` (`SourceBackedWorkbook`), RTF `source_backed::tail_append` (`source_backed::Package`); commit writes nothing, `write_to`/`save` stream the former plan byte for byte.
- F3 accepted under 0351 §2.1 rule 6; F4 `parts::{dofr, saved_selection}::{Edit, Commit, Patch}`.
- DOC body run coalesces paragraph `bold`/`italic`: 1k sets 6.6 s/15.7 GB → 1.3–1.6 s/0.13 GB, 0.91–1.04× the leaf (budget 2.0×); 256-set batch 0.13×.
- ABBA vs the 0359 tip: source-backed routes 0.44–1.00×, outputs identical; four one-run misses are the same callgrind instructions; gates all 0; `p0359a`.
