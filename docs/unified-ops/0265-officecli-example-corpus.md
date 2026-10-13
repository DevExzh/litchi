# 0265 — The OfficeCLI example-corpus harness

Date: 2026-10-13. Status: implements record 0248 §6 Phase D step 17
(record 0015 W8, the OfficeCLI example corpus). Branch `uo2/s17-corpus`, cut
from trunk `4d1afe92bc`.

Every OfficeCLI example script now replays against the `litchi` binary in
CI. Each `officecli` call either passes or is classified as a gap, a
recorded delta, a cascade of an earlier refusal, or a bug. Panics and
`internal_error` fail the test, and so does any regression in the passed
calls.

## 1. What landed

| Piece | Path |
|---|---|
| The corpus: all 96 example scripts (`examples/{word,excel,ppt}/**/*.sh`, 1,168,550 bytes) and the four small pictures two scripts need, from the pinned upstream commit `dced0d74ff` | `test-data/officecli/examples/` (provenance in `test-data/officecli/PROVENANCE.md`) |
| The `officecli` stand-in that translates and records each call | `crates/litchi-cli/tests/officecli_corpus/officecli` |
| The replay test (CI): replays, compares and fails on panic, `internal_error` or a regression | `crates/litchi-cli/tests/officecli_corpus.rs` |
| The recorded outcomes, one row per call (6,808 calls, 257,386 bytes) | `crates/litchi-cli/tests/officecli_corpus/outcomes.tsv` |
| The classifier and table generator, with `--check` | `tools/officecli_corpus.py`, `tools/test_officecli_corpus.py` |
| The generated table: summary per family, format and script; deltas; gaps with registry cross-checks; bugs; record 0256 follow-up counts | [`officecli-corpus-deltas.md`](officecli-corpus-deltas.md) |
| CLI options `add --from`, `set --find/--replace`, `remove --prop` | `crates/litchi-cli/src/{cli,run,lib}.rs`, test `officecli_option_spellings_reach_the_command_channel` |

## 2. How a script is replayed

The scripts are real bash: they use loops, functions, `$(...)` captures,
heredoc batches and Python helpers. They are not parsed. bash runs them
with `officecli` replaced, so every line runs as OfficeCLI's own CI runs
it:

1. The test copies the script and its sibling inputs into a fresh scratch
   directory. It replaces the shebang with
   `trap 'export OFFICECLI_LINE=$LINENO' DEBUG`, so line numbers stay those
   of the vendored file, and it turns `set -e` lines into `set +e`, so one
   refused line does not end the replay.
2. `bin/officecli` (the stand-in, §3) comes first on `PATH`, so it also
   receives `$CLI` and `command officecli` calls. `TMPDIR` and `HOME` point
   into the scratch directory.
3. `bin/python3` and `bin/python` run the real interpreter with `-I -S`,
   so helper snippets see only the standard library and every machine
   replays the same lines (no Pillow, no numpy; delta H1).
4. After the script, each file it created is read back with
   `view <file> text` and `get <file> / --depth 2`, recorded as line 0.
5. Six workers replay scripts in parallel. The suite takes 19 s on the dev
   profile (§7).

Each call becomes a row: line, verb, format (the document's extension),
`--type` (or `(from)`), result (`ok`, `skip`, `usage`, or the error code from
stderr), detail (the refused property, the offending argument, or
`translated:<delta>`), and the last path step. A line inside a loop or a
function counts once per call.

The test then compares the fresh rows with `outcomes.tsv`. Keys are the
line, the verb and the occurrence. It **fails** when:

- any call panics, dies by a signal, is lost, or returns `internal_error`;
- a script runs longer than 900 s;
- a call that passed now fails or is no longer replayed;
- a script family passes fewer calls.

Other differences, such as newly passing lines or changed refusal codes,
are printed but do not fail the test. Re-record them with
`LITCHI_BLESS=1 cargo test -p litchi-cli --test officecli_corpus`, then run
`python3 tools/officecli_corpus.py`. Two consecutive runs here, and a run
with `--all-features`, produced byte-identical outcomes.

## 3. Translation

The stand-in passes every call through unchanged except these:

| OfficeCLI call | Becomes | Why |
|---|---|---|
| `open`, `close` | nothing, recorded `skip` | resident mode is absent (record 0256 D5); each command opens and saves |
| `batch FILE` reading stdin | `batch FILE --input -` | Litchi needs the input named |
| `add F / --type slide` without `layout` | `... --prop "layout=Title Slide"` | OfficeCLI picks the deck's first layout. Litchi requires an exact name (record 0256 D4), and `Title Slide` is the first layout of Litchi's blank deck. Rows are tagged `translated:D4`. |
| `--prop layout=blank` / `layout=title` on a slide | `layout=Blank` / `layout=Title Slide` | OfficeCLI also matches layout types; tagged D4 |
| `set`, `add`, `batch` | `... --unsupported warn` (added by record 0269) | OfficeCLI applies the supported properties of a command and warns about the rest by default; Litchi refuses unless asked (record 0015 Q2/Δ1). A call that skipped properties exits 9 and is recorded as `warned` with the skipped properties as detail; the delta table lists each skipped property in its gap table with refusal `warned`. |

Without the D4 translation every PPTX script stopped at its first slide:
342 refusals, with almost every later PPTX line a cascade. With it, 344
slide insertions pass as translated and the rest of each deck is exercised.

## 4. Classification

`tools/officecli_corpus.py` classifies each refusal in this order:

1. The named rules: deltas D1, D4, D5, D7, D8, U1, U2, H1; gaps G1–G7;
   cascade C1.
2. A **cascade**: `path_not_found`, `missing_argument`, `unsupported_path`
   or `out_of_range` after an earlier refused `add` or `batch` in the same
   script.
3. **Gaps** cross-checked against the registry export (the golden of
   `schema()`, which `capabilities()` shares). The check covers
   `unsupported_*` element refusals and `unsupported_property`:
   - an element or property gap (with its follow-up), or one the registry
     does not list, is a gap;
   - a refusal of a property the registry marks add- or set-able is a
     **bug**.

   Property rows also say whether OfficeCLI's help schema documents the
   property.
4. Anything else is a **bug**.

A separate check lists passed calls on elements whose registry entry lacks
the verb, which would mean an under-claim. There are none.

## 5. Results

From the generated table:

| Family | Lines | Passed | of which translated | Gap | Delta | Cascade | Bug | Resident | Read-back |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| excel | 1633 | 869 | 0 | 749 | 4 | 11 | 0 | 20 | 20/20 |
| excel/charts | 743 | 103 | 0 | 635 | 5 | 0 | 0 | 32 | 32/32 |
| ppt | 310 | 43 | 32 | 139 | 0 | 128 | 0 | 10 | 12/12 |
| ppt/charts | 735 | 127 | 113 | 605 | 0 | 3 | 0 | 28 | 28/28 |
| ppt/ole | 26 | 13 | 2 | 11 | 0 | 2 | 0 | 6 | 6/6 |
| ppt/pictures | 68 | 6 | 5 | 49 | 7 | 6 | 0 | 2 | 2/2 |
| ppt/shapes | 220 | 23 | 19 | 193 | 0 | 4 | 0 | 8 | 8/8 |
| ppt/tables | 457 | 40 | 33 | 179 | 0 | 238 | 0 | 14 | 14/14 |
| ppt/templates | 452 | 17 | 0 | 150 | 117 | 168 | 0 | 11 | 24/24 |
| ppt/textboxes | 132 | 12 | 10 | 105 | 0 | 15 | 0 | 4 | 4/4 |
| ppt/transitions | 522 | 136 | 127 | 386 | 0 | 0 | 0 | 18 | 18/18 |
| word | 1133 | 614 | 3 | 415 | 6 | 98 | 0 | 24 | 32/32 |
| **total** | **6431** | **2003** | **344** | **3616** | **139** | **673** | **0** | **177** | **200/200** |

Per format: DOCX 949 lines, 466 passed; XLSX 2,523 lines, 1,119 passed; PPTX
2,959 lines, 418 passed (344 of them translated slide insertions). No call
panicked or returned `internal_error`. Every created file read back (200/200).
The script's own read lines (`get`, `query`, `view`) pass, or fail as gaps,
for example `query revision` and `get` on XLSX shapes.

Bugs: none. Every refusal is a recorded delta, a gap the registry agrees
with, or a cascade. Before the rules existed, six lines were
unexplained. Four were XLSX `formula="=..."`; OfficeCLI's schema says
"without leading =" and Litchi follows it, so these became delta D8. Two
were PPTX table-cell `fill=none` and a gradient string, which OfficeCLI
documents and Litchi's colour coercion lacks; these became gap G6.

The gaps that block the most lines, all outside this step's ownership
(`litchi-ops` registry and leaf writers):

| Lines | Format | Gap | Follow-up |
|---:|---|---|---|
| 906 | xlsx | `set` cell `text=`: OfficeCLI's cell handler takes `text` as `value` but its help schema does not list it (G7) | add `text` as an alias of the XLSX cell `value` |
| 461 | pptx | `add shape` / `textbox` with `size` (font size): registry `size` is get-only | shape text-run size on add |
| 446 | pptx | `add chart`: no add on chart | record 0248 step 19 (typed chart CRUD) |
| 317 | xlsx | `add chart`: registry gap L-XLSX-R7 | step 19 |
| 162 + 93 | pptx | `add shape` `fill`, `geometry`: registry gap L-PPTX-R6b | L-PPTX-R6b |
| 136 + 25 | pptx | slide `transition` (set and batch): gap M1-extended | slide transitions |
| 138 | all | verbs `validate`, `raw-set`, `add-part`, `dump` (G1) | record 0256 D6 |
| 80 | docx | `add equation`: no add on equation | equation writes |
| 51 | pptx | `add table`: no add on table | PPTX table insertion |

The cascades are mostly PPTX table and shape lines (238 in `ppt/tables`)
whose table or shape `add` was refused.

### Record 0256 follow-ups hit

| Follow-up | Lines | Done here |
|---|---:|---|
| F3 property-granularity "apply the rest and warn" | 2,172 calls refused for one unsupported property | no (ops API) |
| F4 view options | 0 | n/a |
| F6 `add --from`, `set --find/--replace`, `remove --shift`, `get --save` | 40 (`--from`); 0 reach find/replace (§6) | CLI spellings for `--from`, `--find`, `--replace` |
| D6 missing verbs | 138 | no (ops APIs) |
| F1, F2, F5, F7, F8 | not exercised by the corpus | no |

## 6. CLI changes (the trivial part of F6)

These options used to stop at the argument parser with exit 2. Each now
reaches `File::apply`, which refuses it with a typed error, so a later ops
implementation lights them up without another CLI change:

- `add <file> <parent> --from <path>` sets `BatchItem::from`. `--type` is
  now required unless `--from` is given, and the two conflict. Ops answers
  `unsupported_capability` ("add-from copying is not integrated", exit 5).
- `set --find T` and `--replace T` become `find=T` and `replace=T`, ahead
  of the other `--prop`s, as OfficeCLI defines them. `--prop` is optional
  when either is given. Ops answers `unsupported_property` with property
  `find` (exit 4) until the Phase E engine lands.
- `remove --prop k=v` sets the item's props. Ops answers `invalid_argument`
  (remove admits no props; OfficeCLI uses them for tracked deletions; G4).

The corpus's `--find` lines in `word/revisions.sh` do not reach the new
spelling. Their path is captured from OfficeCLI's `/body/p[@paraId=...]`
output, and Litchi prints positional paths, so they are delta D1.

The library API and the glossary are unchanged.

## 7. Measurements

No library path changed. The CLI change adds two optional fields to argument
parsing. The performance bar's leaf-vs-unified budgets therefore do not
apply, and nothing was measured against them.

Harness cost on this host: 6,808 calls (6,431 script calls, 177 resident
no-ops and 200 read-backs). Six runs on the dev profile, with six workers,
took 18.8–21.6 s wall; one serial prototype run took 73 s. This is a
duration, not a performance claim.

## 8. Findings (not fixed here; outside this step's ownership)

- **X1, misleading message.** XLSX `add <sheet> --type chart` (and
  `conditionalformatting`, `validation`, `shape`, `pivottable`,
  `sparkline`, `slicer`, `cell`) answers "an XLSX sheet can only be added
  directly under the workbook". The code (`unsupported_capability`) is
  right; the message names the wrong element.
- **X2, inconsistent codes.** Adding an element the format cannot add yet
  is `unsupported_element` in DOCX, `unsupported_type` in PPTX and
  `unsupported_capability` in XLSX. The CLI exit status is 5 for all three.
- **X3, registry under-documentation.** OfficeCLI's cell handler aliases
  `text` to `value` without documenting it (G7). The registry follows the
  help schema, so the alias is missing.

## 9. Gates

Run with `CARGO_TARGET_DIR=<worktree>/target CARGO_INCREMENTAL=0
CARGO_PROFILE_DEV_DEBUG=0 CARGO_PROFILE_TEST_DEBUG=0`.

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-cli -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo test -p litchi-cli --all-features` (22 passed: 3 unit, 15 `cli`, 3 `officecli_corpus`, 1 doc) | 0 |
| `cargo test -p litchi-ops --all-features` (1,878 passed) | 0 |
| `cargo test -p litchi --all-features` (260 passed) | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-cli` | 0 |
| `python3 tools/check_crate_boundaries.py` | 0 |
| `python3 tools/non_iwork_gate.py verify` | 0 |
| `RUSTC_BOOTSTRAP=1 python3 tools/check_api_glossary.py` (21 profiles, 1,043 bans, 33 families) | 0 |
| `python3 tools/officecli_corpus.py --check` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_no_3rdparty_references tools.test_officecli_capability_delta` (41 tests) | 0 |

Python bindings, the registry goldens and the capability delta table are
unaffected (no ops change).

## 10. Not claimed

- The harness does not compare document content with OfficeCLI's output.
  "Passed" means `litchi` accepted the call and the file reads back. Values
  are checked only by the seed smoke script (`tests/officecli_smoke.sh`,
  record 0256) and the scripts' own read lines.
- Batches count as one line each. Their items are not split out.
- Lines behind refused captures (`$(officecli add ...)`) are counted where
  they land. They are cascades or delta D1, never passes.
