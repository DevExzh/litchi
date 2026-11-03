# 0430 — macOS queue runbook: generated packages, Mac runner, result folding

- Status: merge-ready on branch `uo2/mac-runbook` (cut from trunk
  `feat/unified-ops-api` at `27b4840684`).
- Date: 2026-11-03.
- Purpose: `docs/unified-ops/0363-mac-queue.md` lists 766 migrated iWork
  routes. Each needs a native open, save, close and reopen in Pages, Numbers
  or Keynote, with the result and the artifact SHA-256 recorded, before W3
  deletes `litchi-iwa`. This record makes that run mechanical. A Linux tool
  generates every Litchi-mutated package. A shell script runs them on the
  Mac. A second tool writes the results back into the queue.

## 1. What landed

| Piece | Path | Role |
|---|---|---|
| Parity dump hook | `crates/litchi-iwa/tests/parity/harness.rs` (module `mac_queue`) | With `LITCHI_MAC_QUEUE_DUMP=<dir>`, every harness case writes its input, host output, leaf output and a JSON sidecar (case, test function, occurrence, format, tier, fixture, read-only, passed). Without the variable it does nothing. |
| File-route example | `crates/litchi-ops/examples/mac_queue_file_routes.rs` | Writes the `litchi_ops::File` rows' packages through `File::set` (cell comment set and clear, chart point-size set and reset, for each format), a Numbers empty-cell value write and two Pages documents built with `Identity::generate()`. |
| Row mapping | `tools/mac_queue/rows.py` | Parses the queue, gives each row a content id (`q-<12 hex>` over its first five cells, so filling Result keeps the id), and maps each row to its sources or to an explicit group. |
| Generator | `tools/mac_queue/generate.py` | Runs both producers, deduplicates packages by SHA-256, and writes `manifest.json`, `queue.tsv`, `packages/`, `inputs/` and `host-built-inputs.tsv`. `--check` only checks that every row maps. |
| Mac runner | `tools/mac_queue/run_on_mac.sh` | Runs each package through osascript: open, Save As, close, reopen, export to PDF, close. Writes `results.tsv`. Supports `--dry-run`, `--resume`, `--only` and `--app`. |
| Result folding | `tools/mac_queue/apply_results.py` | Fills the Result and Artifact SHA-256 cells of `0363-mac-queue.md` from the manifest and `results.tsv`. |
| Tests | `tools/test_mac_queue.py` (in `tools/gate/run_gate.sh` `TOOL_MODULES`; the `tooling-unit` row also runs `generate.py --check`) | Parsing, citations, mapping rules, every tracked row mapped, generator grouping and host-input detection, result folding, and a runner dry run. |

### 1.1 Mapping rules (`rows.map_row`, first match wins)

1. **Native authoring** (2 rows). The row asks the operator to create a
   document in the native app: the W1-F volatile highlight and the W4-K pivot
   `Aggregate` names. These rows have no generated package and are grouped
   `native-authoring`.
2. **`litchi_ops::File` rows** (18, wave D7). These map to the example's
   `<format>-cell-comment-{set,clear}` and
   `<format>-chart-font-size-{set,reset}` packages. A `get` row checks the
   package its `set` sibling wrote, where the value is present.
3. **Builder identity** (`litchi_pages::builder::Identity::from_parts`).
   Maps to the two example documents, which have fresh identities.
4. **Retired host route** (`litchi_iwa::numbers::Draft::{set_cell,
   set_cells}`, now crate-private). Maps to the example's
   `numbers-cell-value-set`: `File::set` on an empty cell of `basic.numbers`
   goes through the same shared TST cell writer.
5. **Host rows** (`host …` or `litchi_iwa::…`; 15 rows). These map to the
   **host** output of an explicitly named parity case (`HOST_CASES`). Examples
   are the record 0413 and 0426 host fixes and the W1-G host-only comment
   routes.
6. **Explicit leaf cases** (`LEAF_CASES`). `formula_graph::next_owner_id`
   is proven by `numbers_owner_allocation.rs`, which is not a harness case,
   so the row maps to the Numbers table-duplicate cases.
7. **Citations** (580 rows). The row text names parity tests. The parser
   accepts exact names, brace alternatives (`x_{set,clear}_parity`), `*` for
   the format prefix (`*_chart_pie_start_angle_parity`) and "`name` and
   siblings" (that test family on every format). It finds tests declared with
   `#[test]` and tests declared by `macro_rules!` tables. Leaf rows take the
   leaf output; host rows take the host output.
8. **Exit ledger** (148 rows, mostly "as above" fixtures). The row's
   destination is looked up in
   `docs/unified-ops/results/0363/iwa-exit-ledger.tsv` to find its
   `parity_test`.

If a row matches no rule, or a source names a test or example stem that does
not exist, `generate.py --check` fails, and so does the unit test
`RealQueueTests.test_every_queue_row_maps`.

At generation the generator looks up each parity source among the dumped
runs in this order:

1. by case name;
2. by the test function that ran the case. The sidecar records libtest's
   thread name, `module::test_fn`, so a case named differently from its
   test (for example `numbers_chart_text_title_axis_set_parity` inside
   `numbers_chart_title_and_axis_title_set_parity`) is still found;
3. as a **family fallback**. Most `*_read_parity` tests and some corpus
   loops compare reads directly instead of running a harness case, so they
   write no package. For these, the row takes the packages of the same test
   module, restricted to the test's format. Those packages carry the feature
   the read covers. The manifest lists such sources under
   `rows[].family_fallback`.

A row that still has no package makes generation fail.

### 1.2 Grouping

The generator groups packages by their SHA-256. Rows whose mutations produce
byte-identical packages share one package. This covers read rows, whose
read-only cases save the input unchanged, a `get` row and its `set` sibling,
and the tier-A cases cited by several rows. A row that cites several cases
gets one package per distinct output, and it passes only if all of them pass.
Package ids (`P0001`, …) follow the order in which queue rows first use them.

### 1.3 Representative runs

Many cases repeat one route over dozens of property values, cells or
fixtures. `text_paragraph_property_set_parity` alone writes 192 packages. With
every run kept, the queue needs 3 785 distinct packages. The generator keeps
`--runs-per-source` runs per parity source (default 2). It takes the first
run on each distinct fixture first, so a corpus case keeps fixture coverage,
then runs spread evenly over the rest (`generate.select_runs`). A family
fallback prefers the runs that the module's direct sources already selected,
so a read row adds no package when a write row of its family already has
one. With `--runs-per-source 1` the queue needs 836 packages; with `0` it
keeps every run.

### 1.4 Generation result (this branch)

Generated from the parity run of this branch: `cargo test -p litchi-iwa
--all-features --test parity`, 1190 passed, 0 failed, 210 ignored, about
33 minutes. The generator dropped the 13 runs whose case reported failures:
they are the negative controls of the `*_rejects_*` tests.

| | Count |
|---|---|
| Queue rows | 766 |
| Rows with packages | 764 (580 cited, 148 ledger, 18 File, 15 host, 3 other explicit) |
| Rows that use a family fallback | 172 |
| Rows grouped without a package (`native-authoring`) | 2 |
| Distinct packages (default `--runs-per-source 2`) | **1153**: Keynote 390, Numbers 369, Pages 394 |
| Packages per row | 1 for 229 rows, 2 for 348, 3 to 36 for the rest |
| Host-built inputs kept for W3 | 798 (175 MB) |
| `packages/` size | 262 MB; largest package 803 212 bytes |

None of the generated files is 1 MiB or more. They live outside the
repository, in `/home/zhuhe/code/litchi-worktrees/.coord/mac-runbook/out/`.
A Linux dry run of the runner over that directory verified all 1153 hashes in
17 s.

## 2. Prerequisites

On the Linux host (generation):

- this branch, or trunk once merged, with `litchi-iwa` still present;
- `source /home/zhuhe/code/litchi-worktrees/.coord/mac-runbook/env.sh` (build
  environment, OOM shim);
- about 5 GB free: the raw dumps (`out/raw`) take 3.5 GB and can be deleted
  once the manifest is written. The parity run takes about 33 minutes.

On the Mac (run):

- macOS 14 or later, with **Keynote, Pages and Numbers 14.4 or later**.
  14.4 is the version in the tracked native receipts
  (`test-data/iwork/**/*-receipt.json`). Later 14.x and 15.x versions are
  accepted, and their version is recorded on every result line. The runner
  warns about anything older than 14;
- the stock `/bin/bash`, `osascript`, `shasum` and `defaults`, all of which
  ship with macOS;
- **Automation permission.** The first run asks to allow Terminal (or your
  terminal app) to control Keynote, Pages and Numbers (System Settings →
  Privacy & Security → Automation). Allow all three;
- the apps signed in and past their first-launch screens. Open each app once
  by hand and dismiss the template chooser and any "What's New" sheet;
- free disk about three times the size of `packages/` (re-saved copies and
  PDFs are written next to them).

## 3. Generate (Linux)

```sh
cd /home/zhuhe/code/litchi-worktrees/uo2-mac-runbook   # or the trunk checkout
source /home/zhuhe/code/litchi-worktrees/.coord/mac-runbook/env.sh
python3 tools/mac_queue/generate.py --check            # every row mapped
python3 tools/mac_queue/generate.py --out /home/zhuhe/code/litchi-worktrees/.coord/mac-runbook/out
rm -rf /home/zhuhe/code/litchi-worktrees/.coord/mac-runbook/out/raw   # optional: 3 GB of raw dumps
```

The generator refuses an output directory inside the repository. It runs
cargo with `check=True`, so a failing parity test stops generation. It drops
runs whose case reported failures, which in a green run are only the
negative controls. `--reuse-dumps` rebuilds the manifest from an existing
`out/raw` without running cargo. `--runs-per-source N` sets how many
representative runs to keep (§1.3).

Host-built inputs (`out/inputs/`, listed in `out/host-built-inputs.tsv`) are
inputs that the host route made before the leaf ran, such as "host-added
Column2d chart" fixtures. They are not tracked files. The host draws random
UUIDs, so they differ on every generation, and so do the leaf packages built
from them. Keep the `out/` directory of the run you send to the Mac. W3
snapshots these inputs from that directory (note in
`.coord/mac-runbook/w3-host-built-inputs.md`).

## 4. Copy to the Mac

Copy the whole output directory. It holds `manifest.json`, `queue.tsv`,
`packages/`, `inputs/` and a copy of `run_on_mac.sh`. Leave out `raw/`.

```sh
# on Linux
cd /home/zhuhe/code/litchi-worktrees/.coord/mac-runbook
tar --exclude=out/raw -czf mac-queue-0430.tgz out
sha256sum mac-queue-0430.tgz
# on the Mac
scp linux-host:/home/zhuhe/code/litchi-worktrees/.coord/mac-runbook/mac-queue-0430.tgz ~/
shasum -a 256 ~/mac-queue-0430.tgz        # must match
mkdir -p ~/mac-queue && tar -xzf ~/mac-queue-0430.tgz -C ~/mac-queue
```

## 5. Run (Mac)

```sh
cd ~/mac-queue/out
bash run_on_mac.sh --dry-run .        # platform, app versions, every package hash
bash run_on_mac.sh --app Keynote .    # one app at a time is easier to watch
bash run_on_mac.sh --app Numbers --resume .
bash run_on_mac.sh --app Pages --resume .
bash run_on_mac.sh --resume .         # retries whatever has not passed
```

- Each package is copied to `mac-run/<id>/` and opened from there. The copy
  is saved as `<id>-resaved.<ext>`, closed, reopened, exported to
  `<id>.pdf` and closed again. The original in `packages/` is never touched.
- `results.tsv` gets one line per attempt. The last line for a package
  counts, so `--resume` retries anything that has not passed. Dry-run lines
  are ignored when the results are applied.
- `--timeout SECONDS` (default 180) bounds each AppleScript stage. If a
  modal dialog appears, such as "can't be opened" or a repair or update
  prompt, the stage times out, and the runner closes the documents and
  records `fail` with the stage name. Keep the dialog's text from the screen
  for the failure analysis (§7).
- If the app saves the re-saved copy as a package directory, its SHA-256 is
  the hash of the sorted `<sha256>  <relative path>` listing of its files,
  and `resaved_kind` says `directory`.
- To spot-check what a row expects, look it up in `manifest.json`
  (`rows[].fixture_note`, the queue's Fixture column). Open
  `mac-run/<id>/<id>.pdf` to check visually.

## 6. Bring the results back and apply them

```sh
# on the Mac
cd ~/mac-queue/out && tar -czf ~/mac-queue-results.tgz results.tsv
# on Linux, in the checkout that holds the queue
tar -xzf mac-queue-results.tgz -C /home/zhuhe/code/litchi-worktrees/.coord/mac-runbook/out
python3 tools/mac_queue/apply_results.py \
    --manifest /home/zhuhe/code/litchi-worktrees/.coord/mac-runbook/out/manifest.json \
    --results /home/zhuhe/code/litchi-worktrees/.coord/mac-runbook/out/results.tsv --dry-run
python3 tools/mac_queue/apply_results.py \
    --manifest /home/zhuhe/code/litchi-worktrees/.coord/mac-runbook/out/manifest.json \
    --results /home/zhuhe/code/litchi-worktrees/.coord/mac-runbook/out/results.tsv    # writes the queue
mkdir -p docs/unified-ops/results/0430
cp /home/zhuhe/code/litchi-worktrees/.coord/mac-runbook/out/results.tsv docs/unified-ops/results/0430/
```

- A row is filled only when every one of its packages has a result:
  - if all passed, Result is `pass (0430: <App> <version>, <date>)` and
    Artifact SHA-256 is the re-saved hash (`P0001 <sha>, P0002 <sha>` when
    the row has several packages);
  - if any failed, Result is `FAIL (0430: <id> <stage>, …)`.
- Only `pending` rows are rewritten. `--force` also rewrites the others.
- A result whose package hash differs from the manifest is an error, and so
  is a manifest row that is no longer in the queue. The fix in both cases is
  to regenerate and rerun the packages concerned. The results TSV is small;
  commit it as evidence (`docs/unified-ops/results/0430/`). Never commit
  packages, PDFs or re-saved copies.
- The two `native-authoring` rows stay `pending`. Do them by hand as their
  Fixture or Reason column describes, then commit the new native fixture and
  its receipt as the earlier `*-receipt.json` files were.

## 7. Failures: what to reopen

For every `FAIL` row:

1. Find the row in `manifest.json` (`rows[]` by line or id). The entry gives:
   - `packages`, the failing package ids;
   - `sources`, the parity case, or the example stem and role, that wrote
     the package;
   - `records`, the leaf record the row cites, with its tracked path, for
     example `0375b` → `docs/unified-ops/0375b-…md`.
2. Reopen that leaf record. Its route is not native-proven. Add a follow-up
   record (`<NNNN>a`, …) with the failing stage, the app version, the dialog
   text and the package SHA-256.
3. Reproduce on Linux. Run the source case alone with the dump on, for
   example
   `LITCHI_MAC_QUEUE_DUMP=/tmp/x cargo test -p litchi-iwa --all-features --test parity <case>`,
   and compare the leaf output with the host output (`archive_diff` in the
   harness) and with a package the app saved natively.
4. Rows whose packages all come from host-route rows (§1.1 rule 5) fail a
   host fix, not a leaf: reopen the cited host-fix record (0371, 0413, 0413b
   or 0426).
5. A `verify` failure means the package was changed in transit. Copy it
   again; do not edit `queue.tsv`.

Do not delete `litchi-iwa` (W3) while any row is `FAIL` or `pending`, except
the two `native-authoring` rows if the owner waives them.

## 8. Gaps and follow-ups

- **The runner has not run on a Mac.** This host is Linux. Only `bash -n`
  and the `--dry-run` path (hash verification, plan, results format) were
  exercised, by the unit test and over the real 1153-package output. The
  AppleScript (`open`, `save … in`, `export … as PDF`, `close saving no`)
  uses the documented Keynote, Pages and Numbers dictionary verbs. Run it
  first on a handful of packages
  (`--only P0001,P0002,P0003` and one per app), check `mac-run/<id>/` by
  hand, and fix the script before the full run.
- **Representative runs, not every run** (§1.3). The default keeps 1153 of
  3785 distinct packages. If the owner wants exhaustive coverage, regenerate
  with `--runs-per-source 0`.
- **Family fallback for read rows** (172 rows). A read test that compares
  reads directly writes no package, so its row is proven by its module's
  write packages, which carry the feature, rather than by a package of its
  own. The manifest lists every such source under `family_fallback`.
- **What the run proves.** It proves open, save, close, reopen and PDF
  export. It does not compare displayed values with the Litchi read. That
  comparison stays a manual spot check against `fixture_note` and the PDF.
  The builder-identity row also needs a manual look: the two documents must
  not trigger a version-conflict prompt when both are open.
- **The two `native-authoring` rows** are not generated (§6).
- **The host-built inputs change on every generation** (§3). Run the Mac
  against the same `out/` that W3 snapshots.

## 9. Gates (this branch, before commit)

| Command | Exit |
|---|---|
| `cargo fmt --all --check` | 0 |
| `cargo clippy -p litchi-iwa -p litchi-ops -p litchi -p litchi-py --all-targets --all-features -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --all-targets -- -D warnings` | 0 |
| `cargo clippy -p litchi-ops --no-default-features --features pages,keynote,numbers --all-targets -- -D warnings` | 0 |
| `RUSTDOCFLAGS=-Dwarnings cargo doc --no-deps --all-features -p litchi-iwa -p litchi-ops` | 0 |
| `LITCHI_MAC_QUEUE_DUMP=… cargo test -p litchi-iwa --all-features --test parity` (twice) | 0 (1190 passed, 210 ignored) |
| `cargo run -p litchi-ops --features pages,keynote,numbers --example mac_queue_file_routes -- …` | 0 |
| `python3 -m unittest tools.test_officecli_corpus tools.test_officecli_capability_delta tools.test_no_3rdparty_references tools.test_check_crate_boundaries tools.test_non_iwork_gate tools.test_mac_queue tools.test_gate_runner tools.test_iwa_exit_ledger tools.test_check_example_targets` | 0 |
| `python3 tools/officecli_capability_delta.py --check`, `tools/officecli_corpus.py --check` | 0, 0 |
| `python3 tools/unified_bench_matrix.py --check` (after regenerating for the new record) | 0 |
| `python3 tools/check_feature_matrix.py` | 0 (31 ok) |
| `python3 tools/iwa_exit_ledger.py --check`, `tools/mac_queue/generate.py --check` | 0, 0 |
| `python3 tools/check_crate_boundaries.py`, `tools/non_iwork_gate.py verify`, `tools/check_example_targets.py` | 0, 0, 0 |
| `bash tools/mac_queue/run_on_mac.sh --dry-run` over the generated `out/` | 0 (1153 dry-run) |

No library code, public API, dependency or capability changed. The change
adds a test-only harness module, an example and tools. For that reason the
full `litchi-ops` and `litchi` test suites, the glossary gate, the pyo3
tests, the perf-baseline check and the golden re-bless were not rerun.
