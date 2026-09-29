# 0832 column-assignment promotion guard

This packet supplements the frozen nine-case 0832 experiment. It exercises
the parser and resolver when a worksheet contains more than one physical
`<col>` record, while keeping the parent experiment's sources, driver, and
nine-case matrix unchanged. Comparative timing is limited to
`xlsx_real_file_ordinary_save_edit`; the lifecycle selector is used only as a
one-sample output-preservation admission check.

The two generated inputs start from the pinned
`test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx` fixture.
The generator copies ZIP records at the raw local-payload level and replaces
only `xl/worksheets/sheet1.xml`:

* `disjoint-two-records.xlsx` retains the original complete record for column
  B and adds one disjoint complete record for column D.
* `overlap-128-wide-complete.xlsx` contains 128 valid complete records with
  nested overlap from each starting column through XFD. Every record supplies
  width, style, all modeled boolean flags, and outline level.

The generator independently records compressed and uncompressed payload hashes
for every member. The reader repeats the ZIP payload comparison and verifies
that bytes before and after the `<cols>` element, including the eight-cell
sheet data, MCE declaration, date styles, shared strings, and autofilter
markers, remain present. The lifecycle harness checks and then deletes its
published file, so the supplement's output oracle is cross-leg and cross-mode
equality of the retained one-sample `output_sha256`. It is not an independent
Office validation or a retained output-member comparison.

The retained stages ran in this order. These are historical commands, not
rerun-in-place commands: outputs are write-once and runtime roots have been
removed. Numbered attempt receipts also retain the three preflight executions.

```text
python3 -B docs/performance/results/change-0832/run_guard.py generate_fixtures.py
python3 -B docs/performance/results/change-0832/run_guard.py reader.py --preflight
python3 -B docs/performance/results/change-0832/run_guard.py driver.py prepare
# The parent 0832 driver performs its unchanged comparative capture here.
python3 -B docs/performance/results/change-0832/run_guard.py driver.py qualification
python3 -B docs/performance/results/change-0832/run_guard.py reader.py --admit
# Retained missing-SOURCES failure; prepare its additive correction.
python3 -B docs/performance/results/change-0832/run_guard.py replay_reader.py prepare
python3 -B docs/performance/results/change-0832/run_guard.py replay_reader.py --admit
# Retained missing-write_once failure; supply both omitted globals.
python3 -B docs/performance/results/change-0832/run_guard.py replay_reader_v2.py prepare
python3 -B docs/performance/results/change-0832/run_guard.py replay_reader_v2.py --admit
python3 -B docs/performance/results/change-0832/run_guard.py driver.py capture
python3 -B docs/performance/results/change-0832/run_guard.py replay_reader_v2.py
python3 -B docs/performance/results/change-0832/run_guard.py replay_reader_v2.py --check
```

`prepare` must run after the parent packet has produced both before and after
native/observer binaries and before any comparative capture. It freezes those
binary hashes, both candidate source archives, the original and generated
fixtures, the parent packet inputs, and all three guard scripts. The later
stages use those binaries directly and do not build, install sources, invoke
Git, or alter the parent packet.

The protocol is 16 qualification reports (two phases × two fixtures × two
legs × two binaries), then 24 native edit reports (six A/B and B/A blocks ×
two fixtures × two legs) and eight observer edit reports (two A/B and B/A
blocks × two fixtures × two legs). Native reports use 500 samples and three
warmups; observer reports use three samples and no warmup; all workloads run
on CPU 12. The reader derives nearest-rank within-process p50/p95/p99,
midpoint block medians, and six-block paired bootstrap intervals with seed
832128, 10,000 resamples, and ranks 250/9749.

`analysis.json` always has structural status `pass` when the packet is
complete. Any native p50 lower-confidence-bound increase above 1.05, paired
RSS increase above 1.05, or observer increase in allocation calls, requested
bytes, raw region-peak bytes, or peak-above-entry bytes is retained in `regression_flags`; the parent
adoption decision must reject a nonempty list. Observer elapsed times are
diagnostic only, and this packet makes no nonempty-column-action, cross-format,
or whole-save performance claim.
It also makes no cache-miss or cache-locality claim.

The owned runtime scratch directory is
`../litchi-fs-0832-promotion-guard`; remove it only after `reader.py` and
`reader.py --check` pass. Retain the fixture archives, command receipts,
logs, RSS files, reports, and `analysis.json` as packet evidence.

The root launcher retains every script attempt, terminal status, log, and
source hashes. Reader admission checks all 16 qualification reports before
comparative capture. After root cleanup, replay uses the retained binary
descriptors and explicit scratch-removal receipt.

Two admission failures are retained: the frozen reader omitted `SOURCES`
and `write_once`. `replay_reader_v2.py` supplies only those two globals,
checks the complete unresolved-global set, and binds the original freeze,
qualification reports, earlier correction, and both failure receipts. All
16 reports passed admission before comparative capture; no workload was
rerun and no statistic, threshold, parser, or frozen script was changed.
Replay the completed packet with `replay_reader_v2.py --check`; the original
reader remains as frozen evidence and cannot run standalone. The numbered
launcher attempts preserve the actual historical command order.

The compact analysis field `full_report_payloads_retained: false` means that
raw reports are not embedded in the summary JSON. All original report files
are retained separately under `qualification/`, `native/`, and `observer/`.

Replay the completed packet without creating new evidence files:

```sh
python3 -B docs/performance/results/change-0832/promotion-guard/replay_reader_v2.py --check
```
