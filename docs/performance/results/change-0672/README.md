# Change 0672 evidence packet

Record: [`../../0672-xlsx-stored-cell-allocation.md`](../../0672-xlsx-stored-cell-allocation.md).
This packet measures the stored-route `SourceWorksheet::cells` reservation
residue from queue row 15 of change 0651. `performance_claim: none`.

## Provenance

| | |
|---|---|
| base | `5fa92d7ced5a78f3c2c84a6afe9d0ba404253717` |
| before | `/home/zhuhe/code/litchi-worktrees/before-5fa92d7ce` |
| after | `/home/zhuhe/code/litchi-worktrees/0672` |
| branch | `perf/0672-xlsx-read-allocation` |
| host | AMD EPYC 9R45, 32 cores, Linux 7.0.0-1012-aws, x86_64 |
| toolchain | rustc 1.95.0 (59807616e 2026-04-14), cargo 1.95.0 |
| timing pin | `taskset -c 17`, 20 warmups and 30 samples per leg |
| build | release probe, two build jobs; dev/test validation used `CARGO_PROFILE_DEV_DEBUG=0`, `CARGO_PROFILE_TEST_DEBUG=0`, two jobs |

The release probe binaries were built separately from the before and after
path dependencies and staged outside their Cargo target directories. Their
SHA-256 values are in [`binaries.sha256`](binaries.sha256). The generated
workbook's SHA-256 is retained beside them.

## Corpus and scope

* `test-data/poi/test-data/spreadsheet/no_drawing_patriarch.xlsx`, worksheet
  `Лист 1`, 75,770 stored cells, stored-route materialization.
* `corpus/dense-wide-probe.xlsx`, two generated 256×256 integer worksheets,
  used to exercise the power-of-two dense result and the selected scanner.

The counting allocator reports allocation calls, allocated bytes, peak live
bytes and retained live bytes for `cells`, `cells-warm`, and the unchanged
`visit-warm` control. The before and after differential tables use the probe's
four-way value/refusal oracle. `timing/summary.txt` records the paired timing
and same-binary A/A floor; raw samples and their hashes are under
`timing/runs/`.

## Reproduction

The probe manifests use absolute path dependencies so each leg is explicit:

```sh
for leg in before after; do
  cargo build --release --manifest-path probe/Cargo.$leg.toml
done
```

Cargo requires the conventional filename `Cargo.toml`; the measurement run
copied each manifest into a temporary probe directory before building. The
probe commands were:

```sh
xlsx0672 corpus corpus/dense-wide-probe.xlsx 2 256 256
xlsx0672_alloc cells-warm test-data/poi/test-data/spreadsheet/no_drawing_patriarch.xlsx 'Лист 1'
xlsx0672 diff test-data/poi/test-data/spreadsheet/no_drawing_patriarch.xlsx corpus/dense-wide-probe.xlsx
```

## Contents

| path | purpose |
|---|---|
| `decision.json` | scoped disposition, evidence and limitations |
| `log-sections.md` | four coordinator log sections |
| `probe/` | copied and repointed 0642 probe plus counting allocator |
| `corpus/dense-wide-probe.xlsx` | generated dense control workbook |
| `alloc/alloc-raw.txt` | before/after allocation counts |
| `differential/diff-before.tsv`, `differential/diff-after.tsv` | three-worksheet four-way oracle |
| `timing/summary.txt` | A/A floor and paired warm timing |
| `timing/runs/` | raw timing samples |
| `timing/runs-sha256.txt` | raw timing file hashes |
| `binaries.sha256` | measured binary and corpus hashes |
| `gates.txt` | validation commands and results |
