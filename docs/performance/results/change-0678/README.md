# Change 0678 evidence packet

This packet prices the bounded weighted cache design in
[`../../0678-xls-query-cache-design.md`](../../0678-xls-query-cache-design.md).
It contains no production change and no performance claim.

## Provenance

| | |
|---|---|
| repository revision | `389167b38c8b43c84c0d1799369e17360a59209c` |
| corpus | 126 `.xls`/`.xlt` files under `test-data/` |
| probe | [`probe/xls_index_shape.py`](probe/xls_index_shape.py) |
| Python | 3.14.4 |
| stream reader | `olefile` 0.47 |
| source build | none; no Cargo command was run |
| result class | wire-level logical size and frame-work census |

The probe extracts a Workbook/Book stream with `olefile` and does not prove
that the Rust source-backed owner admits a fixture. It applies the current
cheap SST count/length checks when charging SST locator weight; malformed or
encrypted declarations remain in the per-fixture JSONL with
`invalid_reason`. Fixture SHA-256 values are included in every JSONL row.

## Reproduce

```sh
python3 docs/performance/results/change-0678/probe/xls_index_shape.py \
  test-data \
  --output docs/performance/results/change-0678/probe/index-shape.jsonl \
  2> docs/performance/results/change-0678/probe/summary.json
```

The checked-in output reports 126 analyzed physical streams, 372 worksheet
substreams, 127,072 recognized stored-cell occurrences at 126,152 distinct
coordinates, and 3,073,536 logical worksheet-index bytes under the proposed
24-byte occurrence slot plus 64-byte entry charge. Every occurrence is
retained because an earlier duplicate can raise a target-specific SST read/decode or
formula failure before a later duplicate wins. An out-of-range SST index instead
produces a typed cell-error value that a later duplicate can overwrite. The largest worksheet is
`test-data/poi/test-data/spreadsheet/54016.xls`, sheet `Sheet1`: 38,950
occurrence slots and 934,864 bytes. Its SST has 28 segments and 7,893 unique
entries, charging 127,024 bytes at 24 bytes per segment, 16 bytes per entry,
and 64 bytes overhead. Together they charge 1,061,888 bytes.
The fixture hash recorded by the JSONL row is
`2e050f1fbb31868b097aa6d4d0fe0a16af8e39c252af82d01cd8f93c4f9a911a`.

The admissible SST groups in this wire census total 290,520 logical bytes;
the largest is `54016.xls`. Four declared SST groups fail the probe's
count/length admission check and receive zero cache weight. The semantic
source-backed open count remains the 113 fixtures recorded by change 0648;
this packet does not replace that differential.

## Integrity

The output files are deterministic for the checked-in corpus. Recompute their
hashes after rerunning the command if the corpus changes:

```text
probe/xls_index_shape.py  8094f0b86c98ae151e022aca71f3cd462cd881fa20adf208a4d2ba01964e9980
probe/index-shape.jsonl   6f559c014369b528d5e91f2f640d1024ceadd25abcf34e56f41b90990ded74ce
probe/summary.json        f835ebdb892550ba769830ba1c458596993dd0806ca20e7004380763b0e67ea2
```

The earlier repeated-query price is from the accepted 0605 evidence: a second
`54016.xls` query adds 25 reads, 615,822 bytes, and 19 source observations;
the proposed index serves the target from one indexed frame and any required
SST read. This packet does not claim that model as a new timing result.
