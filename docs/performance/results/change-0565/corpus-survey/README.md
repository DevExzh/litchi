# Corpus survey: CFB geometry and BIFF globals read schedules

This survey exists because it changed the design of change 0565. The first draft
windowed the globals scan only after the contiguous BoundSheet8 run had ended, so
that the smallest `lbPlyPos` was always known before any window was issued and no
byte past the globals end could ever be read. The survey showed that schedule is
worth almost nothing: the first BoundSheet8 record is late in *record* order but
early in *byte* order — record 526 of 621 but byte 20,783 of 551,377 for
`ConditionalFormattingSamples.xls` — so an exact phase running to the end of the
run still costs one read per record for 85% of the records.

The implemented schedule therefore starts filling right after a short exact
prologue, and accepts a bounded read past the globals end. The survey quantifies
that bound over the corpus, and models both the chosen and the rejected fill
sizes through the same generator so the comparison that drove the design is
reproducible.

## Contents

| Path | What it is |
| --- | --- |
| `survey.py` | The generator. Deterministic; standard library plus `olefile`. |
| `survey.json` | Per-file records: CFB geometry, BIFF globals framing, and both modelled read schedules. |
| `survey.md` | The tables and the summary, generated from the same run. |

## Replay

```sh
python3 -B docs/performance/results/change-0565/corpus-survey/survey.py \
  --repo . --out docs/performance/results/change-0565/corpus-survey
```

The rejected 4 KiB draft is reproduced from the same generator with
`--first-window 4096`, and the superseded clamp-first draft with
`--prologue-records 0`.

The output is byte-identical on rerun; paths are recorded repository-relative so
the artifact does not depend on how `--repo` was spelled.

## What it establishes

Scope: 104 files under `test-data/ole`, of which 98 open as CFB and 54 carry a
`Workbook` or `Book` stream, plus 72 `*.xls` fixtures elsewhere as an appendix.

- **Read counts.** Aggregate logical stream reads for one open of all 54 files:
  **8,543 today, 434 modelled**, or 5.08%.
- **No file gets worse.** The prologue fetches each record's payload together
  with the next record's header, so `password.xls`, the only encrypted fixture
  and the only file where the globals scan stops early, stays at 2 reads. An
  earlier draft that read the payload separately cost it 3.
- **Over-read.** 32,532 bytes across the corpus, worst 3,949
  (`WithChart.xls`), with a hard ceiling of one fill. Slack between the globals
  end and the smallest `lbPlyPos` is **zero in all 54 files**, so once the clamp
  is known a fill stops exactly at the globals end, and every over-read byte is
  read before the first BoundSheet8 is framed.
- **The clamp is never dropped on this corpus.** Production drops it only when
  the globals frame past the smallest declared sheet position, which no fixture
  does. A fill that merely overtakes the clamp does not drop it, because once
  the bytes are resident no fill is required.
- **Why 512 bytes and not 4,096.** The same generator at `--first-window 4096`
  gives 331 reads but 109,045 bytes of over-read with a worst file of 11,010.
  512 costs about two more reads per file and removes 70% of the over-read, and
  it is one CFB sector.
- **FILEPASS position.** Every encrypted fixture in the repository carries
  FILEPASS at record index 1, byte offset 20: `test-data/ole/xls/password.xls`,
  `test-data/poi/test-data/spreadsheet/35897-type4.xls` and
  `test-data/poi/test-data/spreadsheet/xor-encryption-abc.xls`. A four-record
  prologue covers all of them with no payload byte read.
- **BoundSheet8 shape.** The run is contiguous in 54 of 54 files and `lbPlyPos`
  is strictly ascending in 54 of 54. This is **empirical, not normative**:
  `[MS-XLS]` 2.4.28 imposes no ordering, which is why the production clamp is a
  running minimum rather than a minimum over the whole stream. No fixture has a
  corrupt `lbPlyPos`, so the corrupt-position tests in `litchi-xls` are
  synthetic and are the only coverage of that path.
- **A separate CFB lead, not taken in this change.** `load_fat` reads one sector
  per call while `read_sectors_batched` already exists in the same file. Batching
  contiguous FAT runs saves reads on 5 of 98 files, 30 reads in total, dominated
  by `picture.doc` (23 reads to 3). `load_minifat` batching saves nothing on this
  corpus. The first FAT sector is physically adjacent to the header in 34 of 98
  files.

## What is not here

Every read count is a model of stream-level logical reads and of the physically
contiguous spans they split into. None is a measured syscall count. No timing,
allocation or cold-cache result is in this directory.
