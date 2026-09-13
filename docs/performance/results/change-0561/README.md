# change-0561 evidence packet: repeated positional reads in file-backed OPC

Change record: pending. This packet is **attribution only** — it retains a
reproducible diagnostic, no production change and no performance claim.

## What it shows

Every capture is one perf-harness child with `--warmup 0 --samples 1`, traced
with `strace -f -e trace=pread64`, reading a file-backed OOXML package.

| Capture | `pread64` calls | distinct (offset, length) ranges | calls that re-read a range | immediate back-to-back duplicates | mean read |
| --- | ---: | ---: | ---: | ---: | ---: |
| `docx_file_source_open` | 266 | 23 | 91.4% | 22.9% | 219.7 B |
| `docx_file_source_full_text` | 282 | 23 | 91.8% | 23.0% | 228.3 B |
| `pptx_file_source_open` | 13,592 | 1,262 | 90.7% | 25.0% | 120.6 B |
| `pptx_file_source_list_slides` | 16,792 | 1,262 | 92.5% | 25.0% | 125.6 B |
| `pptx_file_source_selected_slide` | 13,608 | 1,262 | 90.7% | 25.0% | 120.6 B |

Roughly nine of every ten positional reads on these paths re-read bytes the same
child already read, and a quarter of all calls repeat the immediately preceding
range verbatim. In the PPTX captures, individual ranges are read 6, 10, 12, 20
and in one case 24 times. The read-size histogram for
`pptx_file_source_selected_slide` is 6,760 calls of 16 bytes or fewer, 3,400 of
32 or fewer, 2,136 of 256 or fewer and 1,232 of 1,024 or fewer.

The repeating per-entry shape is a 30-byte read, a variable-size read nearby,
then a 16-byte read, then **the same 16-byte read again at the identical
offset**.

Note that `pptx_file_source_open` already accounts for almost all of the calls
that `pptx_file_source_selected_slide` makes, so the repetition is in the open
and index path rather than in the slide query.

## Per-operation syscall isolation

`isolation/` holds `strace -f -c` summaries for the same binary at 5 and at 25
samples. Differencing the two and dividing by 20 isolates the per-operation
syscall count from process start-up and the harness's in-process corpus
generation:

| Case | `statx` per operation | `pread64` per operation |
| --- | ---: | ---: |
| `docx_file_source_full_text` | 598.0 | 262.0 |
| `pptx_file_source_selected_slide` | 9,696.0 | 13,588.0 |

The owned-source selectors `xlsx_source_first_cell` and
`opc_source_open_main_read` isolate to zero of both, which confirms the counts
come from the file adapter rather than from the harness.

## Reproduce

```sh
cargo build --release --locked --manifest-path tools/perf-baseline/Cargo.toml
strace -f -e trace=pread64 -o capture.pread.txt \
  tools/perf-baseline/target/release/litchi-perf-baseline \
  --case pptx_file_source_open --warmup 0 --samples 1 --json /dev/null
python3 -B docs/performance/results/change-0561/summarize_reads.py \
  --traces <directory holding the capture> --output read-repetition.json
```

`summarize_reads.py` reads plain or gzipped captures and reports only properties
of the retained trace; it measures no time.

## What is not here

No production change, no latency or resource measurement, and no claim that the
repeated reads are removable — only that they occur. The captures are warm-cache
and single-sample; they count syscalls, not physical device I/O.
