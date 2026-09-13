# 0562: one data-descriptor read per verified entry read

Status: retained. `performance_claim: none` — this record carries deterministic
syscall counts and paired latency, not a registry claim.

OLE2 and OOXML remain the active priority. ODF optimization stays deferred until
that goal completes; iWork is excluded.

## What was removed

`ZipVerifier::read` performs its completion work when
`read == 0 || self.size >= self.wayfinder.uncompressed_size_hint()`. For an entry
framed with a data descriptor, that work resolves the descriptor from the source.

`Read::read_to_end` must observe one `Ok(0)` after a member's payload is
complete, and the archive read path gives it a read limit of
`uncompressed_size + 1` so `Take` never short-circuits. Both halves of the
condition therefore fire for one member read: first on the read that completes
the payload, then again on the terminating read. The second resolution read the
same 16 bytes at the same offset, immediately after the first.

The verifier now remembers the observation it last accepted and returns early
when the current CRC and size are identical to it. A read that adds bytes
changes the observation and is verified normally.

## Why it is exactly equivalent

The skipped call would compute `expected` and `observed` from the same
`self.crc`, the same `self.size` and the same immutable entry, and compare them
the same way. Verifying an identical pair twice can only reach the same
conclusion. The memo is written only after a successful verification, so a
failure still propagates — and stops `read_to_end` before a second attempt could
occur anyway.

Change [0561](0561-opc-repeated-positional-reads.md) records the separate proof
that the descriptor's CRC cannot differ from the entry's retained CRC; this
change does not rely on that, only on the idempotence of the comparison.

## Measured effect

One perf-harness child per case, `--warmup 0 --samples 1`, traced with
`strace -f -e trace=pread64`. Counts are properties of the retained traces.

| Case | `pread64` calls | change | immediate back-to-back duplicates |
| --- | ---: | ---: | --- |
| `docx_file_source_open` | 266 → 224 | **−15.79%** | 22.9% → 8.5% |
| `docx_file_source_full_text` | 282 → 236 | **−16.31%** | 23.0% → 8.1% |
| `pptx_file_source_open` | 13,592 → 10,216 | **−24.84%** | 25.0% → 0.2% |
| `pptx_file_source_selected_slide` | 13,608 → 10,228 | **−24.84%** | 25.0% → 0.2% |

The distinct-range count is unchanged at 23 and 1,262, so no range stopped being
read; only repeats were removed. Mean read size rises from 120.6 to 155.2 bytes
on the PPTX captures because the removed calls were the smallest ones.

Dividing the bytes no longer read by the calls no longer made identifies exactly
what was removed:

| Case | calls removed | bytes removed | bytes per removed call |
| --- | ---: | ---: | ---: |
| `docx_file_source_open` | 42 | 672 | 16.0 |
| `docx_file_source_full_text` | 46 | 736 | 16.0 |
| `pptx_file_source_open` | 3,376 | 54,016 | 16.0 |
| `pptx_file_source_selected_slide` | 3,380 | 54,080 | 16.0 |

Every removed call was exactly sixteen bytes — the data-descriptor width for a
non-ZIP64 entry — so the change removed the duplicate descriptor resolutions and
nothing else.

### Paired latency

An A1/G1/G2/A2 matrix over nine OOXML selectors with ASLR disabled, 3 warmups
and 30 samples per child: 36 children, 9 case/corpus rows, 36 statistic
comparisons, **17 improving in both directions and zero adverse in both
directions by more than 5%**. The baseline binary was built from the tree this
change is applied to, so the matrix isolates this change rather than the batch.

| Selector | p50, first direction | p50, second direction |
| --- | ---: | ---: |
| `opc_file_source_open` | −5.21% | −1.70% |
| `pptx_file_source_list_slides` | −3.51% | −3.68% |
| `docx_file_source_open_full_text_lifecycle` | −2.71% | −0.46% |
| `pptx_file_source_open` | −1.64% | −1.47% |
| `docx_file_source_open` | −0.34% | −2.03% |
| `pptx_file_source_selected_slide` | −1.29% | +0.90% |
| `xlsx_file_open` | −1.21% | +0.04% |
| `docx_file_source_full_text` | −0.09% | +0.80% |
| `opc_file_eager_open` (owned control) | +1.52% | −0.92% |

A 24.84% cut in syscalls buys only a few percent of warm-cache latency, which is
the expected shape: a 16-byte `pread64` served from the page cache is cheap. The
syscall count is the durable result, and it is the quantity that matters for
cold-cache reads, for caller-supplied range sources where each read is a round
trip, and for sandboxes that filter or trace syscalls — none of which this host
measures. The owned-source control behaves as predicted, moving in neither
direction consistently.

The DOCX captures retain 8.1–8.5% immediate duplicates after the change. They
are a different phenomenon: the retained trace shows one 840-byte read at offset
64 repeated two to four times consecutively, and the per-member descriptor reads
now appear exactly once each. That repeated span is a structural member being
re-read rather than a descriptor, which change 0561 already lists as an open
follow-up. It is recorded here, not addressed.

## Correctness evidence

`crates/soapberry-zip/tests/descriptor_read_once.rs` adds three tests that count
positional reads through an instrumented `ReaderAt`:

- `one_verified_read_resolves_the_data_descriptor_once` asserts exactly one
  payload-bearing read at the descriptor offset and **zero** immediate duplicate
  ranges for a whole verified read;
- `repeated_verified_reads_each_resolve_the_descriptor_once` asserts three
  verified reads of the same entry resolve it three times, so the memo is
  per-verifier and does not leak across reads;
- `a_corrupt_descriptor_still_fails_the_verified_read` flips a bit in the
  descriptor's CRC field and requires `InvalidData`.

The first two were confirmed to **fail** with the memoization removed and pass
with it, so they are a real regression guard rather than a restatement.

`soapberry-zip` passes its full suite with zero failures.

## Limitations

No cold-cache, physical-device, remote/range-source, peak-RSS or allocation
result is claimed. The captures are warm-cache and single-sample; they count
syscalls, not physical device I/O. The remaining ~88% of reads that re-read a
range are untouched by this change — memoizing the resolved descriptor and the
local-header framing per entry, coalescing small members' spans, and giving
structural members a retention policy all remain open, and change 0561 records
what each would remove.
