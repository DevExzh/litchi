# Reproducing the 0683 selected-record probe

Build the same probe source and Cargo.lock inside two checkouts: baseline
`0a292428d598b5235ca0075f2b1a5d5b931af31d` and the candidate. Path dependencies
resolve five levels above this directory. The exact commands, binary hashes,
source hashes and raw-output hashes are retained by `measure.py`.

The retained measurement run used these paths:

```sh
# Copy this probe directory into the identical relative location of the
# baseline checkout before building. Keep its Cargo.lock identical.
cargo build --release --offline --locked \
  --manifest-path /home/zhuhe/code/litchi-0683-before/docs/performance/results/change-0683/probe/Cargo.toml \
  --target-dir /home/zhuhe/code/litchi-target-0683-before
cargo build --release --offline --locked \
  --manifest-path /home/zhuhe/code/litchi/docs/performance/results/change-0683/probe/Cargo.toml \
  --target-dir /home/zhuhe/code/litchi-target-0683-after
python3 generate.py /home/zhuhe/code/litchi-0683-corpus
python3 measure.py baseline /home/zhuhe/code/litchi-target-0683-before/release/xlsx0683 /home/zhuhe/code/litchi-target-0683-after/release/xlsx0683 /home/zhuhe/code/litchi-0683-corpus
python3 measure.py candidate /home/zhuhe/code/litchi-target-0683-before/release/xlsx0683 /home/zhuhe/code/litchi-target-0683-after/release/xlsx0683 /home/zhuhe/code/litchi-0683-corpus
```

Run from this directory; adjust the explicit paths and CPU 12 in the driver for
another host. The baseline matrix/A/A runs before building the candidate so
build activity does not overlap measurements. `generate.py` uses deterministic
ZIP metadata and emits a corpus manifest with member/archive sizes, hashes,
physical record count N and shared-string reference count K. Generate once and
reuse the same files for both binaries. ZIP bytes can depend on Python/zlib
versions; retained hashes bind the actual input.

Cases are dense numeric (65,536 records), sparse, 4 KiB inline text, repeated
shared strings, formulas with cached numbers, dense records followed by valid
`pageMargins` that forces fallback, and an unterminated late tail that refuses.
The public raw scanner separately proves eligible/ineligible routes. The real
POI `no_drawing_patriarch.xlsx`, worksheet `Лист 1`, is a materialized fallback
control. The differential compares five public read routes within each binary
and compares before/after outputs exactly, including actual callback counts on
refusal. Digests use FNV-1a over canonical Debug/address strings as a test oracle,
not as a cryptographic integrity proof; corpus and artifact bindings use SHA-256.

## Measurement boundaries

The driver pins each child to CPU 12. Each timing group has three warmups and
20 samples per leg. Baseline A/A uses A1, A2; candidate comparisons use A1, B1,
B2, A2 within each group. No geometric mean or tail-latency claim is implied.
`*-cold` means a newly opened package on each sample, with ordinary warm OS
file caching; it is not physical cold-cache evidence. `*-selected` prepares
an owner outside the timer and repeats queries on it; eligible scans remain
uncached, while an ineligible first call warms its fallback store. `*-warm`
materializes the store before measurement. Warm late-refusal rows are omitted
because setup itself refuses; the differential still records that outcome.

Semantic rendering/digest work runs separately from timing and allocation
counts. Timed rows contain actual counts and success/refusal; mismatched or
missing samples must fail the final audit. `cells` timing includes dropping the
returned vector; allocation measurements retain that vector through gauge
reads. Prepared owners also survive gauge reads. Newly opened cold owners are
dropped before the retained gauge, while their returned cell vectors survive.

The standalone counting allocator delegates storage to `System`; this is
measurement-only unsafe code, outside production crates. A successful alloc
or realloc counts as one allocation call; cumulative requested bytes add the
full new realloc request. Logical live bytes change by the size delta only
after success. Peak/live gauges exclude allocator headers, transient internal
realloc copy overlap, and RSS. Deallocation counters cover explicit dealloc
calls, not allocator-internal realloc release. No admission/budget claim follows.

`run-diff.sh`, `run-alloc.sh`, `run-aa.sh`, and `run-timing.sh` remain smaller
manual alternatives for generated cases. `measure.py` is the authoritative
recorded driver and also includes the real control. `summarize.py` reads the
headerless key=value TSV rows; the final packet audit supplies cross-leg and
source-binding checks.
