# RTF logical-tail append evidence

Status: opt-in harness and correctness tranche; no performance claim

The RTF performance harness now distinguishes logical append to an existing
document from streaming creation. It adds these selectors:

- `rtf_logical_tail_append`
- `rtf_logical_tail_noop_save`

Both selectors use the existing default-formatted plain lifecycle corpus, not
the forward-only creation writer. `--semantic-shape tiny,medium,large` appends
4, 64, or 256 bounded one-run plain paragraphs. The deterministic append text,
input byte count, output limits, inserted-byte limits, and durable JSON limits
are all generated before timed samples.

The timed interval includes borrowed paragraph staging, the logical-tail
candidate commit, its validation, and sequential publication to a non-seek
hashing sink. The sink accepts at most a fixed 16 KiB window per `Write` call,
retains zero output bytes, hashes every accepted byte, and records accepted
bytes, write calls, and largest write. The append summary additionally records
source, caller-input, inserted, and output bytes plus paragraph/run counts.
The API's validated candidate snapshot is intentionally retained by the
transaction; the sink window is therefore not a transaction-memory or RSS
bound.

Before timing, the runner verifies deterministic output and complete reopen
projection, exact sequential bytes, the empty-append shared-snapshot no-op,
in-memory patch replay and inverse, durable deterministic-JSON
encode/decode/apply and inverse, and changed/foreign source conflict refusal.
The result's `sink.rtf_tail_append` object exposes those five gate booleans so
machine-readable reports cannot be mistaken for timing-only output.

Representative one-sample smoke output from the current source is:

| Shape | Existing source bytes | Appended paragraphs | Input bytes | Inserted bytes | Published bytes |
|---|---:|---:|---:|---:|---:|
| tiny | 1,304 | 4 | 168 | 273 | 1,577 |
| medium | 10,808 | 64 | 2,816 | 4,421 | 15,229 |
| large | 540,008 | 256 | 11,008 | 17,413 | 557,421 |

The large append publishes through 35 bounded sink writes; the tiny and medium
artifacts fit in one 16 KiB sink window. These values are deterministic corpus
and sink counters, not latency results. No release-build CPU-pinned ABBA run,
allocation profile, peak-RSS measurement, or speedup claim is introduced here.

Run the tranche with:

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 30 --semantic-shape tiny,medium,large \
  --case rtf_logical_tail_append,rtf_logical_tail_noop_save \
  --json target/perf/rtf-logical-tail-append.json
```
