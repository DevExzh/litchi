# 0362yc — ODP streamed-reader correction and attribution

This record corrects the returned 0362yb candidate. The streamed text sink now
clears its `quick_xml` event buffer before every read. `read_resolved_event_into`
appends into that buffer, and clearing at the top of the loop covers the early
`continue` paths for notes, non-slide elements, and nested shapes. The sink
therefore retains only the current event and slide state. A scaling regression
test covers 256 and 2,048-slide documents.

`SourceBackedPresentation::write_text_to` also restores the encrypted-member
route. A source reader cannot decrypt an encrypted `content.xml`, so that
specific admission error falls back to the existing `read_content_xml` path,
which decrypts the member and neutralizes a prologue `DOCTYPE` before the
ordinary parser runs. Other archive, callback, and source errors retain their
existing progress and freshness precedence. The encrypted source test compares
the complete sink output with the owned presentation. A second test inserts a
prologue `DOCTYPE` and compares the streamed and owned outputs byte for byte.

## Matched allocator lifecycle

The workload is the accepted 0358a ODP source: 2,000 slides, 591,185-byte
`content.xml`, a 4,194,304-byte incompressible picture, and a 4,215,428-byte
archive. The source clone is allocated before each measurement. The release
probe source is outside the repository at `/tmp/odp_stream_probe-0362yb.rs`
(5,694 bytes, SHA-256
`73daf4691cdbc49b83a9e71eb0a10dc9f988ed53a9a8c3929ef3ce09c83f5e3c`).

| Scenario | Peak | Allocated | Allocations | Reallocations |
|---|---:|---:|---:|---:|
| open | 116,001 | 200,325 | 72 | 7 |
| first slide | 116,001 | 2,714,155 | 20,087 | 21 |
| second slide after first | 990,891 | 6,369,002 | 50,099 | 7,950 |
| first text | 990,891 | 4,175,837 | 34,084 | 7,945 |
| second text after first | 990,891 | 8,257,127 | 68,097 | 15,883 |
| streamed `write_text_to` | 116,001 | 1,853,913 | 18,080 | 19 |
| repeated `write_text_to` | 116,001 | 3,507,501 | 36,088 | 31 |
| append, commit, write | 612,845 | 5,853,743 | 36,185 | 99 |

The earlier 0362yb `write_text_to` row (937,560 peak / 2,705,503 allocated)
was invalid evidence: the event buffer retained the serialized events from
the whole document. After the clear, the sink is back to 116,001 peak and
1,853,913 allocated on the same source. The scaling test also keeps peak
memory below 2 MiB and below its document-size growth bound.

Against the eager-`Arc<String>` reference from 0362ya, the current first-slide
semantic allocation is 2,513,830 bytes after subtracting the current open row,
versus 2,433,510 bytes after subtracting the eager open row: a fixed +80,320
bytes. First text has the same +80,320 delta (3,975,512 versus 3,895,192).
The identical delta across slide and text, while the current open peak is
577,230 bytes lower and no full XML string is retained by the streamed route,
attributes it to the bounded ZIP reader/Deflate decoder scratch and parser
setup, rather than a second semantic model or payload-sized XML cache. It is a
fixed per-pass cost; the repeated slide/text rows remain bounded, but their
allocation and instruction counts do not yet have an accepted 0358a paired
baseline and remain open for coordinator performance review.

The eager reference rows were: open 693,231/777,932; first slide
693,231/3,211,442; first text 1,501,373/4,673,124; and append/commit/write
1,204,071/6,431,822 (peak/allocated). The current append route is lower by
591,226 peak bytes and 578,079 allocated bytes than that reference.

## Paired 0358a reference

The accepted source-backed tail-append measurements remain the authoritative
native reference. `sb-before-1.json` reports 5,513,403 peak bytes and
13,131.943 µs median; both after legs report 6,107,267 peak bytes and
13,296.183 / 13,314.163 µs median. All 15 samples produced 4,215,645 bytes
from the 4,215,428-byte source, with the same SipHash. The original probe
source is `docs/unified-ops/results/0358a/probe_before.rs` (SHA-256
`7dad851981dec5075a67b8ee595efc9b5649cc255a7cf49207d0f29af160a2a4`); the
accepted binary hashes are recorded in
`docs/unified-ops/results/0358a/binaries.sha256`.

## Callgrind diagnostics

The matched five-route diagnostic used the same release probe lifecycle after
the event-buffer fix. Its source is `/tmp/odp_stream_probe-0362yb.rs`, its
release binary SHA-256 is
`273e9efe4bf59275acb8ecbfe975e0655135d11f81eb80ff87181ccd454b80ed`, and its
Callgrind file is `/tmp/callgrind-0362yc-odp-stream-probe.out` (221,768 bytes,
SHA-256
`e1811b1a063d034e2db84b48bcaf6984347dd5ffa2174f4605d97027966efdaf`). It
collected 1,560,857,174 instruction references. A second diagnostic added the
repeated slide, text, and sink routes: source `/tmp/odp_stream_probe-0362yc.rs`
(6,516 bytes, SHA-256
`19ea600a0fcebf340c21baae15dfb1c300b94320e0b500d4c0aed86e56c93359`), binary
SHA-256
`78e65be1fce4393e54b9f102eb5324566b9e9c437b4e48f8de5ebf5c25f447cb`, and
Callgrind file `/tmp/callgrind-0362yc-odp-stream-probe-final.out` (226,434
bytes, SHA-256
`8e67bfaee6034a5d237e931c40dc898f8efa333a3ab1a2667de141d39a7ebf5d`) with
1,984,085,064 instruction references. These are deterministic diagnostics,
not an ABBA or native qualification; no native performance claim is made by
this record.

## Validation

* `cargo test --offline -p litchi-odp --lib -- --test-threads=4` — 223 passed.
* `cargo test --offline -p litchi-odp --test source_backed -- --test-threads=4` — 12 passed.
* `cargo test --offline -p litchi-odp --test sequential_text -- --test-threads=4` — 17 passed.
* `cargo test --offline -p litchi-odp --test sequential_text_scaling -- --exact --nocapture --test-threads=1` — passed.
* `cargo clippy --offline -p litchi-odp --lib --tests -- -D warnings` — passed.
* `cargo fmt --all -- --check` — passed.

The 0358a native residual and repeated-reader performance comparison remain
open; the returned 0362yb evidence is preserved unchanged.
