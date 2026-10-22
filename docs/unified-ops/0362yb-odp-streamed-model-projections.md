# 0362yb — ODP streamed semantic projections

This candidate continues the 0358a ODP source-owner repair. The source-backed
owner still validates `content.xml` at open, and `content_xml()` still provides
the complete lazy raw projection. Semantic slide, indexed-slide, and text
projections now consume the already verified ZIP member through a bounded
`BufRead`. The ordinary parser is generic over that reader; its existing
string entry points remain wrappers over `&[u8]`. The fused text sink uses the
same reader shape and a reusable event buffer.

The source reader is always drained and verified. A parser or source failure is
reported with the existing source/freshness precedence. Encrypted content,
which cannot use the streaming reader, keeps the established materialized
fallback. Selected-slide parsing still runs the complete document traversal,
including validation after the selected page.

## Matched allocation probe

The workload is the 0358a ODP source: 2,000 slides, 591,185-byte
`content.xml`, a 4 MiB incompressible picture, and a 4,215,428-byte archive.
The source clone is allocated before each measurement. Release build, one
probe process, counting allocator; the probe source is retained outside the
repository as `/tmp/odp_stream_probe-0362yb.rs` (5,694 bytes,
SHA-256 `73daf4691cdbc49b83a9e71eb0a10dc9f988ed53a9a8c3929ef3ce09c83f5e3c`).

| Scenario | Peak bytes | Allocated bytes | Allocations | Reallocations |
|---|---:|---:|---:|---:|
| open | 116,001 | 200,325 | 72 | 7 |
| first slide | 116,001 | 2,714,155 | 20,087 | 21 |
| first text | 990,891 | 4,175,837 | 34,084 | 7,945 |
| streamed `write_text_to` | 937,560 | 2,705,503 | 18,080 | 31 |
| append, commit, write | 612,845 | 5,853,743 | 36,185 | 99 |

For comparison, the paired eager-`Arc<String>` reference from the preceding
candidate measured open at 693,231 peak / 777,932 allocated, first slide at
693,231 / 3,211,442, first text at 1,501,373 / 4,673,124, and append/commit/
write at 1,204,071 / 6,431,822. The streamed semantic pass therefore retains
577,230 fewer peak bytes for first slide, 510,482 fewer for first text, and
591,226 fewer for append/commit/write. After subtracting each open row, the
semantic allocation difference is +80,320 bytes for both first slide and first
text; the former +691,175-byte semantic reread from 0362ya is removed.

The accepted 0358a source-backed ABBA remains the separate native reference:
the former route peaked at 5,513,403 bytes and the 0358a route at 6,107,267
bytes over 15 tail-append samples, with identical 4,215,645-byte output. Its
median time was 13,131.943 µs before and 13,296.183/13,314.163 µs on the two
after legs. This candidate's allocator probe is a matched lifecycle check,
not a replacement for that native ABBA or a callgrind measurement.

## Validation

The following checks passed on the candidate tree:

* `cargo test --offline -p litchi-odp --lib -- --test-threads=4` — 223 passed.
* `cargo test --offline -p litchi-odp --test source_backed -- --test-threads=4` — 12 passed.
* `cargo test --offline -p litchi-odp --test sequential_text -- --test-threads=4` — 16 passed.
* `cargo clippy --offline -p litchi-odp --lib --tests -- -D warnings` — passed.

Native ABBA and callgrind for this candidate remain open for coordinator
review. The prior 0362ya evidence is unchanged.
