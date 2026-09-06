# Combined ZIP capture/decode I/O evidence

Each alternative uses an independently indexed identical immutable archive.
Counters begin after indexing and include each path's strict layout validation.
These ten deterministic cases are correctness/I/O assertions, not latency samples.

| Method | Decoded bytes | Compressed bytes | Control calls / bytes | Combined calls / bytes | Returned-byte change |
| --- | ---: | ---: | ---: | ---: | ---: |
| Store | 0 | 0 | 4 / 71 | 2 / 41 | -42.254% |
| Store | 1 | 1 | 6 / 73 | 3 / 42 | -42.466% |
| Store | 65,536 | 65,536 | 9 / 131,143 | 3 / 65,577 | -49.996% |
| Store | 262,181 | 262,181 | 16 / 524,433 | 7 / 262,222 | -49.999% |
| Store | 1,048,613 | 1,048,613 | 40 / 2,097,297 | 19 / 1,048,654 | -50.000% |
| Deflate | 0 | 7 | 8 / 133 | 5 / 80 | -39.850% |
| Deflate | 1 | 9 | 10 / 153 | 5 / 82 | -46.405% |
| Deflate | 65,536 | 65,558 | 13 / 131,251 | 6 / 65,631 | -49.996% |
| Deflate | 262,181 | 262,266 | 22 / 524,667 | 9 / 262,339 | -49.999% |
| Deflate | 1,048,613 | 1,048,938 | 58 / 2,098,011 | 21 / 1,049,011 | -50.000% |

Control reads decoded bytes and then captures/compares compressed bytes. Combined
captures once, decodes once and returns both validated outputs. Every case preserves
decoded bytes and exact compressed token data; the token survives dropping the
source archive and publishes through the preservation writer without recompression.
The pre-existing destination member survives that publication.

Both operations retain complete decoded and compressed payloads. No allocation
count, peak-memory reduction, latency or production PPTX speedup is established.
Empty and one-byte cases include substantial metadata overhead; larger cases remove
roughly half the returned source bytes. OPC/PPTX adoption still requires explicit
budget, cache, token lifetime, semantic-validation and source-freshness integration.
