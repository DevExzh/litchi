# OPC combined capture: deterministic I/O assertions

Identical, independently opened packages; counters begin after opening and end
after cold read/authorization. Maximum provider return is 65,536 bytes. Tests
also compare exact decoded/compressed bytes and whole published outputs.

| Method | Decoded bytes | Compressed bytes | Control calls | Combined calls | Control returned bytes | Combined returned bytes |
|---|---:|---:|---:|---:|---:|---:|
| Store | 0 | 0 | 10 | 8 | 214 | 184 |
| Store | 1 | 1 | 12 | 9 | 216 | 185 |
| Store | 65536 | 65536 | 15 | 9 | 131286 | 65720 |
| Store | 262181 | 262181 | 22 | 13 | 524576 | 262365 |
| Deflate | 0 | 7 | 14 | 11 | 276 | 223 |
| Deflate | 1 | 9 | 16 | 11 | 296 | 225 |
| Deflate | 65536 | 65558 | 19 | 12 | 131394 | 65774 |
| Deflate | 262181 | 262266 | 28 | 15 | 524810 | 262482 |

No latency, RSS, allocation-peak or complete PPTX speedup claim.
