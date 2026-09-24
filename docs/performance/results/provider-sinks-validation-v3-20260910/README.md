Opt-in provider and sink performance axes

Adds positional filesystem reads, atomic saves, simulated range reads and non-seekable writes with explicit source/sink counters and correctness oracles. Per-sample vectors follow elapsed sample order. Procfs I/O includes probe overhead; simulated ranges are not remote/device measurements, and sink-copy counters exclude package-internal copying.

The default 37 cases and 201 rows remain unchanged. The selectable registry now contains 443 cases because four provider axes were added. Root ran the complete V2 library suite (366 passed, one stale enumeration failure, one ignored), verified that V3 changes only the enumeration test, and passed the corrected test, 13 provider tests, strict all-target Clippy and scoped formatting. The preparer reports a complete V3 suite with 367 passed and one ignored; its terminal summary is retained separately. Independent review passed. No performance capture or gain is claimed by this implementation batch.
