# XLSX SVG lifecycle profile results

The top-level results directory remains reserved for the freeze-gated final
run. `exploratory-before/` is a separately labeled, partially optimized source
snapshot: it contains one fresh process/sample for the 16/64/256
same-drawing lanes, raw allocator receipts, `/usr/bin/time -v` files, semantic
checks, source preimage, and binary manifest. Those receipts are exploratory
and do not populate the final report or establish a repeatable performance
comparison with the current implementation.

`exploratory-detach-before/` contains six single-sample lanes: shared and
distinct SVG targets at 16, 64, and 256 pictures. Its source snapshot and
manifests identify the implementation measured; these are not measurements of
later lifecycle corrections.

`exploratory-namespace-before/` is a separate source snapshot for the
32-picture, distinct-target, shared-root-namespace inventory probe. Its
timing and allocator receipt is exploratory. The separate raw-source diagnostic
in `../../xlsx-svg-source/retained-scope-probe.rs` counts retained SVG source
bytes and is not a timing or peak-memory measurement. Receipt `input_bytes`
counts the fixture package bytes, whereas that diagnostic reports drawing XML
bytes; those sizes must not be compared as the same metric.

Executable payloads from all exploratory bundles were removed after their
recorded SHA-256 checks; `binary.sha256`, commands, build provenance, raw
receipts, and source snapshots remain. See [`COMPACTION.md`](COMPACTION.md).
