# 0362dvib — DrawingML native raw probe repair

## Scope

This repair is limited to the source-backed adapters used by the paired
DrawingML raw qualification. It does not change a format owner, package
placement, or the owner replay implementation.

The probe now:

- parses each source and changed fixture into a typed `PreparedValue` before
  timed writer rows; `entry_read` times only the reader, `vec_write` times only
  the writer, and shared-reader rows time only the shared reader;
- compares typed semantic models after a write, retaining exact-source checks
  only for formatcode2, SVG blips, theme-family values/parts, and the strict
  InkML action profile; the image blip lane is explicitly scanner-only;
- uses the production SVG contextual writer for its output-limit preflight,
  with an unbounded observation sink. Vec/String-only routes are explicitly
  N/A for the bounded-sink capability rather than receiving a synthetic cap;
- passes the current typed limits through the `drawingml-raw-current`,
  `drawingml-theme-family-current`, and `drawingml-ink-limits-current`
  feature gates while preserving the historical one-argument baseline APIs.

The validator and source test are updated to require semantic receipts,
typed preflight declarations, and outside-clock fidelity evidence. No
benchmark output or large evidence is committed by this record; the
coordinator must run the locked paired normal and allocator suites on the
assembled candidate before accepting performance claims.

## Source and validation boundary

- Probe repair base: `17b36d7b8a9b3612d395579c74e966787fab2c9b`.
- The owner preparation/replay lane is reviewed separately and is not claimed
  by this record.
- Required static validation: `python3 -m unittest tools.test_drawingml_raw_qualification`.
- Required final validation: the coordinator's pinned, offline, locked
  DrawingML raw paired normal and allocator runs, followed by the existing
  `drawingml_raw_validator.py`.
