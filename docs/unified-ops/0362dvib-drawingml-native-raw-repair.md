# 0362dvib — DrawingML native raw probe repair

## Scope

This repair is limited to the source-backed adapters used by the paired
DrawingML raw qualification. It does not change a format owner, package
placement, or the owner replay implementation.

The probe now:

- parses each source and changed fixture into a typed `PreparedValue` before
  timed writer rows; `entry_read` times only the reader, `vec_write` times only
  the writer, and shared-reader rows time only the shared reader;
- freezes source, changed, and malformed generated-fixture hashes and byte
  lengths before the first warmup, and compares that receipt across every
  paired baseline and candidate leg;
- compares typed semantic models after a write, retaining exact-source checks
  for formatcode2, SVG blips, theme-family values/parts, and InkML action
  profiles. `write_profile_to` retains the profile source, so its receipt
  proves exact unchanged and changed bytes as well as semantic fidelity; the
  image blip lane is explicitly scanner-only;
- uses the production SVG contextual writer for its output-limit preflight,
  with an unbounded observation sink. Every current typed output limit,
  including Vec/String routes, receives an actual refusal check; chart's
  bounded partial output is reported and checked against its cap. Diagram and
  Theme Family part rows are named source-edit-and-serialize operations;
- runs an untimed fixture-preflight child for each source endpoint before any
  warmup or ABBA child, comparing the descriptor hash/byte length and the
  source/changed/malformed generated hashes/byte lengths across baseline and
  candidate. These control events are retained
  in the runner journal with `measured=false` and are excluded from measured
  RSS/timing children. The SVG preflight intentionally uses a contextual
  parsed value and contextual writer; the normal Vec and bounded-sink rows and
  the outside-clock fidelity check use that same contextual route. The
  standalone reader/writer is checked outside timing over the completed
  contextual documents as an explicitly additional route; the existing
  post-run validator comparison across all qualified reports is retained as a
  second handoff check;
- mutates parsed, source-backed formatcode2 and Theme Family values. Their
  changed receipts verify the untouched lexical source splice and decoded
  semantic change. Complete Theme-part add/replace receipts parse the result
  and compare the requested name, id, and variant id after the operation;
- builds all operation detail/refusal receipts before the first warmup. Timed
  rows black-box only the production Result, with cap selection, output sizing,
  error classification, digests, and receipt construction outside the clock;
- keeps image-blip `find_first_embed`/`write_embed` helper work in a separate
  scanner-helper receipt; it is not reported as a typed writer;
- passes the current typed limits through the `drawingml-raw-current`,
  `drawingml-theme-family-current`, and `drawingml-ink-limits-current`
  feature gates while preserving the historical one-argument baseline APIs.

The validator and source test are updated to require semantic receipts,
typed output-preflight declarations, frozen fixture identities, and
outside-clock fidelity evidence. Text-body evidence names its actual
`read(xml, scope, limits)` entry. No
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
