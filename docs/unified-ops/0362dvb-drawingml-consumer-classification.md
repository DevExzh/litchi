# 0362dvb — DrawingML consumer classification follow-up

This follow-up records the consumer repairs in
`9566398bf35fb54c5d4eb9b354e62c60a544743b`. The source commit is based on
the owner-policy receipt commit `3b0ee480b798f3a65c1296fb03e13eaf2f1c3a34`;
the exact source trees, pinned lock, commands, sessions, and external log
hashes are in
[`results/0362dvb/drawingml-consumer-classification.json`](results/0362dvb/drawingml-consumer-classification.json).

## Repairs

- DrawingML `Limit`, `InvalidLimit`, and stale `PatchConflict` now retain
  `limit_exceeded`, `invalid_argument`, and `invalid_input` classifications
  through the DOCX/PPTX host wrappers and the OPS chart and SmartArt
  adapters. Ordinary parse/authoring invalid errors retain their existing
  parse-vs-authoring fallback categories.
- Chart and SmartArt consumers no longer discard typed producer limits;
  modern-comment body conversion and chart-native classification keep the
  typed variants as well.
- Namespace-prefix versus namespace-byte overflow, attempted XML depth,
  Chartex node/depth overflow, theme output policy, and lexical coordinate
  limits now report their real typed dimensions. Transform and DOCX/PPTX
  coordinate adapters preserve those dimensions instead of stringifying the
  parse error.
- A focused overlong-coordinate regression covers the lexical coordinate
  budget, and OPS classification tests cover host, native chart, and parsed
  chart paths.

## Validation

The final source check ran before the source commit on the unchanged staged-
equivalent working tree. It used the pinned workspace lock (`52ed0c…`), the
repository capped shim, one build job, offline locked dependencies, zero
debug profiles, an external persistent `TMPDIR`, and an external target of
2,590,004,731 bytes (<20 GB). The final focused filter passed 2 tests; all
other tests were filtered by the requested `drawingml` name. The first
attempt's cargo compile failure is retained separately from the successful
runs: it exposed the root DrawingML `Common(XmlError)` classifier shape and
was fixed before the final run.

The earlier standalone `unified-chart-add --offline --locked` failure and
its registry-preserving local-package lock repair remain recorded with exact
hashes in the JSON below and in the owner-policy receipt. No build target,
binary, or large evidence file is committed.
