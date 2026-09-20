# 0709 DOCX ordinary-save semantic oracle

This standalone probe supplements the ordinary-save harness. The harness's
determinism checks prove that a reference publication is stable, but they do
not prove that `document_mut().add_paragraph_with_text(...)` changed the
document as intended. This probe uses the public `litchi_docx::Package` API and
checks both documented publication doors for two checked-in fixtures:

* `test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx`
  is the admitted real-producer fixture. It contains nested numbered-list
  tables and a body-final `w:sectPr`; the probe checks the source paragraph
  projection, appends one fixed marker, serializes through `to_stream` and
  `save`, reopens each output, checks the marker exactly once at the end, and
  checks that the pre-existing paragraph projection remains an exact prefix.
* `test-data/libreoffice-core/sw/qa/writerfilter/dmapper/data/alt-chunk-header.docx`
  is the historical 0638 refusal. It is reader-admitted but has `w:altChunk`
  after the body-final `w:sectPr`. The current source therefore refuses
  `document_mut()` with the structural reason `body-final section properties
  are not the final body child`; 0638 recorded the same fixture before the BOM
  correction, when the surfaced error was a syntax/truncated-range refusal.
  The probe verifies that the refusal remains typed and that both publication
  doors still round-trip the unedited package.

For every output it compares ZIP member compressed payloads and uncompressed
size hints. The comparison is a payload identity check: ZIP local headers,
central-directory metadata, and the main document's expected edit are outside
that identity. A successful route may normalize the public text projection;
the exact paragraph-prefix check is the semantic preservation gate, while the
document-text prefix is reported as a weaker diagnostic.

Run it from the repository root after the primary build:

```text
cargo run --release --manifest-path \
  docs/performance/results/change-0709/oracle/Cargo.toml -- \
  --repo-root . \
  --output docs/performance/results/change-0709/oracle/report.json
```

The output report is deliberately outside the timed harness. The binary exits
nonzero if either the admitted semantic edit or the historical refusal
round-trip fails.
