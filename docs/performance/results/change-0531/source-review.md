# 0531 MCE namespace-search source review

This review covers `docs/performance/results/change-0531/candidate.patch` at
revision `b2dffa3d75656f46b36091afcf1e7410a66a90f8` (the source file itself is
still at its baseline SHA-256
`e5911cdcd94116474b09062af96d639b8f08d80a09a26d90c86a931a7344fa74`). The
patch has one hunk in
`crates/litchi-ooxml-common/src/mce/codec.rs`; its SHA-256 is
`a809d1e41f99451c66f4d751bc3c9368ca6f7be6b0a02f1e52f7d01090bb2c8b`, and
`git apply --check --whitespace=error-all` accepts it. I made no Rust source
changes and ran no Rust build or test.

The review applies the accepted constraints in [ADR 0001](../../../adr/0001-priorities-and-api-layers.md),
[ADR 0002](../../../adr/0002-crate-topology.md),
[ADR 0003](../../../adr/0003-snapshots-edits-and-patches.md),
[ADR 0005](../../../adr/0005-io-memory-and-performance.md),
[ADR 0006](../../../adr/0006-validation-security-and-compatibility.md),
[ADR 0008](../../../adr/0008-migration-and-verification.md),
[ADR 0010](../../../adr/0010-facade-archive-ownership.md),
[ADR 0011](../../../adr/0011-ooxml-physical-package-ownership.md), and
[ADR 0024](../../../adr/0024-current-topology.md). The complete accepted-ADR
hash set remains recorded in the 0530
[`adr-manifest.json`](../change-0530/adr-manifest.json).

## Semantic equivalence

The old branch computes

```text
xml.windows(NAMESPACE.len()).any(|w| w == NAMESPACE.as_bytes())
```

and the candidate computes

```text
memchr::memmem::find(xml, NAMESPACE.as_bytes()).is_some()
```

`NAMESPACE` is a fixed, non-empty ASCII string, so its string length equals
the byte-slice length. For a non-empty byte needle, `memmem::find` searches the
same contiguous byte windows as the standard slice iterator, including
arbitrary non-UTF-8 bytes and NULs. Only the presence boolean is consumed, so
the returned offset is irrelevant. The candidate therefore has the same
branch decision for every `&[u8]`, including inputs shorter than the needle,
matches at either boundary, overlapping matches, and near-matches.

The hunk leaves the following behavior and order intact:

- `max_input_bytes` is checked before the scan.
- A no-match result checks `max_output_bytes`, returns `Cow::Borrowed(xml)`,
  and returns `Report::default()`.
- A match enters the same `quick_xml` MCE parser with the same capabilities,
  output bound, errors, parser, and output ownership.
- A namespace occurrence in text, an attribute value, a comment, or CDATA
  still triggers the parser. This existing lexical trigger is broader than
  an XML namespace declaration; the patch does not narrow or reinterpret it.

`process_ooxml`, `process_part`, `process_part_arc`, and `process_str` all
delegate to this function. In particular, `process_part_arc` uses the
borrowed/owned result to decide whether the original part `Arc` can be shared,
and DOCX source-backed paths use that identity to refuse unsupported MCE branch
selection. The candidate preserves that distinction because it changes no
branch condition.

The private `find_bytes` calls used by `active_offsets` and marker handling are
unchanged. The event-oriented MCE processor in `mce/stream.rs` is also
unchanged. This patch consequently makes no claim about those paths.

## Affected OOXML callers

The changed function is shared by all three OOXML verticals and the common
DrawingML layer. Representative direct or transitive callers are:

| owner/path | representative consumers | observable contract |
| --- | --- | --- |
| common OOXML | `custom_xml`, web XML, spreadsheet XML maps | configured MCE limits and borrowed no-op output |
| DOCX | document/section validation, source-backed document and settings paths, comments, footnotes, styles, themes, charts, fonts, text boxes, glossary, web, and modern comments | source preservation, typed MCE refusals, and `Arc` sharing on the no-op path |
| PPTX | presentation and slide parts, shape scene readers, tag/source mapping, actions, media, transitions, notes, fonts, comments, embedded controls/ink/OLE, and presentation-property codecs | borrowed scene/part XML and source-offset mapping when MCE rewrites bytes |
| XLSX | raw workbook/catalog/worksheet/styles/shared-strings readers, cell watches, tables, validations, page/view/setup codecs, chart-sheet/package paths, connections, pivots, ActiveX, and metadata codecs | worksheet and package parsing must retain existing MCE selection and error precedence |
| XLSB and DrawingML | host drawing/cell transfer, DrawingML themes/charts/diagrams/styles | shared parser behavior and UTF-8/borrowed output assumptions |

The direct calls can be inventoried with:

```sh
rg -n '(process_markup_compatibility|process_ooxml|process_part_arc|process_part|process_str)\(' \
  crates/litchi-ooxml-common crates/litchi-drawingml crates/litchi-docx \
  crates/litchi-pptx crates/litchi-xlsx crates/litchi-xlsb -g '*.rs'
```

The candidate introduces no archive type, public API, unsafe code, or new
dependency. `memchr` is already a direct dependency of
`litchi-ooxml-common`, so the accepted crate-ownership and dependency-direction
boundaries remain intact.

## Risks and admission conditions

There is no semantic objection to measuring this hunk. The main performance
risk is that `memmem::find` may have more setup overhead than a simple
`windows` loop for short MCE-free XML, while helping longer payloads. The
candidate must therefore be measured on both MCE-absent and MCE-present
distributions, including the short common parts reached by DOCX and PPTX
guards. The 0530 profile attributed about 5.18% of the selected XLSX planning
interval to the complete `process_ooxml` child edge; that edge includes parsing
and setup, so it is not the removable scan cost and cannot support a 5.18%
end-to-end claim.

Retain the candidate only under the frozen 0531 gates: total p50 and mean
reductions of at least 1%, planning p50 reduction of at least 2%, and planning
instruction reduction of at least 1%, in both repeats of both primary shapes,
with no unexplained allocation regression and passing XLSX, DOCX, and PPTX
guards. Review every adverse metric individually. A failed native gate means
the source hunk remains an unapplied measurement candidate.

The no-match path allocates no output, but the candidate's searcher setup and
the existing output-limit preflight still need allocation and boundary
verification. Do not change the input/output limits, error ordering, MCE
capabilities, strict/transitional policy, `active_offsets`, or streaming
processor as part of this pilot.

## Required correctness coverage after applying the candidate

1. Add a focused MCE fast-path parity test over byte-oriented fixtures:
   empty/short input, exact URI at the beginning and end, overlapping and
   repeated matches, one-byte mutations at every URI position, URI prefix and
   suffix near-matches, embedded NUL, and invalid UTF-8. Assert that the
   borrowed/owned branch and returned bytes agree with the reference
   `windows(...).any(...)` result. Keep the input and output limits in the
   fixture matrix.
2. Exercise the exact URI in an element/namespace attribute, ordinary
   attribute value, text, comment, and CDATA, plus a malformed document that
   contains the URI. These cases must continue to enter the parser and retain
   the old success/error behavior; a near-match must retain the borrowed
   fast path.
3. Preserve the existing semantic MCE suite: choice/fallback selection,
   ignored and processed content, preserved elements/attributes,
   `MustUnderstand`, malformed AlternateContent, exact output limits, deep
   namespace/directive limits, and the POI styles and LibreOffice PPTX
   fixtures. These cover parser behavior after the detector says “present.”
4. Check ownership-sensitive wrappers: `process_str` must remain borrowed for
   MCE-free strings and owned after a real transformation; `process_part_arc`
   must continue sharing the source `Arc` only on the no-op path. Retain the
   DOCX source-backed refusal and PPTX source-offset mapping tests.
5. Run the common and all affected OOXML/DrawingML test targets after the
   source is applied:

   ```sh
   cargo test -p litchi-ooxml-common --lib mce
   cargo test -p litchi-ooxml-common --test markup_compatibility
   cargo test -p litchi-drawingml
   cargo test -p litchi-docx
   cargo test -p litchi-pptx
   cargo test -p litchi-xlsx
   cargo test -p litchi-xlsb
   ```

Disposition: semantically safe to proceed to the frozen candidate measurement,
with the lexical-trigger and ownership cases above treated as required
regressions. No production optimization should be retained from this review
alone; the performance and cross-format gates decide that outcome.
