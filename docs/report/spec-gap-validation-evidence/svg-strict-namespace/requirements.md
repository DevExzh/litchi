# Bounded Strict SVG namespace proof

This retained probe covers one synthetic PPTX package. It starts from
`pptx-svg-lifecycle/outputs/source.pptx`, converts the core PresentationML,
DrawingML, and OfficeDocument relationship namespace URIs to ISO/IEC 29500
Strict with `PhysPkgReader`/`PhysPkgWriter`, and leaves the OPC package
relationships namespace unchanged.

The downstream edit uses the ordinary positional PPTX API:

1. open the converted package and select slide 0, picture 0;
2. attach one SVG, call the ordinary fallible `commit`, and publish it;
3. reopen the saved package, detach the same owner, call `commit`, and publish;
4. apply each retained inverse to its exact in-memory snapshot; and
5. compare the detached slide, relationship part, media, opaque member, and
   member set with the converted source.

The output has three independent checks:

- the complete slide is accepted by the vendored ECMA-376 Part 1 Strict
  `pml.xsd` graph;
- the selected `asvg:svgBlip` is accepted by the unmodified MS-ODRAWXML 5.24
  schema text with its original Transitional `a:AG_Blob` and relationship
  imports; and
- a direct Transitional `r:embed` child passes while the equivalent direct
  Strict-namespace attribute is rejected. The rejection is retained as a
  negative control.

The validator also strictifies the retained LibreOffice XLSX drawing member
and validates the resulting native-derived Strict XDR/DrawingML graph while
restoring Transitional relationship attributes on each SVG child. This is a
schema control, not an XLSX lifecycle implementation or a native Strict
producer claim.

`run_checks.py` performs the compile-first gate. `run_probe.py` runs the
downstream binary and offline validator after that gate. `verify.py` recomputes
source, tool, fixture, generated-output, and raw-result digests. The root
Cargo `target/` is shared; any accidentally created `harness/target/` is
removed by the scripts.

This evidence does not claim native Office acceptance, rendering, SVG
sanitization, or a final feature approval. It also does not rewrite the older
PPTX lifecycle evidence directory.
