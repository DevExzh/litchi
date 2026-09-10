# 0501 native producer coverage review

The next PPTX semantic gap is successful source-backed copying of a real
producer dependency graph. The synthetic source-backed lifecycles already
have measured evidence in 0424, 0431, and 0448, with caller attribution in
0449. Those records cover staged payload reuse, verified compressed transfer,
range pacing, and source-backed plan/publication timing. They should not be
described as missing generic lifecycle or ABBA evidence.

The 0501 hashing candidate is a separate change. Its permitted scope is the
private decoded image/chart payload contribution to `touched_digest`; exact
payload comparisons, source lineage and revision checks, graph and metadata
digests, candidate revalidation, compressed-payload authorization, resource
admission, cancellation chunk checks, and partial-output behavior remain
required. This review records the native-producer gap; it does not claim that
notes closure is implemented or measured in 0501.

## Recommended bounded task

Implement and measure one real-producer notes relationship closure through
`SourceBackedPresentationEditor::plan_cross_slide_copy` and
`publish_cross_slide_copy_to_stream` in
`crates/litchi-pptx/src/presentation/source_cross_copy.rs`.

Use one source slide with exactly:

- one existing slide-layout relationship;
- one `notesSlide` relationship;
- a notes slide with exactly its slide backlink and one `notesMaster`
  relationship; and
- an existing notes master/theme boundary that is equivalent in the
  destination.

The first tranche should copy only the notes slide part. It should allocate a
collision-free notes-slide URI, add the inserted slide's `notesSlide`
relationship, rewrite the copied notes slide's slide backlink to the inserted
slide, and retain the destination's existing notes master and notes theme.
The source notes XML remains an exact source-proven part. Refuse before any
output for comments, VML, external targets, extra notes relationships,
missing or non-equivalent notes master/theme graphs, orphan notes parts, or
any other relationship outside this bounded topology.

## Fixture and provenance

The preferred fixture is:

`test-data/libreoffice-core/oox/qa/unit/data/tdf131082.pptx`

It is 25,881 bytes with SHA-256
`a7287e4303cce471c20f10d7a6f90c0717ff99d1165450033c314b821ac34fa1`.
Its relevant relationship graph is:

```text
ppt/presentation.xml
  -- notesMaster --> ppt/notesMasters/notesMaster1.xml
ppt/slides/slide1.xml
  -- slideLayout --> ppt/slideLayouts/slideLayout1.xml
  -- notesSlide  --> ppt/notesSlides/notesSlide1.xml
ppt/notesSlides/notesSlide1.xml
  -- slide       --> ppt/slides/slide1.xml
  -- notesMaster --> ppt/notesMasters/notesMaster1.xml
ppt/notesMasters/notesMaster1.xml
  -- theme       --> ppt/theme/theme2.xml
```

The fixture has no chart, comment, VML, OLE, diagram, or external dependency,
which keeps this closure bounded. Use the same fixture bytes as source and
destination through two independent `OwnedSource` instances. The existing
harness pattern is in `tools/perf-baseline/src/lib.rs:15990`; the distinct
wrappers satisfy the source-backed requirement that source and destination do
not share one lineage or revision even though their bytes match.

The fixture is a public producer-authored package carried in the LibreOffice
test corpus. The repository records the LibreOffice-corpus provenance and
MPL-2.0 coverage in `test-data/libreoffice-core/METAFILE_PROVENANCE.md`, with
the retained license text at
`test-data/odf/native-resave/source/LICENSE-MPL-2.0.txt`. The package's own
`docProps/app.xml` identifies `Microsoft Office PowerPoint`, AppVersion
`14.0000`; corpus location and license provenance therefore do not establish
that LibreOffice originally generated the package. A future LibreOffice save
through `tools/native_odf_resave.py` is separate compatibility evidence, as
described in `test-data/office-interop/PROVENANCE.md`, followed by independent
Litchi readback.

## Exact remaining closure scope

The production implementation must add a prepared notes relationship/part
alongside the current `PreparedSlideRelationship` variants at
`crates/litchi-pptx/src/presentation/source_cross_copy.rs:302`. Planning
currently admits only layout, image, and chart relationships around line
1056. Publication must add the notes XML and its rewritten relationship sidecar
through the existing `SourceTopologyPlan` APIs:

- `try_add_source_xml_part`;
- `try_add_internal_relationship`; and
- the existing source proof, candidate match, and touched-digest paths.

The notes graph checks should follow
`crates/litchi-pptx/src/notes/package.rs:67` and its index validation: exactly
one presentation notes-master reference, one notes-master theme edge, one
notes-slide backlink, one notes-master edge, valid content types, and no
orphan or external notes parts. A source-backed helper may be needed because
the existing notes loader consumes `OpcPackage` while this planner consumes
`SourceBackedPackage`.

Correctness coverage belongs in
`crates/litchi-pptx/tests/source_backed_cross_copy.rs` and should assert:

1. a one-to-two-slide copy creates `notesSlide2.xml`, rewrites its backlink,
   retains the destination notes master/theme, and reopens through the notes
   graph;
2. deterministic output, source immutability, raw destination preservation,
   and collision-safe relationship/member allocation;
3. stale source/destination revisions, cancellation, limits, and partial sinks
   remain atomic; and
4. malformed notes graphs and unsupported mixed relationships retain typed
   fail-closed refusals.

The performance seam is the existing
`build_pptx_source_backed_cross_copy_corpus` and
`run_pptx_source_backed_cross_copy_lifecycle` in
`tools/perf-baseline/src/lib.rs:15193` and `:48798`. A future opt-in real-notes
selector should carry the fixture hash, expected notes member/relationship
counts, planned bytes, semantic reopen gates, and source/raw-preservation
checks. The plan/publication timings can use the existing bytes/file/range
provider lanes. A LibreOffice save followed by independent Litchi readback is
an additional producer-reopen gate and should remain clearly separate from
the lifecycle timer.

## Deferred chart closure

Chart closure remains a separate, larger task. The real LibreOffice fixture
`test-data/libreoffice-core/sd/qa/unit/data/pptx/tdf114821.pptx` is the
preferred chart candidate because it has the same MPL provenance. It is
50,235 bytes with SHA-256
`939720d41bb739a7fc152d07af323bb15ddfc93f2f1831db28556bff46d804be` and
contains chart style, chart color style, an embedded workbook,
`AlternateContent`, and producer extension namespaces. Supporting it requires
bounded outbound chart dependency admission, opaque workbook/style/color part
copying, relationship-ID and URI remapping, and chart XML validation that
preserves the producer extensions. The current source-backed planner
intentionally refuses outbound chart relationships at
`source_cross_copy.rs:1119`.

The simpler `test-data/ooxml/pptx/line-chart.pptx` is useful for owned-path
comparison (`crates/litchi-pptx/src/opened/tests.rs:3036`), but its independent
fixture-license provenance is less explicit. Neither chart fixture should be
folded into the bounded notes task.

No hashing shortcut is part of this native closure review. The 0449 SHA-256
profile is attribution, not proof that publication revalidation can be
removed; exact source freshness and candidate checks remain mandatory.
