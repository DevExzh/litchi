# Output acceptance scope

The frozen protocol permits newly added member wrappers to differ between
revisions. It does not permit differences in copied media, slide semantics,
dependency closure, or untouched destination members. The existing benchmark
checks these properties independently of its deterministic republication check.
No benchmark timing or oracle code is changed for this optimization.

Source inspected: `tools/perf-baseline/src/lib.rs`, particularly
`build_pptx_source_backed_cross_copy_corpus`,
`verify_pptx_source_backed_lifecycle_output`, and
`verify_pptx_source_backed_cross_copy_lifecycle_gates`.

- Corpus construction independently regenerates and compares source and
  destination archives against the corresponding owned corpus. Owned planning
  supplies expected added part counts, logical bytes, relationship counts, and
  collision remaps.
- Output validation reopens source, destination, and candidate through the
  owned PPTX/OPC readers. It checks ordered slide names/text, insertion position,
  layout reuse, and the complete admitted dependency closure.
- For media, it follows each source and candidate picture relationship,
  requires internal image targets without fragments or queries, checks unique
  target names without destination collisions, and compares content types and
  decoded bytes. Source media is also compared with the deterministic payload
  generator: eight 2 MiB PNG-typed payloads in the media-rich corpus.
- Slide XML comparison permits only the intended embedded relationship-ID
  substitutions. Part and ZIP member sets and content-type semantics are
  checked separately.
- Every untouched destination member is compared both decoded and physically
  through `raw_zip_members`. The only excluded destination members are
  `[Content_Types].xml`, presentation XML, and its relationship member.
  Untouched physical order, central-directory order, and archive comment are
  also compared.
- Additional gates require deterministic republication, stable source bytes
  and revision, and rejection of changed source, changed destination, and
  foreign destination authority. Any false gate aborts corpus construction.

Each retained report records these gate results and its output identity.
The before/after source manifests must show the same harness source hashes.
Exact compressed source-payload reuse is a separate ZIP/OPC implementation-test
obligation; decoded equality alone does not prove it.

Portable replay checks report custody, gate values, matched inputs, and derived
statistics. It does not rerun Rust readers against exported output archives:
the capture does not retain those archives. These source-bound producer gates
must not be described as portable independent archive revalidation. The
synthetic benchmark also supplies no native PPTX cross-slide-copy evidence.
