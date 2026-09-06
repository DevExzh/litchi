# Native PPTX cross-copy gap

This is a source-only audit for change 0432. It inspected PPTX ZIP/XML
structure, producer metadata, the source-backed cross-slide-copy planner, and
the existing fixture provenance notes. It did not import or mutate a fixture,
call the cross-copy API, build, test, profile, or run a CPU workload. There is
no runtime success claim in this report.

## Bounded result and corpus scope

The checked-in inventory in
[`change-0429/native-inventory.json`](../change-0429/native-inventory.json)
contains 78 PPTX files and 146 slides. Sixteen slides contain one or more
direct `p:pic` elements (20 pictures total), and none of the 146 slides has a
nonempty `p:cSld@name`. The four native-labelled inputs in the inventory
(original shapes, LibreOffice-resaved shapes, POI `bug62513`, and POI
`EmbeddedVideo`) cover 32 slides and have no slide names; 11 of those slides
have media relationships.

The local ignored `3rdparty/` checkout contains a larger exploratory corpus
(829 PPTX files, 1,642 slides, 334 media-bearing slides, and one named slide).
The only named slide is `3rdparty/poi/test-data/slideshow/SampleShow.pptx`
slide 0 (`FirstSlide`); it has no direct picture/media closure. The candidate
below is therefore useful for source inspection, but it is not a checked-in
benchmark fixture: `.gitignore` ignores `/3rdparty`, and
`git ls-files '3rdparty/**/*.pptx'` is empty. The tracked
`METAFILE_PROVENANCE.md` describes the vendored metafile subsets, not these
ignored PPTX files. The Microsoft Office application strings in the candidate
archives are producer metadata, not an independently recorded save-chain
provenance.

For the inspected image-bearing slides, no positive native source/destination
cross-copy pair is currently supported by the planner. This is a bounded
finding about the inspected image-bearing scope; it does not establish that
every text-only native pair is impossible.

## Best unmodified native-producer candidate

The strongest source/destination pair found in the ignored LibreOffice QA
checkout is:

| role | archive | package SHA-256 | selector |
| --- | --- | --- | --- |
| source | `3rdparty/libreoffice-core/sd/qa/unit/data/TextFittingComparisonWithMSO_1.pptx` | `c0e3f5c79cc055bc7e8c01bbb120b06c4cb9756dd4b59f41d11ce883e7c0594d` | `source_position=0` (`ppt/slides/slide1.xml`) |
| destination | `3rdparty/libreoffice-core/sd/qa/unit/data/TextFittingComparisonWithMSO_2.pptx` | `6f7edbcdfe083546e1d145a5929c419c74d4fabfcd9622bb193c55b784155a6a` | `destination_slide_position=0`, `insertion_position=1` |

The corresponding API shape is
`plan_cross_slide_copy(source, 0, 0, 1)`. Both archives identify themselves
as Microsoft Office PowerPoint (`AppVersion=16.0000`). The selected slides
are transitional PresentationML, have no `mc:AlternateContent`, and have
exactly two slide relationships: `rId2` to a direct image and `rId1` to
`slideLayout2.xml`. There are no notes, audio, video, or other unreferenced
slide relationships on the selected slides.

The source direct picture is shape position 0, shape id 8, name `Picture 7`,
`rId2`, bounds `(694156,72735,8496905,6858000)`, and
`ppt/media/image1.png` (50,585 bytes,
`0763d6cff329b2daa2bc25bb2d0a8ed5663392a0ac4a47256a71d38c1688f03f`). The
destination direct picture is shape position 0, shape id 10, name `Picture 9`,
`rId2`, bounds `(287432,0,8565622,7004911)`, and
`ppt/media/image1.png` (48,283 bytes,
`bbe40632848e118c7dbae4eee746fc2f8b7a55e1ed4e82038329ac3a542db8e7`). This
is an image-bearing source/destination candidate, rather than an OPC-only
synthetic fixture.

The selected owner graph is exactly equal at the raw part and relationship
level in both archives: `ppt/slideLayouts/slideLayout2.xml` points to
`ppt/slideMasters/slideMaster1.xml`, which points to `ppt/theme/theme1.xml`.
The known closure hashes are:

| closure member | bytes | source SHA-256 | destination SHA-256 |
| --- | ---: | --- | --- |
| `[Content_Types].xml` | 5,149 / 3,842 | `3d183cebbdb25dcf55bb6cbca431ccc489058a3d3a0475f8418ea13949417ea9` | `9a7711b8ae63be258e053a4f985e1dd85c2c6ed5ff2ecd391cae0166db36f41c` |
| `ppt/slides/slide1.xml` | 6,923 / 7,329 | `277dc419c0b3db538dd103e7086be2f71f37f99db951e2770f122f55c1cf6f28` | `6e12ad59fe0f6b43d7ec9701ac12db9a62ddc137e9aa0187f978fbef21fe134c` |
| `ppt/slides/_rels/slide1.xml.rels` | 446 | `2556274e445456b553f6c7ebaa82748f98028b16bb948d01c68f3582386aceba` | `2556274e445456b553f6c7ebaa82748f98028b16bb948d01c68f3582386aceba` |
| `ppt/slideLayouts/slideLayout2.xml` | 3,979 | `f9ba6ecab715d3b0729e881f4bf04625b1c3881cee7e68482995c8977782d76c` | `f9ba6ecab715d3b0729e881f4bf04625b1c3881cee7e68482995c8977782d76c` |
| `ppt/slideLayouts/_rels/slideLayout2.xml.rels` | 311 | `8246d333bf3764cd35563e3df1828c26bbc28890815a2987caf3e592791ba60d` | `8246d333bf3764cd35563e3df1828c26bbc28890815a2987caf3e592791ba60d` |
| `ppt/slideMasters/slideMaster1.xml` | 13,934 | `a78233920045b5cb35955b64030fce7eba5d1d9c8d6e6c4003fb239ee4b8a3b1` | `a78233920045b5cb35955b64030fce7eba5d1d9c8d6e6c4003fb239ee4b8a3b1` |
| `ppt/slideMasters/_rels/slideMaster1.xml.rels` | 1,991 | `e9e503158ddaff4d9afa825a3d1048ffa1c1275291d9ab818812ff8e061ea8fd` | `e9e503158ddaff4d9afa825a3d1048ffa1c1275291d9ab818812ff8e061ea8fd` |
| `ppt/theme/theme1.xml` | 8,399 | `d37627ed1a5bda40fa339cc2115c70db5e019f1b128736b92b267180531434c6` | `d37627ed1a5bda40fa339cc2115c70db5e019f1b128736b92b267180531434c6` |
| `ppt/media/image1.png` | 50,585 / 48,283 | `0763d6cff329b2daa2bc25bb2d0a8ed5663392a0ac4a47256a71d38c1688f03f` | `bbe40632848e118c7dbae4eee746fc2f8b7a55e1ed4e82038329ac3a542db8e7` |

The source layout relationship is `rId1` with target
`../slideMasters/slideMaster1.xml`. The master relationship set has the same
12 exact internal edges in both packages, including `rId12` to
`../theme/theme1.xml`; the theme is a relationship-free leaf. The table is a
static closure oracle only. No planner digest, output archive, or post-copy
semantic oracle was produced because the name gate below prevents an honest
runtime run.

## Narrow implementation gate

[`namespace.rs`](../../../../crates/litchi-pptx/src/namespace.rs:48) reads the
first PresentationML `p:cSld` and returns its optional `name` attribute. A
missing attribute is decoded as an empty string by `presentation_name`.

In [`source_cross_copy.rs`](../../../../crates/litchi-pptx/src/presentation/source_cross_copy.rs:1205):

* `prepare` at approximately lines 1205-1214 filters the source name and
  returns `SlideCopyRefusal::AmbiguousTopology` with
  `source slide name is missing or empty` when it is absent.
* `destination_names` at approximately lines 4257-4290 walks **every** slide
  in the destination presentation. It returns the same refusal kind with
  `destination slide name is missing or empty` for an absent destination name,
  before the source-name collision check.

Both selected slides in the candidate have `<p:cSld>` without `name`; all 16
source slides and all six destination slides are unnamed. Consequently
`plan_cross_slide_copy(source, 0, 0, 1)` would be refused at the
destination-name scan, and the source-name gate would also refuse it. This pair
already has the exact owner graph and direct-image relationship shape required
by the planner, so the missing producer names are the highest-impact compatible
gap visible in the static evidence. Adding names to either archive would make
the input a mutated fixture and would not be a valid native benchmark.

## Recommendation and preservation boundary

Keep this pair as a quarantined, unmodified native-producer candidate until a
checked-in provenance decision is made and a real package with nonempty unique
slide names is available. Do not relabel it as a positive cross-copy oracle.
For a valid future run, use selectors `(0, 0, 1)` and retain the package and
closure hashes above as input metadata. The expected source-backed semantics
would be an immutable source, a new destination slide at insertion position 1,
reuse of the destination owner graph, and exact preservation of the selected
source image payload and untouched destination members. Those are API-contract
expectations, not observations from this audit.

Any follow-up measurement belongs in the existing tool-only harness: caller-owned
immutable snapshots, explicit bounded providers and execution budgets, clocks
around the selected API sequence, and separate package/oracle validation. This
audit authorizes no production change, fixture import, or native cross-copy
claim.

If a true native cross-copy pair cannot be admitted, use the tracked POI
fixtures for a separate native image lifecycle: `bug62513.pptx` slide 4/image
0 is a clean direct JPEG candidate, while `EmbeddedVideo.pptx` slide 0/image 0
and `EmbeddedAudio.pptx` slide 0/image 0 exercise poster-image reads while
retaining their media members. Those cases provide native image/read or
targeted edit/publish coverage; they do not close the native cross-copy gap.
The existing original and LibreOffice-resaved `shapes.pptx` files should remain
separate static/refusal controls, not be used as an untouched native pair.
