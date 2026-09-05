# Native PPTX source audit

This is a source-only audit for change 0429. It does not change a fixture,
production crate, or benchmark implementation, and it does not certify a
runtime result. The checked-in ZIP/XML inventory is
[`native-inventory.json`](native-inventory.json).

## Cross-copy result

No inspected image-bearing native source/destination pair can honestly support
a positive source-backed cross-copy claim. This is a bounded result for the
image-bearing slides inspected here; it is not a claim that every text-only
native pair is impossible.

Every image-bearing slide in the inventory has a missing `p:cSld@name`. The
planner in `crates/litchi-pptx/src/presentation/source_cross_copy.rs` refuses
an empty or missing source name (`AmbiguousTopology`) and also refuses a name
collision with a destination slide. It then proves exact layout/master/theme
graph equality by digest; independently rewritten producer packages are not
considered equivalent. A same-source plan is also rejected by the existing
source-backed tests. The Microsoft `shapes.pptx` and its LibreOffice-resaved
derivative use different layout registrations and are not a compatible pair.

Do not add names to, rewrite, or otherwise mutate a native fixture to make the
planner pass. The bounded native evidence should therefore report this as an
explicit cross-copy gap. A destination chosen only to force a refusal is
negative gate coverage, not a native cross-copy lifecycle.

## Provenance and candidate selectors

The strongest positive source-backed image/edit candidate is the vendored POI
fixture:

* `test-data/poi/test-data/slideshow/bug62513.pptx` (SHA-256
  `cd841112bd5b53f21e8434d080af8b2eeca78a07bc820e9ea0b97a0962c85c79`). Its
  package metadata says Microsoft Office PowerPoint, 19 slides, and it has the
  Apache POI corpus attribution: the retained upstream notice is
  `test-data/poi/NOTICE`, and
  `test-data/libreoffice-core/METAFILE_PROVENANCE.md` records the POI-origin
  corpus as Apache-2.0 with the repository license retained. Select zero-based
  `slide_position=4` (slide5).
  The slide has no markup-compatibility branch, one direct picture at
  `shape_position=1`, `image_position=0`, shape id `37890`, name `Picture 2`,
  `rId2`, and internal target `/ppt/media/image2.jpeg` (`image/jpeg`, 21,997
  bytes, payload SHA-256
  `d5f10480ab75ce1175eea7432e684e9fac7319a7b9bd15512759385eae4842db`). Its
  first text shape is `shape_position=0`, `Title 1`, text `Membership`, so the
  same selector supports a focused text edit and publish probe while retaining
  the image and all untouched package members. The other slides make this a
  media/embedded-object-rich package, while the selected slide has a small,
  clean direct-image closure.

Two additional POI fixtures are useful for a media-rich image-only probe. They
  have no `mc:AlternateContent` on the selected slide and their direct poster
  image parts are leaves:

* `test-data/poi/test-data/slideshow/EmbeddedVideo.pptx`,
  `slide_position=0`, `image_position=0`, shape id `2`, name
  `file_example_MP4_480_1_5MG_Trim`, `rId4`, `/ppt/media/image1.png`, 65,215
  bytes, payload SHA-256
  `f5516c6cae484df63ce03db77fb69b778660916b9207de5a4e04aa5e3b72908d`.
  The package metadata says Microsoft Office PowerPoint and also contains the
  MP4 media member. `images()`/`read_image(0)` exercise the poster image only;
  they must leave the video relationship and bytes untouched.
* `test-data/poi/test-data/slideshow/EmbeddedAudio.pptx`, the analogous
  Microsoft Macintosh PowerPoint fixture, has `slide_position=0`,
  `image_position=0`, shape id `4`, name `sample-3s.mp3`, `/ppt/media/image1.png`,
  4,717 bytes, and payload SHA-256
  `b0151c2c2e3cf64bc37a7bb9d8b8b98d4c4fccf7b6af4c08c4f847a79f9db0da` (the
  package also contains the MP3 member).

The LibreOffice provenance is explicit in
[`test-data/office-interop/PROVENANCE.md`](../../../../test-data/office-interop/PROVENANCE.md):
runtime LibreOffice 26.2.5.2, MPL-2.0, genuine source
`test-data/ooxml/pptx/shapes.pptx`, Litchi changed output, and the checked-in
LibreOffice resave. It is useful as a separately named native-producer
artifact, but the current source image API cannot use either member as a
positive image-read case without a fixture/API decision:

* Original `shapes.pptx` has one `/ppt/media/image1.jpg`, but its `a:blip`
  contains a nested `a:extLst`; `parse_picture_relationship` deliberately
  refuses that unsupported nested blip grammar.
* `libreoffice-resaved/shapes-litchi.pptx` has one `/ppt/media/image1.jpeg`
  and the same 11,988-byte payload hash as the original, but its slide has
  `mc:AlternateContent`. `SourceSlide::images`, `image`, and source-backed
  slide edits reject markup-compatibility branch selection before returning a
  descriptor.

Consequently those two files are valid static producer/payload comparison
inputs and expected refusal cases under the current gates; they are not
evidence that a positive native image lifecycle succeeds. The resaved fixture
must remain distinct and must not be relabeled as an untouched producer
input.

## API and preservation gates

`SourceSlide::images()` reads only the selected slide XML and relationship
metadata. It must return the exact descriptor identity (position, shape
position/id/name/bounds, relationship id, target and content type) without
reading the media payload. `image(position)` repeats the exact zero-based
selection. `read_image(position)` may read only an internal `/ppt/media/`
`image/*` part with no outbound relationships; it retains source-version,
cancellation, cache and managed-budget checks and never follows an external
target. For the media fixtures, audio/video payloads are outside this direct
poster-image closure.

For the `bug62513` edit probe, the safe semantic is one selected text-shape
replacement followed by the existing commit/publish API. The source archive
and provider stay immutable; the output stream is separate. The selected
JPEG, all other slide XML, embedded objects, relationships, content types and
unknown members must remain byte-preserved according to the existing
publication contract. A source image result should be retained and its bytes
revalidated after package-owner drops.

## 0429 harness/ADR fit

Keep this evidence in a new tool-only lifecycle command. No production API,
fixture, or ownership change is needed. Every provider is an explicit caller
adapter (`bytes`, file, or bounded range); range adapters use the protocol's
4,096-byte and 65,536-byte caps and optional fixed delay, with no network or
ambient filesystem access beyond the named file provider. Native original,
LibreOffice, and POI inputs remain separate labels from the synthetic
media-rich corpus.

Start clocks immediately around the selected API sequence (open/plan/publish,
or open/metadata/image-read), and exclude source construction, copying, sink
reservation, observers, oracle checks and drops. Keep source snapshots owned by
the caller/provider and immutable; retain returned image data across owner
drops. Enforce existing finite package, cache, allocation, cancellation and
managed-budget limits. Keep logical reads, cache gauges, budget gauges and RSS
as separate evidence. These boundaries fit ADRs 0003, 0005 and 0006 and the
0429 protocol; they support provider-specific duration and ownership evidence
only. They do not support a native cross-copy, physical-I/O, cold-filesystem,
remote-service, causal speedup, or general leak claim.

The actionable recommendation is therefore: use `bug62513` slide 4/image 0
for the positive native image and optional targeted edit/publish lifecycle;
use `EmbeddedVideo` (or `EmbeddedAudio`) slide 0/image 0 for the media-rich
poster-image variant; retain original/resaved shapes as separately labelled
static/refusal evidence. The formal positive protocol can use the exact
ZIP-derived POI-slide and POI-video oracles; the original/resaved pair remains
a separate static/refusal control. Record native cross-copy as unsupported for
the inspected image-bearing inventory.
