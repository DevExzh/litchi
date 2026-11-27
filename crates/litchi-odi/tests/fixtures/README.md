# ODI fixture provenance

`odf-1.4-normative-synthetic.fodi` is a hand-authored, deterministic fixture
derived from the ODF 1.4 image-document grammar. It is normative synthetic
evidence only. It was not created or resaved by LibreOffice, Apache OpenOffice,
NeoOffice, or another native producer.

The repository and local fixture roots were initially searched for `.odi` and
`.fodi` files and contained no producer artifact. An official Apache OpenOffice 4.1.16
Linux distribution was then downloaded from `downloads.apache.org`; its SHA-256
digest matched the published value
`febd01695bbd9ff68d509dbb973bfd714dff0e0a99e50abb4ea32a37eb6aa2ce`.
The shipped filter registry contained no OpenDocument Image (`ODI`) or flat
OpenDocument Image (`FODI`) export filter. At that audit stage, the corpus had
no genuine producer evidence.

On 2026-08-10, the local check was repeated. No `libreoffice`, `soffice`,
`openoffice`, or `swriter` executable was available on `PATH`, and bounded
searches of the workspace, `/tmp`, `/opt`, `/usr/local`, and the user cache
found only this synthetic `.fodi`. Thus this environment cannot produce or
validate a genuine changed-file resave through a locally installed native
office-suite executable. The later ODFDOM route below is independent of that
negative local-suite result.

The 2026-08-10 authoritative-source search also inspected these producer and
format-project revisions for files ending in `.odi` or `.fodi`. Complete
recursive tree results were available for Apache OpenOffice and Apache Tika;
the relevant filter, drawing, and test subtrees were traversed for the larger
LibreOffice and NeoOffice repositories:

- Apache OpenOffice, revision
  `55c04d1336cb0228ec67f5ea1b5d7cd6aa993e7d`,
  <https://github.com/apache/openoffice/tree/55c04d1336cb0228ec67f5ea1b5d7cd6aa993e7d>,
  Apache-2.0.
- LibreOffice core, revision
  `7d915a0d4ade0e8b9cde3fd23f0fd92c066a78e1`,
  <https://github.com/LibreOffice/core/tree/7d915a0d4ade0e8b9cde3fd23f0fd92c066a78e1>,
  repository-declared GPL-3.0.
- NeoOffice, revision
  `61705f0d1df563b4cbdb4752e550a0f826da8821`,
  <https://github.com/neooffice/NeoOffice/tree/61705f0d1df563b4cbdb4752e550a0f826da8821>;
  its `LICENSE` states GPL-3.0 for the application and MPL-2.0 for most
  NeoOffice-authored source, subject to per-file third-party terms.
- Apache Tika, revision
  `348908214476f466565ea5d770c79f5e1ba55851`,
  <https://github.com/apache/tika/tree/348908214476f466565ea5d770c79f5e1ba55851>,
  Apache-2.0.

No ODI/FODI artifact was found. The authoritative OpenOffice MIME registry at
<https://www.openoffice.org/framework/documentation/mimetypes/mimetypes.html>
confirms the `.odi` media type, but it is documentation rather than a
producer-created document. Nothing from that search was added to the fixture
corpus.

The next 2026-08-10 producer-route audit found no local `libreoffice`,
`soffice`, or `unoconv` executable, no installed `libreoffice-core`,
`libreoffice-draw`, or `python3-uno` package, and no importable Python `uno`
module. Ubuntu offered packages in its configured repositories, but an
available package is not evidence that its producer exposes ODI/FODI output,
so no package was installed and no artifact was synthesized.

LibreOffice core revision
`7d915a0d4ade0e8b9cde3fd23f0fd92c066a78e1` was also checked at its exact
filter-registry directory objects. The `filters` tree
`bca007324c9650ffb870bc27c858c9257e26b5ed` and `types` tree
`f59f3fa951b8b327a0c1a31529841bd23c37b072` expose ODG/FODG entries, including
[`ODG_FlatXML.xcu`](https://github.com/LibreOffice/core/blob/7d915a0d4ade0e8b9cde3fd23f0fd92c066a78e1/filter/source/config/fragments/filters/ODG_FlatXML.xcu)
and
[`draw_ODG_FlatXML.xcu`](https://github.com/LibreOffice/core/blob/7d915a0d4ade0e8b9cde3fd23f0fd92c066a78e1/filter/source/config/fragments/types/draw_ODG_FlatXML.xcu),
but no ODI/FODI filter or type.
The official conversion-filter documentation likewise documents the
[`OutputFileExtension[:OutputFilterName]`](https://help.libreoffice.org/latest/en-US/text/shared/guide/convertfilters.html)
route but lists no OpenDocument Image filter. This is negative capability
evidence, not a producer fixture; that route added nothing to the corpus.

## ODFDOM 0.13.0 producer evidence

The later producer-route audit found a legitimate programmatic producer in
[ODF Toolkit ODFDOM 0.13.0](https://odftoolkit.org/downloads.html), released
2026-01-23 from tag `odftoolkit-0.13.0`. The official API documents
[`OdfImageDocument.newImageDocument()`](https://odftoolkit.org/api/odfdom/org/odftoolkit/odfdom/doc/OdfImageDocument.html)
as creating an empty image document. ODFDOM is Apache-2.0; the exact Maven
artifact was `org.odftoolkit:odfdom-java:0.13.0`, downloaded from
<https://repo1.maven.org/maven2/org/odftoolkit/odfdom-java/0.13.0/odfdom-java-0.13.0.jar>,
with SHA-256
`c98c13fabb2ee67afd89d63177bdfbfc8304962f7ee9b7527d62cbe2000b4a5b`.
Its manifest identifies implementation version `0.13.0`, The Document
Foundation as implementation vendor, and Apache-2.0 as the bundle license.

Using a temporary OpenJDK 17.0.19 runtime, JShell called
`OdfImageDocument.newImageDocument()`, populated the producer's default
`draw:frame` through ODFDOM's namespace-aware DOM with a deterministic inline
1×1 PNG, and called `save` to create `odfdom-0.13.0-original.odi`. A separate
`OdfImageDocument.loadDocument` opened those exact bytes, changed only the
frame's `draw:name` semantic value from `ODFDOM-0.13.0-Original` to
`ODFDOM-0.13.0-Changed`, and called `save` to create the changed round trip.
Both packages contain the uncompressed first-member MIME value
`application/vnd.oasis.opendocument.image` and one image frame with inline
PNG data. The exact API transcript and inline payload are retained in
`odfdom-0.13.0-generate.jsh`; ZIP timestamps mean it is a provenance transcript,
not a promise that a later execution will reproduce the archive checksum.

- `odfdom-0.13.0-original.odi`: SHA-256
  `322b60d333123efe74c93d889ea1b15058a1436e98f3cf0b51889fa70caec074`.
- `odfdom-0.13.0-changed.odi`: SHA-256
  `2e7169a5bf2ff014276d750d4f46f1fdd4f571ff403cf59f9b9b1ae9115dbece`.

Both packages produced no diagnostics under the official ODF Validator 0.13.0
executable JAR. Its published and downloaded SHA-256 both equal
`5684feec5cbdcd5783998978c096ac9ccea53a454e2d6ae803ce482d2336d1dc`.

Across the producer resave, the uncompressed `mimetype`, `styles.xml`,
`settings.xml`, and `META-INF/manifest.xml` members remain byte-identical.
`content.xml` records the requested frame-name change, while ODFDOM also
updates producer-managed `meta.xml` editing metadata. The legacy
`meta:generator` string inherited from ODFDOM's bundled default template is
not used as producer identity; provenance rests on the pinned API call,
released artifact checksum, original bytes, and separate load/change/save
chain. These two generated fixtures are distributed under this repository's
Apache-2.0 license.
