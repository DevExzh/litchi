# ODI fixture provenance

`odf-1.4-normative-synthetic.fodi` is a hand-authored, deterministic fixture
derived from the ODF 1.4 image-document grammar. It is normative synthetic
evidence only. It was not created or resaved by LibreOffice, Apache OpenOffice,
NeoOffice, or another native producer.

The repository and local fixture roots were searched for `.odi` and `.fodi`
files and contained no producer artifact. An official Apache OpenOffice 4.1.16
Linux distribution was then downloaded from `downloads.apache.org`; its SHA-256
digest matched the published value
`febd01695bbd9ff68d509dbb973bfd714dff0e0a99e50abb4ea32a37eb6aa2ce`.
The shipped filter registry contained no OpenDocument Image (`ODI`) or flat
OpenDocument Image (`FODI`) export filter. Consequently, this corpus still has
no genuine producer ODI/FODI evidence. A future producer fixture must include
the producer name/version and an unchanged original file.

On 2026-08-10, the local check was repeated. No `libreoffice`, `soffice`,
`openoffice`, or `swriter` executable was available on `PATH`, and bounded
searches of the workspace, `/tmp`, `/opt`, `/usr/local`, and the user cache
found only this synthetic `.fodi`. Thus this environment cannot produce or
validate a genuine changed-file resave without installing a producer that
actually exposes an ODI/FODI export filter.

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
producer-created document. Therefore there is still no file whose producer,
original bytes, revision, and license can all be proven, and nothing from this
search was added to the fixture corpus.
