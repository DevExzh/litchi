# External LibreOffice QA fixture provenance

The official fixture is pinned to LibreOffice commit
[39ee11db03368a18bc1746a6312307c4cf1b4bdc](https://git.libreoffice.org/core/+/39ee11db03368a18bc1746a6312307c4cf1b4bdc),
which renamed `sd/qa/unit/data/a.pptx` to `smoketest.pptx` without changing its
content. It was introduced in
[83ea2162b80b4d03e4cca90d354b5d56ea3f5898](https://git.libreoffice.org/core/+/83ea2162b80b4d03e4cca90d354b5d56ea3f5898)
as part of the initial sd filter tests.

- Bytes: 29,956.
- SHA-256: `88a4755fa90815802c8f439c9e0488772e5e7d8db63cfd0326e4d3f35fdeaa44`.
- Git blob: `e0cfe49009c9d735b5dd6ea774dda2e7a6710ae8`.
- Selector: source slide 0, independently opened destination anchor 0, insertion 1.

fetch-fixture.py downloads the immutable official Gitiles artifact, validates its
length, SHA-256 and Git blob identity, and writes it only into exclusive task
scratch. external-fixture.json binds that execution and the fetch driver.
No new binary fixture is imported into tracked source. The upstream binary has
no app/core producer properties or per-file license notice; source-tree license
notices do not prove an original producer/save chain. This evidence therefore
uses the precise label **unmodified LibreOffice QA fixture, same-archive
source/destination clone**. It establishes neither an independent producer pair
nor a native Office/LibreOffice application roundtrip.

The fixture contains one unnamed slide with two text shapes, one direct picture,
and exactly one image plus one layout relationship. Its relationship XML is
noncanonical. No input name, XML spelling, relationship, ZIP framing or payload
is changed to make it acceptable to the API.
