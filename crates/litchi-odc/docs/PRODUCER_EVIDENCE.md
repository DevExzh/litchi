# Standalone ODC/FODC producer evidence

Updated on 2026-08-10 after the ODC remediation review.

Genuine standalone producer-created `.odc` evidence is checked in under
`test-data/odf/odc-producer-evidence/`. ODFDOM 0.13.0 created the first package
with `OdfChartDocument.newChartDocument()` and a separate JVM loaded, changed,
and saved the second package. Both artifacts have the correct standalone chart
MIME, pinned producer/library/source hashes, clean ODF Validator transcripts,
and independent semantic reopen results. The exact commands, member hashes,
licensing, and deliberately removed generator metadata are recorded in that
directory's `PROVENANCE.md`; document metadata alone is not treated as producer
identity.

ODFDOM 0.13.0 exposed no standalone FODC route, and the native LibreOffice and
Apache OpenOffice chart-filter experiments below still did not yield a changed
standalone save. The positive claim is therefore genuine programmatic-producer
ODC create/change/save evidence, not current LibreOffice interoperability and
not FODC support.

The repository does contain genuine LibreOffice chart subdocuments inside
producer-created `.fods` and `.fodt` files. Those embedded XML fragments are
useful interoperability inputs for the shared chart reader, but they are not
standalone ODC/FODC packages and are deliberately not copied, repackaged, or
renamed as standalone fixture evidence.

## Primary-source search record

The 2026-08-10 follow-up checked these authoritative sources:

| Source | Pinned revision | Search and result | License/provenance |
|---|---|---|---|
| [LibreOffice core](https://github.com/LibreOffice/core) | `7d915a0d4ade0e8b9cde3fd23f0fd92c066a78e1` (`master`) | Complete, non-truncated recursive trees for `chart2`, `filter`, `odk`, `sc`, `sd`, `sw`, `test`, and `xmloff` contained no `.odc`, `.fodc`, or `.otc`. The repository-wide tree was truncated and therefore is not claimed as complete. | The upstream repository carries its own [license notice](https://github.com/LibreOffice/core/blob/master/LICENSE); no artifact was copied. |
| [Apache OpenOffice](https://github.com/apache/openoffice) | `55c04d1336cb0228ec67f5ea1b5d7cd6aa993e7d` (`trunk`) | Complete, non-truncated recursive `test` and `extras` trees contained no `.odc`, `.fodc`, or `.otc`. The repository-wide tree was truncated and therefore is not claimed as complete. | Apache OpenOffice is distributed under the repository's [Apache license notice](https://github.com/apache/openoffice/blob/trunk/LICENSE); no artifact was copied. |
| [OASIS ODF 1.4 OS Relax NG schema](https://docs.oasis-open.org/office/OpenDocument/v1.4/os/schemas/OpenDocument-v1.4-schema.rng) | ODF 1.4 OASIS Standard, SHA-256 `4034ec6be29205d5fc1ee5f42468ac6ef824287b3aba6d9289032af4fafbda7f` | Defines the optional empty `chart:coordinate-region` under `chart:plot-area` with common draw position and size attributes. A compact crate-local test case is derived from that rule. | Authoritative standards source. No schema or document artifact is vendored, and the derived XML is explicitly not producer evidence. |

Before the ODFDOM chain was produced, general web searches and the LibreOffice
OpenGrok path search exposed no reusable standalone artifact with producer and
license provenance. Search-engine absence was not treated as proof; the current
positive evidence instead rests on the pinned executable producer chain.

## Native standalone persistence audit

The 2026-08-10 follow-up also exercised two verified official producer builds
through their bundled UNO bridges and isolated user profiles:

| Producer | Verified distribution | Native route and result |
|---|---|---|
| Apache OpenOffice 4.1.16, build 9816 | [ASF Linux x86-64 DEB archive](https://archive.apache.org/dist/openoffice/4.1.16/binaries/en-US/Apache_OpenOffice_4.1.16_Linux_x86-64_install-deb_en-US.tar.gz), SHA-256 `febd01695bbd9ff68d509dbb973bfd714dff0e0a99e50abb4ea32a37eb6aa2ce` | `private:factory/schart` returned `com.sun.star.comp.chart2.ChartModel`, and the filter factory exposed `chart8` as “ODF Chart”. Saving a populated unbacked chart with `storeAsURL()` disposed the bridge and produced no file. |
| LibreOffice 26.2.5.2, build `cd7284b4cbbfeb507e630c1aac019f4157393acb` | [TDF Linux x86-64 DEB archive](https://download.documentfoundation.org/libreoffice/stable/26.2.5/deb/x86_64/LibreOffice_26.2.5_Linux_x86-64_deb.tar.gz), SHA-256 `2f03bfb2ac9f33ea7c77331b4b7a23300fb0ed7443566046bf8b5bc51c1bed1e` | The same native factory and filter route returned a chart model, but `storeAsURL()` disposed the bridge. Loading a transparently authored standalone bootstrap and calling in-place `store()` completed after title/data mutation and explicit modified state, but the output remained byte-identical. |

Both suites were run with the native `chart8` filter, not by converting or
repackaging an embedded chart. Because neither route yielded producer-written
changed bytes, no fixture from this audit is checked in or described as
producer-created evidence.

A future native-application interoperability claim qualifies only if its
provenance records the native suite saving a changed standalone ODC/FODC
artifact and Litchi fully reopening the result. The ODFDOM evidence satisfies
the separate programmatic-producer claim but is not relabeled as LibreOffice,
OpenOffice, or FODC evidence.
